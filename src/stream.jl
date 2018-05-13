
"""
Open a file for streaming.
"""
function open_internal_file_stream(xf::XLSXFile, filename::String) :: Tuple{ZipFile.Reader, EzXML.StreamReader}
    @assert internal_file_exists(xf, filename) "Couldn't find $filename in $(xf.filepath)."
    @assert isfile(xf.filepath) "Can't open internal file $filename for streaming because the XLSX file $(xf.filepath) was not found."
    io = ZipFile.Reader(xf.filepath)

    for f in io.files
        if f.name == filename
            return io, EzXML.StreamReader(f)
        end
    end

    error("Couldn't find $filename in $(xf.filepath).")
end

Base.isopen(i::InternalFileStream) = !isnull(i.io)

function Base.close(i::InternalFileStream)
    if !isnull(i.stream_reader)
        close(get(i.stream_reader))
        i.stream_reader = Nullable{EzXML.StreamReader}()
    end

    if !isnull(i.io)
        close(get(i.io))
        i.io = Nullable{ZipFile.Reader}()
    end

    nothing
end

function open_worksheet_stream!(ws::Worksheet; force_reopen::Bool=false)
    if force_reopen || !isopen(ws.cache.internal_file_stream)
        wb = get_workbook(ws)
        xf = get_xlsxfile(ws)
        target_file = "xl/" * get_relationship_target_by_id(wb, ws.relationship_id)
        cache = ws.cache
        stream = cache.internal_file_stream

        stream.io, stream.stream_reader = open_internal_file_stream(xf, target_file)

        reader = get(stream.stream_reader)

        # read Worksheet dimension
        while !EzXML.done(reader)
            if EzXML.nodetype(reader) == EzXML.READER_ELEMENT && EzXML.nodename(reader) == "dimension"
                @assert EzXML.nodedepth(reader) == 1 "Malformed Worksheet \"$(ws.name)\": unexpected node depth for dimension node: $(EzXML.nodedepth(reader))."
                ref_str = reader["ref"]
                if is_valid_cellname(ref_str)
                    cache.dimension = CellRange("$(ref_str):$(ref_str)")
                else
                    cache.dimension = CellRange(ref_str)
                end

                break
            end
        end

        # The reader will be positioned in the first row element inside sheetData
        while !EzXML.done(reader)
            if EzXML.nodetype(reader) == EzXML.READER_ELEMENT && EzXML.nodename(reader) == "sheetData"
                @assert EzXML.nodedepth(reader) == 1 "Malformed Worksheet \"$(ws.name)\": unexpected node depth for sheetData node: $(EzXML.nodedepth(reader))."
                break
            end
        end

        @assert EzXML.nodename(reader) == "sheetData" "Malformed Worksheet \"$(ws.name)\": Couldn't find sheetData element."

        while !EzXML.done(reader) # go next node
            if EzXML.nodetype(reader) == EzXML.READER_ELEMENT && EzXML.nodename(reader) == "row"
                break
            elseif is_end_of_sheet_data(reader)
                ws.cache.done_reading = true
                close(cache.internal_file_stream)
                break
            end
        end
    end

    nothing
end

@inline function get_stream_reader(ws::Worksheet) :: EzXML.StreamReader
    open_worksheet_stream!(ws)
    return get(ws.cache.internal_file_stream.stream_reader)
end

"""
Detects a closing sheetData element
"""
@inline is_end_of_sheet_data(r::EzXML.StreamReader) = (EzXML.nodedepth(r) <= 1) || (EzXML.nodetype(r) == EzXML.READER_END_ELEMENT && EzXML.nodename(r) == "sheetData")

@inline function push_row!(wc::WorksheetCache, row::Int)
    if !haskey(wc.cells, row)
        # add new row to the cache
        wc.cells[row] = Dict{Int, Cell}()
        push!(wc.rows_in_cache, row)
        wc.row_index[row] = length(wc.rows_in_cache)
    end
    nothing
end

# Add cell to cache
@inline function push_cell!(wc::WorksheetCache, cell::Cell)
    r, c = row_number(cell), column_number(cell)
    wc.cells[r][c] = cell
    nothing
end

#
# SheetRowIterator
#

@inline worksheet(r::SheetRow) = r.sheet
@inline worksheet(itr::SheetRowIterator) = itr.sheet

# state is the row number. At the start state, no row has been ready, so let's set to 0.
function Base.start(itr::SheetRowIterator)
    open_worksheet_stream!(worksheet(itr))
    return 0
end

function Base.done(itr::SheetRowIterator, row_from_last_iteration::Int)
    ws_cache = worksheet(itr).cache

    if ws_cache.done_reading
        # if cache is done reading from stream, we're done if the sheetData is empty, or if there are no more rows in cache
        if row_from_last_iteration == 0 && isempty(ws_cache.rows_in_cache)
            # sheetData is empty (no rows in worksheet)
            return true
        else
            if row_from_last_iteration == ws_cache.rows_in_cache[end]
                # no more rows in cache left to read
                return true
            end
        end
    end

    return false
end

#(i, state) = next(I, state)
function Base.next(itr::SheetRowIterator, row_from_last_iteration::Int)

    ws_cache = worksheet(itr).cache

    # fetches the next row
    if row_from_last_iteration == 0 && !isempty(ws_cache.rows_in_cache)
        # the next row is in cache, and it's the first one
        current_row = ws_cache.rows_in_cache[1]
    elseif row_from_last_iteration != 0 && ws_cache.row_index[row_from_last_iteration] < length(ws_cache.rows_in_cache)
        # the next row is in cache
        current_row = ws_cache.rows_in_cache[ws_cache.row_index[row_from_last_iteration] + 1]
    else
        # will read next row from stream.
        # The stream should be already positioned in the next row
        reader = get_stream_reader(worksheet(itr))
        @assert EzXML.nodename(reader) == "row"
        current_row = parse(Int, reader["r"])
        push_row!(ws_cache, current_row)

        # iterate thru row cells
        while !done(reader)

            # If this is the end of this row, will point to the next row or set the end of this stream
            if EzXML.nodetype(reader) == EzXML.READER_END_ELEMENT && EzXML.nodename(reader) == "row"
                done(reader) # go to the next node
                if is_end_of_sheet_data(reader)
                    # mark end of stream
                    ws_cache.done_reading = true
                    close(ws_cache.internal_file_stream)
                else
                    # make sure we're pointing to the next row node
                    @assert EzXML.nodetype(reader) == EzXML.READER_ELEMENT && EzXML.nodename(reader) == "row"
                end
                break
            elseif EzXML.nodetype(reader) == EzXML.READER_ELEMENT && EzXML.nodename(reader) == "c"
                cell = Cell( EzXML.expandtree(reader) )
                @assert row_number(cell) == current_row "Inconsistent state: expected row number $(current_row), but cell has row number $(row_number(cell))"

                # let's put this cell in the cache
                push_cell!(ws_cache, cell)
            elseif EzXML.nodetype(reader) == EzXML.READER_ELEMENT && EzXML.nodename(reader) == "row"
                # last row has no child elements, so we're already pointing to the next row
                break
            end
        end
    end

    return SheetRow(worksheet(itr), current_row, ws_cache.cells[current_row]), current_row
end

function find_row(itr::SheetRowIterator, row::Int) :: SheetRow
    for r in itr
        if row_number(r) == row
            return r
        end
    end
    error("Row $row not found.")
end

row_number(sr::SheetRow) = sr.row

function getcell(r::SheetRow, column_index::Int) :: AbstractCell
    if haskey(r.rowcells, column_index)
        return r.rowcells[column_index]
    else
        return EmptyCell(CellRef(row_number(r), column_index))
    end
end

function getcell(r::SheetRow, column_name::AbstractString)
    @assert is_valid_column_name(column_name) "$column_name is not a valid column name."
    return getcell(r, decode_column_number(column_name))
end

getdata(r::SheetRow, column) = getdata(worksheet(r), getcell(r, column))
Base.getindex(r::SheetRow, x) = getdata(r, x)

"""
    eachrow(sheet)

Creates a row iterator for a worksheet.

Example: Query all cells from columns 1 to 4.

```julia
left = 1  # 1st column
right = 4 # 4th column
for sheetrow in XLSX.eachrow(sheet)
    for column in left:right
        cell = XLSX.getcell(sheetrow, column)

        # do something with cell
    end
end
```
"""
eachrow(ws::Worksheet) = SheetRowIterator(ws)

function Base.isempty(sr::SheetRow)
    return isempty(sr.rowcells)
end
