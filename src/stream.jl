
#=
https://docs.julialang.org/en/v1/base/collections/#lib-collections-iteration-1

for i in iter   # or  "for i = iter"
    # body
end

is translated into:

next = iterate(iter)
while next != nothing
    (i, state) = next
    # body
    next = iterate(iter, state)
end
=#

#=
# About Iterators

* `SheetRowIterator` is an abstract iterator that has `SheetRow` as its elements. `SheetRowStreamIterator` and `WorksheetCache` implements `SheetRowIterator` interface.
* `SheetRowStreamIterator` is a dumb iterator for row elements in sheetData XML tag of a worksheet.
* `WorksheetCache` has a `SheetRowStreamIterator` and caches all values read from the stream.
* `TableRowIterator` is a smart iterator that looks for tabular data, but uses a SheetRowIterator under the hood.

The implementation of `SheetRowIterator` will be chosen automatically by `XLSX.eachrow` method,
based on the `enable_cache` option used in `XLSX.openxlsx` method.

=#

#=
# SheetRowStreamIterator

It's state is the SheetRowStreamIteratorState.
The iterator element is a SheetRow.
=#

# strip off namespace prefix of nodename
function nodename(x::EzXML.StreamReader)
    split(EzXML.nodename(x), ':')[end]
end

@inline get_worksheet(itr::SheetRowStreamIterator) = itr.sheet
@inline row_number(state::SheetRowStreamIteratorState) = state.row

Base.show(io::IO, state::SheetRowStreamIteratorState) = print(io, "SheetRowStreamIteratorState( is open = $(state.is_open) , row = $(state.row) )")

# Opens a file for streaming.
@inline function open_internal_file_stream(xf::XLSXFile, filename::String) :: Tuple{ZipFile.Reader, EzXML.StreamReader}
    @assert internal_xml_file_exists(xf, filename) "Couldn't find $filename in $(xf.source)."
    @assert xf.source isa IO || isfile(xf.source) "Can't open internal file $filename for streaming because the XLSX file $(xf.filepath) was not found."
    io = ZipFile.Reader(xf.source)

    for f in io.files
        if f.name == filename
            return io, EzXML.StreamReader(f)
        end
    end

    error("Couldn't find $filename in $(xf.source).")
end

@inline Base.isopen(s::SheetRowStreamIteratorState) = s.is_open

@inline function Base.close(s::SheetRowStreamIteratorState)
    if isopen(s)
        s.is_open = false
        close(s.xml_stream_reader)
        close(s.zip_io)
    end
    nothing
end

# Creates a reader for row elements in the Worksheet's XML.
# Will return a stream reader positioned in the first row element if it exists.
# If there's no row element inside sheetData XML tag, it will close all streams and return `nothing`.
function Base.iterate(itr::SheetRowStreamIterator, state::Union{Nothing, SheetRowStreamIteratorState}=nothing)

    ws = get_worksheet(itr)

    if state == nothing # first iteration. Will open a stream and create the first state instance
        state = let
            target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
            zip_io, reader = open_internal_file_stream(get_xlsxfile(ws), target_file)

            # The reader will be positioned in the first row element inside sheetData
            # First, let's look for sheetData opening element
            while EzXML.iterate(reader) != nothing
                if EzXML.nodetype(reader) == EzXML.READER_ELEMENT && nodename(reader) == "sheetData"
                    @assert EzXML.nodedepth(reader) == 1 "Malformed Worksheet \"$(ws.name)\": unexpected node depth for sheetData node: $(EzXML.nodedepth(reader))."
                    break
                end
            end

            @assert nodename(reader) == "sheetData" "Malformed Worksheet \"$(ws.name)\": Couldn't find sheetData element."

            # Now let's look for a row element, if it exists
            while EzXML.iterate(reader) != nothing # go next node
                if EzXML.nodetype(reader) == EzXML.READER_ELEMENT && nodename(reader) == "row"
                    break
                elseif is_end_of_sheet_data(reader)
                    # this Worksheet has no rows
                    close(reader)
                    close(zip_io)
                    return nothing
                end
            end

            # row number is set to 0 in the first state
            SheetRowStreamIteratorState(zip_io, reader, true, 0)
        end
    end

    # given that the first iteration case is done in the code above, we shouldn't get it again in here
    @assert state != nothing "Error processing Worksheet $(ws.name): shouldn't get first iteration case again."

    reader = state.xml_stream_reader
    if is_end_of_sheet_data(reader)
        @assert !isopen(state)
        return nothing
    else
        @assert isopen(state) "Error processing Worksheet $(ws.name): Can't fetch rows from a closed workbook."
    end

    # will read next row from stream.
    # The stream should be already positioned in the next row
    @assert nodename(reader) == "row"
    current_row = parse(Int, reader["r"])
    rowcells = Dict{Int, Cell}() # column -> cell

    # iterate thru row cells
    while EzXML.iterate(reader) != nothing

        if is_end_of_sheet_data(reader)
            close(state)
            break
        end

        # If this is the end of this row, will point to the next row or set the end of this stream
        if EzXML.nodetype(reader) == EzXML.READER_END_ELEMENT && nodename(reader) == "row"

            while true
                if is_end_of_sheet_data(reader)
                    close(state)
                    break
                elseif EzXML.nodetype(reader) == EzXML.READER_ELEMENT && nodename(reader) == "row"
                    break
                end

                @assert EzXML.iterate(reader) != nothing
            end


            # breaks while loop to return current row
            break

        elseif EzXML.nodetype(reader) == EzXML.READER_ELEMENT && nodename(reader) == "c"

            # reads current cell to rowcells
            cell = Cell( EzXML.expandtree(reader) )
            @assert row_number(cell) == current_row "Error processing Worksheet $(ws.name): Inconsistent state: expected row number $(current_row), but cell has row number $(row_number(cell))"
            rowcells[column_number(cell)] = cell

        elseif EzXML.nodetype(reader) == EzXML.READER_ELEMENT && nodename(reader) == "row"
            # last row has no child elements, so we're already pointing to the next row
            break
        end
    end

    sheet_row = SheetRow(get_worksheet(itr), current_row, rowcells)

    # update state
    state.row = current_row
    return sheet_row, state
end

# Detects a closing sheetData element
@inline is_end_of_sheet_data(r::EzXML.StreamReader) = (EzXML.nodedepth(r) <= 1) || (EzXML.nodetype(r) == EzXML.READER_END_ELEMENT && nodename(r) == "sheetData")

#
# WorksheetCache
#

# Indicates whether worksheet cache will be fed while reading worksheet cells.
@inline is_cache_enabled(ws::Worksheet) = is_cache_enabled(get_xlsxfile(ws))
@inline is_cache_enabled(wb::Workbook) = is_cache_enabled(get_xlsxfile(wb))
@inline is_cache_enabled(xl::XLSXFile) = xl.use_cache_for_sheet_data
@inline is_cache_enabled(itr::SheetRowIterator) = is_cache_enabled(get_worksheet(itr))

@inline function push_sheetrow!(wc::WorksheetCache, sheet_row::SheetRow)
    r = row_number(sheet_row)

    if !haskey(wc.cells, r)
        # add new row to the cache
        wc.cells[r] = sheet_row.rowcells
        push!(wc.rows_in_cache, r)
        wc.row_index[r] = length(wc.rows_in_cache)
    end
    nothing
end

#
# WorksheetCache iterator
#
# The state is the row number. The element is a SheetRow.
#

function WorksheetCache(ws::Worksheet)
    itr = SheetRowStreamIterator(ws)
    return WorksheetCache(CellCache(), Vector{Int}(), Dict{Int, Int}(), itr, nothing, true)
end

@inline get_worksheet(r::SheetRow) = r.sheet
@inline get_worksheet(itr::WorksheetCache) = get_worksheet(itr.stream_iterator)

# In the WorksheetCache iterator, the element is a SheetRow, the state is the row number
function Base.iterate(ws_cache::WorksheetCache, row_from_last_iteration::Int=0)

    #the sorting operation are very costly when adding row  and only needed if we use the row iterator
    if ws_cache.dirty
        sort!(ws_cache.rows_in_cache)
        ws_cache.row_index = Dict{Int, Int}(ws_cache.rows_in_cache[i] => i for i in 1:length(ws_cache.rows_in_cache))
        ws_cache.dirty = false
    end

    if row_from_last_iteration == 0 && !isempty(ws_cache.rows_in_cache)

        # the next row is in cache, and it's the first one
        current_row_number = ws_cache.rows_in_cache[1]
        sheet_row_cells = ws_cache.cells[current_row_number]
        return SheetRow(get_worksheet(ws_cache), current_row_number, sheet_row_cells), current_row_number

    elseif row_from_last_iteration != 0 && ws_cache.row_index[row_from_last_iteration] < length(ws_cache.rows_in_cache)

        # the next row is in cache
        current_row_number = ws_cache.rows_in_cache[ws_cache.row_index[row_from_last_iteration] + 1]
        sheet_row_cells = ws_cache.cells[current_row_number]
        return SheetRow(get_worksheet(ws_cache), current_row_number, sheet_row_cells), current_row_number

    else

        next = iterate(ws_cache.stream_iterator, ws_cache.stream_state)
        if next == nothing
            return nothing
        end

        sheet_row, next_stream_state = next

        # add new row to WorkSheetCache
        push_sheetrow!(ws_cache, sheet_row)

        # update stream state
        ws_cache.stream_state = next_stream_state

        return sheet_row, row_number(sheet_row)
    end
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

"""
    getcell(xlsxfile, cell_reference_name) :: AbstractCell
    getcell(worksheet, cell_reference_name) :: AbstractCell
    getcell(sheetrow, column_name) :: AbstractCell
    getcell(sheetrow, column_number) :: AbstractCell

Returns the internal representation of a worksheet cell.

Returns `XLSX.EmptyCell` if the cell has no data.
"""
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

getdata(r::SheetRow, column) = getdata(get_worksheet(r), getcell(r, column))
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
function eachrow(ws::Worksheet) :: SheetRowIterator
    if is_cache_enabled(ws)
        if ws.cache == nothing
            ws.cache = WorksheetCache(ws)
        end
        return ws.cache
    else
        return SheetRowStreamIterator(ws)
    end
end

function Base.isempty(sr::SheetRow)
    return isempty(sr.rowcells)
end
