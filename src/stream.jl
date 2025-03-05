
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
# SheetRowIterator

It's state is the SheetRowStreamIteratorState.
The iterator element is a SheetRow.
=#

# strip off namespace prefix of nodename
function nodename(x::XML.LazyNode)
    split(XML.tag(x), ':')[end]
end

@inline get_worksheet(itr::SheetRowIterator) = itr.sheet
@inline row_number(state::SheetRowStreamIteratorState) = state.row

Base.show(io::IO, state::SheetRowStreamIteratorState) = print(io, "SheetRowStreamIteratorState( itr = $(state.itr), itr_state = $(state.itr_state), row = $(state.row) )")

# Opens a file for streaming.
@inline function open_internal_file_stream(xf::XLSXFile, filename::String) :: XML.LazyNode
    @assert internal_xml_file_exists(xf, filename) "Couldn't find $filename in $(xf.source)."
    @assert xf.source isa IO || isfile(xf.source) "Can't open internal file $filename for streaming because the XLSX file $(xf.filepath) was not found."
    XML.LazyNode(XML.Raw(ZipArchives.zip_readentry(xf.io, filename)))
end

# Creates a reader for row elements in the Worksheet's XML.
# Will return a stream reader positioned in the first row element if it exists.
# If there's no row element inside sheetData XML tag, it will close all streams and return `nothing`.
function Base.iterate(itr::SheetRowStreamIterator, state::Union{Nothing, SheetRowStreamIteratorState}=nothing)
    local current_row
    local current_row_ht
    local sheet_row
    local next_element = ""
    local nc = 0
    local cell_no=0

    ws = get_worksheet(itr)

    if isnothing(state) # first iteration. Will open a LazyNode for iteration, find the first row and create the first state instance

        state = let 

            target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
            reader = open_internal_file_stream(get_xlsxfile(ws), target_file)

            @assert length(reader) > 0 "Couldn't open reader for Worksheet $(ws.name)."
            @assert XML.tag(reader[end]) == "worksheet" "Expecting to find a worksheet node.: Found a $(XML.tag(reader[end]))."
            ws_elements = XML.children(reader[end])
            idx = findfirst(y -> y=="sheetData", [XML.tag(x) for x in ws_elements])
            next_element= idx===nothing ? "" : (ws_elements[idx+1])
            
            next = iterate(reader)
            while next !== nothing
                (lznode, lzstate) = next

                if XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == "sheetData"
                    nrows = length(filter(n -> XML.nodetype(n) == XML.Element && XML.tag(n) == "row", XML.children(lznode)))
                    if nrows == 0
                        return nothing
                    end
                    @assert XML.depth(lznode) == 1 "Malformed Worksheet \"$(ws.name)\": unexpected node depth for sheetData node: $(XML.depth(lznode))."
                    break
                end

                next = iterate(reader, lzstate)
            end
            if next === nothing
                return nothing
            end

            # Now let's look for a row element, if it exists
            next = iterate(reader, lzstate)
            while next !== nothing # go next node
                (lznode, lzstate) = next
                if XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == "row" # This is the first row
                    a = XML.attributes(lzstate)
                    current_row = parse(Int, a["r"])
                    current_row_ht = haskey(a, "ht") ? parse(Float64, a["ht"]) : nothing
                    nc = length(filter(n -> XML.nodetype(n) == XML.Element && XML.tag(n) == "c", XML.children(lzstate))) # number of cells in this row
                    cell_no = 0
                    break
                end
                next = iterate(reader, lzstate)
            end
            SheetRowStreamIteratorState(reader, lzstate, current_row, current_row_ht)
        end
    end

    # given that the first iteration case is done in the code above, we shouldn't get it again in here
    @assert state !== nothing "Error processing Worksheet $(ws.name): shouldn't get first iteration case again."

    reader = state.itr
    lzstate = state.itr_state

    # Expecting iterator to be at the first row element
    current_row = state.row
    current_row_ht = state.ht
    rowcells = Dict{Int, Cell}() # column -> cell
    isnothing(lzstate) && return nothing
    next = iterate(reader, lzstate) # iterate through row cells
    while next !== nothing
        (lznode, lzstate) = next
        if XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == next_element # This is the end of sheetData
            return nothing
        elseif XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == "row" # This is the next row
            a = XML.attributes(lzstate)
            current_row = parse(Int, a["r"])
            current_row_ht = haskey(a, "ht") ? parse(Float64, a["ht"]) : nothing
            nc = length(filter(n -> XML.nodetype(n) == XML.Element && XML.tag(n) == "c", XML.children(lzstate))) # number of cells in this row
            cell_no = 0
        elseif XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == "c" # This is a cell
            cell_no += 1
            cell = Cell(lznode)
            @assert row_number(cell) == current_row "Error processing Worksheet $(ws.name): Inconsistent state: expected row number $(current_row), but cell has row number $(row_number(cell))"
            rowcells[column_number(cell)] = cell
            if cell_no == nc # when all cells found

                sheet_row = SheetRow(get_worksheet(itr), current_row, current_row_ht, rowcells) # put them together in a sheet_row

                break
            end
        end
        next = iterate(reader, lzstate)
        if next === nothing
            return nothing
        end
    end

    # update state
    state.row = current_row

    state.ht = current_row_ht
    state.itr_state = lzstate



    return sheet_row, state
end


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
        wc.row_ht[r] = sheet_row.ht
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
    return WorksheetCache(CellCache(), Vector{Int}(), Dict{Int, Union{Float64, Nothing}}(), Dict{Int, Int}(), itr, nothing, true)
end

@inline get_worksheet(r::SheetRow) = r.sheet
@inline get_worksheet(itr::WorksheetCache) = get_worksheet(itr.stream_iterator)

# In the WorksheetCache iterator, the element is a SheetRow, the state is the row number
function Base.iterate(ws_cache::WorksheetCache, row_from_last_iteration::Int=0)

    #the sorting operation is very costly when adding row and only needed if we use the row iterator
    if ws_cache.dirty
        sort!(ws_cache.rows_in_cache)
        ws_cache.row_index = Dict{Int, Int}(ws_cache.rows_in_cache[i] => i for i in 1:length(ws_cache.rows_in_cache))
        ws_cache.dirty = false
       end

    if row_from_last_iteration == 0 && !isempty(ws_cache.rows_in_cache)
        # the next row is in cache, and it's the first one
        current_row_number = ws_cache.rows_in_cache[1]
        current_row_ht = ws_cache.row_ht[current_row_number]
        sheet_row_cells = ws_cache.cells[current_row_number]
        return SheetRow(get_worksheet(ws_cache), current_row_number, current_row_ht, sheet_row_cells), current_row_number

    elseif row_from_last_iteration != 0 && ws_cache.row_index[row_from_last_iteration] < length(ws_cache.rows_in_cache)
        # the next row is in cache
        current_row_number = ws_cache.rows_in_cache[ws_cache.row_index[row_from_last_iteration] + 1]
        current_row_ht = ws_cache.row_ht[current_row_number]
        sheet_row_cells = ws_cache.cells[current_row_number]
        return SheetRow(get_worksheet(ws_cache), current_row_number, current_row_ht, sheet_row_cells), current_row_number

    else
        next = iterate(ws_cache.stream_iterator, ws_cache.stream_state)

       if next === nothing
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


@inline row_number(sr::SheetRow) = sr.row

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
        if ws.cache === nothing
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
