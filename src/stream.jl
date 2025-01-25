
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

#    zip = ZipArchives.ZipReader(xf)
    if filename in ZipArchives.zip_names(xf.io)
        return XML.parse(XML.LazyNode, ZipArchives.zip_readentry(xf.io, filename, String))
    end 
#    return xf.io, XML.parse(XML.LazyNode, ZipArchives.zip_readentry(xf.io, filename, String))

    error("Couldn't find $filename in $(xf.source).")
end

@inline Base.isopen(s::SheetRowStreamIteratorState) = s.is_open

@inline function Base.close(s::SheetRowStreamIteratorState)
#    if isopen(s)
#        s.is_open = false
        #close(s.xml_stream_reader)
        # close(s.zip_io)
#    end
    nothing
end

# Creates a reader for row elements in the Worksheet's XML.
# Will return a stream reader positioned in the first row element if it exists.
# If there's no row element inside sheetData XML tag, it will close all streams and return `nothing`.

function Base.iterate(itr::SheetRowStreamIterator, state::Union{Nothing, SheetRowStreamIteratorState}=nothing)
    local current_row
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
            @assert XML.tag(reader[2]) == "worksheet"
            ws_elements = XML.children(reader[2])
            idx = findfirst(y -> y=="sheetData", [XML.tag(x) for x in ws_elements])
            next_element= idx===nothing ? "" : (ws_elements[idx+1])
#            println(next_element)
            
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
                    current_row = parse(Int, XML.attributes(lzstate)["r"])
                    nc = length(filter(n -> XML.nodetype(n) == XML.Element && XML.tag(n) == "c", XML.children(lzstate))) # number of cells in this row
                    cell_no = 0
                    break
                end
                next = iterate(reader, lzstate)
            end
            SheetRowStreamIteratorState(reader, lzstate, #=nelements, el_no, =# current_row)
        end
    end

    # given that the first iteration case is done in the code above, we shouldn't get it again in here
    @assert state !== nothing "Error processing Worksheet $(ws.name): shouldn't get first iteration case again."

    reader = state.itr
    lzstate = state.itr_state

    # Expecting iterator to be at the first row element
#    @assert XML.tag(lzstate) == "row" "Expecting a row element, but got $(XML.tag(lzstate))"


    current_row = state.row
    
    rowcells = Dict{Int, Cell}() # column -> cell
 
#    cell_no = 0
#    nc=0
    next = iterate(reader, lzstate) # iterate through row cells
    while next !== nothing
        (lznode, lzstate) = next
        if XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == next_element # This is the end of sheetData
            println("Next element : = ", next_element )
            return nothing
        elseif XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == "row" # This is the next row
            current_row = parse(Int, XML.attributes(lznode)["r"])
            nc = length(filter(n -> XML.nodetype(n) == XML.Element && XML.tag(n) == "c", XML.children(lzstate))) # number of cells in this row
            cell_no = 0
#            println(nc, " cells found in this row : ", lzstate)
        #            break
        elseif XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == "c" # This is a cell
            cell_no += 1
            cell = Cell(lznode)
            @assert row_number(cell) == current_row "Error processing Worksheet $(ws.name): Inconsistent state: expected row number $(current_row), but cell has row number $(row_number(cell))"
            rowcells[column_number(cell)] = cell
            if cell_no == nc
                sheet_row = SheetRow(get_worksheet(itr), current_row, rowcells) # does this only after all cells found.
                break
            end
        end
        next = iterate(reader, lzstate)
        if next === nothing
            return nothing
        end
    end

    if nc==0
        sheet_row=SheetRow(get_worksheet(itr), current_row, rowcells)
    end
#    println("4 : ",sheet_row)
#    println(next)

    # update state
    state.row = current_row
    state.itr_state = lzstate
#    println("stream160\n", state)
#    println("5 : ",current_row)#, " ", sheet_row)
    return sheet_row, state

end

#=
function Base.iterate(itr::SheetRowStreamIterator, state::Union{Nothing, SheetRowStreamIteratorState}=nothing)

    ws = get_worksheet(itr)

    if state === nothing # first iteration. Will open a stream and create the first state instance
        state = let
            target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
            zip_io, lzreader = open_internal_file_stream(get_xlsxfile(ws), target_file)

            # The reader will be positioned in the first row element inside sheetData
            # First, let's look for sheetData opening element
            nextlznode = iterate(lzreader)
            while nextlznode !== nothing
                (lznode, lzstate) = nextlznode
                if XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == "sheetData"
                    @assert XML.depth(lznode) == 1 "Malformed Worksheet \"$(ws.name)\": unexpected node depth for sheetData node: $(XML.depth(lznode))."
                    break
                end
                nextlznode = iterate(lzreader, lzstate)
            end

            @assert XML.tag(nextlznode[begin]) == "sheetData" "Malformed Worksheet \"$(ws.name)\": Couldn't find sheetData element."

            se=XML.children(nextlznode[begin])
            last_sheet_child = length(se)>0 ? filter(n -> XML.nodetype(n) == XML.Element, se)[end] : nothing

            nextlznode = iterate(lzreader, nextlznode[end])
            # Now let's look for a row element, if it exists
            while nextlznode !== nothing # go next node
                (lznode, lzstate) = nextlznode
                if XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == "row"
                    break
                end
                    if is_end_of_sheet_data(lznode, last_sheet_child)
                       # this Worksheet has no rows
                       # close(reader) # No longer necessary??
                       # close(zip_io) # No longer necessary??
                        return nothing
                    end
                nextlznode = iterate(lzreader, lzstate)
            end
            println("stream188 : ", lzreader)
            println("stream188 : ", nextlznode)
            # row number is set to 0 in the first state
            SheetRowStreamIteratorState(zip_io, lzreader, isnothing(nextlznode) ? nothing : nextlznode[end], last_sheet_child, true, 0)
        end
    end

    # given that the first iteration case is done in the code above, we shouldn't get it again in here
    @assert state !== nothing "Error processing Worksheet $(ws.name): shouldn't get first iteration case again."


    lzreader = state.xml_stream_reader
    lznode = state.lzstate
#    nexttnode = iterate(reader, state)
#    println(lznode)
#    exit()
println("stream207 : ", lznode)
println("stream207 : ", state.sheet_end)
    if is_end_of_sheet_data(lznode, state.sheet_end)
#        @assert !isopen(state)
        return nothing
#    else
#        @assert isopen(state) "Error processing Worksheet $(ws.name): Can't fetch rows from a closed workbook."
    end

    # will read next row from stream.
    # The stream should be already positioned in the next row
    @assert XML.tag(lznode) == "row"
    current_row = parse(Int, XML.attributes(lznode)["r"])
    rowcells = Dict{Int, Cell}() # column -> cell

    row_end = XML.children(lznode)[end]

    nextlznode = iterate(lzreader, lznode)

    # iterate thru row cells
    while nextlznode !== nothing
#        nexttnode = iterate(reader, state)
        if is_end_of_sheet_data(lznode, state.sheet_end)
#            close(state)
            break
        end

        # If this is the end of this row, will point to the next row or set the end of this stream
        if lznode==row_end

            while true
                if is_end_of_sheet_data(lznode, state.sheet_end)
#                    close(state)
                    break
                elseif XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == "row"
                    break
                end
                nextlznode = iterate(reader, lznode)
                @assert nextlznode !== nothing
            end

            # breaks while loop to return current row
            break

        elseif XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == "c"

            # reads current cell to rowcells
            cell = Cell(lznode)
            @assert row_number(cell) == current_row "Error processing Worksheet $(ws.name): Inconsistent state: expected row number $(current_row), but cell has row number $(row_number(cell))"
            rowcells[column_number(cell)] = cell

        elseif XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == "row"
            # last row has no child elements, so we're already pointing to the next row
            break
        end
    end

    sheet_row = SheetRow(get_worksheet(itr), current_row, rowcells)

    # update state
    return sheet_row, state
end
=#
#=
function Base.iterate(itr::SheetRowStreamIterator, state::Union{Nothing, SheetRowStreamIteratorState}=nothing)
    ws = get_worksheet(itr)
    sheetData_rows = Vector{XML.LazyNode}()

    if isnothing(state)
        state = let 
            target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
            doc = open_internal_file_stream(get_xlsxfile(ws), target_file)
            if length(doc) > 0
                for child in doc
                    if XML.tag(child) == "sheetData"
                        sheetData_rows = filter(n -> XML.nodetype(n) == XML.Element && XML.tag(n) == "row", XML.children(child))
                        break
                    end
                end
            end

            nrows=length(sheetData_rows)
            if nrows == 0
                return nothing
            end

            SheetRowStreamIteratorState(sheetData_rows, nrows, 0, 0)

        end
    end

    @assert state !== nothing "Error processing Worksheet $(ws.name): shouldn't get first iteration case again."

    state.vrow += 1
    if state.vrow > state.nrows
        return nothing
    end

    current_row = state.sheetData_rows[state.vrow]

    @assert XML.tag(current_row) == "row"

    state.wsrow = parse(Int, XML.attributes(current_row)["r"])
    rowcells = Dict{Int, Cell}() # column -> cell

    for c in XML.children(current_row)
        if XML.tag(c) == "c"
            cell = Cell(c)
            @assert row_number(cell) == state.wsrow "Error processing Worksheet $(ws.name): Inconsistent state: expected row number $(state.wsrow), but cell has row number $(row_number(cell))"
            rowcells[column_number(cell)] = cell
        end
    end
    sheet_row = SheetRow(ws, state.wsrow, state.vrow, rowcells)

    # update state
    return sheet_row, state

end
=#    


# Detects a closing sheetData element
#@inline is_end_of_sheet_data(n::XML.LazyNode, e::Union{Nothing, XML.LazyNode}) = isnothing(e)==0 ? true : (n == e)

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
#    println(r)
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
 #   println("stream390 : ", row_from_last_iteration)

    #the sorting operation is very costly when adding row and only needed if we use the row iterator
    if ws_cache.dirty
#        println("sorting dirty cache")
        sort!(ws_cache.rows_in_cache)
        ws_cache.row_index = Dict{Int, Int}(ws_cache.rows_in_cache[i] => i for i in 1:length(ws_cache.rows_in_cache))
        ws_cache.dirty = false
    end

    if row_from_last_iteration == 0 && !isempty(ws_cache.rows_in_cache)
#        println("stream400 : ",ws_cache.rows_in_cache)
#        println(ws_cache.row_index)
#        println(ws_cache.cells[current_row_number])

        # the next row is in cache, and it's the first one
        current_row_number = ws_cache.rows_in_cache[1]
#        println("stream313 : ",ws_cache.rows_in_cache)
#        println(ws_cache.row_index)
#        println(ws_cache.cells[current_row_number])
        sheet_row_cells = ws_cache.cells[current_row_number]
        return SheetRow(get_worksheet(ws_cache), current_row_number, sheet_row_cells), current_row_number

    elseif row_from_last_iteration != 0 && ws_cache.row_index[row_from_last_iteration] < length(ws_cache.rows_in_cache)
#        println("stream411")
        # the next row is in cache
        current_row_number = ws_cache.rows_in_cache[ws_cache.row_index[row_from_last_iteration] + 1]
        sheet_row_cells = ws_cache.cells[current_row_number]
        return SheetRow(get_worksheet(ws_cache), current_row_number, sheet_row_cells), current_row_number

    else
#        println("stream420")
#        println("printing ws_cache for $(ws_cache.stream_iterator)")
#        println("cells          : ", ws_cache.cells)
#        println("rows in cache  : ", ws_cache.rows_in_cache )
#        println("row index      : ", ws_cache.row_index)
#        println("stream420")
#        println(ws_cache.stream_state)
        next = iterate(ws_cache.stream_iterator, ws_cache.stream_state)
        if next === nothing
#            println("stream423")
            return nothing
        end
#        println("stream425", next)
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

#cache_sr(wsc::WorksheetCache) = wsc.stream_state.vrow

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
#    println("stream498 ")
    if is_cache_enabled(ws)
#        println("stream500 ")
        if ws.cache === nothing
            ws.cache = WorksheetCache(ws)
        end
        return ws.cache
    else
#        println("stream506 ")
        return SheetRowStreamIterator(ws)
    end
end

function Base.isempty(sr::SheetRow)
    return isempty(sr.rowcells)
end
