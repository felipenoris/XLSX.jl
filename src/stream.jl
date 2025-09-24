
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
* `SheetRowStreamIterator` is a dumb iterator for row elements in sheetData XML tag of a worksheet. Empty rows are not represented in the XML file so cannot be seen by the iterator.
* `WorksheetCache` has a `SheetRowStreamIterator` and caches all values read from the stream.
* `TableRowIterator` is a smart iterator that looks for tabular data, but uses a SheetRowIterator under the hood.

The implementation of `SheetRowIterator` will be chosen automatically by `eachrow` method,
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

    !internal_xml_file_exists(xf, filename) && throw(XLSXError("Couldn't find $filename in $(xf.source)."))

    return XML.LazyNode(XML.Raw(ZipArchives.zip_readentry(xf.io, filename)))

end

function get_rowcells!(rowcells::Dict{Int, Cell}, row::XML.LazyNode, ws::Worksheet)
#=
    @assert row.tag == "row" "Not a row node"

    sst_count=0
    d=row.depth

    row_cellnodes = Channel{XML.LazyNode}(1 << 20)
    row_cells = Channel{XLSX.Cell}(1 << 20)

    # consumer task
    consumer = @async begin
        for cell in row_cells        
            sst_count += cell.datatype == "s" ? 1 : 0
            @inbounds rowcells[column_number(cell)] = cell
        end
    end

    # Feed row_cellnodes
    cellnode=XML.next(row)
    while cellnode.depth > d
        if cellnode.tag == "c" # This is a cell
            put!(row_cellnodes, cellnode)
#            cell = Cell(cellnode, ws) # construct an XLSX.Cell from an XML.LazyNode
#            if row_number(cell) != current_row
#                throw(XLSXError("Error processing Worksheet $(ws.name): Inconsistent state: expected row number $(current_row), but cell has row number $(row_number(cell))"))
#            end
        end
        cellnode = XML.next(cellnode)
    end
    close(row_cellnodes)

    # Producer tasks
    @sync for _ in 1:Threads.nthreads()
        Threads.@spawn begin
            for cn in row_cellnodes
                cell = Cell(cn, ws)
#                println(cell)
                put!(row_cells, cell)
            end
        end
    end

    close(row_cells)

#=
    cellnode=XML.next(row)
    while cellnode.depth > d
        if cellnode.tag == "c" # This is a cell
            put!(row_cellnodes, copy(cellnode))
#            cell = Cell(cellnode, ws) # construct an XLSX.Cell from an XML.LazyNode
#            if row_number(cell) != current_row
#                throw(XLSXError("Error processing Worksheet $(ws.name): Inconsistent state: expected row number $(current_row), but cell has row number $(row_number(cell))"))
#            end
        end
        cellnode = XML.next(cellnode)
    end

    close(row_cellnodes)
=#
    wait(consumer)  # ensure consumer is done

    if cellnode.tag == "row" # have reached the end og last row, beginning of next
        return cellnode, sst_count
    else
        return nothing, sst_count
    end
=#
    @assert row.tag == "row" "Not a row node"

    sst_count=0

    d=row.depth

    cellnode=XML.next(row)

    while cellnode.depth > d
        if cellnode.tag == "c" # This is a cell
            cell = Cell(cellnode, ws) # construct an XLSX.Cell from an XML.LazyNode
            sst_count += cell.datatype == "s" ? 1 : 0
            @inbounds rowcells[column_number(cell)] = cell
        end
        cellnode = XML.next(cellnode)
    end
    if !isnothing(cellnode) && cellnode.tag == "row" # have reached the end of last row, beginning of next
        return cellnode, sst_count
    else                                             # no more rows
        return nothing, sst_count
    end
end

# Creates an iterator for row elements in the Worksheet's XML.
function Base.iterate(itr::SheetRowStreamIterator)
    ws = get_worksheet(itr)
    target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
    sheetnode = open_internal_file_stream(get_xlsxfile(ws), target_file) # worksheet target files are LazyNodes

    length(sheetnode) <= 0 && throw(XLSXError("Couldn't open reader for Worksheet $(ws.name)."))
    XML.tag(sheetnode[end]) != "worksheet" && throw(XLSXError("Expecting to find a worksheet node.: Found a $(XML.tag(sheetnode[end]))."))

    sheetnode=XML.next(sheetnode)

    while XML.tag(sheetnode) != "sheetData" # Check for `sheetData`
        sheetnode = XML.next(sheetnode)
        sheetnode === nothing && throw(XLSXError("No `sheetData` node found in worksheet"))
    end

    XML.depth(sheetnode) != 2 && throw(XLSXError("Malformed Worksheet \"$(ws.name)\": unexpected node depth for sheetData node: $(XML.depth(lznode))."))

    rownode=XML.next(sheetnode)

    while XML.tag(rownode) != "row" # Check for at least one `row`
        rownode = XML.next(rownode)
        rownode === nothing && return nothing # no rows found
    end

    # rownode is the now the first row
    a = XML.attributes(rownode) # get row number and row heigth (if specified)
    current_row = parse(Int, a["r"])
    current_row_ht = haskey(a, "ht") ? parse(Float64, a["ht"]) : nothing

    # collect all cells in this row
    rowcells = Dict{Int, Cell}()
    next_rownode, sst_count = get_rowcells!(rowcells, rownode, ws)
    
    itr.sheet.sst_count += sst_count

    sheet_row = SheetRow(ws, current_row, current_row_ht, rowcells) # create the sheet_row

    # debug
#    @assert sheetnode.raw.data == next_rownode.raw.data "LazyNode data don't match"

    return sheet_row, SheetRowStreamIteratorState(next_rownode, rowcells)
end

function Base.iterate(itr::SheetRowStreamIterator, state::Union{Nothing, SheetRowStreamIteratorState})
    ws = get_worksheet(itr)
    rownode = state.next_rownode
    rowcells = state.rowcells
    empty!(rowcells)

    if rownode === nothing # there is no next_rownode - all rows processed
        return nothing
    end

    # get row number and row heigth (if specified)
    a = XML.attributes(rownode)
    current_row = parse(Int, a["r"])
    current_row_ht = haskey(a, "ht") ? parse(Float64, a["ht"]) : nothing

    # collect all cells in this row
    next_rownode, sst_count = get_rowcells!(rowcells, rownode, ws)
    
    itr.sheet.sst_count += sst_count

    sheet_row = SheetRow(ws, current_row, current_row_ht, rowcells) # create the sheet_row

    return sheet_row, SheetRowStreamIteratorState(next_rownode, rowcells)
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
# The state is the row number and a flag for if the cache is full or being filled. The element is a SheetRow.
#
function WorksheetCache(ws::Worksheet)
    itr = SheetRowStreamIterator(ws)
    return WorksheetCache(false, CellCache(), Vector{Int}(), Dict{Int, Union{Float64, Nothing}}(), Dict{Int, Int}(), itr, nothing, true)
end

@inline get_worksheet(r::SheetRow) = r.sheet
@inline get_worksheet(itr::WorksheetCache) = get_worksheet(itr.stream_iterator)

# In the WorksheetCache iterator, the element is a SheetRow, the state is the row number and a flag on whether the cache is already full or not
function Base.iterate(ws_cache::WorksheetCache, state::Union{Nothing, WorksheetCacheIteratorState}=nothing)

    # If first iteration, check if cache is full
    if isnothing(state)
        if ws_cache.is_full
            state=WorksheetCacheIteratorState(0, true)
        else
            state=WorksheetCacheIteratorState(0, false)
        end
    end

    # the sorting operation is very costly when adding row and only needed if we use the row iterator
    if ws_cache.dirty
        sort!(ws_cache.rows_in_cache)
        ws_cache.row_index = Dict{Int, Int}(ws_cache.rows_in_cache[i] => i for i in 1:length(ws_cache.rows_in_cache))
        ws_cache.dirty = false
    end


    # read from cache
    if state.row_from_last_iteration == 0 && !isempty(ws_cache.rows_in_cache)
        # the next row is in cache, and it's the first one
        current_row_number = ws_cache.rows_in_cache[1]
        current_row_ht = ws_cache.row_ht[current_row_number]
        sheet_row_cells = ws_cache.cells[current_row_number]
        state.row_from_last_iteration=current_row_number
        return SheetRow(get_worksheet(ws_cache), current_row_number, current_row_ht, sheet_row_cells), state

    elseif state.row_from_last_iteration != 0 && ws_cache.row_index[state.row_from_last_iteration] < length(ws_cache.rows_in_cache)
        # the next row is in cache
        current_row_number = ws_cache.rows_in_cache[ws_cache.row_index[state.row_from_last_iteration] + 1]
        current_row_ht = ws_cache.row_ht[current_row_number]
        sheet_row_cells = ws_cache.cells[current_row_number]
        state.row_from_last_iteration=current_row_number
        return SheetRow(get_worksheet(ws_cache), current_row_number, current_row_ht, sheet_row_cells), state

    end

    if !state.full_cache
        # cache not yet full, read from file.
        # NOTE: cache is always full here now 
#        next = iterate(ws_cache.stream_iterator, ws_cache.stream_state)

#        if next === nothing
#            ws_cache.is_full = true
#            return nothing
#        end
#        sheet_row, next_stream_state = next

        # add new row to WorkSheetCache
#        push_sheetrow!(ws_cache, sheet_row)

        # update stream state
#        ws_cache.stream_state = next_stream_state

#        state.row_from_last_iteration=row_number(sheet_row)

#        return sheet_row, state
    end
end

function find_row(itr::SheetRowIterator, row::Int) :: SheetRow
    for r in itr
        if row_number(r) == row
            return r
        end
    end
    throw(XLSXError("Row $row not found."))
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
    !is_valid_column_name(column_name) && throw(XLSXError("$column_name is not a valid column name."))
    return getcell(r, decode_column_number(column_name))
end

getdata(r::SheetRow, column::Union{Vector{T}, UnitRange{T}}) where {T<:Integer} = [getdata(get_worksheet(r), getcell(r, x)) for x in column]
getdata(r::SheetRow, column) = getdata(get_worksheet(r), getcell(r, column))
Base.getindex(r::SheetRow, x) = getdata(r, x)

"""
    eachrow(sheet)

Creates a row iterator for a worksheet.

Example: Query all cells from columns 1 to 4.

```julia
left = 1  # 1st column
right = 4 # 4th column
for sheetrow in eachrow(sheet)
    for column in left:right
        cell = XLSX.getcell(sheetrow, column)

        # do something with cell
    end
end
```

Note: The `eachrow` row iterator will not return any row that 
consists entirely of `EmptyCell`s. These are simply not seen 
by the iterator. The `length(eachrow(sheet))` function therefore 
defines the number of rows that are not entirely empty and will, 
in any case, only succeed if the worksheet cache is in use.
"""
function eachrow(ws::Worksheet) :: SheetRowIterator
    if is_cache_enabled(ws)
        if ws.cache === nothing # fill cache if enabled but empty on first use of eachrow iterator
            target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
            lznode = open_internal_file_stream(get_xlsxfile(ws), target_file)
            first_cache_fill!(ws, lznode, Threads.nthreads())
        end
        return ws.cache
    else
        return SheetRowStreamIterator(ws)
    end
end

function Base.isempty(sr::SheetRow)
    return isempty(sr.rowcells)
end

Base.length(r::WorksheetCache)=length(r.cells)

#--------------------------------------------------------------------- Fill cache on first read (multi-threaded)
function stream_rows(n::XML.LazyNode, chunksize::Int; channel_size::Int=1 << 20)

    rows = Vector{XML.LazyNode}(undef, chunksize)
    pos=0
    Channel{Vector{XML.LazyNode}}(channel_size) do out
        while !isnothing(n)
            if n.tag == "row"
                pos += 1
                rows[pos] = n
            end
            if pos >= chunksize
                put!(out, copy(rows))
                pos=0
            end
            n = XML.next(n)
        end
        if pos>0 # handle last incomplete chunk
            put!(out, rows[1:pos])
        end
    end
end

function process_row(row::XML.LazyNode, handled_attributes::Set{String}, ws::Worksheet, mylock::ReentrantLock)
    unhandled_attributes = Dict{String,String}()

    atts = XML.attributes(row)
    if !isnothing(atts)
        current_row_ht = haskey(atts, "ht") ? parse(Float64, atts["ht"]) : nothing
        row_num = haskey(atts, "r") ? parse(Int, atts["r"]) : nothing
        unhandled_attributes = Dict(filter(attr -> !in(first(attr), handled_attributes), atts))
    end

    # Process cells
    rowcells = Dict{Int,Cell}()
    cells = XML.children(row)
    sst_count=0
    for c in cells
        if c.tag == "c"
            cell = Cell(c, ws; mylock)
            col_num = column_number(cell)
            if cell.datatype=="s"
                sst_count += 1
            end

            # Verify row consistency
            if row_number(cell) != row_num
                @warn "Row number mismatch: expected $row_num, got $(row_number(cell))"
            end

            rowcells[col_num] = cell
        end
    end

    return sst_count, SheetRow(ws, row_num, current_row_ht, rowcells), unhandled_attributes

end

function first_cache_fill!(ws::Worksheet, lznode::XML.LazyNode, nthreads::Int)
    chunksize=1000

    handled_attributes = Set{String}([
        "r",            # the row number
        "spans",        # the columns the row spans
        "ht",           # the row height
        "customHeight"  # flag for when custom height is defined
    ])
    unhandled_attributes = Dict{Int,Dict{String,String}}() # Row number => (name, value)

    if ws.cache === nothing
        ws.cache = WorksheetCache(ws)
    else
        throw(XLSXError("Expecting empty cache but cache not empty!"))
    end

    sheet_rows = Channel{Vector{Tuple{Int, SheetRow, Dict{String,String}}}}(1 << 20)

    consumer = @async begin
        sst_total=0
        for rows in sheet_rows
            for (row_sst_count, sheet_row, unatt) in rows
                if !isempty(unatt)
                    unhandled_attributes[row_number(sheet_row)] = unatt
                end
                push_sheetrow!(ws.cache, sheet_row)
                sst_total += row_sst_count
            end
        end
        ws.sst_count = sst_total
        ws.unhandled_attributes = isempty(unhandled_attributes) ? nothing : unhandled_attributes
    end

    streamed_rows = stream_rows(lznode, chunksize)

    # Producer tasks
    mylock = ReentrantLock() # lock for thread-safe access to shared string table in case of inlineStrings
    @sync for _ in 1:nthreads
        Threads.@spawn begin
            for rows in streamed_rows
                row_count=0
                chunk=Vector{Tuple{Int, SheetRow, Dict{String,String}}}(undef, chunksize)
                for row in rows
                    row_count += 1
                    chunk[row_count] = process_row(row, handled_attributes, ws, mylock) # process <row> LazyNodes into SheetRows
                    if row_count == chunksize
                        put!(sheet_rows, copy(chunk))
                        row_count=0
                    end
                end
                if row_count>0 # handle last incomplete chunk
                    put!(sheet_rows, chunk[1:row_count])
                end
            end
        end
    end
    close(sheet_rows)

    wait(consumer) # ensure consumer is done

    ws.cache.is_full=true
end

# Materialise specific rows from a worksheet.xml file into SheetRows
# (faster than using eachrow which materialises every row).
function match_rows(ws::Worksheet, rows_to_match::Vector{Int})::Vector{SheetRow}
    matched_rows=Vector{SheetRow}()

    sort!(rows_to_match)
    i=1

    target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
    lznode = open_internal_file_stream(get_xlsxfile(ws), target_file)
    n = XML.next(lznode)
    while !isnothing(n)
        if n.tag == "row" # find each row
            atts = XML.attributes(n)
            if !isnothing(atts)
                row_num = haskey(atts, "r") ? parse(Int, atts["r"]) : nothing
            end
            if !isnothing(row_num) && row_num == rows_to_match[i] # process matching rows into SheetRows
                current_row_ht = haskey(atts, "ht") ? parse(Float64, atts["ht"]) : nothing

                # Process cells
                rowcells = Dict{Int,Cell}()
                cells = XML.children(n)
                for c in cells
                    if c.tag == "c"
                        cell = Cell(c, ws)
                        col_num = column_number(cell)

                        # Verify row consistency
                        if row_number(cell) != row_num
                            @warn "Row number mismatch: expected $row_num, got $(row_number(cell))"
                        end

                        rowcells[col_num] = cell
                    end
                end

                sheetrow = SheetRow(ws, row_num, current_row_ht, rowcells)
                push!(matched_rows, sheetrow)
                i+=1
                i>length(rows_to_match) && break # stop once all rows matched
            end
        end
        n = XML.next(n)
    end

    return matched_rows
end
