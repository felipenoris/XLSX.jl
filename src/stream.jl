
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

# Creates a reader for row elements in the Worksheet's XML.
# Will return a stream reader positioned in the first row element if it exists.
# If there's no row element inside sheetData XML tag, it will close all streams and return `nothing`.
function Base.iterate(itr::SheetRowStreamIterator, state::Union{Nothing, SheetRowStreamIteratorState}=nothing)
    local current_row
    local current_row_ht
    local sheet_row
    local nc = 0
    local cell_no=0

    ws = get_worksheet(itr)

    if isnothing(state) # first iteration. Will open a LazyNode for iteration, find the first row and create the first state instance

        state = let 

            target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
            reader = open_internal_file_stream(get_xlsxfile(ws), target_file)

            length(reader) <= 0 && throw(XLSXError("Couldn't open reader for Worksheet $(ws.name)."))
            XML.tag(reader[end]) != "worksheet" && throw(XLSXError("Expecting to find a worksheet node.: Found a $(XML.tag(reader[end]))."))
            next_element=XML.next(reader)

            while XML.tag(next_element) != "sheetData" # Check for `sheetData`
                next_element = XML.next(next_element)
                next_element === nothing && throw(XLSXError("No `sheetData` node found in worksheet"))
            end

            next = iterate(reader)

            while next !== nothing
                (lznode, lzstate) = next

                if XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == "sheetData"

                    XML.depth(lznode) != 2 && throw(XLSXError("Malformed Worksheet \"$(ws.name)\": unexpected node depth for sheetData node: $(XML.depth(lznode))."))

                    while XML.tag(lznode) != "row" # Check for at least one `row`
                        lznode = XML.next(lznode)
                        lznode === nothing && return nothing # no rows found
                    end

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
                    nc = 0
                    for child in XML.children(lzstate)
                        XML.nodetype(child) == XML.Element && XML.tag(child) == "c" && (nc += 1)
                    end
                    #nc = length(filter(n -> XML.nodetype(n) == XML.Element && XML.tag(n) == "c", XML.children(lzstate))) # number of cells in this row
                    cell_no = 0
                    break
                end
                next = iterate(reader, lzstate)
            end
            SheetRowStreamIteratorState(reader, lzstate, current_row, current_row_ht)
        end
    end

    # given that the first iteration case is done in the code above, we shouldn't get it again in here
    state === nothing && throw(XLSXError("Error processing Worksheet $(ws.name): shouldn't get first iteration case again."))

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

        if XML.nodetype(lznode) == XML.Element && XML.depth(lznode) == 2 # This is the end of sheetData
            return nothing

        elseif XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == "row" # This is the next row
            a = XML.attributes(lzstate)
            current_row = parse(Int, a["r"])
            current_row_ht = haskey(a, "ht") ? parse(Float64, a["ht"]) : nothing
            nc = 0
            for child in XML.children(lzstate)
                XML.nodetype(child) == XML.Element && XML.tag(child) == "c" && (nc += 1)
            end
            cell_no = 0

        elseif XML.nodetype(lznode) == XML.Element && XML.tag(lznode) == "c" # This is a cell
            cell_no += 1
            cell = Cell(lznode)
            if row_number(cell) != current_row
                throw(XLSXError("Error processing Worksheet $(ws.name): Inconsistent state: expected row number $(current_row), but cell has row number $(row_number(cell))"))
            end
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
function Base.eachrow(ws::Worksheet) :: SheetRowIterator
    if is_cache_enabled(ws)
        if ws.cache === nothing
#            ws.cache = WorksheetCache(ws) # the old way - fill cache incrementally only as far as needed using iterator
            target_file = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
            lznode = open_internal_file_stream(get_xlsxfile(ws), target_file)
            first_cache_fill!(ws, lznode, Threads.nthreads()) # fill cache if enabled but empty on first use of eachrow iterator
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
function stream_rows(n::XML.LazyNode, handled_attributes::Set{String}; channel_size::Int=1 << 20)
    n = XML.next(n)
    Channel{Tuple{XML.LazyNode, Set{String}}}(channel_size) do out
        while !isnothing(n)
            if n.tag == "row"
                put!(out, (n, handled_attributes))
            end
            n = XML.next(n)
        end
    end
end

function process_row(row::XML.LazyNode, handled_attributes::Set{String}, ws::Worksheet)
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
            cell = Cell(c)
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
    handled_attributes = Set{String}([
        "r",     # the row number
        "spans", # the columns the row spans
        "ht",    # the row height
    ])
    unhandled_attributes = Dict{Int,Dict{String,String}}() # Row number => (name, value)

    if ws.cache === nothing
        ws.cache = WorksheetCache(ws)
    else
        throw(XLSXError("Expecting empty cache but cache not empty!"))
    end

    sheet_rows = Channel{Tuple{Int, SheetRow, Dict{String,String}}}(1 << 20)

    consumer = @async begin
        sst_total=0
        for (row_sst_count, sheet_row, unatt) in sheet_rows
            if !isempty(unatt)
                unhandled_attributes[row_number(sheet_row)] = unatt
            end
            push_sheetrow!(ws.cache, sheet_row)
            sst_total += row_sst_count
        end
        ws.sst_count = sst_total
        ws.unhandled_attributes = isempty(unhandled_attributes) ? nothing : unhandled_attributes
    end

    rows = stream_rows(lznode, handled_attributes)

    # Producer tasks
    @sync for _ in 1:nthreads
        Threads.@spawn begin
            for (row, handled_attributes) in rows
                sheetrow = process_row(row, handled_attributes, ws) # process <row> LazyNodes into SheetRows
                put!(sheet_rows, sheetrow)
            end
        end
    end
    close(sheet_rows)

    wait(consumer)

    ws.cache.is_full=true
end

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
                        cell = Cell(c)
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
