
#=
https://docs.julialang.org/en/stable/manual/interfaces/#man-interface-iteration-1

for i = I   # or  "for i in I"
    # body
end

is translated into:

state = start(I)
while !done(I, state)
    (i, state) = next(I, state)
    # body
end
=#

@inline worksheet(r::SheetRow) = r.sheet
@inline worksheet(itr::SheetRowIterator) = itr.sheet

# creates SheetRow with unpopulated rowcells
SheetRow(ws::Worksheet, row::Int, xml_element::LightXML.XMLElement) = SheetRow(ws, row, xml_element, Dict{Int, Cell}(), false)

function populate_row_cells!(r::SheetRow)
    if !r.is_rowcells_populated
        for c in r.row_xml_element["c"]
            cell = Cell(c)
            @assert row_number(cell) == r.row "Malformed Excel file. range_row = $(r.row), cell.ref = $(cell.ref)"
            r.rowcells[column_number(cell)] = cell
        end
        r.is_rowcells_populated = true
    end
    nothing
end

Base.start(itr::SheetRowIterator) = start(itr.xml_rows_iterator)
Base.done(itr::SheetRowIterator, state) = done(itr.xml_rows_iterator, state)

#(i, state) = next(I, state)
function Base.next(itr::SheetRowIterator, state)
    xml_element, next_state = next(itr.xml_rows_iterator, state)
    row = parse(Int, LightXML.attribute(xml_element, "r"))
    return SheetRow(worksheet(itr), row, xml_element), next_state
end

function find_row(itr::SheetRowIterator, row::Int) :: SheetRow
    for r in itr
        if row_number(r) == row
            return r
        end
    end
    error("Row $row not found.")
end

function SheetRowIterator(ws::Worksheet)
    xroot = LightXML.root(ws.data)
    @assert LightXML.name(xroot) == "worksheet" "Malformed sheet $(ws.name)."
    vec_sheetdata = xroot["sheetData"]
    @assert length(vec_sheetdata) <= 1 "Malformed sheet $(ws.name)."
    return SheetRowIterator(ws, LightXML.child_elements(vec_sheetdata[1]))
end

row_number(sr::SheetRow) = sr.row

function getcell(r::SheetRow, column_index::Int) :: AbstractCell
    populate_row_cells!(r)

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

getdata(r::SheetRow, column) = celldata(worksheet(r), getcell(r, column))

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

#
# Table
#

function Base.isempty(sr::SheetRow)
    populate_row_cells!(sr)
    return isempty(sr.rowcells)
end

"""
   column_bounds(sr::SheetRow)

Returns a tuple with the first and last index of the columns for a `SheetRow`.
"""
function column_bounds(sr::SheetRow)

    @assert !isempty(sr) "Can't get column bounds from an empty row."

    local first_column_index::Int = first(keys(sr.rowcells))
    local last_column_index::Int = first_column_index

    for k in keys(sr.rowcells)
        if k < first_column_index
            first_column_index = k
        end

        if k > last_column_index
            last_column_index = k
        end
    end

    return (first_column_index, last_column_index)
end

# anchor_column will be the leftmost column of the column_bounds
function last_column_index(sr::SheetRow, anchor_column::Int) :: Int

    @assert !isempty(getcell(sr, anchor_column)) "Can't get column bounds based on an empty anchor cell."

    local first_column_index::Int = anchor_column
    local last_column_index::Int = first_column_index

    if length(keys(sr.rowcells)) == 1
        return anchor_column
    end

    columns = sort(collect(keys(sr.rowcells)))
    first_i = findfirst(colindex -> colindex == anchor_column, columns)
    last_column_index = anchor_column

    for i in (first_i+1):length(columns)
        if columns[i] - 1 != last_column_index
            return last_column_index
        end

        last_column_index = columns[i]
    end

    return last_column_index
end

"""
    TableRowIterator(sheet, [column_range], [first_row]; [names], [header])

`header` is a boolean indicating wether the first row of the table is a table header.

If `header == false` and no `names` were supplied, column names will be generated following the column names found in the Excel file.
Also, the column range will be inferred by the non-empty contiguous cells in the first row of the table.

The user can replace column names by assigning the optional `names` input variable with a `Vector{Symbol}`.
"""
function TableRowIterator(sheet::Worksheet, cols::Union{ColumnRange, AbstractString}; first_row::Int=_find_first_row_with_data(sheet, convert(ColumnRange, cols).start), column_labels::Vector{Symbol}=Vector{Symbol}(), header::Bool=true)
    itr = SheetRowIterator(sheet)
    column_range = convert(ColumnRange, cols)

    if isempty(column_labels)
        if header
            # will use celldata to get column names
            for column_index in column_range.start:column_range.stop
                sheet_row = find_row(itr, first_row)
                cell = getcell(sheet_row, column_index)
                @assert !isempty(cell) "Header cell can't be empty."
                push!(column_labels, Symbol(celldata(sheet, cell)))
            end
        else
            # generate column_labels if there's no header information anywhere
            for c in column_range
                push!(column_labels, c)
            end
        end
    else
        # check consistency for column_range and column_labels
        @assert length(column_labels) == length(column_range) "`column_range` (length=$(length(column_range))) and `column_labels` (length=$(length(column_labels))) must have the same length."
    end

    first_data_row = header ? first_row + 1 : first_row
    return TableRowIterator(itr, Index(column_range, column_labels), first_data_row)
end

function TableRowIterator(sheet::Worksheet; first_row::Int = 1, column_labels::Vector{Symbol}=Vector{Symbol}(), header::Bool=true)
    for r in eachrow(sheet)

        # skip rows until we reach first_row
        if row_number(r) < first_row
            continue
        end

        if !isempty(r)
            columns_ordered = sort(collect(keys(r.rowcells)))
            for (ci, cn) in enumerate(columns_ordered)
                if !Missings.ismissing(getdata(r, cn))
                    # found a row with data. Will get ColumnRange from non-empty consecutive cells
                    first_row = row_number(r)
                    column_start = cn
                    column_stop = cn

                    if length(columns_ordered) == 1
                        # there's only one column
                        column_range = ColumnRange(column_start, column_stop)
                        return TableRowIterator(sheet, column_range; first_row=first_row, column_labels=column_labels, header=header)
                    else
                        # will figure out the column range
                        for ci_stop in (ci+1):length(columns_ordered)
                            cn_stop = columns_ordered[ci_stop]
                            if cn_stop - 1 != column_stop
                                column_range = ColumnRange(column_start, column_stop)
                                return TableRowIterator(sheet, column_range; first_row=first_row, column_labels=column_labels, header=header)
                            end
                            column_stop = cn_stop
                        end
                    end

                    # if got here, it's because all columns are non-empty
                    column_range = ColumnRange(column_start, column_stop)
                    return TableRowIterator(sheet, column_range; first_row=first_row, column_labels=column_labels, header=header)
                end
            end
        end
    end

    error("Couldn't find a table in sheet $(sheet.name)")
end

function _find_first_row_with_data(sheet::Worksheet, column_number::Int)
    # will find first_row
    for r in eachrow(sheet)
        if !Missings.ismissing(getdata(r, column_number))
            return row_number(r)
        end
    end
    error("Column $(encode_column_number(column_number)) has no data.")
end

@inline worksheet(tri::TableRowIterator) = tri.itr.sheet
@inline worksheet(r::TableRow) = worksheet(r.itr)

"""
Returns real sheet column numbers (based on cellref)
"""
@inline sheet_column_numbers(i::Index) = values(i.column_map)
@inline table_column_numbers(i::Index) = eachindex(i.column_labels)
@inline table_column_numbers(r::TableRow) = table_column_numbers(r.itr.index)

"""
Maps table column index (1-based) -> sheet column index (cellref based)
"""
@inline table_column_to_sheet_column_number(index::Index, table_column_number::Int) = index.column_map[table_column_number]
@inline table_columns_count(i::Index) = length(i.column_labels)
@inline table_columns_count(itr::TableRowIterator) = table_columns_count(itr.index)
@inline table_columns_count(r::TableRow) = table_columns_count(r.itr)
@inline table_row_number(r::TableRow) = r.table_row_index
@inline sheet_row_number(r::TableRow) = row_number(r.sheet_row)
@inline get_column_labels(index::Index) = index.column_labels
@inline get_column_labels(itr::TableRowIterator) = get_column_labels(itr.index)
@inline get_column_labels(r::TableRow) = get_column_labels(r.itr)
@inline get_column_label(r::TableRow, table_column_number::Int) = get_column_labels(r)[table_column_number]

# iterate into TableRow to get each column value
Base.start(r::TableRow) = start(table_column_numbers(r))
Base.done(r::TableRow, state) = done(table_column_numbers(r), state)

function Base.next(r::TableRow, state)
    (next_column_number, next_state) = next(table_column_numbers(r), state)
    return (r[next_column_number], next_state)
end

function getcell(r::TableRow, table_column_number::Int)
    table_row_iterator = r.itr
    index = table_row_iterator.index
    sheet_row_iterator = table_row_iterator.itr
    sheet = worksheet(r)
    sheet_column = table_column_to_sheet_column_number(index, table_column_number)
    sheet_row = r.sheet_row
    return getcell(sheet_row, sheet_column)
end

getdata(r::TableRow, table_column_number::Int) = celldata(worksheet(r), getcell(r, table_column_number))

function getdata(r::TableRow, column_label::Symbol)
    index = r.itr.index
    if haskey(index.lookup, column_label)
        return getindex(r, index.lookup[column_label])
    else
        error("Invalid column label: $column_label.")
    end
end

Base.getindex(r::TableRow, x) = getdata(r, x)

struct TableRowIteratorState
    state::Any
    sheet_row::SheetRow
    table_row_index::Int
    last_sheet_row_number::Int
    is_done::Bool
end

function Base.start(itr::TableRowIterator)
    last_state = start(itr.itr)

    # go to the first_data_row
    while !done(itr.itr, last_state)
        (sheet_row, state) = next(itr.itr, last_state)
        if row_number(sheet_row) == itr.first_data_row
            return TableRowIteratorState(state, sheet_row, 1, itr.first_data_row, false)
        end
        last_state = state
    end

    return TableRowIteratorState(state, sheet_row, 1, itr.first_data_row, true)
end

function Base.next(itr::TableRowIterator, state::TableRowIteratorState)
    sheet_row_itr_is_done = done(itr.itr, state.state)
    current_sheet_row = state.sheet_row
    current_sheet_row_number = row_number(current_sheet_row)

    if !sheet_row_itr_is_done
        next_sheet_row, next_state = next(itr.itr, state.state)
    else
        next_sheet_row, next_state = current_sheet_row, state
    end

    return TableRow(itr, current_sheet_row, state.table_row_index), TableRowIteratorState(next_state, next_sheet_row, state.table_row_index+1, current_sheet_row_number, sheet_row_itr_is_done)
end

function Base.done(itr::TableRowIterator, state::TableRowIteratorState)
    # empty table case
    if state.is_done
        return true
    end

    # check skipping rows
    if row_number(state.sheet_row) != itr.first_data_row && row_number(state.sheet_row) != (state.last_sheet_row_number + 1)
        return true
    end

    # check empty rows
    if isempty(state.sheet_row)
        return true
    end

    # check if there are any data inside column range
    for c in sheet_column_numbers(itr.index)
        if !Missings.ismissing(celldata(worksheet(itr), getcell(state.sheet_row, c)))
            return false
        end
    end

    return true
end

function infer_eltype(v::Vector)
    local hasmissing::Bool = false
    local t::DataType = Any

    if isempty(v)
        return eltype(v)
    end

    for i in 1:length(v)
        if Missings.ismissing(v[i])
            hasmissing = true
        else
            if t != Any && typeof(v[i]) != t
                return Any
            else
                t = typeof(v[i])
            end
        end
    end

    if t == Any
        return Any
    else
        if hasmissing
            return Union{Missings.Missing, t}
        else
            return t
        end
    end
end

function gettable(itr::TableRowIterator; infer_eltypes::Bool=false)
    column_labels = get_column_labels(itr)
    columns_count = length(column_labels)
    data = Vector{Any}(columns_count)
    for c in 1:columns_count
        data[c] = Vector{Any}()
    end

    for r in itr
        for c in 1:columns_count
            push!(data[c], getdata(r, c))
        end
    end

    if infer_eltypes
        rows = length(data[1])
        for c in 1:columns_count
            new_column_data = Vector{infer_eltype(data[c])}(rows)
            for r in 1:rows
                new_column_data[r] = data[c][r]
            end
            data[c] = new_column_data
        end
    end

    return data, column_labels
end
