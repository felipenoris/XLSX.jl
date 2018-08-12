
#
# Table
#

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
    eachtablerow(sheet, [columns]; [first_row], [column_labels], [header], [stop_in_empty_row], [stop_in_row_function])

Constructs an iterator of table rows. Each element of the iterator is of type `TableRow`.

`header` is a boolean indicating wether the first row of the table is a table header.

If `header == false` and no `names` were supplied, column names will be generated following the column names found in the Excel file.
Also, the column range will be inferred by the non-empty contiguous cells in the first row of the table.

The user can replace column names by assigning the optional `names` input variable with a `Vector{Symbol}`.

`stop_in_empty_row` is a boolean indicating wether an empty row marks the end of the table.
If `stop_in_empty_row=false`, the iterator will continue to fetch rows until there's no more rows in the Worksheet.
The default behavior is `stop_in_empty_row=true`. Empty rows may be returned by the iterator when `stop_in_empty_row=false`.

`stop_in_row_function` is a Function that receives a `TableRow` and returns a `Bool` indicating if the end of the table was reached.

Example for `stop_in_row_function`:

```
function stop_function(r)
    v = r[:col_label]
    return !Missings.ismissing(v) && v == "unwanted value"
end
```

Example code:
```
for r in XLSX.eachtablerow(sheet)
    # r is a `TableRow`. Values are read using column labels or numbers.
    rn = XLSX.row_number(r) # `TableRow` row number.
    v1 = r[1] # will read value at table column 1.
    v2 = r[:COL_LABEL2] # will read value at column labeled `:COL_LABEL2`.
end
```

See also `gettable`.
"""
function eachtablerow(sheet::Worksheet, cols::Union{ColumnRange, AbstractString}; first_row::Int=_find_first_row_with_data(sheet, convert(ColumnRange, cols).start), column_labels::Vector{Symbol}=Vector{Symbol}(), header::Bool=true, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Function, Void}=nothing) :: TableRowIterator
    itr = eachrow(sheet)
    column_range = convert(ColumnRange, cols)

    if isempty(column_labels)
        if header
            # will use getdata to get column names
            for column_index in column_range.start:column_range.stop
                sheet_row = find_row(itr, first_row)
                cell = getcell(sheet_row, column_index)
                @assert !isempty(cell) "Header cell can't be empty."
                push!(column_labels, Symbol(getdata(sheet, cell)))
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
    return TableRowIterator(itr, Index(column_range, column_labels), first_data_row, stop_in_empty_row, stop_in_row_function)
end

function eachtablerow(sheet::Worksheet; first_row::Int = 1, column_labels::Vector{Symbol}=Vector{Symbol}(), header::Bool=true, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Function, Void}=nothing) :: TableRowIterator
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
                        return eachtablerow(sheet, column_range; first_row=first_row, column_labels=column_labels, header=header, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function)
                    else
                        # will figure out the column range
                        for ci_stop in (ci+1):length(columns_ordered)
                            cn_stop = columns_ordered[ci_stop]

                            # Will stop if finds an empty cell or a skipped column
                            if Missings.ismissing(getdata(r, cn_stop)) || (cn_stop - 1 != column_stop)
                                column_range = ColumnRange(column_start, column_stop)
                                return eachtablerow(sheet, column_range; first_row=first_row, column_labels=column_labels, header=header, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function)
                            end
                            column_stop = cn_stop
                        end
                    end

                    # if got here, it's because all columns are non-empty
                    column_range = ColumnRange(column_start, column_stop)
                    return eachtablerow(sheet, column_range; first_row=first_row, column_labels=column_labels, header=header, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function)
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

@inline get_worksheet(tri::TableRowIterator) = get_worksheet(tri.itr)

"""
Returns real sheet column numbers (based on cellref)
"""
@inline sheet_column_numbers(i::Index) = values(i.column_map)

"""
Returns an iterator for table column numbers.
"""
@inline table_column_numbers(i::Index) = eachindex(i.column_labels)
@inline table_column_numbers(r::TableRow) = table_column_numbers(r.index)

"""
Maps table column index (1-based) -> sheet column index (cellref based)
"""
@inline table_column_to_sheet_column_number(index::Index, table_column_number::Int) = index.column_map[table_column_number]
@inline table_columns_count(i::Index) = length(i.column_labels)
@inline table_columns_count(itr::TableRowIterator) = table_columns_count(itr.index)
@inline table_columns_count(r::TableRow) = table_columns_count(r.index)
@inline row_number(r::TableRow) = r.row
@inline get_column_labels(index::Index) = index.column_labels
@inline get_column_labels(itr::TableRowIterator) = get_column_labels(itr.index)
@inline get_column_labels(r::TableRow) = get_column_labels(r.index)
@inline get_column_label(r::TableRow, table_column_number::Int) = get_column_labels(r)[table_column_number]

# iterate into TableRow to get each column value
Base.start(r::TableRow) = start(table_column_numbers(r))
Base.done(r::TableRow, state) = done(table_column_numbers(r), state)

function Base.next(r::TableRow, state)
    (next_column_number, next_state) = next(table_column_numbers(r), state)
    return (r[next_column_number], next_state)
end

Base.getindex(r::TableRow, x) = getdata(r, x)

function TableRow(itr::TableRowIterator, sheet_row::SheetRow, row::Int)
    ws = get_worksheet(itr)
    index = itr.index
    cell_values = Vector{CellValueType}()

    for table_column_number in table_column_numbers(index)
        sheet_column = table_column_to_sheet_column_number(index, table_column_number)
        cell = getcell(sheet_row, sheet_column)
        push!(cell_values, getdata(ws, cell))
    end

    return TableRow(row, index, cell_values)
end

getdata(r::TableRow, table_column_number::Int) = r.cell_values[table_column_number]

function getdata(r::TableRow, column_label::Symbol)
    index = r.index
    if haskey(index.lookup, column_label)
        return getdata(r, index.lookup[column_label])
    else
        error("Invalid column label: $column_label.")
    end
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

    # user asked to stop
    if isa(itr.stop_in_row_function, Function) && itr.stop_in_row_function(TableRow(itr, state.sheet_row, state.table_row_index))
        return true
    end

    # empty table case
    if state.is_done
        return true
    elseif !itr.stop_in_empty_row
        return false
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
        if !Missings.ismissing(getdata(get_worksheet(itr), getcell(state.sheet_row, c)))
            return false
        end
    end

    return true
end

function infer_eltype(v::Vector{Any})
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

infer_eltype(v::Vector{T}) where T = T

function gettable(itr::TableRowIterator; infer_eltypes::Bool=false)
    column_labels = get_column_labels(itr)
    columns_count = table_columns_count(itr)
    data = Vector{Any}(columns_count)
    for c in 1:columns_count
        data[c] = Vector{Any}()
    end

    for r in itr # r is a TableRow
        is_empty_row = true

        for (ci, cv) in enumerate(r) # iterate a TableRow to get column data
            push!(data[ci], cv)
            if !Missings.ismissing(cv)
                is_empty_row = false
            end
        end

        # undo insert row in case of empty rows
        if is_empty_row
            for c in 1:columns_count
                pop!(data[c])
            end
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

"""
    gettable(sheet, [columns]; [first_row], [column_labels], [header], [infer_eltypes], [stop_in_empty_row], [stop_in_row_function]) -> data, column_labels

Returns tabular data from a spreadsheet as a tuple `(data, column_labels)`.
`data` is a vector of columns. `column_labels` is a vector of symbols.
Use this function to create a `DataFrame` from package `DataFrames.jl`.

Use `columns` argument to specify which columns to get.
For example, `columns="B:D"` will select columns `B`, `C` and `D`.
If `columns` is not given, the algorithm will find the first sequence
of consecutive non-empty cells.

Use `first_row` to indicate the first row from the table.
`first_row=5` will look for a table starting at sheet row `5`.
If `first_row` is not given, the algorithm will look for the first
non-empty row in the spreadsheet.

`header` is a `Bool` indicating if the first row is a header.
If `header=true` and `column_labels` is not specified, the column labels
for the table will be read from the first row of the table.
If `header=false` and `column_labels` is not specified, the algorithm
will generate column labels. The default value is `header=true`.

Use `column_labels` as a vector of symbols to specify names for the header of the table.

Use `infer_eltypes=true` to get `data` as a `Vector{Any}` of typed vectors.
The default value is `infer_eltypes=false`.

`stop_in_empty_row` is a boolean indicating wether an empty row marks the end of the table.
If `stop_in_empty_row=false`, the `TableRowIterator` will continue to fetch rows until there's no more rows in the Worksheet.
The default behavior is `stop_in_empty_row=true`.

`stop_in_row_function` is a Function that receives a `TableRow` and returns a `Bool` indicating if the end of the table was reached.

Example for `stop_in_row_function`:

```
function stop_function(r)
    v = r[:col_label]
    return !Missings.ismissing(v) && v == "unwanted value"
end
```

Rows where all column values are equal to `Missing.missing` are dropped.

Example code for `gettable`:

```julia
julia> using DataFrames, XLSX

julia> xf = XLSX.openxlsx("myfile.xlsx")

julia> df = DataFrame(XLSX.gettable(xf["mysheet"])...)

julia> close(xf)

See also: `readtable`.
```
"""
function gettable(sheet::Worksheet, cols::Union{ColumnRange, AbstractString}; first_row::Int=_find_first_row_with_data(sheet, convert(ColumnRange, cols).start), column_labels::Vector{Symbol}=Vector{Symbol}(), header::Bool=true, infer_eltypes::Bool=false, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Function, Void}=nothing)
    itr = eachtablerow(sheet, cols; first_row=first_row, column_labels=column_labels, header=header, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function)
    return gettable(itr; infer_eltypes=infer_eltypes)
end

function gettable(sheet::Worksheet; first_row::Int = 1, column_labels::Vector{Symbol}=Vector{Symbol}(), header::Bool=true, infer_eltypes::Bool=false, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Function, Void}=nothing)
    itr = eachtablerow(sheet; first_row=first_row, column_labels=column_labels, header=header, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function)
    return gettable(itr; infer_eltypes=infer_eltypes)
end
