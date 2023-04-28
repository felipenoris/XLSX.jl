
#
# Table
#

# Returns a tuple with the first and last index of the columns for a `SheetRow`.
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

_colname_prefix_symbol(sheet::Worksheet, cell::Cell) = Symbol(getdata(sheet, cell))
_colname_prefix_symbol(sheet::Worksheet, ::EmptyCell) = Symbol("#Empty")

# helper function to manage problematic column labels
# Empty cell -> "#Empty"
# No_unique_label -> No_unique_label_2
function push_unique!(vect::Vector{Symbol}, sheet::Worksheet, cell::AbstractCell, iter::Int=1)
    name = _colname_prefix_symbol(sheet, cell)

    if iter > 1
        name = Symbol(name, '_', iter)
    end

    if name in vect
        push_unique!(vect, sheet, cell, iter + 1)
    else
        push!(vect, name)
    end

    nothing
end

"""
    eachtablerow(sheet, [columns]; [first_row], [column_labels], [header], [stop_in_empty_row], [stop_in_row_function], [keep_empty_rows])

Constructs an iterator of table rows. Each element of the iterator is of type `TableRow`.

`header` is a boolean indicating whether the first row of the table is a table header.

If `header == false` and no `column_labels` were supplied, column names will be generated following the column names found in the Excel file.

The `columns` argument is a column range, as in `"B:E"`.
If `columns` is not supplied, the column range will be inferred by the non-empty contiguous cells in the first row of the table.

The user can replace column names by assigning the optional `column_labels` input variable with a `Vector{Symbol}`.

`stop_in_empty_row` is a boolean indicating whether an empty row marks the end of the table.
If `stop_in_empty_row=false`, the iterator will continue to fetch rows until there's no more rows in the Worksheet.
The default behavior is `stop_in_empty_row=true`. Empty rows may be returned by the iterator when `stop_in_empty_row=false`.

`stop_in_row_function` is a Function that receives a `TableRow` and returns a `Bool` indicating if the end of the table was reached.

Example for `stop_in_row_function`:

```
function stop_function(r)
    v = r[:col_label]
    return !ismissing(v) && v == "unwanted value"
end
```

`keep_empty_rows` determines whether rows where all column values are equal to `missing` are kept (`true`) or skipped (`false`) by the row iterator.
`keep_empty_rows` never affects the *bounds* of the iterator; the number of rows read from a sheet is only affected by `first_row`, `stop_in_empty_row` and `stop_in_row_function` (if specified).
`keep_empty_rows` is only checked once the first and last row of the table have been determined, to see whether to keep or drop empty rows between the first and the last row.

Example code:
```
for r in XLSX.eachtablerow(sheet)
    # r is a `TableRow`. Values are read using column labels or numbers.
    rn = XLSX.row_number(r) # `TableRow` row number.
    v1 = r[1] # will read value at table column 1.
    v2 = r[:COL_LABEL2] # will read value at column labeled `:COL_LABEL2`.
end
```

See also [`XLSX.gettable`](@ref).
"""
function eachtablerow(
            sheet::Worksheet,
            cols::Union{ColumnRange, AbstractString};
            first_row::Union{Nothing, Int}=nothing,
            column_labels=nothing,
            header::Bool=true,
            stop_in_empty_row::Bool=true,
            stop_in_row_function::Union{Nothing, Function}=nothing,
            keep_empty_rows::Bool=false,
        ) :: TableRowIterator

    if first_row == nothing
        first_row = _find_first_row_with_data(sheet, convert(ColumnRange, cols).start)
    end

    itr = eachrow(sheet)
    column_range = convert(ColumnRange, cols)

    if column_labels == nothing
        column_labels = Vector{Symbol}()
        if header
            # will use getdata to get column names
            for column_index in column_range.start:column_range.stop
                sheet_row = find_row(itr, first_row)
                cell = getcell(sheet_row, column_index)
                push_unique!(column_labels, sheet, cell)
            end
        else
            # generate column_labels if there's no header information anywhere
            for c in column_range
                push!(column_labels, Symbol(c))
            end
        end
    else
        # check consistency for column_range and column_labels
        @assert length(column_labels) == length(column_range) "`column_range` (length=$(length(column_range))) and `column_labels` (length=$(length(column_labels))) must have the same length."
    end

    first_data_row = header ? first_row + 1 : first_row
    return TableRowIterator(sheet, Index(column_range, column_labels), first_data_row, stop_in_empty_row, stop_in_row_function, keep_empty_rows)
end

function TableRowIterator(sheet::Worksheet, index::Index, first_data_row::Int, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Nothing, Function}=nothing, keep_empty_rows::Bool=false)
    return TableRowIterator(eachrow(sheet), index, first_data_row, stop_in_empty_row, stop_in_row_function, keep_empty_rows)
end

function eachtablerow(
            sheet::Worksheet;
            first_row::Union{Nothing, Int}=nothing,
            column_labels=nothing,
            header::Bool=true,
            stop_in_empty_row::Bool=true,
            stop_in_row_function::Union{Nothing, Function}=nothing,
            keep_empty_rows::Bool=false,
        ) :: TableRowIterator

    if first_row == nothing
        # if no columns were given,
        # first_row must be provided and cannot be inferred.
        # If it was not provided, will use first row as default value
        first_row = 1
    end

    for r in eachrow(sheet)

        # skip rows until we reach first_row, and if !keep_empty_rows then skip empty rows
        if row_number(r) < first_row || isempty(r) && !keep_empty_rows
            continue
        end

        columns_ordered = sort(collect(keys(r.rowcells)))

        for (ci, cn) in enumerate(columns_ordered)
            if !ismissing(getdata(r, cn))
                # found a row with data. Will get ColumnRange from non-empty consecutive cells
                first_row = row_number(r)
                column_start = cn
                column_stop = cn

                if length(columns_ordered) == 1
                    # there's only one column
                    column_range = ColumnRange(column_start, column_stop)
                    return eachtablerow(sheet, column_range; first_row=first_row, column_labels=column_labels, header=header, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function, keep_empty_rows=keep_empty_rows)
                else
                    # will figure out the column range
                    for ci_stop in (ci+1):length(columns_ordered)
                        cn_stop = columns_ordered[ci_stop]

                        # Will stop if finds an empty cell or a skipped column
                        if ismissing(getdata(r, cn_stop)) || (cn_stop - 1 != column_stop)
                            column_range = ColumnRange(column_start, column_stop)
                            return eachtablerow(sheet, column_range; first_row=first_row, column_labels=column_labels, header=header, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function, keep_empty_rows=keep_empty_rows)
                        end
                        column_stop = cn_stop
                    end
                end

                # if got here, it's because all columns are non-empty
                column_range = ColumnRange(column_start, column_stop)
                return eachtablerow(sheet, column_range; first_row=first_row, column_labels=column_labels, header=header, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function, keep_empty_rows=keep_empty_rows)
            end
        end
    end

    error("Couldn't find a table in sheet $(sheet.name)")
end

function _find_first_row_with_data(sheet::Worksheet, column_number::Int)
    # will find first_row
    for r in eachrow(sheet)
        if !ismissing(getdata(r, column_number))
            return row_number(r)
        end
    end
    error("Column $(encode_column_number(column_number)) has no data.")
end

@inline get_worksheet(tri::TableRowIterator) = get_worksheet(tri.itr)

# Returns real sheet column numbers (based on cellref)
@inline sheet_column_numbers(i::Index) = values(i.column_map)

# Returns an iterator for table column numbers.
@inline table_column_numbers(i::Index) = eachindex(i.column_labels)
@inline table_column_numbers(r::TableRow) = table_column_numbers(r.index)

# Maps table column index (1-based) -> sheet column index (cellref based)
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

function Base.iterate(r::TableRow)
    next = iterate(table_column_numbers(r))
    if next == nothing
        return nothing
    else
        next_column_number, next_state = next
        return r[next_column_number], next_state
    end
end

function Base.iterate(r::TableRow, state)
    next = iterate(table_column_numbers(r), state)
    if next == nothing
        return nothing
    else
        next_column_number, next_state = next
        return r[next_column_number], next_state
    end
end

Base.getindex(r::TableRow, x) = getdata(r, x)

function TableRow(table_row::Int, index::Index, sheet_row::SheetRow)
    ws = get_worksheet(sheet_row)

    cell_values = Vector{CellValueType}()
    for table_column_number in table_column_numbers(index)
        sheet_column = table_column_to_sheet_column_number(index, table_column_number)
        cell = getcell(sheet_row, sheet_column)
        push!(cell_values, getdata(ws, cell))
    end

    return TableRow(table_row, index, cell_values)
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

Base.IteratorSize(::Type{<:TableRowIterator}) = Base.SizeUnknown()
Base.eltype(::TableRowIterator) = TableRow

function Base.iterate(itr::TableRowIterator)
    next = iterate(itr.itr)

    # go to the first_data_row
    while next != nothing
        (sheet_row, sheet_row_iterator_state) = next

        if row_number(sheet_row) == itr.first_data_row
            table_row_index = 1
            return TableRow(table_row_index, itr.index, sheet_row), TableRowIteratorState(table_row_index, row_number(sheet_row), sheet_row_iterator_state)
        else
            next = iterate(itr.itr, sheet_row_iterator_state)
        end
    end

    # no rows for this table
    return nothing
end

function Base.iterate(itr::TableRowIterator, state::TableRowIteratorState)
    table_row_index = state.table_row_index + 1
    next = iterate(itr.itr, state.sheet_row_iterator_state) # iterate the SheetRowIterator

    if next == nothing
        return nothing
    end

    sheet_row, sheet_row_iterator_state = next

    #
    # checks if we're done reading this table
    #

    # check skipping rows
    # The XML can skip rows if there's no data in it,
    # so this is why is_empty_table_row function below wouldn't catch this case
    if itr.stop_in_empty_row && row_number(sheet_row) != itr.first_data_row && row_number(sheet_row) != (state.sheet_row_index + 1)
        return nothing
    end

    # checks if there are any data inside column range
    function is_empty_table_row(sheet_row::SheetRow) :: Bool

        if isempty(sheet_row)
            return true
        end

        for c in sheet_column_numbers(itr.index)
            if !ismissing(getdata(get_worksheet(itr), getcell(sheet_row, c)))
                return false
            end
        end
        return true
    end

    if is_empty_table_row(sheet_row)
        if itr.stop_in_empty_row
            # user asked to stop fetching table rows if we find an empty row
            return nothing
        elseif !itr.keep_empty_rows
            # keep looking for a non-empty row
            next = iterate(itr.itr, sheet_row_iterator_state)
            while next != nothing
                sheet_row, sheet_row_iterator_state = next
                if !is_empty_table_row(sheet_row)
                    break
                end
                next = iterate(itr.itr, sheet_row_iterator_state)
            end

            if next == nothing
                # end of file
                return nothing
            end
        end
    end

    # if the `is_empty_table_row` check above was successful, we can't get empty sheet_row here
    @assert !is_empty_table_row(sheet_row) || itr.keep_empty_rows
    table_row = TableRow(table_row_index, itr.index, sheet_row)

    # user asked to stop
    if itr.stop_in_row_function != nothing && itr.stop_in_row_function(table_row)
        return nothing
    end

    # we got a row to return
    return table_row, TableRowIteratorState(table_row_index, row_number(sheet_row), sheet_row_iterator_state)
end

function infer_eltype(v::Vector{Any})
    local hasmissing::Bool = false
    local t::DataType = Any

    if isempty(v)
        return eltype(v)
    end

    for i in 1:length(v)
        if ismissing(v[i])
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
            return Union{Missing, t}
        else
            return t
        end
    end
end

infer_eltype(v::Vector{T}) where T = T

function check_table_data_dimension(data::Vector)

    # nothing to check
    isempty(data) && return

    # all columns should be vectors
    for (colindex, colvec) in enumerate(data)
        @assert isa(colvec, Vector) "Data type at index $colindex is not a vector. Found: $(typeof(colvec))."
    end

    # no need to check row count
    length(data) == 1 && return

    # check all columns have the same row count
    col_count = length(data)
    row_count = length(data[1])
    for colindex in 2:col_count
        @assert length(data[colindex]) == row_count "Not all columns have the same number of rows. Check column $colindex."
    end

    nothing
end

function gettable(itr::TableRowIterator; infer_eltypes::Bool=false) :: DataTable
    column_labels = get_column_labels(itr)
    columns_count = table_columns_count(itr)
    data = Vector{Any}(undef, columns_count)
    for c in 1:columns_count
        data[c] = Vector{Any}()
    end

    for r in itr # r is a TableRow
        is_empty_row = true

        for (ci, cv) in enumerate(r) # iterate a TableRow to get column data
            push!(data[ci], cv)
            if !ismissing(cv)
                is_empty_row = false
            end
        end

        # undo insert row in case of empty rows
        if is_empty_row && !itr.keep_empty_rows
            for c in 1:columns_count
                pop!(data[c])
            end
        end
    end

    if infer_eltypes
        rows = length(data[1])
        for c in 1:columns_count
            new_column_data = Vector{infer_eltype(data[c])}(undef, rows)
            for r in 1:rows
                new_column_data[r] = data[c][r]
            end
            data[c] = new_column_data
        end
    end

    check_table_data_dimension(data)

    return DataTable(data, column_labels)
end

"""
    gettable(
        sheet,
        [columns];
        [first_row],
        [column_labels],
        [header],
        [infer_eltypes],
        [stop_in_empty_row],
        [stop_in_row_function],
        [keep_empty_rows]
    ) -> DataTable

Returns tabular data from a spreadsheet as a struct `XLSX.DataTable`.
Use this function to create a `DataFrame` from package `DataFrames.jl`.

Use `columns` argument to specify which columns to get.
For example, `"B:D"` will select columns `B`, `C` and `D`.
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

`stop_in_empty_row` is a boolean indicating whether an empty row marks the end of the table.
If `stop_in_empty_row=false`, the `TableRowIterator` will continue to fetch rows until there's no more rows in the Worksheet.
The default behavior is `stop_in_empty_row=true`.

`stop_in_row_function` is a Function that receives a `TableRow` and returns a `Bool` indicating if the end of the table was reached.

# Example for `stop_in_row_function`

```julia
function stop_function(r)
    v = r[:col_label]
    return !ismissing(v) && v == "unwanted value"
end
```

`keep_empty_rows` determines whether rows where all column values are equal to `missing` are kept (`true`) or dropped (`false`) from the resulting table.
`keep_empty_rows` never affects the *bounds* of the table; the number of rows read from a sheet is only affected by `first_row`, `stop_in_empty_row` and `stop_in_row_function` (if specified).
`keep_empty_rows` is only checked once the first and last row of the table have been determined, to see whether to keep or drop empty rows between the first and the last row.

# Example

```julia
julia> using DataFrames, XLSX

julia> df = XLSX.openxlsx("myfile.xlsx") do xf
        DataFrame(XLSX.gettable(xf["mysheet"]))
    end
```

See also: [`XLSX.readtable`](@ref).
"""
function gettable(sheet::Worksheet, cols::Union{ColumnRange, AbstractString}; first_row::Union{Nothing, Int}=nothing, column_labels=nothing, header::Bool=true, infer_eltypes::Bool=false, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Function, Nothing}=nothing, keep_empty_rows::Bool=false)
    itr = eachtablerow(sheet, cols; first_row=first_row, column_labels=column_labels, header=header, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function, keep_empty_rows=keep_empty_rows)
    return gettable(itr; infer_eltypes=infer_eltypes)
end

function gettable(sheet::Worksheet; first_row::Union{Nothing, Int}=nothing, column_labels=nothing, header::Bool=true, infer_eltypes::Bool=false, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Function, Nothing}=nothing, keep_empty_rows::Bool=false)
    itr = eachtablerow(sheet; first_row=first_row, column_labels=column_labels, header=header, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function, keep_empty_rows=keep_empty_rows)
    return gettable(itr; infer_eltypes=infer_eltypes)
end
