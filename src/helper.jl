
#
# Helper Functions
#

function readdata(filepath::AbstractString, sheet::Union{AbstractString, Int}, ref)
    xf = openxlsx(filepath, enable_cache=false)
    c = getdata(getsheet(xf, sheet), ref )
    close(xf)
    return c
end

function readdata(filepath::AbstractString, sheetref::AbstractString)
    xf = openxlsx(filepath, enable_cache=false)
    c = getdata(xf, sheetref)
    close(xf)
    return c
end

"""
    readtable(filepath, sheet, [columns]; [first_row], [column_labels], [header], [infer_eltypes], [stop_in_empty_row], [stop_in_row_function]) -> data, column_labels

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

Example code for `readtable`:

```julia
julia> using DataFrames, XLSX

julia> df = DataFrame(XLSX.readtable("myfile.xlsx", "mysheet")...)

See also: `gettable`.
```
"""
function readtable(filepath::AbstractString, sheet::Union{AbstractString, Int}; first_row::Int = 1, column_labels::Vector{Symbol}=Vector{Symbol}(), header::Bool=true, infer_eltypes::Bool=false, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Function, Void}=nothing, enable_cache::Bool=false)
    xf = openxlsx(filepath, enable_cache=enable_cache)
    c = gettable(getsheet(xf, sheet); first_row=first_row, column_labels=column_labels, header=header, infer_eltypes=infer_eltypes, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function)
    close(xf)
    return c
end

"""
    writetable(filename, data, columnnames; [rewrite], [sheetname])

`data` is a vector of columns.
`columnames` is a vector of column labels.
`rewrite` is a `Bool` to control if `filename` should be rewrited if already exists.
`sheetname` is the name for the worksheet.

Example using `DataFrames.jl`:

```julia
import DataFrames, XLSX
df = DataFrames.DataFrame(integers=[1, 2, 3, 4], strings=["Hey", "You", "Out", "There"], floats=[10.2, 20.3, 30.4, 40.5])
XLSX.writetable("df.xlsx", DataFrames.columns(df), DataFrames.names(df))
```
"""
function writetable(filename::AbstractString, data, columnnames; rewrite::Bool=false, sheetname::AbstractString="", anchor_cell::Union{String, CellRef}=CellRef("A1"))

    if !rewrite
        @assert !isfile(filename) "$filename already exists."
    end

    xf = open_default_template()
    sheet = xf[1]

    if sheetname != ""
        rename!(sheet, sheetname)
    end

    if isa(anchor_cell, String)
        anchor_cell = CellRef(anchor_cell)
    end

    writetable!(sheet, data, columnnames; anchor_cell=anchor_cell)

    # write output file
    writexlsx(filename, xf, rewrite=rewrite)
    nothing
end

"""
    writetable(filename::AbstractString; rewrite::Bool=false, kw...)
    writetable(filename::AbstractString, tables::Vector{Tuple{String, Vector{Any}, Vector{String}}}; rewrite::Bool=false)

Write multiple tables.

`kw` is a variable keyword argument list. Each element should be in this format: `sheetname=( data, column_names )`,
where `data` is a vector of columns and `column_names` is a vector of column labels.

Example:

```julia
import DataFrames, XLSX

df1 = DataFrames.DataFrame(COL1=[10,20,30], COL2=["Fist", "Sec", "Third"])
df2 = DataFrames.DataFrame(AA=["aa", "bb"], AB=[10.1, 10.2])

XLSX.writetable("report.xlsx", REPORT_A=( DataFrames.columns(df1), DataFrames.names(df1) ), REPORT_B=( DataFrames.columns(df2), DataFrames.names(df2) ))
```
"""
function writetable(filename::AbstractString; rewrite::Bool=false, kw...)
    if !rewrite
        @assert !isfile(filename) "$filename already exists."
    end

    xf = open_default_template()

    is_first = true

    for (sheetname, (data, column_names)) in kw
        if is_first
            # first sheet already exists in template file
            sheet = xf[1]
            rename!(sheet, string(sheetname))
            writetable!(sheet, data, column_names)

            is_first = false
        else
            sheet = addsheet!(xf, string(sheetname))
            writetable!(sheet, data, column_names)
        end
    end

    # write output file
    writexlsx(filename, xf, rewrite=rewrite)
    nothing
end

function writetable(filename::AbstractString, tables::Vector{Tuple{String, Vector{Any}, Vector{T}}}; rewrite::Bool=false) where {T<:Union{String, Symbol}}
    if !rewrite
        @assert !isfile(filename) "$filename already exists."
    end

    xf = open_default_template()

    is_first = true

    for (sheetname, data, column_names) in tables
        if is_first
            # first sheet already exists in template file
            sheet = xf[1]
            rename!(sheet, string(sheetname))
            writetable!(sheet, data, column_names)

            is_first = false
        else
            sheet = addsheet!(xf, string(sheetname))
            writetable!(sheet, data, column_names)
        end
    end

    # write output file
    writexlsx(filename, xf, rewrite=rewrite)
end

"""
    emptyfile(sheetname::AbstractString="")

Returns an empty, writable `XLSXFile` with 1 worksheet.

`sheetname` is the name of the worksheet, defaults to `Sheet1`.
"""
function emptyfile(sheetname::AbstractString="")
    xf = open_default_template()

    if sheetname != ""
        rename!(xf[1], sheetname)
    end

    return xf
end

function Base.open(f::Function, filename::AbstractString, sheetname::AbstractString="")
    xf = open_default_template()

    if sheetname != ""
        rename!(xf[1], sheetname)
    end

    try
        f(xf)
    finally
        closefile(xf, filename)
    end
end

"""
    writerow(xf::XLSXFile, data, row, start_column::String, sheet_num::Int=1)

Write one row to `XLSXFile`.

`data` is a vector.
`row` is the row index.
`start_column` is the `String` representation of the first column index to write.
`sheet_num` is the `Integer` representation of the target worksheet.

Example:

```julia
using XLSX

data = 1, 2, 3]
xf = XLSX.emptyfile()

XLSX.writerow(xf, data, 1, "A", 1)
```
"""
function writerow(xf::XLSXFile, data, row, start_column::String, sheet_num::Int=1)
    @assert sheetcount(xf) >= sheet_num "Sheet number $sheet_num not found."

    sheet = xf[sheet_num]
    anchor_cell = "$start_column$row"

    writerow(sheet, data; anchor_cell=CellRef(anchor_cell))
end

"""
    writerow(xf::XLSXFile, data, row, start_column::Int, sheet_num::Int=1)

Write one row to `XLSXFile`.

`data` is a vector.
`row` is the row index.
`start_column` is the `Integer` representation of the first column index to write.
`sheet_num` is the `Integer` representation of the target worksheet.

Example:

```julia
using XLSX

data = [1, 2, 3]
xf = XLSX.emptyfile()

XLSX.writerow(xf, data, 1, 1, 1)
```
"""
function writerow(xf::XLSXFile, data, row, start_column::Int, sheet_num::Int=1)
    writerow(xf, data, row, encode_column_number(start_column), sheet_num)
end

"""
    writerow(xf::XLSXFile, data, row, start_column::String, sheetname::String)

Write one row to `XLSXFile`.

`data` is a vector.
`row` is the row index.
`start_column` is the `String` representation of the first column index to write.
`sheetname` is the name of the target worksheet.

Example:

```julia
using XLSX

data = [1, 2, 3]
xf = XLSX.emptyfile("New Sheet")

XLSX.writerow(xf, data, 1, "A", "New Sheet")
```
"""
function writerow(xf::XLSXFile, data, row, start_column::String, sheetname::String)
    @assert hassheet(xf, sheetname) "Sheet name \"$sheetname\" not found."

    sheet = xf[sheetname]
    anchor_cell = "$start_column$row"

    writerow(sheet, data; anchor_cell=CellRef(anchor_cell))
end

"""
    writerow(xf::XLSXFile, data, row, start_column::Int, sheetname::String)

Write one row to `XLSXFile`.

`data` is a vector.
`row` is the row index.
`start_column` is the `Integer` representation of the first column index to write.
`sheetname` is the name of the target worksheet.

Example:

```julia
using XLSX

data = [1, 2, 3]
xf = XLSX.emptyfile("New Sheet")

XLSX.writerow(xf, data, 1, 1, "New Sheet")
```
"""
function writerow(xf::XLSXFile, data, row, start_column::Int, sheetname::String,)
    writerow(xf, data, row, encode_column_number(start_column), sheetname)
end

"""
    writerow(sheet::Worksheet, data; anchor_cell::CellRef=CellRef("A1"))

Write one row to `Worksheet`.

`data` is a vector.
`anchor_cell` is the first cell of the row to write, defaults to "A1".

Example:

```julia
using XLSX

data = [1, 2, 3]
xf = XLSX.emptyfile("New Sheet")
sheet = xf["New Sheet"]

# write the second row starting at the second column
XLSX.writerow(sheet, data, anchor_cell=XLSX(CellRef("B2")))
```
"""
function writerow(sheet::Worksheet, data; anchor_cell::CellRef=CellRef("A1"))
    @assert is_writable(get_xlsxfile(sheet)) "XLSXFile instance is not writable."
    writerow!(sheet, data; anchor_cell=CellRef(anchor_cell))

end

"""
    closefile(xf::XLSXFile, filepath::AbstractString; rewrite::Bool=false)

Writes content and closes `XLSXFile`.
"""
function closefile(xf::XLSXFile, filepath::AbstractString; rewrite::Bool=false)
    writexlsx(filepath, xf, rewrite=rewrite)
    nothing
end
