
function Worksheet(xf::XLSXFile, sheet_element::EzXML.Node)
    @assert EzXML.nodename(sheet_element) == "sheet"

    sheetId = parse(Int, sheet_element["sheetId"])
    relationship_id = sheet_element["r:id"]
    name = sheet_element["name"]

    target = "xl/" * get_relationship_target_by_id(xf.workbook, relationship_id)
    sheet_data = xf.data[target]

    return Worksheet(xf, sheetId, relationship_id, name, sheet_data)
end

isdate1904(ws::Worksheet) = isdate1904(ws.package)

"""
Retuns the dimension of this worksheet as a CellRange.
"""
function dimension(ws::Worksheet) :: CellRange
    xroot = EzXML.root(ws.data)
    @assert EzXML.nodename(xroot) == "worksheet" "Unicorn!"

    for dimension_element in EzXML.eachelement(xroot)
        if EzXML.nodename(dimension_element) == "dimension"
            ref_str = dimension_element["ref"]
            if is_valid_cellname(ref_str)
                return CellRange("$(ref_str):$(ref_str)")
            else
                return CellRange(ref_str)
            end
        end
    end

    error("Malformed Worksheet $(ws.name): dimension not found.")
end

"""
    getdata(sheet, ref)
    getdata(filepath, sheet, ref)

Returns a escalar or a matrix with values from a spreadsheet.
`ref` can be a cell reference or a range.

Example:

```julia
julia> v = XLSX.getdata("myfile.xlsx", "mysheet", "A1:B4")
```

Indexing in a `Worksheet` will dispatch to `getdata` method.
So the following example will have the same effect as the first example.

```julia
julia> f = XLSX.read("myfile.xlsx")

julia> sheet = f["mysheet"]

julia> v = sheet["A1:B4"]
```
"""
getdata(ws::Worksheet, single::CellRef) = celldata(ws, getcell(ws, single))

function getdata(ws::Worksheet, rng::CellRange) :: Array{Any,2}
    result = Array{Any, 2}(size(rng))
    fill!(result, Missings.missing)

    top = row_number(rng.start)
    bottom = row_number(rng.stop)
    left = column_number(rng.start)
    right = column_number(rng.stop)

    for sheetrow in eachrow(ws)
        if top <= sheetrow.row && sheetrow.row <= bottom
            for column in left:right
                cell = getcell(sheetrow, column)
                if !isempty(cell)
                    (r, c) = relative_cell_position(cell.ref, rng)
                    result[r, c] = celldata(ws, cell)
                end
            end
        end
    end

    return result
end

function getdata(ws::Worksheet, ref::AbstractString) :: Union{Array{Any,2}, Any}
    if is_valid_cellname(ref)
        return getdata(ws, CellRef(ref))
    elseif is_valid_cellrange(ref)
        return getdata(ws, CellRange(ref))
    else
        error("$ref is not a valid cell or range reference.")
    end
end

getdata(ws::Worksheet) = getdata(ws, dimension(ws))

Base.getindex(ws::Worksheet, r) = getdata(ws, r)
Base.getindex(ws::Worksheet, ::Colon) = getdata(ws)

Base.show(io::IO, ws::Worksheet) = print(io, "XLSX.Worksheet: \"$(ws.name)\". Dimension: $(dimension(ws)).")

"""
    getcell(sheet, ref)
    getcell(filepath, sheet, ref)

Returns an `AbstractCell` that represents a cell in the spreadsheet.

Example:

```julia
julia> sheet = XLSX.getsheet("myfile.xlsx", "mysheet")

julia> cell = XLSX.getcell(sheet, "A1")
```
"""
function getcell(ws::Worksheet, single::CellRef) :: AbstractCell

    for sheetrow in eachrow(ws)
        if row_number(sheetrow) == row_number(single)
            return getcell(sheetrow, column_number(single))
        end
    end

    return EmptyCell(single)
end

function getcell(ws::Worksheet, ref::AbstractString)
    if is_valid_cellname(ref)
        return getcell(ws, CellRef(ref))
    else
        error("$ref is not a valid cell reference.")
    end
end

"""
    getcellrange(sheet, rng)
    getcellrange(filepath, sheet, rng)

Returns a matrix with cells as `Array{AbstractCell, 2}`.
`rng` must be a valid cell range, as in `"A1:B2"`.

Example:

```julia
julia> XLSX.getcellrange("myfile.xlsx", "mysheet", "A1:B4")
```
"""
function getcellrange(ws::Worksheet, rng::CellRange) :: Array{AbstractCell,2}
    result = Array{AbstractCell, 2}(size(rng))
    for cellref in rng
        (r, c) = relative_cell_position(cellref, rng)
        result[r, c] = EmptyCell(cellref)
    end

    top = row_number(rng.start)
    bottom = row_number(rng.stop)
    left = column_number(rng.start)
    right = column_number(rng.stop)

    for sheetrow in eachrow(ws)
        if top <= sheetrow.row && sheetrow.row <= bottom
            for column in left:right
                cell = getcell(sheetrow, column)
                if !isempty(cell)
                    (r, c) = relative_cell_position(cell.ref, rng)
                    result[r, c] = cell
                end
            end
        end
    end

    return result
end

function getcellrange(ws::Worksheet, rng::AbstractString)
    if is_valid_cellrange(rng)
        return getcellrange(ws, CellRange(rng))
    else
        error("$rng is not a valid cell range.")
    end
end

"""
    gettable(sheet, [columns]; [first_row], [column_labels], [header], [infer_eltypes]) -> data, column_labels
    gettable(filepath, sheet, [columns]; [first_row], [column_labels], [header], [infer_eltypes]) -> data, column_labels

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

Example:

```julia
julia> using DataFrames, XLSX

julia> df = DataFrame(XLSX.gettable("myfile.xlsx", "mysheet")...)
```
"""
function gettable(sheet::Worksheet, cols::Union{ColumnRange, AbstractString}; first_row::Int=_find_first_row_with_data(sheet, convert(ColumnRange, cols).start), column_labels::Vector{Symbol}=Vector{Symbol}(), header::Bool=true, infer_eltypes::Bool=false)
    itr = TableRowIterator(sheet, cols; first_row=first_row, column_labels=column_labels, header=header)
    return gettable(itr; infer_eltypes=infer_eltypes)
end

function gettable(sheet::Worksheet; first_row::Int = 1, column_labels::Vector{Symbol}=Vector{Symbol}(), header::Bool=true, infer_eltypes::Bool=false)
    itr = TableRowIterator(sheet; first_row=first_row, column_labels=column_labels, header=header)
    return gettable(itr; infer_eltypes=infer_eltypes)
end

#
# Helper functions
#
getcell(filepath::AbstractString, sheet::Union{AbstractString, Int}, ref) = getcell( read(filepath)[sheet], ref )
getcell(filepath::AbstractString, sheetref::AbstractString) = getcell(read(filepath), sheetref)
getcellrange(filepath::AbstractString, sheet::Union{AbstractString, Int}, rng) = getcellrange( read(filepath)[sheet], rng )
getcellrange(filepath::AbstractString, sheetref::AbstractString) = getcellrange(read(filepath), sheetref)
getdata(filepath::AbstractString, sheet::Union{AbstractString, Int}, ref) = getdata( read(filepath)[sheet], ref )
getdata(filepath::AbstractString, sheetref::AbstractString) = getdata(read(filepath), sheetref)
gettable(filepath::AbstractString, sheet::Union{AbstractString, Int}; first_row::Int = 1, column_labels::Vector{Symbol}=Vector{Symbol}(), header::Bool=true, infer_eltypes::Bool=false) = gettable( read(filepath)[sheet]; first_row=first_row, column_labels=column_labels, header=header, infer_eltypes=infer_eltypes )
