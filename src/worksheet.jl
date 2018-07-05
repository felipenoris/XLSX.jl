
function Worksheet(xf::XLSXFile, sheet_element::EzXML.Node)
    @assert EzXML.nodename(sheet_element) == "sheet"
    sheetId = parse(Int, sheet_element["sheetId"])
    relationship_id = sheet_element["r:id"]
    name = sheet_element["name"]
    dim = read_worksheet_dimension(xf, relationship_id, name)

    return Worksheet(xf, sheetId, relationship_id, name, dim)
end

function read_worksheet_dimension(xf::XLSXFile, relationship_id, name) :: CellRange
    wb = get_workbook(xf)
    target_file = "xl/" * get_relationship_target_by_id(wb, relationship_id)
    zip_io, reader = open_internal_file_stream(xf, target_file)

    local result::Nullable{CellRange} = Nullable{CellRange}()

    # read Worksheet dimension
    while !EzXML.done(reader)
        if EzXML.nodetype(reader) == EzXML.READER_ELEMENT && EzXML.nodename(reader) == "dimension"
            @assert EzXML.nodedepth(reader) == 1 "Malformed Worksheet \"$(ws.name)\": unexpected node depth for dimension node: $(EzXML.nodedepth(reader))."
            ref_str = reader["ref"]
            if is_valid_cellname(ref_str)
                result = CellRange("$(ref_str):$(ref_str)")
            else
                result = CellRange(ref_str)
            end

            break
        end
    end

    close(reader)
    close(zip_io)

    if isnull(result)
        error("Couldn't parse worksheet $name dimension.")
    else
        return get(result)
    end
end

@inline isdate1904(ws::Worksheet) = isdate1904(get_workbook(ws))

"""
Retuns the dimension of this worksheet as a CellRange.
"""
@inline get_dimension(ws::Worksheet) :: CellRange = ws.dimension

function set_dimension!(ws::Worksheet, rng::CellRange)
    ws.dimension = rng
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
julia> f = XLSX.openxlsx("myfile.xlsx")

julia> sheet = f["mysheet"]

julia> v = sheet["A1:B4"]

julia> close(f)
```
"""
getdata(ws::Worksheet, single::CellRef) = getdata(ws, getcell(ws, single))

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
                    (r, c) = relative_cell_position(cell, rng)
                    result[r, c] = getdata(ws, cell)
                end
            end
        end

        # don't need to read new rows
        if sheetrow.row > bottom
            break
        end
    end

    return result
end

function getdata(ws::Worksheet, rng::ColumnRange) :: Array{Any,2}
    columns_count = length(rng)
    columns = Vector{Vector{Any}}(columns_count)
    for i in 1:columns_count
        columns[i] = Vector{Any}()
    end

    left, right = column_bounds(rng)

    for sheetrow in eachrow(ws)
        for column in left:right
            cell = getcell(sheetrow, column)
            c = relative_column_position(cell, rng) # r will be ignored
            push!(columns[c], getdata(ws, cell))
        end
    end

    rows = length(columns[1])
    for i in 1:columns_count
        @assert length(columns[i]) == rows "Inconsistent state: Each column should have the same number of rows."
    end

    return hcat(columns...)
end

function getdata(ws::Worksheet, ref::AbstractString) :: Union{Array{Any,2}, Any}
    if is_valid_cellname(ref)
        return getdata(ws, CellRef(ref))
    elseif is_valid_cellrange(ref)
        return getdata(ws, CellRange(ref))
    elseif is_valid_column_range(ref)
        return getdata(ws, ColumnRange(ref))
    elseif is_worksheet_defined_name(ws, ref)
        v = get_defined_name_value(ws, ref)
        if is_defined_name_value_a_constant(v)
            return v
        elseif is_defined_name_value_a_reference(v)
            return getdata(ws, v)
        else
            error("Unexpected defined name value: $v.")
        end
    elseif is_workbook_defined_name(get_workbook(ws), ref)
        wb = get_workbook(ws)
        v = get_defined_name_value(wb, ref)
        if is_defined_name_value_a_constant(v)
            return v
        elseif is_defined_name_value_a_reference(v)
            return getdata(get_xlsxfile(ws), v)
        else
            error("Unexpected defined name value: $v.")
        end
    else
        error("$ref is not a valid cell or range reference.")
    end
end

getdata(ws::Worksheet, rng::SheetCellRange) = getdata(get_xlsxfile(ws), rng)
getdata(ws::Worksheet) = getdata(ws, get_dimension(ws))

Base.getindex(ws::Worksheet, r) = getdata(ws, r)
Base.getindex(ws::Worksheet, ::Colon) = getdata(ws)

function Base.show(io::IO, ws::Worksheet) 
    rg = get_dimension(ws)
    nrow, ncol = size(rg)
    @printf(io, "%d×%d %s: [\"%s\"](%s)", nrow, ncol, typeof(ws), ws.name, rg)
end
"""
    getcell(sheet, ref)
    getcell(filepath, sheet, ref)

Returns an `AbstractCell` that represents a cell in the spreadsheet.

Example:

```julia
julia> xf = XLSX.openxlsx("myfile.xlsx")

julia> sheet = xf["mysheet"]

julia> cell = XLSX.getcell(sheet, "A1")

julia> close(xf)
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
                    (r, c) = relative_cell_position(cell, rng)
                    result[r, c] = cell
                end
            end
        end

        # don't need to read new rows
        if sheetrow.row > bottom
            break
        end
    end

    return result
end

function getcellrange(ws::Worksheet, rng::ColumnRange) :: Array{AbstractCell,2}
    columns_count = length(rng)
    columns = Vector{Vector{AbstractCell}}(columns_count)
    for i in 1:columns_count
        columns[i] = Vector{AbstractCell}()
    end

    left, right = column_bounds(rng)

    for sheetrow in eachrow(ws)
        for column in left:right
            cell = getcell(sheetrow, column)
            c = relative_column_position(cell, rng) # r will be ignored
            push!(columns[c], cell)
        end
    end

    rows = length(columns[1])
    for i in 1:columns_count
        @assert length(columns[i]) == rows "Inconsistent state: Each column should have the same number of rows."
    end

    return hcat(columns...)
end

function getcellrange(ws::Worksheet, rng::AbstractString)
    if is_valid_cellrange(rng)
        return getcellrange(ws, CellRange(rng))
    elseif is_valid_column_range(rng)
        return getcellrange(ws, ColumnRange(rng))
    else
        error("$rng is not a valid cell range.")
    end
end
