
function Worksheet(xf::XLSXFile, sheet_element::EzXML.Node)
    @assert EzXML.nodename(sheet_element) == "sheet"
    sheetId = parse(Int, sheet_element["sheetId"])
    relationship_id = sheet_element["r:id"]
    name = sheet_element["name"]
    is_hidden = haskey(sheet_element, "state") && sheet_element["state"] in ["hidden", "veryHidden"]
    dim = read_worksheet_dimension(xf, relationship_id, name)

    return Worksheet(xf, sheetId, relationship_id, name, dim, is_hidden)
end

function Base.axes(ws::Worksheet, d)
    dim = get_dimension(ws)
    if dim === nothing
        throw(DimensionMismatch("Worksheet $ws has no dimension"))
    elseif d == 1
        return dim.start.row_number:dim.stop.row_number
    elseif d == 2
        return dim.start.column_number:dim.stop.column_number
    else
        throw(ArgumentError("Unsupported dimension $d"))
    end
end

# 18.3.1.35 - dimension (Worksheet Dimensions). This is optional, and not required.
function read_worksheet_dimension(xf::XLSXFile, relationship_id, name) :: Union{Nothing, CellRange}
    local result::Union{Nothing, CellRange} = nothing

    wb = get_workbook(xf)
    target_file = get_relationship_target_by_id("xl", wb, relationship_id)
    zip_io, reader = open_internal_file_stream(xf, target_file)

    try
        # read Worksheet dimension
        while EzXML.iterate(reader) != nothing
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
    finally
        close(reader)
        close(zip_io)
    end

    return result
end

@inline isdate1904(ws::Worksheet) = isdate1904(get_workbook(ws))

# Returns the dimension of this worksheet as a CellRange.
# Returns `nothing` if the dimension is unknown.
@inline get_dimension(ws::Worksheet) :: Union{Nothing, CellRange} = ws.dimension

function set_dimension!(ws::Worksheet, rng::CellRange)
    ws.dimension = rng
    nothing
end

"""
    getdata(sheet, ref)
    getdata(sheet, row, column)

Returns a scalar or a matrix with values from a spreadsheet.
`ref` can be a cell reference or a range.

Indexing in a `Worksheet` will dispatch to `getdata` method.

# Example

```julia
julia> f = XLSX.readxlsx("myfile.xlsx")

julia> sheet = f["mysheet"]

julia> matrix = sheet["A1:B4"]

julia> single_value = sheet[2, 2] # B2
```

See also [`XLSX.readdata`](@ref).
"""
getdata(ws::Worksheet, single::CellRef) = getdata(ws, getcell(ws, single))
getdata(ws::Worksheet, row::Integer, col::Integer) = getdata(ws, CellRef(row, col))
getdata(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = getdata(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))))
function getdata(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon)
    dim = get_dimension(ws)
    return if dim === nothing
        @warn "No worksheet dimension found"
        []
    else
        getdata(ws, CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number)))
    end
end
function getdata(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}})
    dim = get_dimension(ws)
    return if dim === nothing
        @warn "No worksheet dimension found"
        []
    else
        getdata(ws, CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col))))
    end
end

function getdata(ws::Worksheet, rng::CellRange) :: Array{Any,2}
    result = Array{Any, 2}(undef, size(rng))
    fill!(result, missing)

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
    columns = Vector{Vector{Any}}(undef, columns_count)
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

function getdata(ws::Worksheet)
    if ws.dimension != nothing
        return getdata(ws, get_dimension(ws))
    else
        error("Worksheet dimension is unknown.")
    end
end

Base.getindex(ws::Worksheet, r) = getdata(ws, r)
Base.getindex(ws::Worksheet, r, c) = getdata(ws, r, c)
Base.getindex(ws::Worksheet, ::Colon) = getdata(ws)

function Base.show(io::IO, ws::Worksheet)
    hidden_string = ws.is_hidden ? "(hidden)" : ""
    if get_dimension(ws) != nothing
        rg = get_dimension(ws)
        nrow, ncol = size(rg)
        @printf(io, "%d×%d %s: [\"%s\"](%s) %s", nrow, ncol, typeof(ws), ws.name, rg, hidden_string)
    else
        @printf(io, "%s: [\"%s\"] %s", typeof(ws), ws.name, hidden_string)
    end
end

"""
    getcell(sheet, ref)

Returns an `AbstractCell` that represents a cell in the spreadsheet.

Example:

```julia
julia> xf = XLSX.readxlsx("myfile.xlsx")

julia> sheet = xf["mysheet"]

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

getcell(ws::Worksheet, row::Integer, col::Integer) = getcell(ws, CellRef(row, col))

"""
    getcellrange(sheet, rng)

Returns a matrix with cells as `Array{AbstractCell, 2}`.
`rng` must be a valid cell range, as in `"A1:B2"`.
"""
function getcellrange(ws::Worksheet, rng::CellRange) :: Array{AbstractCell,2}
    result = Array{AbstractCell, 2}(undef, size(rng))
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
    columns = Vector{Vector{AbstractCell}}(undef, columns_count)
    for i in 1:columns_count
        columns[i] = Vector{AbstractCell}()
    end

    let
        left, right = column_bounds(rng)

        for sheetrow in eachrow(ws)
            for column in left:right
                cell = getcell(sheetrow, column)
                c = relative_column_position(cell, rng) # r will be ignored
                push!(columns[c], cell)
            end
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
