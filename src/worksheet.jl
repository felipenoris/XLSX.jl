
isdate1904(ws::Worksheet) = isdate1904(ws.package)

"""
Retuns the dimension of this worksheet as a CellRange.
"""
function dimension(ws::Worksheet) :: CellRange
    xroot = LightXML.root(ws.data)
    @assert LightXML.name(xroot) == "worksheet" "Unicorn!"

    vec_dimension = xroot["dimension"]
    @assert length(vec_dimension) == 1 "Malformed Worksheet $(ws.name): only one `dimension` tag is allowed in worksheet data file."

    dimension_element = vec_dimension[1]
    ref_str = LightXML.attribute(dimension_element, "ref")

    if is_valid_cellname(ref_str)
        return CellRange("$(ref_str):$(ref_str)")
    else
        return CellRange(ref_str)
    end
end

function Base.getindex(ws::Worksheet, single::CellRef) :: Any
    xroot = LightXML.root(ws.data)
    @assert LightXML.name(xroot) == "worksheet" "Unicorn!"
    vec_sheetdata = xroot["sheetData"]
    @assert length(vec_sheetdata) <= 1 "Malformed sheet $(ws.name)."
    if length(vec_sheetdata) == 0
        return Missings.missing
    end

    rows = vec_sheetdata[1]["row"] # rows is a Vector{LightXML.XMLElement}

    for r in rows
        current_row_index = parse(Int, LightXML.attribute(r, "r"))

        if current_row_index != row_number(single)
            continue
        end

        # iterate over row -> c elements
        for c in r["c"]

            ref = CellRef(LightXML.attribute(c, "r"))
            if column_number(ref) != column_number(single)
                continue
            else
                cell = Cell(c)
                @assert row_number(cell.ref) == current_row_index "Malformed Excel file."
                return cellvalue(ws, cell)
            end
        end
    end

    return Missings.missing
end

function Base.getindex(ws::Worksheet, rng::CellRange) :: Array{Any,2}
    result = Array{Any, 2}(size(rng))
    fill!(result, Missings.missing)

    top = row_number(rng.start)
    bottom = row_number(rng.stop)
    left = column_number(rng.start)
    right = column_number(rng.stop)

    xroot = LightXML.root(ws.data)
    @assert LightXML.name(xroot) == "worksheet" "Unicorn!"
    vec_sheetdata = xroot["sheetData"]
    @assert length(vec_sheetdata) <= 1 "Malformed sheet $(ws.name)."
    if length(vec_sheetdata) == 0
        return result
    end

    rows = vec_sheetdata[1]["row"] # rows is a Vector{LightXML.XMLElement}

    for r in rows
        current_row_index = parse(Int, LightXML.attribute(r, "r"))

        if current_row_index < top || bottom < current_row_index
            continue
        end

        # iterate over row -> c elements
        for c in r["c"]

            ref = CellRef(LightXML.attribute(c, "r"))
            if column_number(ref) < left || right < column_number(ref)
                continue
            else
                cell = Cell(c)
                @assert row_number(cell.ref) == current_row_index "Malformed Excel file."
                (r, c) = relative_cell_position(cell.ref, rng)
                result[r, c] = cellvalue(ws, cell)
            end
        end
    end

    return result
end

function Base.getindex(ws::Worksheet, ref::AbstractString) :: Union{Array{Any,2}, Any}
    if is_valid_cellname(ref)
        return getindex(ws, CellRef(ref))
    elseif is_valid_cellrange(ref)
        return getindex(ws, CellRange(ref))
    else
        error("$ref is not a valid cell or range reference.")
    end
end

Base.getindex(ws::Worksheet, ::Colon) = getindex(ws, dimension(ws))

getdata(ws::Worksheet, r) = getindex(ws, r)
getdata(ws::Worksheet) = getindex(ws, dimension(ws))

Base.show(io::IO, ws::Worksheet) = println(io, "XLSX.Worksheet: \"$(ws.name)\". Dimension: $(dimension(ws)).")

function getcell(ws::Worksheet, single::CellRef) :: Cell
    xroot = LightXML.root(ws.data)
    @assert LightXML.name(xroot) == "worksheet" "Unicorn!"
    vec_sheetdata = xroot["sheetData"]
    @assert length(vec_sheetdata) <= 1 "Malformed sheet $(ws.name)."
    if length(vec_sheetdata) == 0
        return Missings.missing
    end

    rows = vec_sheetdata[1]["row"] # rows is a Vector{LightXML.XMLElement}

    for r in rows
        current_row_index = parse(Int, LightXML.attribute(r, "r"))

        if current_row_index != row_number(single)
            continue
        end

        # iterate over row -> c elements
        for c in r["c"]
            ref = CellRef(LightXML.attribute(c, "r"))

            if column_number(ref) == column_number(single)
                return Cell(c)
            end
        end
    end

    error("Cell $ref not found in worksheet $(ws.name).")
end

function getcell(ws::Worksheet, ref::AbstractString)
    if is_valid_cellname(ref)
        return getcell(ws, CellRef(ref))
    else
        error("$ref is not a valid cell or range reference.")
    end
end

function getcellrange(ws::Worksheet, rng::CellRange) :: Array{Union{Cell, Missings.Missing},2}
    result = Array{Union{Cell, Missings.Missing},2}(size(rng))
    fill!(result, Missings.missing)

    top = row_number(rng.start)
    bottom = row_number(rng.stop)
    left = column_number(rng.start)
    right = column_number(rng.stop)

    xroot = LightXML.root(ws.data)
    @assert LightXML.name(xroot) == "worksheet" "Unicorn!"
    vec_sheetdata = xroot["sheetData"]
    @assert length(vec_sheetdata) <= 1 "Malformed sheet $(ws.name)."
    if length(vec_sheetdata) == 0
        return result
    end

    rows = vec_sheetdata[1]["row"] # rows is a Vector{LightXML.XMLElement}

    for r in rows
        current_row_index = parse(Int, LightXML.attribute(r, "r"))

        if current_row_index < top || bottom < current_row_index
            continue
        end

        # iterate over row -> c elements
        for c in r["c"]

            ref = CellRef(LightXML.attribute(c, "r"))
            if column_number(ref) < left || right < column_number(ref)
                continue
            else
                cell = Cell(c)
                @assert row_number(cell.ref) == current_row_index "Malformed Excel file."
                (r, c) = relative_cell_position(cell.ref, rng)
                result[r, c] = cell
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
