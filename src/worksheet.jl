
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
getdata(ws::Worksheet) = ws[:]

function Cell(c::LightXML.XMLElement)
    # c (Cell) element is defined at section 18.3.1.4
    # t (Cell Data Type) is an enumeration representing the cell's data type. The possible values for this attribute are defined by the ST_CellType simple type (§18.18.11).
    # s (Style Index) is the index of this cell's style. Style records are stored in the Styles Part.

    @assert LightXML.name(c) == "c" "`cellvalue` function expects a `c` (cell) XMLElement."

    ref = CellRef(LightXML.attribute(c, "r"))

    # type
    if LightXML.has_attribute(c, "t")
        t = LightXML.attribute(c, "t")
    else
        t = ""
    end

    # style
    if LightXML.has_attribute(c, "s")
        s = LightXML.attribute(c, "s")
    else
        s = ""
    end

    vs = c["v"] # Vector{LightXML.XMLElement}
    @assert length(vs) <= 1 "Unsupported: cell $(ref) has $(length(vs)) `v` tags."

    if length(vs) == 0
        v = ""
    else
        v = LightXML.content(vs[1])
    end

    fs = c["f"]
    if length(fs) == 0
        f = ""
    else
        @assert length(fs) == 1 "Unsupported..."
        f = LightXML.content(fs[1])
    end

    return Cell(ref, t, s, v, f)
end

function cellvalue(ws::Worksheet, cell::Cell) :: Any

    if cell.datatype == "inlineStr"
        error("datatype inlineStr not supported...")
    end

    if cell.datatype == "s"
        # use sst
        return sst_unformatted_string(ws, cell.value)

    elseif (cell.datatype == "" || cell.datatype == "n")

        if cell.value == ""
            return Missings.missing
        end

        if cell.style != "" && styles_is_datetime(ws, cell.style)
            # datetime
            return _cellvalue_datetime(cell.value, isdate1904(ws))

        elseif cell.style != "" && styles_is_float(ws, cell.style)

            # float
            return parse(Float64, cell.value)

        else
            # fallback to unformatted number
            if contains(cell.value, ".")
                v_num = parse(Float64, cell.value)
            else
                v_num = parse(Int64, cell.value)
            end

            return v_num
        end
    elseif cell.datatype == "b"
        if cell.value == "0"
            return false
        elseif cell.value == "1"
            return true
        else
            error("Unknown boolean value: $(cell.value).")
        end
    end

    error("Couldn't parse cellvalue for $cell.")
end

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
