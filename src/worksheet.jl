
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

function gettable(sheet::Worksheet, cols::Union{ColumnRange, AbstractString}; first_row::Int=_find_first_row_with_data(sheet, convert(ColumnRange, cols).start), column_labels::Vector{Symbol}=Vector{Symbol}(), header::Bool=true, infer_eltypes::Bool=false)
    itr = TableRowIterator(sheet, cols; first_row=first_row, column_labels=column_labels, header=header)
    return gettable(itr; infer_eltypes=infer_eltypes)
end

function gettable(sheet::Worksheet; first_row::Int = 1, column_labels::Vector{Symbol}=Vector{Symbol}(), header::Bool=true, infer_eltypes::Bool=false)
    itr = TableRowIterator(sheet; first_row=first_row, column_labels=column_labels, header=header)
    return gettable(itr; infer_eltypes=infer_eltypes)
end
