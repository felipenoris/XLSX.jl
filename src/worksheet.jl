
function Worksheet(xf::XLSXFile, sheet_element::XML.Node)
    XML.tag(sheet_element) != "sheet" && throw(XLSXError("Something wrong here!"))
    a = XML.attributes(sheet_element)
    sheetId = parse(Int, a["sheetId"])
    relationship_id = a["r:id"]
    name = a["name"]
    is_hidden = haskey(a, "state") && a["state"] in ["hidden", "veryHidden"]
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
function read_worksheet_dimension(xf::XLSXFile, relationship_id, name)::Union{Nothing,CellRange}
    local result::Union{Nothing,CellRange} = nothing

    wb = get_workbook(xf)
    target_file = get_relationship_target_by_id("xl", wb, relationship_id)
    zip_io, doc = open_internal_file_stream(xf, target_file)

    reader = iterate(doc)
    # Now let's look for a row element, if it exists
    while reader !== nothing # go next node
        (sheet_row, state) = reader
        if XML.nodetype(sheet_row) == XML.Element && XML.tag(sheet_row) == "dimension"
            XML.depth(sheet_row) != 2 && throw(XLSXError("Malformed Worksheet \"$name\": unexpected node depth for dimension node: $(XML.depth(sheet_row))."))
            ref_str = XML.attributes(sheet_row)["ref"]
            if is_valid_cellname(ref_str)
                result = CellRange("$(ref_str):$(ref_str)")
            else
                result = CellRange(ref_str)
            end

            break
        end
        reader = iterate(doc, state)
    end

    return result
end

@inline isdate1904(ws::Worksheet) = isdate1904(get_workbook(ws))

# Returns the dimension of this worksheet as a CellRange.
# Returns `nothing` if the dimension is unknown.
@inline get_dimension(ws::Worksheet)::Union{Nothing,CellRange} = ws.dimension

function set_dimension!(ws::Worksheet, rng::CellRange)
    ws.dimension = rng
    nothing
end

"""
    getdata(sheet, ref)
    getdata(sheet, row, column)

Returns a scalar, vector or a matrix with values from a spreadsheet.
`ref` can be a cell reference or a range or a valid defined name.

Indexing in a `Worksheet` will dispatch to `getdata` method.

# Example

```julia
julia> f = XLSX.readxlsx("myfile.xlsx")

julia> sheet = f["mysheet"] # Worksheet

julia> matrix = sheet["A1:B4"] # CellRange

julia> matrix = sheet["A:B"] # Column range

julia> matrix = sheet["1:4"] # Row range

julia> matrix = sheet["Contiguous"] # Named range

julia> matrix = sheet[1:30, 1] # use unit ranges to define rows and/or columns

julia> matrix = sheet[[1, 2, 3], 1] # vectors of integers to define rows and/or columns

julia> vector = sheet["A1:A4,C1:C4,G5"] # Non-contiguous range

julia> vector = sheet["Location"] # Non-contiguous named range

julia> single_value = sheet[2, 2] # Cell "B2"
```

See also [`XLSX.readdata`](@ref).
"""
getdata(ws::Worksheet, single::CellRef) = getdata(ws, getcell(ws, single))
getdata(ws::Worksheet, row::Integer, col::Integer) = getdata(ws, CellRef(row, col))
getdata(ws::Worksheet, row::Int, col::Vector{Int}) = [getdata(ws, a, b) for a in [row], b in col]
getdata(ws::Worksheet, row::Vector{Int}, col::Int) = [getdata(ws, a, b) for a in row, b in [col]]
getdata(ws::Worksheet, row::Vector{Int}, col::Vector{Int}) = [getdata(ws, a, b) for a in row, b in col]
getdata(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = getdata(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))))
function getdata(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        getdata(ws, CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number)))
    end
end
function getdata(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}})
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        getdata(ws, CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col))))
    end
end

function getdata(ws::Worksheet, rng::CellRange)::Array{Any,2}
    result = Array{Any,2}(undef, size(rng))
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

function getdata(ws::Worksheet, rng::ColumnRange)::Array{Any,2}
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        start = CellRef(dim.start.row_number, rng.start)
        stop = CellRef(dim.stop.row_number, rng.stop)
        return getdata(ws, CellRange(start, stop))
    end
end
function getdata(ws::Worksheet, rng::RowRange)::Array{Any,2}
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        start = CellRef(rng.start, dim.start.column_number,)
        stop = CellRef(rng.stop, dim.stop.column_number)
        return getdata(ws, CellRange(start, stop))
    end
end

#=
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
        length(columns[i]) != rows && throw(XLSXError("Inconsistent state: Each column should have the same number of rows."))
    end

    return hcat(columns...)
end

function getdata(ws::Worksheet, rng::RowRange) :: Array{Any,2}
    rows_count = length(rng)
    dim = get_dimension(ws)

    rows = Vector{Vector{Any}}(undef, rows_count)
    for i in 1:rows_count
        rows[i] = Vector{Any}()
    end

    let
        top, bottom = row_bounds(rng)
        left = dim.start.column_number
        right = dim.stop.column_number

        for sheetrow in eachrow(ws)
            if sheetrow.row > bottom
                break
            end
            if top > sheetrow.row
                continue
            else
                row_index=sheetrow.row-top+1
                for column in left:right
                    cell = getcell(sheetrow, column)
                    push!(rows[row_index], getdata(ws, cell))
                end
            end
        end
    end

    cols = length(rows[1])
    for r in rows
        length(r) != cols && throw(XLSXError("Inconsistent state: Each row should have the same number of columns."))
    end

    return permutedims(hcat(rows...))
end
=#

function getdata(ws::Worksheet, rng::NonContiguousRange)::Vector{Any}
    results = Vector{Any}()
    for r in rng.rng
        if r isa CellRef
            push!(results, getdata(ws, r))
        else
            for cell in r
                push!(results, getdata(ws, cell))
            end
        end
    end
    return results
end

# Needed for definedName references
getdata(ws::Worksheet, s::SheetCellRef) = getdata(ws, s.cellref)
getdata(ws::Worksheet, s::SheetCellRange) = getdata(ws, s.rng)
getdata(ws::Worksheet, s::SheetColumnRange) = getdata(ws, s.colrng)
getdata(ws::Worksheet, s::SheetRowRange) = getdata(ws, s.rowrng)

function getdata(ws::Worksheet, ref::AbstractString)::Union{Array{Any,2},Any}
    if is_worksheet_defined_name(ws, ref)
        v = get_defined_name_value(ws, ref)
        if is_defined_name_value_a_constant(v)
            return v
        elseif is_defined_name_value_a_reference(v)
            return getdata(ws, v)
        else
            throw(XLSXError("Unexpected defined name value: $v."))
        end
    elseif is_workbook_defined_name(get_workbook(ws), ref)
        wb = get_workbook(ws)
        v = get_defined_name_value(wb, ref)
        if is_defined_name_value_a_constant(v)
            return v
        elseif is_defined_name_value_a_reference(v)
            return getdata(get_xlsxfile(ws), v)
        else
            throw(XLSXError("Unexpected defined name value: $v."))
        end
    elseif is_valid_cellname(ref)
        return getdata(ws, CellRef(ref))
    elseif is_valid_cellrange(ref)
        return getdata(ws, CellRange(ref))
    elseif is_valid_column_range(ref)
        return getdata(ws, ColumnRange(ref))
    elseif is_valid_row_range(ref)
        return getdata(ws, RowRange(ref))
    elseif is_valid_sheet_cellname(ref)
        return getdata(ws, SheetCellRef(ref))
    elseif is_valid_sheet_cellrange(ref)
        return getdata(ws, SheetCellRange(ref))
    elseif is_valid_sheet_column_range(ref)
        return getdata(ws, SheetColumnRange(ref))
    elseif is_valid_sheet_row_range(ref)
        return getdata(ws, SheetRowRange(ref))
    elseif is_valid_non_contiguous_range(ref)
        return getdata(ws, NonContiguousRange(ws, ref))
    else
        throw(XLSXError("`$ref` is not a valid cell or range reference."))
    end
end

function getdata(ws::Worksheet)
    if ws.dimension !== nothing
        return getdata(ws, get_dimension(ws))
    else
        throw(XLSXError("Worksheet dimension is unknown."))
    end
end

Base.getindex(ws::Worksheet, r) = getdata(ws, r)
Base.getindex(ws::Worksheet, r, c) = getdata(ws, r, c)
Base.getindex(ws::Worksheet, ::Colon) = getdata(ws)

function Base.show(io::IO, ws::Worksheet)
    hidden_string = ws.is_hidden ? "(hidden)" : ""
    if get_dimension(ws) !== nothing
        rg = get_dimension(ws)
        nrow, ncol = size(rg)
        @printf(io, "%d×%d %s: [\"%s\"](%s) %s", nrow, ncol, typeof(ws), ws.name, rg, hidden_string)
    else
        @printf(io, "%s: [\"%s\"] %s", typeof(ws), ws.name, hidden_string)
    end
end

"""
    getcell(sheet, ref)
    getcell(sheet, row, col)

Return an `AbstractCell` that represents a cell in the spreadsheet.
Return a matrix with cells as `Array{AbstractCell, 2}` if called 
with a reference tomore than one cell.

If `ref` is a range, `getcell` dispatches to `getcellrange`.

Example:

```julia
julia> xf = XLSX.readxlsx("myfile.xlsx")

julia> sheet = xf["mysheet"]

julia> cell = XLSX.getcell(sheet, "A1")

julia> cell = XLSX.getcell(sheet, 1:3, [2,4,6])

Other examples are as [`getdata()`](@ref).
```

"""
function getcell(ws::Worksheet, single::CellRef)::AbstractCell

    # Access cache directly if it exists and if file `isread` - much faster!
    if is_cache_enabled(ws) && ws.cache !== nothing
        if haskey(get_xlsxfile(ws).files, "xl/worksheets/sheet$(ws.sheetId).xml") && get_xlsxfile(ws).files["xl/worksheets/sheet$(ws.sheetId).xml"] == true
            if haskey(ws.cache.cells, single.row_number)
                if haskey(ws.cache.cells[single.row_number], single.column_number)
                    return ws.cache.cells[single.row_number][single.column_number]
                end
            end
            return EmptyCell(single)
        end
    end

    # If can't use cache then iterate sheetrows
    for sheetrow in eachrow(ws)
        if row_number(sheetrow) == row_number(single)
            return getcell(sheetrow, column_number(single))
        end
    end

    return EmptyCell(single)
end
getcell(ws::Worksheet, s::SheetCellRef) = getcell(ws, s.cellref)
getcell(ws::Worksheet, s::SheetCellRange) = getcellrange(ws, s.rng)
getcell(ws::Worksheet, s::SheetColumnRange) = getcellrange(ws, s.colrng)
getcell(ws::Worksheet, s::SheetRowRange) = getcellrange(ws, s.rowrng)
getcell(ws::Worksheet, s::CellRange) = getcellrange(ws, s.rng)
getcell(ws::Worksheet, s::ColumnRange) = getcellrange(ws, s.colrng)
getcell(ws::Worksheet, s::RowRange) = getcellrange(ws, s.rowrng)

getcell(ws::Worksheet, row::Integer, col::Integer) = getcell(ws, CellRef(row, col))
getcell(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Vector{Int}) = [getcell(ws, a, b) for a in collect(row), b in col]
getcell(ws::Worksheet, row::Vector{Int}, col::Union{Integer,UnitRange{<:Integer}}) = [getcell(ws, a, b) for a in row, b in collect(col)]
getcell(ws::Worksheet, row::Vector{Int}, col::Vector{Int}) = [getcell(ws, a, b) for a in row, b in col]
getcell(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = getcellrange(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))))
function getcell(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        getcellrange(ws, CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number)))
    end
end
function getcell(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}})
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        getcellrange(ws, CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col))))
    end
end

function getcell(ws::Worksheet, ref::AbstractString)
    if is_worksheet_defined_name(ws, ref)
        v = get_defined_name_value(ws, ref)
        if is_defined_name_value_a_reference(v)
            return getcell(ws, v)
        else
            throw(XLSXError("`$ref` is not a valid cell or range reference."))
        end
    elseif is_workbook_defined_name(get_workbook(ws), ref)
        wb = get_workbook(ws)
        v = get_defined_name_value(wb, ref)
        if is_defined_name_value_a_reference(v)
            return isa(v, SheetCellRef) ? getcell(get_xlsxfile(ws), v) : getcellrange(get_xlsxfile(ws), v)
        else
            throw(XLSXError("`$ref` is not a valid cell or range reference."))
        end
    elseif is_valid_cellname(ref)
        return getcell(ws, CellRef(ref))
    elseif is_valid_sheet_cellname(ref)
        return getcell(ws, SheetCellRef(ref))
    elseif is_valid_cellrange(ref)
        return getcellrange(ws, CellRange(ref))
    elseif is_valid_column_range(ref)
        return getcellrange(ws, ColumnRange(ref))
    elseif is_valid_row_range(ref)
        return getcellrange(ws, RowRange(ref))
    elseif is_valid_non_contiguous_range(ref)
        return getcellrange(ws, NonContiguousRange(ws, ref))
    elseif is_valid_sheet_cellrange(ref)
        return getcellrange(ws, SheetCellRange(ref))
    elseif is_valid_sheet_column_range(ref)
        return getcellrange(ws, SheetColumnRange(ref))
    elseif is_valid_sheet_row_range(ref)
        return getcellrange(ws, SheetRowRange(ref))
    elseif is_valid_non_contiguous_range(ref)
        return getcellrange(ws, NonContiguousRange(ref))
    end
    throw(XLSXError("`$ref` is not a valid cell or range reference."))
end

"""
    getcellrange(sheet, rng)

Return a matrix with cells as `Array{AbstractCell, 2}`.
`rng` must be a valid cell range, column range or row range,
as in `"A1:B2"`, `"A:B"` or `"1:2"`, or a non-contiguous range.
For row and column ranges, the extent of the range in the other 
dimension is determined by the worksheet's dimension.
A non-contiguous range (which is not rectangular) will return a vector.

For example usage, see [`getdata()`](@ref).

"""
function getcellrange(ws::Worksheet, rng::CellRange)::Array{AbstractCell,2}
    result = Array{AbstractCell,2}(undef, size(rng))
    for cellref in rng
        (r, c) = relative_cell_position(cellref, rng)
        cell = getcell(ws, cellref)
        result[r, c] = isempty(cell) ? EmptyCell(cellref) : cell
    end
    #=
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
    =#
    return result
end

getcellrange(ws::Worksheet, s::SheetCellRef) = getcellrange(ws, s.cellref)
getcellrange(ws::Worksheet, s::SheetCellRange) = getcellrange(ws, s.rng)
getcellrange(ws::Worksheet, s::SheetColumnRange) = getcellrange(ws, s.colrng)
getcellrange(ws::Worksheet, s::SheetRowRange) = getcellrange(ws, s.rowrng)

getcellrange(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Vector{Int}) = [getcell(ws, a, b) for a in collect(row), b in col]
getcellrange(ws::Worksheet, row::Vector{Int}, col::Union{Integer,UnitRange{<:Integer}}) = [getcell(ws, a, b) for a in row, b in collect(col)]
getcellrange(ws::Worksheet, row::Vector{Int}, col::Vector{Int}) = [getcell(ws, a, b) for a in row, b in col]
getcellrange(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = getcell(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))))
getcellrange(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon) = getcell(ws, row, :)
getcellrange(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}) = getcell(ws, :, col)

function getcellrange(ws::Worksheet, rng::ColumnRange)::Array{AbstractCell,2}
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        start = CellRef(dim.start.row_number, rng.start)
        stop = CellRef(dim.stop.row_number, rng.stop)
        return getcellrange(ws, CellRange(start, stop))
    end
end
function getcellrange(ws::Worksheet, rng::RowRange)::Array{AbstractCell,2}
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        start = CellRef(rng.start, dim.start.column_number,)
        stop = CellRef(rng.stop, dim.stop.column_number)
        return getcellrange(ws, CellRange(start, stop))
    end
end

#=
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
        length(columns[i]) != rows && throw(XLSXError("Inconsistent state: Each column should have the same number of rows."))
    end

    return hcat(columns...)
end

function getcellrange(ws::Worksheet, rng::RowRange) :: Array{AbstractCell,2}
    dim = get_dimension(ws)

    rows = Vector{Vector{AbstractCell}}()

    let
        top, bottom = row_bounds(rng)
        left = dim.start.column_number
        right = dim.stop.column_number

        for (i, sheetrow) in enumerate(eachrow(ws))
            push!(rows, Vector{AbstractCell}())
            if top <= sheetrow.row && sheetrow.row <= bottom
                for column in left:right
                    cell = getcell(sheetrow, column)
                    push!(rows[i], cell)
                end
            end
            if sheetrow.row > bottom
                break
            end
        end
    end

    cols = length(rows[1])
    for r in rows
        length(r) != cols && throw(XLSXError("Inconsistent state: Each row should have the same number of columns."))
    end

    return permutedims(hcat(rows...))
end
=#
function getcellrange(ws::Worksheet, rng::NonContiguousRange)::Vector{AbstractCell}
    results = Vector{AbstractCell}()
    for r in rng.rng
        if r isa CellRef
            push!(results, getcell(ws, r))
        else
            for cell in r
                push!(results, getcell(ws, cell))
            end
        end
    end
    return results
end

function getcellrange(ws::Worksheet, rng::AbstractString)
    if is_worksheet_defined_name(ws, rng)
        v = get_defined_name_value(ws, rng)
        if is_defined_name_value_a_reference(v)
            return getcellrange(ws, v)
        else
            throw(XLSXError("$rng is not a valid cell range."))
        end
    elseif is_workbook_defined_name(get_workbook(ws), rng)
        wb = get_workbook(ws)
        v = get_defined_name_value(wb, rng)
        if is_defined_name_value_a_reference(v)
            isa(v, SheetCellRef) && throw(XLSXError("`$rng` is not a valid cell range."))
            return getcellrange(get_xlsxfile(ws), v)
        else
            throw(XLSXError("`$rng` is not a valid cell range."))
        end
    elseif is_valid_cellrange(rng)
        return getcellrange(ws, CellRange(rng))
    elseif is_valid_column_range(rng)
        return getcellrange(ws, ColumnRange(rng))
    elseif is_valid_row_range(rng)
        return getcellrange(ws, RowRange(rng))
    elseif is_valid_non_contiguous_range(rng)
        return getcellrange(ws, NonContiguousRange(ws, rng))
    elseif is_valid_sheet_cellrange(rng)
        return getcellrange(s, SheetCellRange(rng))
    elseif is_valid_sheet_column_range(rng)
        return getcellrange(s, SheetColumnRange(rng))
    elseif is_valid_sheet_row_range(rng)
        return getcellrange(s, SheetRowRange(rng))
    elseif is_valid_non_contiguous_range(rng)
        return getcellrange(s, NonContiguousRange(rng))
    end
    throw(XLSXError("`$rng` is not a valid cell range."))
end
