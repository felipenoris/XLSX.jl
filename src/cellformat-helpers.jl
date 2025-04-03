
const font_tags = ["b", "i", "u", "strike", "outline", "shadow", "condense", "extend", "sz", "color", "name", "scheme"]
const border_tags = ["left", "right", "top", "bottom", "diagonal"]
const fill_tags = ["patternFill"]
const builtinFormats = Dict(
    "0" => "General",
    "1" => "0",
    "2" => "0.00",
    "3" => "#,##0",
    "4" => "#,##0.00",
    "5" => "\$#,##0_);(\$#,##0)",
    "6" => "\$#,##0_);Red",
    "7" => "\$#,##0.00_);(\$#,##0.00)",
    "8" => "\$#,##0.00_);Red",
    "9" => "0%",
    "10" => "0.00%",
    "11" => "0.00E+00",
    "12" => "# ?/?",
    "13" => "# ??/??",
    "14" => "m/d/yyyy",
    "15" => "d-mmm-yy",
    "16" => "d-mmm",
    "17" => "mmm-yy",
    "18" => "h:mm AM/PM",
    "19" => "h:mm:ss AM/PM",
    "20" => "h:mm",
    "21" => "h:mm:ss",
    "22" => "m/d/yyyy h:mm",
    "37" => "#,##0_);(#,##0)",
    "38" => "#,##0_);Red",
    "39" => "#,##0.00_);(#,##0.00)",
    "40" => "#,##0.00_);Red",
    "45" => "mm:ss",
    "46" => "[h]:mm:ss",
    "47" => "mmss.0",
    "48" => "##0.0E+0",
    "49" => "@"
)
const builtinFormatNames = Dict(
    "General" => 0,
    "Number" => 2,
    "Currency" => 7,
    "Percentage" => 9,
    "ShortDate" => 14,
    "LongDate" => 15,
    "Time" => 21,
    "Scientific" => 48
)
const floatformats = r"""
\.[0#?]|
[0#?]e[+-]?[0#?]|
[0#?]/[0#?]|
%
"""ix

#
# -- A bunch of helper functions ...
#

function copynode(o::XML.Node)
    n = XML.parse(XML.Node, XML.write(o))[1]
    n = XML.Node(n.nodetype, n.tag, isnothing(n.attributes) ? XML.OrderedDict{String,String}() : n.attributes, n.value, isnothing(n.children) ? Vector{XML.Node}() : n.children)
    return n
end
function buildNode(tag::String, attributes::Dict{String,Union{Nothing,Dict{String,String}}})::XML.Node
    if tag == "font"
        attribute_tags = font_tags
    elseif tag == "border"
        attribute_tags = border_tags
    elseif tag == "fill"
        attribute_tags = fill_tags
    else
        throw(XLSXError("Unknown tag: $tag"))
    end
    new_node = XML.Element(tag)
    for a in attribute_tags # Use this as a device to keep ordering constant for Excel
        if tag == "font"
            if haskey(attributes, a)
                if isnothing(attributes[a])
                    cnode = XML.Element(a)
                else
                    cnode = XML.Node(XML.Element, a, XML.OrderedDict{String,String}(), nothing, tag ∈ ["border", "fill"] ? Vector{XML.Node}() : nothing)
                    for (k, v) in attributes[a]
                        cnode[k] = v
                    end
                end
                push!(new_node, cnode)
            end
        elseif tag == "border"
            if haskey(attributes, a)
                if isnothing(attributes[a])
                    cnode = XML.Element(a)
                else
                    cnode = XML.Node(XML.Element, a, XML.OrderedDict{String,String}(), nothing, tag ∈ ["border", "fill"] ? Vector{XML.Node}() : nothing)
                    color = XML.Element("color")
                    for (k, v) in attributes[a]
                        if k == "style" && v != "none"
                            cnode[k] = v
                        elseif k == "direction"
                            if v in ["up", "both"]
                                new_node["diagonalUp"] = "1"
                            end
                            if v in ["down", "both"]
                                new_node["diagonalDown"] = "1"
                            end
                        else
                            color[k] = v
                         end
                    end
                    if length(XML.attributes(color)) > 0 # Don't push an empty color.
                        push!(cnode, color)
                    end
                end
                push!(new_node, cnode)
            end
        elseif tag == "fill"
            if haskey(attributes, a)
                if isnothing(attributes[a])
                    cnode = XML.Element(a)
                else
                    cnode = XML.Node(XML.Element, a, XML.OrderedDict{String,String}(), nothing, tag ∈ ["border", "fill"] ? Vector{XML.Node}() : nothing)
                    patternfill = XML.Element("patternFill")
                    fgcolor = XML.Element("fgColor")
                    bgcolor = XML.Element("bgColor")
                    for (k, v) in attributes[a]
                        if k == "patternType"
                            patternfill[k] = v
                        elseif first(k, 2) == "fg"
                            fgcolor[k[3:end]] = v
                        elseif first(k, 2) == "bg"
                            bgcolor[k[3:end]] = v
                        end
                    end
                    if !haskey(patternfill, "patternType")
                        throw(XLSXError("No `patternType` attribute found."))
                    end
                    length(XML.attributes(fgcolor)) > 0 && push!(patternfill, fgcolor)
                    length(XML.attributes(bgcolor)) > 0 && push!(patternfill, bgcolor)
                end
                push!(new_node, patternfill)
            end
        else
        end
    end
    return new_node
end

function update_template_xf(ws::Worksheet, existing_style::CellDataFormat, attributes::Vector{String}, vals::Vector{String})::CellDataFormat
    old_cell_xf = styles_cell_xf(ws.package.workbook, Int(existing_style.id))
    new_cell_xf = copynode(old_cell_xf)
    if length(attributes) != length(vals)
        throw(XLSXError("Attributes and values must be of the same length."))
    end
    for (a, v) in zip(attributes, vals)
        new_cell_xf[a] = v
    end
    return styles_add_cell_xf(ws.package.workbook, new_cell_xf)
end
function update_template_xf(ws::Worksheet, existing_style::CellDataFormat, alignment::XML.Node)::CellDataFormat
    old_cell_xf = styles_cell_xf(ws.package.workbook, Int(existing_style.id))
    new_cell_xf = copynode(old_cell_xf)
    if length(XML.children(new_cell_xf)) == 0
        push!(new_cell_xf, alignment)
    else
        new_cell_xf[1] = alignment
    end
    return styles_add_cell_xf(ws.package.workbook, new_cell_xf)
end

# Only used in testing!
function styles_add_cell_font(wb::Workbook, attributes::Dict{String,Union{Dict{String,String},Nothing}})::Int
    new_font = buildNode("font", attributes)
    return styles_add_cell_attribute(wb, new_font, "fonts")
end

# Used by setFont(), setBorder(), setFill(), setAlignment() and setNumFmt()
function styles_add_cell_attribute(wb::Workbook, new_att::XML.Node, att::String)::Int
    xroot = styles_xmlroot(wb)
    i, j = get_idces(xroot, "styleSheet", att)
    existing_elements_count = length(XML.children(xroot[i][j]))
    if parse(Int, xroot[i][j]["count"]) != existing_elements_count
        throw(XLSXError("Wrong number of elements elements found: $existing_elements_count. Expected $(parse(Int, xroot[i][j]["count"]))."))
    end

    # Check new_att doesn't duplicate any existing att. If yes, use that rather than create new.
    for (k, node) in enumerate(XML.children(xroot[i][j]))
        if XML.tag(new_att) == "numFmt" # mustn't compare numFmtId attribute for formats
            if XML.parse(XML.Node, XML.write(node))[1]["formatCode"] == XML.parse(XML.Node, XML.write(new_att))[1]["formatCode"] # XML.jl defines `Base.:(==)`
                return k - 1 # CellDataFormat is zero-indexed
            end
        else
            if XML.parse(XML.Node, XML.write(node))[1] == XML.parse(XML.Node, XML.write(new_att))[1] # XML.jl defines `Base.:(==)`
                return k - 1 # CellDataFormat is zero-indexed
            end
        end
    end

    push!(xroot[i][j], new_att)
    xroot[i][j]["count"] = string(existing_elements_count + 1)

    return existing_elements_count # turns out this is the new index (because it's zero-based)
end
function process_sheetcell(f::Function, xl::XLSXFile, sheetcell::String; kw...)::Int
    if is_workbook_defined_name(xl, sheetcell)
        v = get_defined_name_value(xl.workbook, sheetcell)
        if is_defined_name_value_a_constant(v)
            throw(XLSXError("Can only assign attributes to cells but `$(sheetcell)` is a constant: $(sheetcell)=$v."))
        elseif is_defined_name_value_a_reference(v)
            newid = process_ranges(f, xl, string(v); kw...)
        else
            throw(XLSXError("Unexpected defined name value: $v."))
        end
    elseif is_valid_non_contiguous_sheetcellrange(sheetcell)
        sheetncrng = NonContiguousRange(sheetcell)
        !hassheet(xl, sheetncrng.sheet) && throw(XLSXError("Sheet $(sheetncrng.sheet) not found."))
        newid = f(xl[sheetncrng.sheet], sheetncrng; kw...)
    elseif is_valid_sheet_column_range(sheetcell)
        sheetcolrng = SheetColumnRange(sheetcell)
        !hassheet(xl, sheetcolrng.sheet) && throw(XLSXError("Sheet $(sheetcolrng.sheet) not found."))
        newid = f(xl[sheetcolrng.sheet], sheetcolrng.colrng; kw...)
    elseif is_valid_sheet_row_range(sheetcell)
        sheetrowrng = SheetRowRange(sheetcell)
        !hassheet(xl, sheetrowrng.sheet) && throw(XLSXError("Sheet $(sheetrowrng.sheet) not found."))
        newid = f(xl[sheetrowrng.sheet], sheetrowrng.rowrng; kw...)
    elseif is_valid_sheet_cellrange(sheetcell)
        sheetcellrng = SheetCellRange(sheetcell)
        !hassheet(xl, sheetcellrng.sheet) && throw(XLSXError("Sheet $(sheetcellrng.sheet) not found."))
        newid = f(xl[sheetcellrng.sheet], sheetcellrng.rng; kw...)
    elseif is_valid_sheet_cellname(sheetcell)
        ref = SheetCellRef(sheetcell)
        !hassheet(xl, ref.sheet) && throw(XLSXError("Sheet $(ref.sheet) not found."))
        newid = f(getsheet(xl, ref.sheet), ref.cellref; kw...)
    else
        throw(XLSXError("Invalid sheet cell reference: $sheetcell"))
    end
    return newid
end
function process_ranges(f::Function, ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int
    # Moved the tests for defined names to be first in case a name looks like a column name (e.g. "ID")
    if is_worksheet_defined_name(ws, ref_or_rng)
        v = get_defined_name_value(ws, ref_or_rng)
        if is_defined_name_value_a_constant(v)
            throw(XLSXError("Can only assign attributes to cells but `$(ref_or_rng)` is a constant: $(ref_or_rng)=$v."))
        elseif is_defined_name_value_a_reference(v)
            wb = get_workbook(ws)
            newid = f(get_xlsxfile(wb), string(v); kw...)
        else
            throw(XLSXError("Unexpected defined name value: $v."))
        end
    elseif is_workbook_defined_name(get_workbook(ws), ref_or_rng)
        wb = get_workbook(ws)
        v = get_defined_name_value(wb, ref_or_rng)
        if is_defined_name_value_a_constant(v)
            throw(XLSXError("Can only assign attributes to cells but `$(ref_or_rng)` is a constant: $(ref_or_rng)=$v."))
        elseif is_defined_name_value_a_reference(v)
            if is_valid_non_contiguous_range(string(v))
                _ = f.(Ref(get_xlsxfile(wb)), replace.(split(string(v), ","), "'" => "", "\$" => ""); kw...)
                newid = -1
            else
                newid = f(get_xlsxfile(wb), replace(string(v), "'" => "", "\$" => ""); kw...)
            end
        else
            throw(XLSXError("Unexpected defined name value: $v."))
        end
    elseif is_valid_column_range(ref_or_rng)
        colrng = ColumnRange(ref_or_rng)
        newid = f(ws, colrng; kw...)
    elseif is_valid_row_range(ref_or_rng)
        rowrng = RowRange(ref_or_rng)
        newid = f(ws, rowrng; kw...)
    elseif is_valid_cellrange(ref_or_rng)
        rng = CellRange(ref_or_rng)
        newid = f(ws, rng; kw...)
    elseif is_valid_cellname(ref_or_rng)
        newid = f(ws, CellRef(ref_or_rng); kw...)
    else
        throw(XLSXError("Invalid cell reference or range: $ref_or_rng"))
    end
    return newid
end
function process_columnranges(f::Function, ws::Worksheet, colrng::ColumnRange; kw...)::Int
    bounds = column_bounds(colrng)
    dim = (get_dimension(ws))
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        left = bounds[begin]
        right = bounds[end]
        top = dim.start.row_number
        bottom = dim.stop.row_number

        OK = dim.start.column_number <= left
        OK &= dim.stop.column_number >= right
        OK &= dim.start.row_number <= top
        OK &= dim.stop.row_number >= bottom

        if OK
            rng = CellRange(top, left, bottom, right)
            return f(ws, rng; kw...)
        else
            throw(XLSXError("Column range $colrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
        end
    end
end
function process_rowranges(f::Function, ws::Worksheet, rowrng::RowRange; kw...)::Int
    bounds = row_bounds(rowrng)
    dim = (get_dimension(ws))
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        top = bounds[begin]
        bottom = bounds[end]
        left = dim.start.column_number
        right = dim.stop.column_number

        OK = dim.start.column_number <= left
        OK &= dim.stop.column_number >= right
        OK &= dim.start.row_number <= top
        OK &= dim.stop.row_number >= bottom

        if OK
            rng = CellRange(top, left, bottom, right)
            return f(ws, rng; kw...)
        else
            throw(XLSXError("Row range $rowrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
        end
    end
end
function process_ncranges(f::Function, ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int
    bounds = nc_bounds(ncrng)
    dim = (get_dimension(ws))
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        OK = dim.start.column_number <= bounds.start.column_number
        OK &= dim.stop.column_number >= bounds.stop.column_number
        OK &= dim.start.row_number <= bounds.start.row_number
        OK &= dim.stop.row_number >= bounds.stop.row_number
        if OK
            for r in ncrng.rng
                _ = f(ws, r; kw...)
            end
            return -1
        else
            throw(XLSXError("Non-contiguous range $ncrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
        end
    end
end
function process_cellranges(f::Function, ws::Worksheet, rng::CellRange; kw...)::Int
    if length(rng)==1
        single=true
    else
        single=false
    end
    for cellref in rng
        if getcell(ws, cellref) isa EmptyCell
            single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellref.name). Set the value first."))
            continue
        end
        _ = f(ws, cellref; kw...)
    end
    return -1 # Each cell may have a different attribute Id so we can't return a single value.
end

function process_get_sheetcell(f::Function, xl::XLSXFile, sheetcell::String; kw...)
    ref = SheetCellRef(sheetcell)
    !hassheet(xl, ref.sheet) && throw(XLSXError("Sheet $(ref.sheet) not found."))
    ws = getsheet(xl, ref.sheet)
    d = get_dimension(ws)
    if ref.cellref ∉ d
        throw(XLSXError("Cell specified is outside sheet dimension \"$d\""))
    end
    return f(ws, ref.cellref; kw...)
end
function process_get_cellref(f::Function, ws::Worksheet, cellref::CellRef; kw...)
    wb = get_workbook(ws)
    cell = getcell(ws, cellref)
    d = get_dimension(ws)
    if cellref ∉ d
        throw(XLSXError("Cell specified is outside sheet dimension \"$d\""))
    end
    if cell isa EmptyCell || cell.style == ""
        return nothing
    end
    cell_style = styles_cell_xf(wb, parse(Int, cell.style))
    return f(wb, cell_style; kw...)
end
function process_get_cellname(f::Function, ws::Worksheet, ref_or_rng::AbstractString; kw...)
    if is_workbook_defined_name(get_workbook(ws), ref_or_rng)
        wb = get_workbook(ws)
        v = get_defined_name_value(wb, ref_or_rng)
        if is_defined_name_value_a_constant(v) # Can these have fonts?
            throw(XLSXError("Can only assign borders to cells but `$(ref_or_rng)` is a constant: $(ref_or_rng)=$v."))
        elseif is_defined_name_value_a_reference(v)
            new_att = f(get_xlsxfile(wb), replace(string(v), "'" => ""); kw...)
        else
            throw(XLSXError("Unexpected defined name value: $v."))
        end
    elseif is_valid_cellname(ref_or_rng)
        new_att = f(ws, CellRef(ref_or_rng); kw...)
    else
        throw(XLSXError("Invalid cell reference or range: $ref_or_rng"))
    end
    return new_att
end

#
# - Used for indexing `setAttribute` family of functions
#
function process_colon(f::Function, ws::Worksheet, ::Colon; kw...)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        return f(ws, dim; kw...)
    end
end
function process_intcolon(f::Function, ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        rng=CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number))
        return f(ws, rng; kw...)
    end
end
function process_colonint(f::Function, ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        rng=CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col)))
        return f(ws, rng; kw...)
    end
end
function process_veccolon(f::Function, ws::Worksheet, row::Vector{Int}, ::Colon; kw...)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        if length(row)==1 && dim.start.column_number == dim.stop.column_number
            single=true
        else
            single=false
        end
        for a in row
            for b in dim.start.column_number:dim.stop.column_number
                cellref = CellRef(a, b)
                if getcell(ws, cellref) isa EmptyCell
                    single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellref.name). Set the value first."))
                    continue
                end
                f(ws, cellref; kw...)
            end
        end
    end
    return -1
end
function process_colonvec(f::Function, ws::Worksheet, ::Colon, col::Vector{Int}; kw...)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        if length(col)==1 && dim.start.row_number == dim.stop.row_number
            single=true
        else
            single=false
        end
        for b in col
            for a in dim.start.row_number:dim.stop.row_number
                cellref = CellRef(a, b)
                if getcell(ws, cellref) isa EmptyCell
                    single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellref.name). Set the value first."))
                    continue
                end
                f(ws, cellref; kw...)
            end
        end
    end
    return -1
end
function process_intvec(f::Function, ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Vector{Int}; kw...)
    if length(col)==1 && length(row)==1
        single=true
    else
        single=false
    end
    for a in collect(row), b in col
        cellref = CellRef(a, b)
        if getcell(ws, cellref) isa EmptyCell
            single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellref.name). Set the value first."))
            continue
        end
        f(ws, cellref; kw...)
    end
    return -1
end
function process_vecint(f::Function, ws::Worksheet, row::Vector{Int}, col::Union{Integer,UnitRange{<:Integer}}; kw...)
    if length(col)==1 && length(row)==1
        single=true
    else
        single=false
    end
    for a in row, b in collect(col)
        cellref = CellRef(a, b)
        if getcell(ws, cellref) isa EmptyCell
            single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellref.name). Set the value first."))
            continue
        end
        f(ws, cellref; kw...)
    end
    return -1
end
function process_vecvec(f::Function, ws::Worksheet, row::Vector{Int}, col::Vector{Int}; kw...)
    if length(col)==1 && length(row)==1
        single=true
    else
        single=false
    end
    for a in row, b in col
        cellref = CellRef(a, b)
        if getcell(ws, cellref) isa EmptyCell
            continue
            single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellref.name). Set the value first."))
        end
        f(ws, cellref; kw...)
    end
    return -1
end

#
# - Used for indexing `setUniformAttribute` family of functions
#

#
# Most set functions
#
function process_uniform_core(f::Function, ws::Worksheet, cellref::CellRef, atts::Vector{String}, newid::Union{Int,Nothing}, first::Bool; kw...)
    cell = getcell(ws, cellref)
    if cell isa EmptyCell # Can't add a attribute to an empty cell.
        return newid, first
    end
    if first                           # Get the attribute of the first cell in the range.
        newid = f(ws, cellref; kw...)
        first = false
    else                               # Apply the same attribute to the rest of the cells in the range.
        if cell.style == ""
            cell.style = string(get_num_style_index(ws, 0).id)
        end
        cell.style = string(update_template_xf(ws, CellDataFormat(parse(Int, cell.style)), atts, ["$newid", "1"]).id)
    end
    return newid, first
end
function process_uniform_attribute(f::Function, ws::Worksheet, rng::CellRange, atts::Vector{String}; kw...)

    if !get_xlsxfile(ws).use_cache_for_sheet_data
        throw(XLSXError("Cannot set uniform attributes because cache is not enabled."))
    end

    let newid::Union{Int,Nothing}, first::Bool
        newid = nothing
        first = true
        for cellref in rng
            newid, first = process_uniform_core(f, ws, cellref, atts, newid, first; kw...)
        end
        if first
            newid = -1
        end
        return newid
    end
end
function process_uniform_colon(f::Function, ws::Worksheet, ::Colon; kw...)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        f(ws, dim; kw...)
    end
end
function process_uniform_intcolon(f::Function, ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        f(ws, CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number)); kw...)
    end
end
function process_uniform_colonint(f::Function, ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        f(ws, CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col))); kw...)
    end
end
function process_uniform_veccolon(f::Function, ws::Worksheet, row::Vector{Int}, ::Colon, atts::Vector{String}; kw...)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        let newid::Union{Int,Nothing}, first::Bool
            newid = nothing
            first = true
            for a in row
                for b in dim.start.column_number:dim.stop.column_number
                    cellref = CellRef(a, b)
                    if getcell(ws, cellref) isa EmptyCell
                        continue
                    end
                    newid, first = process_uniform_core(f, ws, cellref, atts, newid, first; kw...)
                end
            end
            if first
                newid = -1
            end
            return newid
        end
    end
end
function process_uniform_colonvec(f::Function, ws::Worksheet, ::Colon, col::Vector{Int}, atts::Vector{String}; kw...)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        let newid::Union{Int,Nothing}, first::Bool
            newid = nothing
            first = true
            for b in col
                for a in dim.start.row_number:dim.stop.row_number
                    cellref = CellRef(a, b)
                    if getcell(ws, cellref) isa EmptyCell
                        continue
                    end
                    newid, first = process_uniform_core(f, ws, cellref, atts, newid, first; kw...)
                end
            end
            if first
                newid = -1
            end
            return newid
        end
    end
end
function process_uniform_intvec(f::Function, ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Vector{Int}, atts::Vector{String}; kw...)
    let newid::Union{Int,Nothing}, first::Bool
        newid = nothing
        first = true
        for a in collect(row), b in col
            cellref = CellRef(a, b).name
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first = process_uniform_core(f, ws, cellref, atts, newid, first; kw...)
        end
        if first
            newid = -1
        end
        return newid
    end
end
function process_uniform_vecint(f::Function, ws::Worksheet, row::Vector{Int}, col::Union{Integer,UnitRange{<:Integer}}, atts::Vector{String}; kw...)
    let newid::Union{Int,Nothing}, first::Bool
        newid = nothing
        first = true
        for a in row, b in collect(col)
            cellref = CellRef(a, b)
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first = process_uniform_core(f, ws, cellref, atts, newid, first; kw...)
        end
        if first
            newid = -1
        end
        return newid
    end
end
function process_uniform_vecvec(f::Function, ws::Worksheet, row::Vector{Int}, col::Vector{Int}, atts::Vector{String}; kw...)
    let newid::Union{Int,Nothing}, first::Bool
        newid = nothing
        first = true
        for a in row, b in col
            cellref = CellRef(a, b)
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first = process_uniform_core(f, ws, cellref, atts, newid, first; kw...)
        end
        if first
            newid = -1
        end
        return newid
    end
end

#
# UniformStyles
#
function process_uniform_core(ws::Worksheet, cellref::CellRef, newid::Union{Int,Nothing}, first::Bool)
    cell = getcell(ws, cellref)
    if cell isa EmptyCell # Can't add a attribute to an empty cell.
        return newid, first
    end
    if first                           # Get the style of the first cell in the range.
        newid = parse(Int, cell.style)
        first = false
    else                               # Apply the same style to the rest of the cells in the range.
        cell.style = string(newid)
    end
    return newid, first
end
function process_uniform_intcolon(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        setUniformStyle(ws, CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number)))
    end
end
function process_uniform_colonint(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}})
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        setUniformStyle(ws, CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col))))
    end
end
function process_uniform_colon(ws::Worksheet, ::Colon)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        setUniformStyle(ws, dim)
    end
end
function process_uniform_veccolon(ws::Worksheet, row::Vector{Int}, ::Colon)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        let newid::Union{Int,Nothing}, first::Bool
            newid = nothing
            first = true
            for a in row
                for b in dim.start.column_number:dim.stop.column_number
                    cellref = CellRef(a, b)
                    if getcell(ws, cellref) isa EmptyCell
                        continue
                    end
                    newid, first = process_uniform_core(ws, cellref, newid, first)
                end
            end
            if first
                newid = -1
            end
            return newid
        end
    end
end
function process_uniform_colonvec(ws::Worksheet, ::Colon, col::Vector{Int})
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        let newid::Union{Int,Nothing}, first::Bool
            newid = nothing
            first = true
            for b in col
                for a in dim.start.row_number:dim.stop.row_number
                    cellref = CellRef(a, b)
                    if getcell(ws, cellref) isa EmptyCell
                        continue
                    end
                    newid, first = process_uniform_core(ws, cellref, newid, first)
                end
            end
            if first
                newid = -1
            end
            return newid
        end
    end
end
function process_uniform_intvec(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Vector{Int})
    let newid::Union{Int,Nothing}, first::Bool
        newid = nothing
        first = true
        for a in collect(row), b in col
            cellref = CellRef(a, b).name
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first = process_uniform_core(ws, cellref, newid, first)
        end
        if first
            newid = -1
        end
        return newid
    end
end
function process_uniform_vecint(ws::Worksheet, row::Vector{Int}, col::Union{Integer,UnitRange{<:Integer}})
    let newid::Union{Int,Nothing}, first::Bool
        newid = nothing
        first = true
        for a in row, b in collect(col)
            cellref = CellRef(a, b)
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first = process_uniform_core(ws, cellref, newid, first)
        end
        if first
            newid = -1
        end
        return newid
    end
end
function process_uniform_vecvec(ws::Worksheet, row::Vector{Int}, col::Vector{Int})
    let newid::Union{Int,Nothing}, first::Bool
        newid = nothing
        first = true
        for a in row, b in col
            cellref = CellRef(a, b)
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first = process_uniform_core(ws, cellref, newid, first)
        end
        if first
            newid = -1
        end
        return newid
    end
end

#
# Alignment is different
#
function process_uniform_core(f::Function, ws::Worksheet, cellref::CellRef, newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}; kw...) # setUniformAlignment is different
    cell = getcell(ws, cellref)
    if cell isa EmptyCell # Can't add a attribute to an empty cell.
        return newid, first, alignment_node
    end
    if first                           # Get the attribute of the first cell in the range.
        newid = f(ws, cellref; kw...)
        new_alignment = getAlignment(ws, cellref).alignment["alignment"]
        alignment_node = XML.Node(XML.Element, "alignment", new_alignment, nothing, nothing)
        first = false
    else                               # Apply the same attribute to the rest of the cells in the range.
        if cell.style == ""
            cell.style = string(get_num_style_index(ws, 0).id)
        end
        cell.style = string(update_template_xf(ws, CellDataFormat(parse(Int, cell.style)), alignment_node).id)
    end
    return newid, first, alignment_node
end
function process_uniform_attribute(f::Function, ws::Worksheet, rng::CellRange; kw...)

    if !get_xlsxfile(ws).use_cache_for_sheet_data
        throw(XLSXError("Cannot set uniform attributes because cache is not enabled."))
    end

    let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
        newid = nothing
        first = true
        alignment_node = nothing
        for cellref in rng
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first, alignment_node = process_uniform_core(f, ws, cellref, newid, first, alignment_node; kw...)
        end
        if first
            newid = -1
        end
        return newid
    end
end
function process_uniform_veccolon(f::Function, ws::Worksheet, row::Vector{Int}, ::Colon; kw...)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
            newid = nothing
            first = true
            alignment_node = nothing
            for a in row
                for b in dim.start.column_number:dim.stop.column_number
                    cellref = CellRef(a, b)
                    if getcell(ws, cellref) isa EmptyCell
                        continue
                    end
                    newid, first, alignment_node = process_uniform_core(f, ws, cellref, newid, first, alignment_node; kw...)
                end
            end
            if first
                newid = -1
            end
            return newid
        end
    end
end
function process_uniform_colonvec(f::Function, ws::Worksheet, ::Colon, col::Vector{Int}; kw...)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
            newid = nothing
            first = true
            alignment_node = nothing
            for b in col
                for a in dim.start.row_number:dim.stop.row_number
                    cellref = CellRef(a, b)
                    if getcell(ws, cellref) isa EmptyCell
                        continue
                    end
                    newid, first, alignment_node = process_uniform_core(f, ws, cellref, newid, first, alignment_node; kw...)
                end
            end
            if first
                newid = -1
            end
            return newid
        end
    end
end
function process_uniform_intvec(f::Function, ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Vector{Int}; kw...)
    let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
        newid = nothing
        first = true
        alignment_node = nothing
        for a in collect(row), b in col
            cellref = CellRef(a, b)
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first, alignment_node = process_uniform_core(f, ws, cellref, newid, first, alignment_node; kw...)
        end
        if first
            newid = -1
        end
        return newid
    end
end
function process_uniform_vecint(f::Function, ws::Worksheet, row::Vector{Int}, col::Union{Integer,UnitRange{<:Integer}}; kw...)
    let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
        newid = nothing
        first = true
        alignment_node = nothing
        for a in row, b in collect(col)
            cellref = CellRef(a, b)
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first, alignment_node = process_uniform_core(f, ws, cellref, newid, first, alignment_node; kw...)
        end
        if first
            newid = -1
        end
        return newid
    end
end
function process_uniform_vecvec(f::Function, ws::Worksheet, row::Vector{Int}, col::Vector{Int}; kw...)
    let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
        newid = nothing
        first = true
        for a in row, b in col
            cellref = CellRef(a, b)
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first, alignment_node = process_uniform_core(f, ws, cellref, newid, first, alignment_node; kw...)
        end
        if first
            newid = -1
        end
        return newid
    end
end

# Check if a string is a valid named color in Colors.jl and convert to "FFRRGGBB" if it is.
function get_colorant(color_string::String)
    try
        c = Colors.parse(Colors.Colorant, color_string)
        rgb = Colors.hex(c, :RRGGBB)
        return "FF" * rgb
    catch
        return nothing
    end
end
function get_color(s::String)::String
    if occursin(r"^[0-9A-F]{8}$", s) # is a valid 8 digit hexadecimal color
        return s
    end
    c = get_colorant(s)
    if isnothing(c)
        throw(XLSXError("Invalid color specified: $s. Either give a valid color name (from Colors.jl) or an 8-digit rgb color in the form AARRGGBB"))
    end
    return c
end
