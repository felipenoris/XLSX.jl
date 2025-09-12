
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
    n = XML.Node(o.nodetype, o.tag, o.attributes, o.value, isnothing(o.children) ? nothing : [copynode(x) for x in o.children])
#    n = deepcopy(o)
    return n
end
function do_sheet_names_match(ws::Worksheet, rng::T) where {T<:Union{SheetCellRef,AbstractSheetCellRange}}
    if ws.name == rng.sheet
        return true
    else
        throw(XLSXError("Worksheet `$(ws.name)` does not match sheet in cell reference: `$(rng.sheet)`"))
    end
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
            #else
        end
    end
    return new_node
end
function isInDim(ws::Worksheet, dim::CellRange, rng::CellRange)
    if !issubset(rng, dim)
        throw(XLSXError("Cell range $rng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
    return true
end
function isInDim(ws::Worksheet, dim::CellRange, row, col)
    if maximum(row) > dim.stop.row_number || minimum(row) < dim.start.row_number
        throw(XLSXError("Row range $row is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
    if maximum(col) > dim.stop.column_number || minimum(col) < dim.start.column_number
        throw(XLSXError("Column range $col is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
    return true
end
function get_new_formatId(wb::Workbook, format::String)::Int
    if haskey(builtinFormatNames, uppercasefirst(format)) # User specified a format by name
        return builtinFormatNames[format]
    else                                      # user specified a format code
        code = lowercase(format)
        code = remove_formatting(code)
        if !occursin(floatformats, code) && !any(map(x -> occursin(x, code), DATETIME_CODES)) # Only a very weak test!
            throw(XLSXError("Specified format is not a valid numFmt: $format"))
        end

        xroot = styles_xmlroot(wb)
        i, j = get_idces(xroot, "styleSheet", "numFmts")
        if isnothing(j) # There are no existing custom formats
            return styles_add_numFmt(wb, format)
        else
            existing_elements_count = length(XML.children(xroot[i][j]))
            if parse(Int, xroot[i][j]["count"]) != existing_elements_count
                throw(XLSXError("Wrong number of font elements found: $existing_elements_count. Expected $(parse(Int, xroot[i][j]["count"]))."))
            end

            format_node = XML.Element("numFmt";
                numFmtId=string(existing_elements_count + PREDEFINED_NUMFMT_COUNT),
                formatCode=XML.escape(format)
            )

            return styles_add_cell_attribute(wb, format_node, "numFmts") + PREDEFINED_NUMFMT_COUNT
        end
    end
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
function update_template_xf(ws::Worksheet, allXfNodes::Vector{XML.Node}, existing_style::CellDataFormat, attributes::Vector{String}, vals::Vector{String})::CellDataFormat
    old_cell_xf = styles_cell_xf(allXfNodes, Int(existing_style.id))
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
    if isnothing(new_cell_xf.children)
        new_cell_xf=XML.Node(new_cell_xf, alignment)
    elseif length(XML.children(new_cell_xf)) == 0
        push!(new_cell_xf, alignment)
    else
        new_cell_xf[1] = alignment
    end
    return styles_add_cell_xf(ws.package.workbook, new_cell_xf)
end
function update_template_xf(ws::Worksheet, allXfNodes::Vector{XML.Node}, existing_style::CellDataFormat, alignment::XML.Node)::CellDataFormat
    old_cell_xf = styles_cell_xf(allXfNodes, Int(existing_style.id))
    new_cell_xf = copynode(old_cell_xf)
    if isnothing(new_cell_xf.children)
        new_cell_xf=XML.Node(new_cell_xf, alignment)
    elseif length(XML.children(new_cell_xf)) == 0
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
            if node["formatCode"] == new_att["formatCode"]
                return k - 1 # CellDataFormat is zero-indexed
            end
        else
            if node == new_att
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
            newid = process_sheetcell(f, xl, string(v); kw...)
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
        newid = f(ws, ColumnRange(ref_or_rng); kw...)
    elseif is_valid_row_range(ref_or_rng)
        newid = f(ws, RowRange(ref_or_rng); kw...)
    elseif is_valid_cellrange(ref_or_rng)
        newid = f(ws, CellRange(ref_or_rng); kw...)
    elseif is_valid_cellname(ref_or_rng)
        newid = f(ws, CellRef(ref_or_rng); kw...)
    elseif is_valid_sheet_cellname(ref_or_rng)
        newid = f(ws, SheetCellRef(ref_or_rng); kw...)
    elseif is_valid_sheet_cellrange(ref_or_rng)
        newid = f(ws, SheetCellRange(ref_or_rng); kw...)
    elseif is_valid_sheet_column_range(ref_or_rng)
        newid = f(ws, SheetColumnRange(ref_or_rng); kw...)
    elseif is_valid_sheet_row_range(ref_or_rng)
        newid = f(ws, SheetRowRange(ref_or_rng); kw...)
    elseif is_valid_non_contiguous_cellrange(ref_or_rng)
        newid = f(ws, NonContiguousRange(ws, ref_or_rng); kw...)
    elseif is_valid_non_contiguous_sheetcellrange(ref_or_rng)
        nc = NonContiguousRange(ref_or_rng)
        newid = do_sheet_names_match(ws, nc) && f(ws, nc; kw...)
    else
        throw(XLSXError("Invalid cell reference or range: $ref_or_rng"))
    end
    return newid
end
function process_columnranges(f::Function, ws::Worksheet, colrng::ColumnRange; kw...)::Int
    bounds = column_bounds(colrng)
    dim = (get_dimension(ws))
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
function process_rowranges(f::Function, ws::Worksheet, rowrng::RowRange; kw...)::Int
    bounds = row_bounds(rowrng)
    dim = (get_dimension(ws))
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
function process_ncranges(f::Function, ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int
    bounds = nc_bounds(ncrng)
    if length(ncrng) == 1
        single = true
    else
        single = false
    end
    dim = (get_dimension(ws))
    OK = dim.start.column_number <= bounds.start.column_number
    OK &= dim.stop.column_number >= bounds.stop.column_number
    OK &= dim.start.row_number <= bounds.start.row_number
    OK &= dim.stop.row_number >= bounds.stop.row_number
    if OK
        for r in ncrng.rng
            if r isa CellRef && getcell(ws, r) isa EmptyCell
                single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(r.name). Set the value first."))
                continue
            end
            _ = f(ws, r; kw...)
        end
        return -1
    else
        throw(XLSXError("Non-contiguous range $ncrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
end
function process_cellranges(f::Function, ws::Worksheet, rng::CellRange; kw...)::Int
    if length(rng) == 1
        single = true
    else
        single = false
    end
    isInDim(ws, get_dimension(ws), rng)
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
        if is_defined_name_value_a_constant(v)
            throw(XLSXError("Can only assign attributes to cells but `$(ref_or_rng)` is a constant: $(ref_or_rng)=$v."))
        elseif is_defined_name_value_a_reference(v)
            new_att = f(get_xlsxfile(wb), replace(string(v), "'" => ""); kw...)
        else
            throw(XLSXError("Unexpected defined name value: $v."))
        end
    elseif is_valid_cellname(ref_or_rng)
        new_att = f(ws, CellRef(ref_or_rng); kw...)
    else
        throw(XLSXError("Invalid cell reference: $ref_or_rng"))
    end
    return new_att
end

#
# - Used for indexing `setAttribute` family of functions
#
function process_colon(f::Function, ws::Worksheet, row, col; kw...)
    dim = get_dimension(ws)
    @assert isnothing(row) || isnothing(col) "Something wrong here!"
    if isnothing(row) && isnothing(col)
        return f(ws, dim; kw...)
    elseif isnothing(col)
        rng = CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number))
    else
        rng = CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col)))
#    else
#        throw(XLSXError("Something wrong here!"))
    end

    return f(ws, rng; kw...)
end
function process_veccolon(f::Function, ws::Worksheet, row, col; kw...)
    dim = get_dimension(ws)
    @assert isnothing(row) || isnothing(col) "Something wrong here!"
    if isnothing(col)
        col = dim.start.column_number:dim.stop.column_number
    else
        row = dim.start.row_number:dim.stop.row_number
#    else
#        throw(XLSXError("Something wrong here!"))
    end
    isInDim(ws, dim, row, col)
    if length(row) == 1 && length(col) == 1
        single = true
    else
        single = false
    end
    for a in row
        for b in col
            cellref = CellRef(a, b)
            if getcell(ws, cellref) isa EmptyCell
                single && throw(XLSXError("Cannot set attribute for an `EmptyCell`: $(cellref.name). Set the value first."))
                continue
            end
            f(ws, cellref; kw...)
        end
    end
    return -1
end
function process_vecint(f::Function, ws::Worksheet, row, col; kw...)
    if length(col) == 1 && length(row) == 1
        single = true
    else
        single = false
    end
    dim = get_dimension(ws)
    isInDim(ws, dim, row, col)
    for a in row, b in col
        cellref = CellRef(a, b)
        if getcell(ws, cellref) isa EmptyCell
            single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(cellref.name). Set the value first."))
            continue
        end
        f(ws, cellref; kw...)
    end
    return -1
end

#
# - Used for indexing `setUniformAttribute` family of functions
#

#
# Most setUniform functions (but not Style or Alignment - see below)
#
function process_uniform_core(f::Function, ws::Worksheet, allXfNodes::Vector{XML.Node}, cellref::CellRef, atts::Vector{String}, newid::Union{Int,Nothing}, first::Bool; kw...)
    cell = getcell(ws, cellref)
    if cell isa EmptyCell # Can't add a attribute to an empty cell.
        return newid, first
    end
    if first                           # Get the attribute of the first cell in the range.
        newid = f(ws, cellref; kw...)
        first = false
    else                               # Apply the same attribute to the rest of the cells in the range.
        if cell.style == ""
            cell.style = string(get_num_style_index(ws, allXfNodes, 0).id)
        end
        cell.style = string(update_template_xf(ws, allXfNodes, CellDataFormat(parse(Int, cell.style)), atts, [string(newid), "1"]).id)
    end
    return newid, first
end
function process_uniform_attribute(f::Function, ws::Worksheet, rng::CellRange, atts::Vector{String}; kw...)
    if !get_xlsxfile(ws).use_cache_for_sheet_data
        throw(XLSXError("Cannot set uniform attributes because cache is not enabled."))
    end
    allXfNodes=find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf", styles_xmlroot(get_workbook(ws)))
    let newid::Union{Int,Nothing}, first::Bool
        newid = nothing
        first = true
        isInDim(ws, get_dimension(ws), rng)
        for cellref in rng
            newid, first = process_uniform_core(f, ws, allXfNodes, cellref, atts, newid, first; kw...)
        end
        if first
            newid = -1
        end
        return newid
    end
end
function process_uniform_ncranges(f::Function, ws::Worksheet, ncrng::NonContiguousRange, atts::Vector{String}; kw...)::Int
    allXfNodes=find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf", styles_xmlroot(get_workbook(ws)))
    bounds = nc_bounds(ncrng)
    if length(ncrng) == 1
        single = true
    else
        single = false
    end
    dim = (get_dimension(ws))
    OK = dim.start.column_number <= bounds.start.column_number
    OK &= dim.stop.column_number >= bounds.stop.column_number
    OK &= dim.start.row_number <= bounds.start.row_number
    OK &= dim.stop.row_number >= bounds.stop.row_number
    if OK
        let newid::Union{Int,Nothing}, first::Bool
            newid = nothing
            first = true
            for r in ncrng.rng
                @assert r isa CellRef || r isa CellRange "Something wrong here"
                if r isa CellRef
                    if getcell(ws, r) isa EmptyCell
                        single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(r.name). Set the value first."))
                        continue
                    end
                    newid, first = process_uniform_core(f, ws, allXfNodes, r, atts, newid, first; kw...)
                else
                    for c in r
                        newid, first = process_uniform_core(f, ws, allXfNodes, c, atts, newid, first; kw...)
                    end
#                else
#                    throw(XLSXError("Something wrong here!"))
                end
            end
            if first
                newid = -1
            end
            return newid
        end
    else
        throw(XLSXError("Non-contiguous range $ncrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
end
function process_uniform_veccolon(f::Function, ws::Worksheet, row, col, atts::Vector{String}; kw...)
    dim = get_dimension(ws)
    @assert isnothing(row) || isnothing(col) "Something wrong here!"
    allXfNodes=find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf", styles_xmlroot(get_workbook(ws)))
    if isnothing(col)
        col = dim.start.column_number:dim.stop.column_number
    else
        row = dim.start.row_number:dim.stop.row_number
#    else
#        throw(XLSXError("Something wrong here!"))
    end
    isInDim(ws, dim, row, col)
    let newid::Union{Int,Nothing}, first::Bool
        newid = nothing
        first = true
        for a in row
            for b in col
                cellref = CellRef(a, b)
                if getcell(ws, cellref) isa EmptyCell
                    continue
                end
                newid, first = process_uniform_core(f, ws, allXfNodes, cellref, atts, newid, first; kw...)
            end
        end
        if first
            newid = -1
        end
        return newid
    end
end
function process_uniform_vecint(f::Function, ws::Worksheet, row, col, atts::Vector{String}; kw...)
    allXfNodes=find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf", styles_xmlroot(get_workbook(ws)))
    let newid::Union{Int,Nothing}, first::Bool
        dim = get_dimension(ws)
        newid = nothing
        first = true
        isInDim(ws, dim, row, col)
        for a in row, b in col
            cellref = CellRef(a, b)
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first = process_uniform_core(f, ws, allXfNodes, cellref, atts, newid, first; kw...)
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
function process_uniform_ncranges(ws::Worksheet, ncrng::NonContiguousRange)::Int
    bounds = nc_bounds(ncrng)
    if length(ncrng) == 1
        single = true
    else
        single = false
    end
    dim = (get_dimension(ws))
    OK = dim.start.column_number <= bounds.start.column_number
    OK &= dim.stop.column_number >= bounds.stop.column_number
    OK &= dim.start.row_number <= bounds.start.row_number
    OK &= dim.stop.row_number >= bounds.stop.row_number
    if OK
        let newid::Union{Int,Nothing}, first::Bool
            newid = nothing
            first = true
            for r in ncrng.rng
                @assert r isa CellRef || r isa CellRange "Something wrong here"
                if r isa CellRef
                    if getcell(ws, r) isa EmptyCell
                        single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(r.name). Set the value first."))
                        continue
                    end
                    newid, first = process_uniform_core(ws, r, newid, first)
                else
                    for c in r
                        newid, first = process_uniform_core(ws, c, newid, first)
                    end
#                else
#                    throw(XLSXError("Something wrong here!"))
                end
            end
            if first
                newid = -1
            end
            return newid
        end
    else
        throw(XLSXError("Non-contiguous range $ncrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
end
function process_colon(ws::Worksheet, row, col)
    dim = get_dimension(ws)
    @assert isnothing(row) || isnothing(col) "Something wrong here!"
    if isnothing(row) && isnothing(col)
        return setUniformStyle(ws, dim)
    elseif isnothing(col)
        rng = CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number))
    else
        rng = CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col)))
#    else
#        throw(XLSXError("Something wrong here!"))
    end

    return setUniformStyle(ws, rng)
end
function process_uniform_veccolon(ws::Worksheet, row, col)
    dim = get_dimension(ws)
    @assert isnothing(row) || isnothing(col) "Something wrong here!"
    if isnothing(col)
        col = dim.start.column_number:dim.stop.column_number
    else
        row = dim.start.row_number:dim.stop.row_number
#    else
#        throw(XLSXError("Something wrong here!"))
    end
    isInDim(ws, dim, row, col)
    let newid::Union{Int,Nothing}, first::Bool
        newid = nothing
        first = true
        for a in row
            for b in col
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
function process_uniform_vecint(ws::Worksheet, row, col)
    let newid::Union{Int,Nothing}, first::Bool
        dim = get_dimension(ws)
        newid = nothing
        first = true
        isInDim(ws, dim, row, col)
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
function process_uniform_core(f::Function, ws::Worksheet, allXfNodes::Vector{XML.Node}, cellref::CellRef, newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}; kw...) # setUniformAlignment is different
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
            cell.style = string(get_num_style_index(ws, allXfNodes, 0).id)
        end
        cell.style = string(update_template_xf(ws, allXfNodes, CellDataFormat(parse(Int, cell.style)), alignment_node).id)
    end
    return newid, first, alignment_node
end
function process_uniform_attribute(f::Function, ws::Worksheet, rng::CellRange; kw...)
    if !get_xlsxfile(ws).use_cache_for_sheet_data
        throw(XLSXError("Cannot set uniform attributes because cache is not enabled."))
    end
    allXfNodes=find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf", styles_xmlroot(get_workbook(ws)))
    let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
        newid = nothing
        first = true
        alignment_node = nothing
        isInDim(ws, get_dimension(ws), rng)
        for cellref in rng
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first, alignment_node = process_uniform_core(f, ws, allXfNodes, cellref, newid, first, alignment_node; kw...)
        end
        if first
            newid = -1
        end
        return newid
    end
end
function process_uniform_ncranges(f::Function, ws::Worksheet, ncrng::NonContiguousRange; kw...)::Int
    allXfNodes=find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf", styles_xmlroot(get_workbook(ws)))
    bounds = nc_bounds(ncrng)
    if length(ncrng) == 1
        single = true
    else
        single = false
    end
    dim = (get_dimension(ws))
    OK = dim.start.column_number <= bounds.start.column_number
    OK &= dim.stop.column_number >= bounds.stop.column_number
    OK &= dim.start.row_number <= bounds.start.row_number
    OK &= dim.stop.row_number >= bounds.stop.row_number
    if OK
        let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
            newid = nothing
            first = true
            alignment_node = nothing
            for r in ncrng.rng
                @assert r isa CellRef || r isa CellRange "Something wrong here"
                if r isa CellRef && getcell(ws, r) isa EmptyCell
                    single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(r.name). Set the value first."))
                    continue
                end
                if r isa CellRef
                    if getcell(ws, r) isa EmptyCell
                        single && throw(XLSXError("Cannot set format for an `EmptyCell`: $(r.name). Set the value first."))
                        continue
                    end
                    newid, first, alignment_node = process_uniform_core(f, ws, allXfNodes, r, newid, first, alignment_node; kw...)
                else
                    for c in r
                        newid, first, alignment_node = process_uniform_core(f, ws, allXfNodes, c, newid, first, alignment_node; kw...)
                    end
#                else
#                    throw(XLSXError("Something wrong here!"))
                end
            end
            if first
                newid = -1
            end
            return newid
        end
    else
        throw(XLSXError("Non-contiguous range $ncrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`."))
    end
end
function process_uniform_veccolon(f::Function, ws::Worksheet, row, col; kw...)
    dim = get_dimension(ws)
    if dim === nothing
        throw(XLSXError("No worksheet dimension found"))
    else
        @assert isnothing(row) || isnothing(col) "Something wrong here!"
        allXfNodes=find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf", styles_xmlroot(get_workbook(ws)))
        if isnothing(col)
            col = dim.start.column_number:dim.stop.column_number
        else
            row = dim.start.row_number:dim.stop.row_number
#        else
#            throw(XLSXError("Something wrong here!"))
        end
        isInDim(ws, dim, row, col)
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
                    newid, first, alignment_node = process_uniform_core(f, ws, allXfNodes, cellref, newid, first, alignment_node; kw...)
                end
            end
            if first
                newid = -1
            end
            return newid
        end
    end
end
function process_uniform_vecint(f::Function, ws::Worksheet, row, col; kw...)
    let newid::Union{Int,Nothing}, first::Bool, alignment_node::Union{XML.Node,Nothing}
        dim = get_dimension(ws)
        if dim === nothing
            throw(XLSXError("No worksheet dimension found"))
        end
        allXfNodes=find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":styleSheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":cellXfs/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":xf", styles_xmlroot(get_workbook(ws)))
        newid = nothing
        first = true
        alignment_node = nothing
        isInDim(ws, dim, row, col)
        for a in row, b in col
            cellref = CellRef(a, b)
            if getcell(ws, cellref) isa EmptyCell
                continue
            end
            newid, first, alignment_node = process_uniform_core(f, ws, allXfNodes, cellref, newid, first, alignment_node; kw...)
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
        throw(XLSXError("Invalid color specified: $s. Either give a valid color name (from Colors.jl) or an 8-digit rgb color in the form FFRRGGBB"))
    end
    return c
end
