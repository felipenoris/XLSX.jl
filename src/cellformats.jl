
const font_tags = ["b", "i", "u", "strike", "outline", "shadow", "condense", "extend", "sz", "color", "name", "scheme"]
const border_tags = ["left", "right", "top", "bottom", "diagonal"]
const fill_tags = ["patternFill"]

copynode(o::XML.Node) = XML.parse(XML.Node, XML.write(o))[1]

function buildNode(tag::String, attributes::Dict{String,Union{Nothing,Dict{String,String}}})::XML.Node
    if tag == "font"
        attribute_tags = font_tags
    elseif tag == "border"
        attribute_tags = border_tags
    elseif tag == "fill"
        attribute_tags = fill_tags
    elseif tag == "alignment"
        attribute_tags = alignment_tags
    else
        error("Unknown tag: $tag")
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
                    @assert haskey(patternfill, "patternType") "No `patternType` attribute found."
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
    @assert length(attributes) == length(vals) "Attributes and values must be of the same length."
    for (a, v) in zip(attributes, vals)
        new_cell_xf[a] = v
    end
    return styles_add_cell_xf(ws.package.workbook, new_cell_xf)
end
function update_template_xf(ws::Worksheet, existing_style::CellDataFormat, alignment::XML.Node)::CellDataFormat
    old_cell_xf = styles_cell_xf(ws.package.workbook, Int(existing_style.id))
    new_cell_xf = copynode(old_cell_xf)
    new_cell_xf[1] = alignment
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
    @assert parse(Int, xroot[i][j]["count"]) == existing_elements_count "Wrong number of font elements found: $existing_elements_count. Expected $(parse(Int, xroot[i][j]["count"]))."

    # Check new_font doesn't duplicate any existing font. If yes, use that rather than create new.
    for (k, node) in enumerate(XML.children(xroot[i][j]))
        #if XML.nodetype(node) == XML.nodetype(new_att) && XML.parse(XML.Node, XML.write(node)) == XML.parse(XML.Node, XML.write(new_att)) # XML.jl defines `Base.:(==)`
        if XML.parse(XML.Node, XML.write(node))[1] == XML.parse(XML.Node, XML.write(new_att))[1] # XML.jl defines `Base.:(==)`
            return k - 1 # CellDataFormat is zero-indexed
        end
    end

    push!(xroot[i][j], new_att)
    xroot[i][j]["count"] = string(existing_elements_count + 1)

    return existing_elements_count # turns out this is the new index (because it's zero-based)
end
#is_non_contiguous_range(v) = occursin(",", string(v)) # Non-contiguous ranges are comma separated `SheetCellRef-like` strings
function process_sheetcell(f::Function, xl::XLSXFile, sheetcell::String; kw...)::Int
    if is_workbook_defined_name(xl, sheetcell)
        v = get_defined_name_value(xl.workbook, sheetcell)
        if is_defined_name_value_a_constant(v)
            error("Can only assign attributes to cells but `$(sheetcell)` is a constant: $(sheetcell)=$v.")
        elseif is_defined_name_value_a_reference(v)
            newid = process_ranges(f, xl, string(v); kw...)
        else
            error("Unexpected defined name value: $v.")
        end
    elseif is_valid_sheet_column_range(sheetcell)
        sheetcolrng = SheetColumnRange(sheetcell)
        newid = f(xl[sheetcolrng.sheet], sheetcolrng.colrng; kw...)
    elseif is_valid_sheet_cellrange(sheetcell)
        sheetcellrng = SheetCellRange(sheetcell)
        newid = f(xl[sheetcellrng.sheet], sheetcellrng.rng; kw...)
    elseif is_valid_sheet_cellname(sheetcell)
        ref = SheetCellRef(sheetcell)
        @assert hassheet(xl, ref.sheet) "Sheet $(ref.sheet) not found."
        newid = f(getsheet(xl, ref.sheet), ref.cellref; kw...)
     else
        error("Invalid sheet cell reference: $sheetcell")
    end
    return newid
end
function process_ranges(f::Function, ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int
    # Moved the tests for defined names to be first in case a name looks like a column name (e.g. "ID")
    if is_worksheet_defined_name(ws, ref_or_rng)
        v = get_defined_name_value(ws, ref_or_rng)
        if is_defined_name_value_a_constant(v)
            error("Can only assign attributes to cells but `$(ref_or_rng)` is a constant: $(ref_or_rng)=$v.")
        elseif is_defined_name_value_a_reference(v)
            wb = get_workbook(ws)
            newid = f(get_xlsxfile(wb), string(v); kw...)
        else
            error("Unexpected defined name value: $v.")
        end
    elseif is_workbook_defined_name(get_workbook(ws), ref_or_rng)
        wb = get_workbook(ws)
        v = get_defined_name_value(wb, ref_or_rng)
        if is_defined_name_value_a_constant(v)
            error("Can only assign attributes to cells but `$(ref_or_rng)` is a constant: $(ref_or_rng)=$v.")
        elseif is_defined_name_value_a_reference(v)
            if is_non_contiguous_range(v)
                _ = f.(Ref(get_xlsxfile(wb)), replace.(split(string(v), ","), "'" => "", "\$" => ""); kw...)
                newid = -1
            else
                newid = f(get_xlsxfile(wb), replace(string(v), "'" => "", "\$" => ""); kw...)
            end
        else
            error("Unexpected defined name value: $v.")
        end
    elseif is_valid_column_range(ref_or_rng)
        colrng = ColumnRange(ref_or_rng)
        newid = f(ws, colrng; kw...)
    elseif is_valid_cellrange(ref_or_rng)
        rng = CellRange(ref_or_rng)
        newid = f(ws, rng; kw...)
    elseif is_valid_cellname(ref_or_rng)
        newid = f(ws, CellRef(ref_or_rng); kw...)
    else
        error("Invalid cell reference or range: $ref_or_rng")
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
        error("Column range $colrng is out of bounds. Worksheet `$(ws.name)` only has dimension `$dim`.")
    end
end
function process_cellranges(f::Function, ws::Worksheet, rng::CellRange; kw...)::Int
    for cellref in rng
        if getcell(ws, cellref) isa EmptyCell
            continue
        end
        _ = f(ws, cellref; kw...)
    end
    return -1 # Each cell may have a different borderId so we can't return a single value.
end
function process_get_sheetcell(f::Function, xl::XLSXFile, sheetcell::String)
    ref = SheetCellRef(sheetcell)
    @assert hassheet(xl, ref.sheet) "Sheet $(ref.sheet) not found."
    return f(getsheet(xl, ref.sheet), ref.cellref)
end
function process_get_cellref(f::Function, ws::Worksheet, cellref::CellRef)
    wb = get_workbook(ws)
    cell = getcell(ws, cellref)

    if cell isa EmptyCell || cell.style == ""
        return nothing
    end

    cell_style = styles_cell_xf(wb, parse(Int, cell.style))
    return f(wb, cell_style)
end
function process_get_cellname(f::Function, ws::Worksheet, ref_or_rng::AbstractString)
    if is_workbook_defined_name(get_workbook(ws), ref_or_rng)
        wb = get_workbook(ws)
        v = get_defined_name_value(wb, ref_or_rng)
        if is_defined_name_value_a_constant(v) # Can these have fonts?
            error("Can only assign borderds to cells but `$(ref_or_rng)` is a constant: $(ref_or_rng)=$v.")
        elseif is_defined_name_value_a_reference(v)
            new_att = f(get_xlsxfile(wb), replace(string(v), "'" => ""))
        else
            error("Unexpected defined name value: $v.")
        end
    elseif is_valid_cellname(ref_or_rng)
        new_att = f(ws, CellRef(ref_or_rng))
    else
        error("Invalid cell reference or range: $ref_or_rng")
    end
    return new_att
end
function process_uniform_attribute(f::Function, ws::Worksheet, rng::CellRange, atts::Vector{String}; kw...)
    let newid
        first = true
        for cellref in rng
            cell = getcell(ws, cellref)
            if cell isa EmptyCell # Can't add a attribute to an empty cell.
                continue
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
        end
        if first
            newid = -1
        end
        return newid
    end
end
function process_uniform_attribute(f::Function, ws::Worksheet, rng::CellRange; kw...)
    let newid, alignment_node
        first = true
        for cellref in rng
            cell = getcell(ws, cellref)
            if cell isa EmptyCell # Can't add a attribute to an empty cell.
                continue
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
        end
        if first
            newid = -1
        end
        return newid
    end
end
#=function process_cell_style(f::Function, sh::Worksheet, cellref::CellRef; kw...)::Int
    wb = get_workbook(sh)
    cell = getcell(sh, cellref)

    @assert !(cell isa EmptyCell) "Cannot set attribute for an `EmptyCell`: $(cellref.name). Set the value first."

    if cell.style == ""
        cell.style = string(get_num_style_index(sh, 0).id)
    end

    return f(sh, wb, cell; kw...)
end
=#

"""
   setFont(sh::Worksheet, cr::String; kw...) -> Int
   setFont(xf::XLSXFile,  cr::String, kw...) -> Int

Set the font used by a single cell, a cell range, a column range or 
a named cell or named range in a worksheet or XLSXfile.

Font attributes are specified using keyword arguments:
- `bold::Bool = nothing`    : set to `true` to make the font bold.
- `italic::Bool = nothing`  : set to `true` to make the font italic.
- `under::String = nothing` : set to `single`, `double` or `none`.
- `strike::Bool = nothing`  : set to `true` to strike through the font.
- `size::Int = nothing`     : set the font size (0 < size < 410).
- `color::String = nothing` : set the font color using an 8-digit hexadecimal RGB value.
- `name::String = nothing`  : set the font name.

Only the attributes specified will be changed. If an attribute is not specified, the current
value will be retained. These are the only attributes supported currently.

No validation of the font names specified is performed. Available fonts will depend
on what your system has installed. If you specify, for example, `name = "badFont"`,
that value will be written to the XLSXfile.

As an expedient to get fonts to work, the `scheme` attribute is simply dropped from
new font definitions.

The `color` attribute can only be defined as rgb values.
- The first two digits represent transparency (α). FF is fully opaque, while 00 is fully transparent.
- The next two digits give the red component.
- The next two digits give the green component.
- The next two digits give the blue component.
So, FF000000 means a fully opaque black color.

Font attributes cannot be set for `EmptyCell`s. Set a cell value first.
If a cell range or column range includes any `EmptyCell`s, they will be
quietly skipped and the font will be set for the remaining cells.

For single cells, the value returned is the font ID of the font applied to the cell.
This can be used to apply the same font to other cells or ranges.

For cell ranges, column ranges and named ranges, the value returned is -1.

Examples:
```julia
julia> setFont(sheet, "A1"; bold=true, italic=true, size=12, name="Arial")          # Single cell

julia> setFont(xfile, "Sheet1!A1"; bold=false, size=14, color="FFB3081F")           # Single cell

julia> setFont(sheet, "A1:B7"; name="Aptos", under="double", strike=true)           # Cell range

julia> setFont(xfile, "Sheet1!A1:B7"; size=24, name="Berlin Sans FB Demi")          # Cell range

julia> setFont(sheet, "A:B"; italic=true, color="FF8888FF", under="single")         # Column range

julia> setFont(xfile, "Sheet1!A:B"; italic=true, color="FF8888FF", under="single")  # Column range

julia> setFont(sheet, "bigred"; size=48, color="FF00FF00")                          # Named cell or range

julia> setFont(xfile, "bigred"; size=48, color="FF00FF00")                          # Named cell or range

```
"""
function setFont end
setFont(ws::Worksheet, rng::CellRange; kw...)::Int = process_cellranges(setFont, ws, rng; kw...)
setFont(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setFont, ws, colrng; kw...)
setFont(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setFont, ws, ref_or_rng; kw...)
setFont(xl::XLSXFile, sheetcell::String; kw...)::Int = process_sheetcell(setFont, xl, sheetcell; kw...)
function setFont(sh::Worksheet, cellref::CellRef; 
        bold::Union{Nothing,Bool}=nothing,
        italic::Union{Nothing,Bool}=nothing,
        under::Union{Nothing,String}=nothing,
        strike::Union{Nothing,Bool}=nothing,
        size::Union{Nothing,Int}=nothing,
        color::Union{Nothing,String}=nothing,
        name::Union{Nothing,String}=nothing
    )::Int

    wb = get_workbook(sh)
    cell = getcell(sh, cellref)

    @assert !(cell isa EmptyCell) "Cannot set attribute for an `EmptyCell`: $(cellref.name). Set the value first."

    if cell.style == ""
        cell.style = string(get_num_style_index(sh, 0).id)
    end

    cell_style = styles_cell_xf(wb, parse(Int, cell.style))

    new_font_atts = Dict{String,Union{Dict{String,String},Nothing}}()
    cell_font = getFont(wb, cell_style)
    old_font_atts = cell_font.font
    old_applyFont = cell_font.applyFont

    for a in font_tags
        if a == "b"
            if isnothing(bold) && haskey(old_font_atts, "b") || bold == true
                new_font_atts["b"] = nothing
            end
        elseif a == "i"
            if isnothing(italic) && haskey(old_font_atts, "i") || italic == true
                new_font_atts["i"] = nothing
            end
        elseif a == "u"
            @assert isnothing(under) || under ∈ ["none", "single", "double"] "Invalid value for under: $under. Must be one of: `none`, `single`, `double`."
            if isnothing(under) && haskey(old_font_atts, "u")
                new_font_atts["u"] = old_font_atts["u"]
            elseif !isnothing(under)
                if under == "single"
                    new_font_atts["u"] = nothing
                elseif under == "double"
                    new_font_atts["u"] = Dict("val" => "double")
                end
            end
        elseif a == "strike"
            if isnothing(strike) && haskey(old_font_atts, "strike") || strike == true
                new_font_atts["strike"] = nothing
            end
        elseif a == "color"
            @assert isnothing(color) || occursin(r"^[0-9A-F]{8}$", color) "Invalid color value: $color. Must be an 8-digit hexadecimal RGB value."
            if isnothing(color) && haskey(old_font_atts, "color")
                new_font_atts["color"] = old_font_atts["color"]
            elseif !isnothing(color)
                new_font_atts["color"] = Dict("rgb" => color)
            end
        elseif a == "sz"
            @assert isnothing(size) || (size > 0 && size < 410) "Invalid size value: $size. Must be between 1 and 409."
            if isnothing(size) && haskey(old_font_atts, "sz")
                new_font_atts["sz"] = old_font_atts["sz"]
            elseif !isnothing(size)
                new_font_atts["sz"] = Dict("val" => string(size))
            end
        elseif a == "name"
            if isnothing(name) && haskey(old_font_atts, "name")
                new_font_atts["name"] = old_font_atts["name"]
            elseif !isnothing(name)
                new_font_atts["name"] = Dict("val" => name)
            end
        elseif a == "scheme" # drop this attribute
        elseif haskey(old_font_atts, a)
            new_font_atts[a] = old_font_atts[a]
        end
    end
    font_node = buildNode("font", new_font_atts)

    new_fontid = styles_add_cell_attribute(wb, font_node, "fonts")

    newstyle = string(update_template_xf(sh, CellDataFormat(parse(Int, cell.style)), ["fontId", "applyFont"], ["$new_fontid", "1"]).id)
    cell.style = newstyle
    return new_fontid
end

"""
   setUniformFont(sh::Worksheet, cr::String; kw...) -> Int
   setUniformFont(xf::XLSXFile,  cr::String, kw...) -> Int

Set the font used by a cell range, a column range or a named range in a 
worksheet or XLSXfile.

First, the font attributes of the first cell in the range (the top-left cell) are
updated according to the given `kw...` (using `setFont()`). The resultant font is 
then applied to each remaining cell in the range.

As a result, every cell in the range will have a uniform font setting.

This differs from `setFont()` which merges the attributes defined by `kw...` into 
the font definition used by each cell individually. For example, if you set the 
font size to 12 for a range of cells, but these cells all use different fonts names 
or colors, etc, `setFont()` will change the font size but leave the font name and 
color unchanged for each cell individually. 

In contrast, `setUniformFont()` will set the font size to 12 for the first cell, but 
will then apply all the font attributes from the updated first cell (ie. name, color, 
etc) to all the other cells in the range.

This can be more efficient when setting the same font for a large number of cells.

The value returned is the font ID of the font uniformly applied to the cells.
If all cells in the range are `EmptyCells` the returned value is -1.

For keyword definitions see `setFont()`@Ref.

Examples:
```julia
julia> setUniformFont(sheet, "A1:B7"; bold=true, italic=true, size=12, name="Arial")       # Cell range

julia> setUniformFont(xfile, "Sheet1!A1:B7"; size=24, name="Berlin Sans FB Demi")          # Cell range

julia> setUniformFont(sheet, "A:B"; italic=true, color="FF8888FF", under="single")         # Column range

julia> setUniformFont(xfile, "Sheet1!A:B"; italic=true, color="FF8888FF", under="single")  # Column range

julia> setUniformFont(sheet, "bigred"; size=48, color="FF00FF00")                          # Named range

julia> setUniformFont(xfile, "bigred"; size=48, color="FF00FF00")                          # Named range
```
"""
function setUniformFont end
setUniformFont(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setUniformFont, ws, colrng; kw...)
setUniformFont(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setUniformFont, xl, sheetcell; kw...)
setUniformFont(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setUniformFont, ws, ref_or_rng; kw...)
setUniformFont(ws::Worksheet, rng::CellRange; kw...)::Int = process_uniform_attribute(setFont, ws, rng, ["fontId", "applyFont"]; kw...)


"""
   getFont(sh::Worksheet, cr::String) -> Union{Nothing, CellFont}
   getFont(xf::XLSXFile, cr::String) -> Union{Nothing, CellFont}
   
Get the font used by a single cell at reference `cr` in a worksheet `sh` or XLSXfile `xf`.

Return a CellFont containing:
- `fontId`    : a 0-based index of the font in the workbook
- `font`      : a dictionary of font attributes: fontAttribute -> (attribute -> value)
- `applyFont` : "1" or "0", indicating whether or not the font is applied to the cell.

Return `nothing` if no cell font is found.

Excel uses several tags to define font properties in its XML structure.
Here's a list of some common tags and their purposes (thanks to Copilot!):
    b        : Indicates bold font.
    i        : Indicates italic font.
    u        : Specifies underlining (e.g., single, double).
    strike   : Indicates strikethrough.
    outline  : Specifies outline text.
    shadow   : Adds a shadow to the text.
    condense : Condenses the font spacing.
    extend   : Extends the font spacing.
    sz       : Sets the font size.
    color    : Sets the font color using RGB values).
    name     : Specifies the font name.
    family   : Defines the font family.
    scheme   : Specifies whether the font is part of the major or minor theme.

Excel defines colours in several ways. Get font will return the colour in any of these
e.g. `"color" => ("theme" => "1")`.

Examples:
```julia
julia> getFont(sh, "A1")

julia> getFont(xf, "Sheet1!A1")
```
"""
function getFont end
getFont(ws::Worksheet, cr::String) = process_get_cellname(getFont, ws, cr)
getFont(xl::XLSXFile, sheetcell::String)::Union{Nothing,CellFont} = process_get_sheetcell(getFont, xl, sheetcell)
getFont(ws::Worksheet, cellref::CellRef)::Union{Nothing,CellFont} = process_get_cellref(getFont, ws, cellref)
getDefaultFont(ws::Worksheet) = getFont(get_workbook(ws), styles_cell_xf(get_workbook(ws), 0))
function getFont(wb::Workbook, cell_style::XML.Node)::Union{Nothing,CellFont}
    if haskey(cell_style, "fontId")
        fontid = cell_style["fontId"]
        applyfont = haskey(cell_style, "applyFont") ? cell_style["applyFont"] : "0"
        xroot = styles_xmlroot(wb)
        font_elements = find_all_nodes("/$SPREADSHEET_NAMESPACE_XPATH_ARG:styleSheet/$SPREADSHEET_NAMESPACE_XPATH_ARG:fonts", xroot)[begin]
        @assert parse(Int, font_elements["count"]) == length(XML.children(font_elements)) "Unexpected number of font definitions found : $(length(XML.children(font_elements))). Expected $(parse(Int, font_elements["count"]))"
        current_font = XML.children(font_elements)[parse(Int, fontid)+1] # Zero based!
        font_atts = Dict{String,Union{Dict{String,String},Nothing}}()
        for c in XML.children(current_font)
            if isnothing(XML.attributes(c)) || length(XML.attributes(c)) == 0
                font_atts[XML.tag(c)] = nothing
            else
                @assert length(XML.attributes(c)) == 1 "Too many font attributes found for $(XML.tag(c)) Expected 1, found $(length(XML.attributes(c)))."
                for (k, v) in XML.attributes(c)
                    font_atts[XML.tag(c)] = Dict(k => v)
                end
            end
        end
        return CellFont(parse(Int, fontid), font_atts, applyfont)
    end

    return nothing
end

"""
   getBorder(sh::Worksheet, cr::String) -> Union{Nothing, CellBorder}
   getBorder(xf::XLSXFile, cr::String) -> Union{Nothing, CellBorder}
   
Get the borders used by a single cell at reference `cr` in a worksheet or XLSXfile.

Return a CellBorder object containing:
- `borderId`    : a 0-based index of the border in the workbook
- `border`      : a dictionary of border attributes: borderAttribute -> (attribute -> value)
- `applyBorder` : "1" or "0", indicating whether or not the border is applied to the cell.

Return `nothing` if no cell border is found.

A cell border has two attributes, `style` and `color`.

Excel defines border using a style and a color in its XML structure.
Here's a list of the available styles (thanks to Copilot!):
- none
- thin
- medium
- dashed
- dotted
- thick
- double
- hair
- mediumDashed
- dashDot
- mediumDashDot
- dashDotDot
- mediumDashDotDot
- slantDashDot

A border postion element (e.g. `top` or `left`) has a `style` attribute, but `color` is a child element.
The color element has one or two attributes (e.g. `rgb`) that define the color of the border.
While the key for the style element will always be `style`, the other keys, for the color element,
will vary depending on how the color is defined (e.g. `rgb`, `indexed`, `auto`, etc.).
Thus, for example, `"top" => Dict("style" => "thin", "rgb" => "FF000000")` would indicate a 
thin black border at the top of the cell while `"top" => Dict("style" => "thin", "auto" => "1")`
would indicate that the color is set automatically by Excel.

The `color` element can have the following attributes:
- auto     : Indicates that the color is automatically defined by Excel
- indexed  : Specifies the color using an indexed color value.
- rgb      : Specifies the rgb color using 8-digit hexadecimal format.
- theme    : Specifies the color using a theme color.
- tint     : Specifies the tint value to adjust the lightness or darkness of the color.

Tint can only be used in conjunction with the theme attribute to derive different shades of the theme color.
For example: <color theme="1" tint="-0.5"/>.

Only the `rgb` attribute can be used in `setBorder` to define a border color.

Examples:
```julia
julia> getBorder(sh, "A1")

julia> getBorder(xf, "Sheet1!A1")
```
"""
function getBorder end
getBorder(xl::XLSXFile, sheetcell::String)::Union{Nothing,CellBorder} = process_get_sheetcell(getBorder, xl, sheetcell)
getBorder(ws::Worksheet, cellref::CellRef)::Union{Nothing,CellBorder} = process_get_cellref(getBorder, ws, cellref)
getBorder(ws::Worksheet, cr::String) = process_get_cellname(getBorder, ws, cr)
getDefaultBorders(ws::Worksheet) = getBorder(get_workbook(ws), styles_cell_xf(get_workbook(ws), 0))
function getBorder(wb::Workbook, cell_style::XML.Node)::Union{Nothing,CellBorder}
    if haskey(cell_style, "borderId")
        borderid = cell_style["borderId"]
        applyborder = haskey(cell_style, "applyBorder") ? cell_style["applyBorder"] : "0"
        xroot = styles_xmlroot(wb)
        border_elements = find_all_nodes("/$SPREADSHEET_NAMESPACE_XPATH_ARG:styleSheet/$SPREADSHEET_NAMESPACE_XPATH_ARG:borders", xroot)[begin]
        @assert parse(Int, border_elements["count"]) == length(XML.children(border_elements)) "Unexpected number of font definitions found : $(length(XML.children(border_elements))). Expected $(parse(Int, border_elements["count"]))"
        current_border = XML.children(border_elements)[parse(Int, borderid)+1] # Zero based!
        border_atts = Dict{String,Union{Dict{String,String},Nothing}}()
        for side in XML.children(current_border)
            if isnothing(XML.attributes(side)) || length(XML.attributes(side)) == 0
                border_atts[XML.tag(side)] = nothing
            else
                @assert length(XML.attributes(side)) == 1 "Too many font attributes found for $(XML.tag(side)) Expected 1, found $(length(XML.attributes(side)))."
                for (k, v) in XML.attributes(side) # style is the only possible attribute of a side
                    border_atts[XML.tag(side)] = Dict(k => v)
                    for subc in XML.children(side) # color is a child of a border element
                        for (k, v) in XML.attributes(subc)
                            border_atts[XML.tag(side)][k] = v
                        end
                    end
                end
            end
        end
        return CellBorder(parse(Int, borderid), border_atts, applyborder)
    end

    return nothing
end

"""
   setBorder(sh::Worksheet, cr::String; kw...) -> Int}
   setBorder(xf::XLSXFile, cr::String; kw...) -> Int
   
Set the borders used used by a single cell, a cell range, a column range or 
a named cell or named range in a worksheet or XLSXfile.

Borders are independently defined for the keywords `left`, `right`, `top` 
and `bottom` for each of the sides of a cell using a vector of pairs 
`attribute => value`. Another keyword, `diagonal`, defines a line running 
bottom-left to top-right across the cell in the same way.

An additional keyword, `allsides`, is provided for convenience. It can be used 
in place of the four side keywords to apply the same border setting to all four 
sides at once. It cannot be used inconjunction with any of the side keywords but 
it can be used together with `diagonal`.

The two attributes that can be set for each keyword are `style` and `rgb`.

Allowed values for `style` are:

- none
- thin
- medium
- dashed
- dotted
- thick
- double
- hair
- mediumDashed
- dashDot
- mediumDashDot
- dashDotDot
- mediumDashDotDot
- slantDashDot

The color is set by specifying an 8-digit hexadecimal value for the `rgb` attribute.
No other color attributes can be applied.

Setting only one of the two attributes leaves the other attribute unchanged for that 
side's border. Omitting one of the keywords leaves the border definition for that
side unchanged, only updating the specified sides.

Border attributes cannot be set for `EmptyCell`s. Set a cell value first.
If a cell range or column range includes any `EmptyCell`s, they will be
quietly skipped and the border will be set for the remaining cells.

For single cells, the value returned is the border ID of the border applied to the cell.
This can be used to apply the same border to other cells or ranges.

For cell ranges, column ranges and named ranges, the value returned is -1.

Examples:
```julia
Julia> setBorder(sh, "D6"; allsides = ["style" => "thick"], diagonal = ["style" => "hair"])

Julia> setBorder(xf, "Sheet1!D4"; left     = ["style" => "dotted", "rgb" => "FF000FF0"],
                                  right    = ["style" => "medium", "rgb" => "FF765000"],
                                  top      = ["style" => "thick",  "rgb" => "FF230000"],
                                  bottom   = ["style" => "medium", "rgb" => "FF0000FF"],
                                  diagonal = ["style" => "dotted", "rgb" => "FF00D4D4"]
                                  )
"""
function setBorder end
setBorder(ws::Worksheet, rng::CellRange; kw...)::Int = process_cellranges(setBorder, ws, rng; kw...)
setBorder(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setBorder, ws, colrng; kw...)
setBorder(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setBorder, ws, ref_or_rng; kw...)
setBorder(xl::XLSXFile, sheetcell::String; kw...)::Int = process_sheetcell(setBorder, xl, sheetcell; kw...)
function setBorder(sh::Worksheet, cellref::CellRef;
    allsides::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    left::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    right::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    top::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    bottom::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    diagonal::Union{Nothing,Vector{Pair{String,String}}}=nothing
)::Int

    kwdict = Dict{String,Union{Dict{String,String},Nothing}}()
    kwdict["allsides"] = isnothing(allsides) ? nothing : Dict{String,String}(p for p in allsides)
    kwdict["left"] = isnothing(left) ? nothing : Dict{String,String}(p for p in left)
    kwdict["right"] = isnothing(right) ? nothing : Dict{String,String}(p for p in right)
    kwdict["top"] = isnothing(top) ? nothing : Dict{String,String}(p for p in top)
    kwdict["bottom"] = isnothing(bottom) ? nothing : Dict{String,String}(p for p in bottom)
    kwdict["diagonal"] = isnothing(diagonal) ? nothing : Dict{String,String}(p for p in diagonal)

    if !isnothing(allsides)
        @assert all(isnothing, [left, right, top, bottom]) "Keyword `allsides` is incompatible with any other keywords except `diagonal`."
        return setBorder(sh, cellref; left=allsides, right=allsides, top=allsides, bottom=allsides, diagonal=diagonal)
    end

    wb = get_workbook(sh)
    cell = getcell(sh, cellref)

    @assert !(cell isa EmptyCell) "Cannot set border for an `EmptyCell`: $(cellref.name). Set the value first."

    if cell.style == ""
        cell.style = string(get_num_style_index(sh, 0).id)
    end

    cell_style = styles_cell_xf(wb, parse(Int, cell.style))
    new_border_atts = Dict{String,Union{Dict{String,String},Nothing}}()

    cell_borders = getBorder(wb, cell_style)
    old_border_atts = cell_borders.border
    old_applyborder = cell_borders.applyBorder

    for a in ["left", "right", "top", "bottom", "diagonal"]
        new_border_atts[a] = Dict{String,String}()
        if isnothing(kwdict[a]) && haskey(old_border_atts, a)
            new_border_atts[a] = old_border_atts[a]
        elseif !isnothing(kwdict[a])
            if !haskey(kwdict[a], "style") && haskey(old_border_atts, a) && haskey(old_border_atts[a], "style")
                new_border_atts[a]["style"] = old_border_atts[a]["style"]
            elseif haskey(kwdict[a], "style")
                @assert kwdict[a]["style"] ∈ ["none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot"] "Invalid style: $v. Must be one of: `none`, `thin`, `medium`, `dashed`, `dotted`, `thick`, `double`, `hair`, `mediumDashed`, `dashDot`, `mediumDashDot`, `dashDotDot`, `mediumDashDotDot`, `slantDashDot`."
                new_border_atts[a]["style"] = kwdict[a]["style"]
            end
            if !haskey(kwdict[a], "rgb") && haskey(old_border_atts, a) && !isnothing(old_border_atts[a])
                for (k, v) in old_border_atts[a]
                    if k != "style"
                        new_border_atts[a][k] = v
                    end
                end
            elseif haskey(kwdict[a], "rgb")
                v = kwdict[a]["rgb"]
                @assert occursin(r"^[0-9A-F]{8}$", v) "Invalid color value: $v. Must be an 8-digit hexadecimal RGB value."
                new_border_atts[a]["rgb"] = v
            end
        end
    end

    border_node = buildNode("border", new_border_atts)

    new_borderid = styles_add_cell_attribute(wb, border_node, "borders")

    newstyle = string(update_template_xf(sh, CellDataFormat(parse(Int, cell.style)), ["borderId", "applyBorder"], ["$new_borderid", "1"]).id)
    cell.style = newstyle
    return new_borderid
end

"""
   setUniformBorder(sh::Worksheet, cr::String; kw...) -> Int
   setUniformBorder(xf::XLSXFile,  cr::String, kw...) -> Int

Set the border used by a cell range, a column range or a named range in a 
worksheet or XLSXfile.

First, the border attributes of the first cell in the range (the top-left cell) are
updated according to the given `kw...` (using `setBorder()`). The resultant border is 
then applied to each remaining cell in the range.

As a result, every cell in the range will have a uniform border setting.

This differs from `setBorder()` which merges the attributes defined by `kw...` into 
the border definition used by each cell individually. For example, if you set the 
border style to `thin` for a range of cells, but these cells all use different border  
colors, `setBorder()` will change the border style but leave the border color unchanged 
for each cell individually. 

In contrast, `setUniformBorder()` will set the border style to `thin` for the first cell,
but will then apply all the border attributes from the updated first cell (ie. both style 
and color) to all the other cells in the range.

This can be more efficient when setting the same border for a large number of cells.

The value returned is the border ID of the border uniformly applied to the cells.
If all cells in the range are `EmptyCells` the returned value is -1.

For keyword definitions see `setBorder()`@Ref.

Examples:
```julia
Julia> setUniformBorder(sh, "B2:D6"; allsides = ["style" => "thick"], diagonal = ["style" => "hair"])

Julia> setUniformBorder(xf, "Sheet1!A1:F20"; left= ["style" => "dotted", "rgb" => "FF000FF0"],
                                              right= ["style" => "medium", "rgb" => "FF765000"],
                                              top= ["style" => "thick", "rgb" => "FF230000"],
                                              bottom= ["style" => "medium", "rgb" => "FF0000FF"],
                                              diagonal= ["style" => "none"]
                                              )
```
"""
function setUniformBorder end
setUniformBorder(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setUniformBorder, ws, colrng; kw...)
setUniformBorder(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setUniformBorder, xl, sheetcell; kw...)
setUniformBorder(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setUniformBorder, ws, ref_or_rng; kw...)
setUniformBorder(ws::Worksheet, rng::CellRange; kw...)::Int = process_uniform_attribute(setBorder, ws, rng, ["borderId", "applyBorder"]; kw...)

"""
   getFill(sh::Worksheet, cr::String) -> Union{Nothing, CellFill}
   getFill(xf::XLSXFile, cr::String) -> Union{Nothing, CellFill}
   
Get the fill used by a single cell at reference `cr` in a worksheet or XLSXfile.

Return a CellFill object containing:
- `fillId`    : a 0-based index of the fill in the workbook
- `fill`      : a dictionary of fill attributes: borderAttribute -> (attribute -> value)
- `applyFill` : "1" or "0", indicating whether or not the fill is applied to the cell.

Return `nothing` if no cell fill is found.

The `fill` element in Excel's XML schema defines the pattern and color 
properties for cell fills. Here are the primary attributes and child elements 
in the `patternFill` element:
- `patternType` : Specifies the type of fill pattern (see below).
- `fgColor`     : Specifies the foreground color of the pattern.
- `bgColor`     : Specifies the background color of the pattern.

The child elements `fgColor` and `bgColor` can each have one or two attributes 
of their own. These color attributes are pushed in to the `DellFill.fill` Dict 
of attributes with either `fg` or `bg` prepended to their names to support later 
reconstruction of the xml element. Thus:
`"patternFill" => Dict("patternType" => "solid", "bgindexed" => "64", "fgtheme" => "0")`
indicates a solid fill with a foreground color defined by theme 0 (in Excel) and 
background color defined by an indexed value. In this case (solid fill), the 
background color will be ignored.

Here is a list of the `patternTypes` Excel uses (thanks to Copilot!):
- `none`
- `solid`
- `mediumGray`
- `darkGray`
- `lightGray`
- `darkHorizontal`
- `darkVertical`
- `darkDown`
- `darkUp`
- `darkGrid`
- `darkTrellis`
- `lightHorizontal`
- `lightVertical`
- `lightDown`
- `lightUp`
- `lightGrid`
- `lightTrellis`
- `gray125`
- `gray0625`

Definitions for neither `fgColor` (foreground color) nor `bgColor` (background color) 
are strictly necessary although certain pattern types are more visually meaningful 
when both are defined. These pattern types include those that create a pattern or grid 
effect, where the contrast between the foreground and background colors enhances the 
visual presentation.

These pattern types include `darkTrellis`, `darkGrid`, `darkHorizontal`, `darkVertical`,
`darkDown`, `darkUp`, `mediumGray`, `lightGray`, `lightTrellis`, `lightGrid`, `lightHorizontal`,
`lightVertical`, `lightDown`, `lightUp`, `gray125` and `gray0625`.

If fgColor (foreground color) and bgColor (background color) are specified when they aren't 
needed, they will simply be ignored by Excel, and the default appearance will be applied.

A fill has a pattern type attribute and two children fgColor and bgColor, each with 
one or two attributes of their own. These color attributes are pushed in to the Dict 
of attributes with either `fg` or `bg` prepended to their name to support later 
reconstruction of the xml element.

Examples:
```julia
julia> getFill(sh, "A1")

julia> getFill(xf, "Sheet1!A1")
```
"""
function getFill end
getFill(xl::XLSXFile, sheetcell::String)::Union{Nothing,CellFill} = process_get_sheetcell(getFill, xl, sheetcell)
getFill(ws::Worksheet, cellref::CellRef)::Union{Nothing,CellFill} = process_get_cellref(getFill, ws, cellref)
getFill(ws::Worksheet, cr::String) = process_get_cellname(getFill, ws, cr)
getDefaultFill(ws::Worksheet) = getFill(get_workbook(ws), styles_cell_xf(get_workbook(ws), 0))
function getFill(wb::Workbook, cell_style::XML.Node)::Union{Nothing,CellFill}
    if haskey(cell_style, "fillId")
        fillid = cell_style["fillId"]
        applyfill = haskey(cell_style, "applyFill") ? cell_style["applyFill"] : "0"
        xroot = styles_xmlroot(wb)
        fill_elements = find_all_nodes("/$SPREADSHEET_NAMESPACE_XPATH_ARG:styleSheet/$SPREADSHEET_NAMESPACE_XPATH_ARG:fills", xroot)[begin]
        @assert parse(Int, fill_elements["count"]) == length(XML.children(fill_elements)) "Unexpected number of font definitions found : $(length(XML.children(fill_elements))). Expected $(parse(Int, fill_elements["count"]))"
        current_fill = XML.children(fill_elements)[parse(Int, fillid)+1] # Zero based!
        fill_atts = Dict{String,Union{Dict{String,String},Nothing}}()
        for pattern in XML.children(current_fill)
            if isnothing(XML.attributes(pattern)) || length(XML.attributes(pattern)) == 0
                fill_atts[XML.tag(pattern)] = nothing
            else
                @assert length(XML.attributes(pattern)) == 1 "Too many font attributes found for $(XML.tag(pattern)) Expected 1, found $(length(XML.attributes(side)))."
                for (k, v) in XML.attributes(pattern) # patternType is the only possible attribute of a fill
                    fill_atts[XML.tag(pattern)] = Dict(k => v)
                    for subc in XML.children(pattern) # foreground and background colors are children of a patternFill element
                        @assert !isnothing(XML.children(subc)) && length(XML.attributes(subc)) > 0 "Too few children found for $(XML.tag(subc)) Expected 1, found 0."
                        @assert length(XML.children(subc)) < 3 "Too many children found for $(XML.tag(subc)) Expected < 3, found $(length(XML.attributes(subc)))."
                        tag = first(XML.tag(subc), 2)
                        for (k, v) in XML.attributes(subc)
                            fill_atts[XML.tag(pattern)][tag*k] = v
                        end
                    end
                end
            end
        end
        return CellFill(parse(Int, fillid), fill_atts, applyfill)
    end

    return nothing
end

"""
   setFill(sh::Worksheet, cr::String; kw...) -> Int}
   setFill(xf::XLSXFile, cr::String; kw...) -> Int
   
Set the fill used used by a single cell, a cell range, a column range or 
a named cell or named range in a worksheet or XLSXfile.

The following keywords are used to define a fill:

- pattern   : Sets the patternType for the fill.
- fgColor   : Sets the foreground color for the fill.
- bgColor   : Sets the background color for the fill.

Here is a list of the `patternTypes` Excel uses (thanks to Copilot!):
- `none`
- `solid`
- `mediumGray`
- `darkGray`
- `lightGray`
- `darkHorizontal`
- `darkVertical`
- `darkDown`
- `darkUp`
- `darkGrid`
- `darkTrellis`
- `lightHorizontal`
- `lightVertical`
- `lightDown`
- `lightUp`
- `lightGrid`
- `lightTrellis`
- `gray125`
- `gray0625`

The two colors are set by specifying an 8-digit hexadecimal value for the `fgColor`
and/or `bgColor`` keywords. No other color attributes can be applied.

Setting only one or two of the attributes leaves the other attribute(s) unchanged 
for that cell's fill.

Fill attributes cannot be set for `EmptyCell`s. Set a cell value first.
If a cell range or column range includes any `EmptyCell`s, they will be
quietly skipped and the fill will be set for the remaining cells.

For single cells, the value returned is the fill ID of the fill applied to the cell.
This can be used to apply the same fill to other cells or ranges.

For cell ranges, column ranges and named ranges, the value returned is -1.

Examples:
```julia
Julia> setFill(sh, "B2"; pattern="gray125", bgColor = "FF000000")

Julia> setFill(xf, "Sheet1!A1:F20"; pattern="none", fgColor = "88FF8800")
```
"""
function setFill end
setFill(ws::Worksheet, rng::CellRange; kw...)::Int = process_cellranges(setFill, ws, rng; kw...)
setFill(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setFill, ws, colrng; kw...)
setFill(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setFill, ws, ref_or_rng; kw...)
setFill(xl::XLSXFile, sheetcell::String; kw...)::Int = process_sheetcell(setFill, xl, sheetcell; kw...)
function setFill(sh::Worksheet, cellref::CellRef;
        pattern::Union{Nothing,String}=nothing,
        fgColor::Union{Nothing,String}=nothing,
        bgColor::Union{Nothing,String}=nothing,
    )::Int

    wb = get_workbook(sh)
    cell = getcell(sh, cellref)

    @assert !(cell isa EmptyCell) "Cannot set fill for an `EmptyCell`: $(cellref.name). Set the value first."

    if cell.style == ""
        cell.style = string(get_num_style_index(sh, 0).id)
    end

    cell_style = styles_cell_xf(wb, parse(Int, cell.style))
    
    new_fill_atts = Dict{String,Union{Dict{String,String},Nothing}}()
    patternFill = Dict{String,String}()

    cell_fill = getFill(wb, cell_style)
    old_fill_atts = cell_fill.fill["patternFill"]
    old_applyFill = cell_fill.applyFill

    for a in ["patternType", "fg", "bg"]
        if a == "patternType"
            if isnothing(pattern) && haskey(old_fill_atts, "patternType")
                patternFill["patternType"] = old_fill_atts["patternType"]
            elseif !isnothing(pattern)
                patternFill["patternType"] = pattern
            end
        elseif a == "fg"
            if isnothing(fgColor)
                for (k, v) in old_fill_atts
                    if occursin(r"^fg.*", k)
                        patternFill[k] = v
                    end
                end
            else
                @assert occursin(r"^[0-9A-F]{8}$", fgColor) "Invalid color value: $fgColor. Must be an 8-digit hexadecimal RGB value."
                patternFill["fgrgb"] = fgColor
            end
        elseif a == "bg"
            if isnothing(bgColor)
                for (k, v) in old_fill_atts
                    if occursin(r"^bg.*", k)
                        patternFill[k] = v
                    end
                end
            else
                @assert occursin(r"^[0-9A-F]{8}$", bgColor) "Invalid color value: $bgColor. Must be an 8-digit hexadecimal RGB value."
                patternFill["bgrgb"] = bgColor
            end
        end
    end
    new_fill_atts["patternFill"] = patternFill

    fill_node = buildNode("fill", new_fill_atts)

    new_fillid = styles_add_cell_attribute(wb, fill_node, "fills")

    newstyle = string(update_template_xf(sh, CellDataFormat(parse(Int, cell.style)), ["fillId", "applyFill"], ["$new_fillid", "1"]).id)
    cell.style = newstyle
    return new_fillid
end

"""
   setUniformFill(sh::Worksheet, cr::String; kw...) -> Int
   setUniformFill(xf::XLSXFile,  cr::String, kw...) -> Int

Set the fill used by a cell range, a column range or a named range in a 
worksheet or XLSXfile.

First, the fill attributes of the first cell in the range (the top-left cell) are
updated according to the given `kw...` (using `setFill()`). The resultant fill is 
then applied to each remaining cell in the range.

As a result, every cell in the range will have a uniform fill setting.

This differs from `setFill()` which merges the attributes defined by `kw...` into 
the fill definition used by each cell individually. For example, if you set the 
fill paternType to `darkGrid` for a range of cells, but these cells all use different fill  
colors, `setFill()` will change the fill patternType but leave the fill color unchanged 
for each cell individually. 

In contrast, `setUniformFill()` will set the fill patternType to `darkGrid` for the first cell,
but will then apply all the fill attributes from the updated first cell (ie. patternType 
and both foreground and background colors) to all the other cells in the range.

This can be more efficient when setting the same fill for a large number of cells.

The value returned is the fill ID of the fill uniformly applied to the cells.
If all cells in the range are `EmptyCells` the returned value is -1.

For keyword definitions see `setFill()`@Ref.

Examples:
```julia
Julia> setUniformFill(sh, "B2:D4"; pattern="gray125", bgColor = "FF000000")

Julia> setUniformFill(xf, "Sheet1!A1:F20"; pattern="none", fgColor = "88FF8800")
```
"""
function setUniformFill end
setUniformFill(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setUniformFill, ws, colrng; kw...)
setUniformFill(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setUniformFill, xl, sheetcell; kw...)
setUniformFill(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setUniformFill, ws, ref_or_rng; kw...)
setUniformFill(ws::Worksheet, rng::CellRange; kw...)::Int = process_uniform_attribute(setFill, ws, rng, ["fillId", "applyFill"]; kw...)

"""
   getAlignment(sh::Worksheet, cr::String) -> Union{Nothing, CellAlignment}
   getAlignment(xf::XLSXFile, cr::String) -> Union{Nothing, CellAlignment}
   
Get the alignment used by a single cell at reference `cr` in a worksheet or XLSXfile.

Return a CellAlignment object containing:
- `alignment`      : a dictionary of alignment attributes: alignmentAttribute -> (attribute -> value)
- `applyAlignment` : "1" or "0", indicating whether or not the alignment is applied to the cell.

Return `nothing` if no cell alignment is found.

The `alignment` element in Excel's XML schema defines horizontal and vertical
alignment of the cell and whether to wrap the cell contents:
- `horizontal`     : Specifies the horizontal alignment of the cell.
- `vertical`       : Specifies the vertical alignment of the cell.
- `wrapText`       : Specifies whether ("1") or not ("0") the cell content wraps
                     in the cell.

Excel supports the following values for the horizontal alignment:
- left             : Aligns the text to the left of the cell.
- center           : Centers the text within the cell.
- right            : Aligns the text to the right of the cell.
- fill             : Repeats the text to fill the entire width of the cell.
- justify          : Justifies the text, spacing it out so that it spans the entire width of the cell.
- centerContinuous : Centers the text across multiple cells (specifically the currrent cell and all 
                     empty cells to the right) as if the text were in a merged cell.
- distributed      : Distributes the text evenly across the width of the cell.

Excel supports the following values for the vertical alignment:
- top              : Aligns the text to the top of the cell.
- center           : Centers the text vertically within the cell.
- bottom           : Aligns the text to the bottom of the cell.
- justify          : Justifies the text vertically, spreading it out evenly within the cell.
- distributed      : Distributes the text evenly from top to bottom in the cell.

For single cells, the value returned is the style ID of the cell.

For cell ranges, column ranges and named ranges, the value returned is -1.

Examples:
```julia
julia> getAlignment(sh, "A1")

julia> getAlignment(xf, "Sheet1!A1")
```

"""
function getAlignment end
getAlignment(xl::XLSXFile, sheetcell::String)::Union{Nothing,CellAlignment} = process_get_sheetcell(getAlignment, xl, sheetcell)
getAlignment(ws::Worksheet, cellref::CellRef)::Union{Nothing,CellAlignment} = process_get_cellref(getAlignment, ws, cellref)
getAlignment(ws::Worksheet, cr::String) = process_get_cellname(getAlignment, ws, cr)
#getDefaultAlignment(ws::Worksheet) = getAlignment(get_workbook(ws), styles_cell_xf(get_workbook(ws), 0))
function getAlignment(wb::Workbook, cell_style::XML.Node)::Union{Nothing,CellAlignment}
    if length(XML.children(cell_style)) == 0 # `alignment` is a child node of the cell `xf`.
        return nothing
    end
    @assert length(XML.children(cell_style)) == 1 "Expected cell `xf` to have 1 child node, found $(length(XML.children(cell_style)))"
    @assert XML.tag(cell_style[1]) == "alignment" "Error cell alignment found but it has no attributes!"
    atts = Dict{String,String}()
    for (k, v) in XML.attributes(cell_style[1])
        atts[k]=v
    end
    alignment_atts = Dict{String,Union{Dict{String,String},Nothing}}()
    alignment_atts["alignment"] = atts
    @assert haskey(cell_style, "applyAlignment") "The `applyAlignment` attribute missing from cell `xf`."
    return CellAlignment(alignment_atts, cell_style["applyAlignment"])
end

"""
   setAlignment(sh::Worksheet, cr::String; kw...) -> Int}
   setAlignment(xf::XLSXFile, cr::String; kw...) -> Int}
   
Set the fill used used by a single cell, a cell range, a column range or 
a named cell or named range in a worksheet or XLSXfile.

The following keywords are used to define a fill:

- horizontal : Sets the horizontal alignment.
- vertical   : Sets the vertical alignment.
- wrapText   : Determines whether the cell content wraps within the cell.

Here are the possible values for the horizontal alignment:
- left             : Aligns the text to the left of the cell.
- center           : Centers the text within the cell.
- right            : Aligns the text to the right of the cell.
- fill             : Repeats the text to fill the entire width of the cell.
- justify          : Justifies the text, spacing it out so that it spans the entire 
                     width of the cell.
- centerContinuous : Centers the text across multiple cells (specifically the currrent 
                     cell and all empty cells to the right) as if the text were in 
                     a merged cell.
- distributed      : Distributes the text evenly across the width of the cell.

Here are the possible values for the vertical alignment:
- top              : Aligns the text to the top of the cell.
- center           : Centers the text vertically within the cell.
- bottom           : Aligns the text to the bottom of the cell.
- justify          : Justifies the text vertically, spreading it out evenly within the cell.
- distributed      : Distributes the text evenly from top to bottom in the cell.

The value of wrapText should be set to `true` or `false` depending if the content is to wrap or not.


Examples:
```julia
julia> setAlignment(s, "D18"; horizontal="center", wrapText=true)

julia> setAlignment(f, "sheet1!D18"; horizontal="right", vertical="top", wrapText=true)

```
"""
function setAlignment end
setAlignment(ws::Worksheet, rng::CellRange; kw...)::Int = process_cellranges(setAlignment, ws, rng; kw...)
setAlignment(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setAlignment, ws, colrng; kw...)
setAlignment(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setAlignment, ws, ref_or_rng; kw...)
setAlignment(xl::XLSXFile, sheetcell::String; kw...)::Int = process_sheetcell(setAlignment, xl, sheetcell; kw...)
function setAlignment(sh::Worksheet, cellref::CellRef; 
    horizontal::Union{Nothing,String}=nothing,
    vertical::Union{Nothing,String}=nothing,
    wrapText::Union{Nothing,Bool}=nothing
    )::Int

    wb = get_workbook(sh)
    cell = getcell(sh, cellref)

    @assert !(cell isa EmptyCell) "Cannot set fill for an `EmptyCell`: $(cellref.name). Set the value first."

    if cell.style == ""
        cell.style = string(get_num_style_index(sh, 0).id)
    end

    cell_style = styles_cell_xf(wb, parse(Int, cell.style))

    atts = XML.OrderedDict{String,String}()
    cell_alignment = getAlignment(wb, cell_style)
    old_alignment_atts = cell_alignment.alignment["alignment"]
    old_applyAlignment = cell_alignment.applyAlignment

    @assert isnothing(horizontal) || horizontal ∈ ["left", "center", "right", "fill", "justify", "centerContinuous", "distributed"] "Invalid horizontal alignment: $horizontal. Must be one of: `left`, `center`, `right`, `fill`, `justify`, `centerContinuous`, `distributed`."
    @assert isnothing(vertical) || vertical ∈ ["top", "center", "bottom", "justify", "distributed"] "Invalid vertical aligment: $vertical. Must be one of: `top`, `center`, `bottom`, `justify`, `distributed`."
    @assert isnothing(wrapText) || wrapText ∈ [true, false] "Invalid wrap option: $wrapText must be one of: `true`, `false`."

    if isnothing(horizontal) && haskey(old_alignment_atts, "horizontal")
        atts["horizontal"] = old_alignment_atts["horizontal"]
    elseif !isnothing(horizontal)
        atts["horizontal"] = horizontal
    end
    if isnothing(vertical) && haskey(old_alignment_atts, "vertical")
        atts["vertical"] = old_alignment_atts["vertical"]
    elseif !isnothing(vertical)
        atts["vertical"] = vertical
    end
    if isnothing(wrapText) && haskey(old_alignment_atts, "wrapText")
        atts["wrapText"] = old_alignment_atts["wrapText"]
    elseif !isnothing(wrapText)
        atts["wrapText"] = wrapText ? "1" : "0"
    end

    alignment_node = XML.Node(XML.Element, "alignment", atts, nothing, nothing)

    newstyle = string(update_template_xf(sh, CellDataFormat(parse(Int, cell.style)), alignment_node).id)
    cell.style = newstyle

    return parse(Int, newstyle)
end

"""
   setUniformAlignment(sh::Worksheet, cr::String; kw...) -> Int
   setUniformAlignment(xf::XLSXFile,  cr::String, kw...) -> Int

Set the alignment used by a cell range, a column range or a named range in a 
worksheet or XLSXfile.

First, the alignment attributes of the first cell in the range (the top-left cell) are
updated according to the given `kw...` (using `setFill()`). The resultant alignment is 
then applied to each remaining cell in the range.

As a result, every cell in the range will have a uniform alignment setting.

This differs from `setAlignment()` which merges the attributes defined by `kw...` into 
the alignment definition used by each cell individually. For example, if you set the 
horizontal alignment to `left` for a range of cells, but these cells all use different 
vertical alignment or wrap, `setAlignment()` will change the horizontal alignment but 
leave the vertical alignment and wrap unchanged for each cell individually. 

In contrast, `setUniformAlignment()` will set the horizontal alignment to `left` for the 
first cell, but will then apply all the alignment attributes from the updated first cell 
(ie. horizontal and vertical alignment and the wrap) to all the other cells in the range.

This can be more efficient when setting the same alignment for a large number of cells.

The value returned is the alignment ID of the alignment uniformly applied to the cells.
If all cells in the range are `EmptyCells`, the returned value is -1.

For keyword definitions see `setAlignment()`@Ref.

Examples:
```julia
Julia> setUniformAlignment(sh, "B2:D4"; horizontal="center", wrap = true)

Julia> setUniformAlignment(xf, "Sheet1!A1:F20"; horizontal="center", vertical="top")
```
"""
function setUniformFill end
setUniformAlignment(ws::Worksheet, colrng::ColumnRange; kw...)::Int = process_columnranges(setUniformAlignment, ws, colrng; kw...)
setUniformAlignment(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setUniformAlignment, xl, sheetcell; kw...)
setUniformAlignment(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setUniformAlignment, ws, ref_or_rng; kw...)
setUniformAlignment(ws::Worksheet, rng::CellRange; kw...)::Int = process_uniform_attribute(setAlignment, ws, rng; kw...)

"""
The <cols> element in an Excel worksheet's XML, which defines the columns and their properties. Here's a breakdown:

<col> elements within <cols> define individual columns.

Attributes like min, max, width, customWidth, and style specify properties for each column.

For example:

min="1" and max="1" indicate the column number (1).

width="2.7109375" sets the column width.

customWidth="1" suggests a custom width is applied.

style attribute might reference specific formatting styles.
"""

"""
this <row> element
Attributes:

r="3": Specifies the row number (3).

spans="1:11": Indicates the row spans from column 1 to column 11.

ht="6": Sets the height of the row to 6 points.

customHeight="1": Indicates that a custom height is applied to the row.

thickBot="1": Suggests that the bottom border of the row is thick.

x14ac:dyDescent="0.3": An attribute specific to certain versions of Excel, likely related to text descent for alignment purposes.
"""