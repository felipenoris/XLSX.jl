
const font_tags = ["b", "i", "u", "strike", "outline", "shadow", "condense", "extend", "sz", "color", "name", "scheme"]

copynode(o::XML.Node) = XML.Node(o.nodetype, o.tag, o.attributes, o.value, o.children)

function buildNode(tag::String, attributes::Dict{String, Union{Nothing, Dict{String, String}}}) :: XML.Node
    new_node = XML.Element(tag)
    for a in font_tags # Use this as a device to keep ordering constant for Excel
        if haskey(attributes, a)
            if isnothing(attributes[a])
                cnode=XML.Element(a)
            else
                cnode = XML.Node(XML.Element, a, XML.OrderedDict{String, String}(), nothing, nothing)
                for (k, v) in attributes[a]
                    cnode[k] = v
                end
            end
            push!(new_node, cnode)
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


"""
   setFont(sh::Worksheet, cr::String; kw...) -> String
   setFont(xf::XLSXFile,  cr::String, kw...) -> String

Set the font used by a single cell, a cell range or a column range in a worksheet or XLSXfile.

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

No validation of the font names specified is performed.
If you specify, for example, `name = badFont`, that value will be written to the XLSXfile.

The `color` attribute can only be defined as rgb values.
- The first two digits represent transparency (α). FF means fully opaque, while 00 means fully transparent.
- The next two digits give the red component.
- The next two digits give the green component.
- The next two digits give the blue component.
So, FF000000 means a fully opaque black color.

As an expedient to get fonts to work, the `scheme` attribute is simply dropped from
new font definitions.

Font attributes cannot be set for `EmptyCell`s. Set a cell value first.
If a cell range or column range includes any `EmptyCell`s, they will be
skipped and the font will be set for the remaining cells.

Examples:
```julia
julia> setFont(sheet, "A1"; bold=true, italic=true, size=12, color="FFFF0000", name="Arial")

julia> setFont(xfile, "Sheet1!A1"; bold=false, size=14, color="FFB3081F", name="Berlin Sans FB Demi")

julia> setFont(sheet, "A1:B7"; name="Aptos", under="double", strike=true)

julia> setFont(xfile, "Sheet1!A1:B7"; size=24, name="Berlin Sans FB Demi", bold=true, italic=true, color="FF0088FF", under="single")

julia> setFont(xfile, "Sheet1!A:B"; italic=true, color="FF8888FF", under="single")

```

The value returned is the style ID of the font applied to the cell(s).
This can be used to apply the same font to other cells or ranges.

"""
function setFont(ws::Worksheet, rng::CellRange; kw...) :: String
    first=true
    let newstyle
        for cell in rng
            if getcell(ws, cell) isa EmptyCell
                continue
            end
            if first # Define the font for the first cell in the range.
                newstyle = setFont(ws, cell; kw...)
                first=false
            else # Apply the font to the rest of the cells in the range.
                @assert !(cell isa EmptyCell) "Cannot set font for an EmptyCell: $(cell.name). Set the value first."
                cell = getcell(ws, cell)
                cell.style = newstyle
            end
        end
        return newstyle
    end
end
function setFont(ws::Worksheet, colrng::ColumnRange; kw...) :: String
    bounds = column_bounds(colrng)
    dim = (get_dimension(ws))

    left=bounds[begin]
    right=bounds[end]
    top=dim.start.row_number
    bottom=dim.stop.row_number

    OK = dim.start.column_number <= left
    OK &= dim.stop.column_number >= right
    OK &= dim.start.row_number <= top
    OK &= dim.stop.row_number >= bottom

    if OK
        rng = CellRange(top, left, bottom, right)
        return setFont(ws, rng; kw...)
    else
        error("Column range $colrng is out of bounds. Worksheet `$(ws.name)` has dimension `$dim`.")
    end
end
function setFont(ws::Worksheet, ref_or_rng::AbstractString; kw...)
    if is_valid_column_range(ref_or_rng)
        colrng=ColumnRange(ref_or_rng)
        newstyle=setFont(ws, colrng; kw...)
    elseif is_valid_cellrange(ref_or_rng)
        rng=CellRange(ref_or_rng)
        newstyle=setFont(ws, rng; kw...)
    elseif is_valid_cellname(ref_or_rng)
        newstyle = setFont(ws, CellRef(ref_or_rng); kw...)
    else
        error("Invalid cell reference or range: $ref_or_rng")
    end
    return newstyle
end
function setFont(xl::XLSXFile, sheetcell::String; kw...)
    if is_valid_sheet_column_range(sheetcell)
        sheetcolrng = SheetColumnRange(sheetcell)
        newstyle = setFont(xl[sheetcolrng.sheet], sheetcolrng.colrng; kw...)
    elseif is_valid_sheet_cellrange(sheetcell)
        sheetcellrng = SheetCellRange(sheetcell)
        newstyle = setFont(xl[sheetcellrng.sheet], sheetcellrng.rng; kw...)
    elseif is_valid_sheet_cellname(sheetcell)
        ref = SheetCellRef(sheetcell)
        @assert hassheet(xl, ref.sheet) "Sheet $(ref.sheet) not found."
        newstyle = setFont(getsheet(xl, ref.sheet), ref.cellref; kw...)
    else
        error("Invalid sheet cell reference: $sheetcell")
    end
    return newstyle
end
function setFont(sh::Worksheet, cellref::CellRef;
        bold::Union{Nothing, Bool}=nothing,
        italic::Union{Nothing, Bool}=nothing,
        under::Union{Nothing, String}=nothing,
        strike::Union{Nothing, Bool}=nothing,
        size::Union{Nothing, Int}=nothing,
        color::Union{Nothing, String}=nothing,
        name::Union{Nothing, String}=nothing
    ) :: Union{Nothing, String}

    wb = get_workbook(sh)
    cell = getcell(sh, cellref)

    @assert !(cell isa EmptyCell) "Cannot set font for an EmptyCell: $(cellref.name). Set the value first."
        
    if cell.style==""
        cell.style = string(get_num_style_index(sh::Worksheet, 0).id)
    end

    cell_style = styles_cell_xf(wb, parse(Int, cell.style))
    new_font_atts = Dict{String, Union{Dict{String, String}, Nothing}}()
 
    cell_font = getFont(wb, cell_style)
    old_font_atts = cell_font.font
    old_applyFont = cell_font.applyFont

    for a in font_tags
        if a == "b"
            if isnothing(bold) && haskey(old_font_atts,"b") || bold == true
                new_font_atts["b"] = nothing
            end
        elseif a == "i"
            if isnothing(italic) && haskey(old_font_atts,"i") || italic == true
                new_font_atts["i"] = nothing
            end
        elseif a == "u"
            @assert isnothing(under) || under ∈ ["none", "single", "double"] "Invalid value for under: $under. Must be one of: `none`, `single`, `double`."
            if isnothing(under) && haskey(old_font_atts,"u")
                new_font_atts["u"] = old_font_atts["u"]
            elseif !isnothing(under)
                if under == "single"
                    new_font_atts["u"] = nothing
                elseif under == "double"
                    new_font_atts["u"] = Dict("val" => "double")
                end
            end
        elseif a == "strike"
            if isnothing(strike) && haskey(old_font_atts,"strike") || strike == true
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
    
    new_fontid = styles_add_cell_font(wb, font_node)
    
    newstyle = string(update_template_xf(sh, CellDataFormat(parse(Int, cell.style)), ["fontId", "applyFont"], ["$new_fontid", "1"]).id)
    cell.style = newstyle
    return newstyle
end

"""
   getFont(sh::Worksheet, cr::String) -> Union{Nothing, CellFont}
   getFont(xf::XLSXFile, cr::String) -> Union{Nothing, CellFont}
   
Get the font used by a single cell at reference `cr` in a worksheet `sh` or XLSXfile `xf`.

Return a CellFont containing:
- `fontId`    : a 0-based index of the font in the workbook
- `font`      : a dictionary of font attributes: fontAttribute -> (attribute -> value)
- `applyFont` : "1" or "0", indicating whether or not the font is applied to the cell.
Return `nothing` if no cell font is found.

Examples:
```julia
julia> getFont(sh, "A1")

julia> getFont(xf, "Sheet1!A1")

```
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

"""
function getFont end

getFont(ws::Worksheet, cr::String) = getFont(ws, CellRef(cr))
function getFont(xl::XLSXFile, sheetcell::String) :: Union{Nothing, CellFont}
    ref = SheetCellRef(sheetcell)
    @assert hassheet(xl, ref.sheet) "Sheet $(ref.sheet) not found."
    return getFont(getsheet(xl, ref.sheet), ref.cellref)
end
function getFont(ws::Worksheet, cellref::CellRef) :: Union{Nothing, CellFont}
    wb = get_workbook(ws)
    cell = getcell(ws, cellref)

    if cell isa EmptyCell || cell.style==""
        return nothing
    end

    cell_style = styles_cell_xf(wb, parse(Int, cell.style))
    return getFont(wb, cell_style)
end
getDefaultFont(ws::Worksheet) = getFont(get_workbook(ws), styles_cell_xf(get_workbook(ws), 0))
function getFont(wb::Workbook, cell_style::XML.Node) :: Union{Nothing, CellFont}
    if haskey(cell_style, "fontId")
        fontid = cell_style["fontId"]
        applyfont= haskey(cell_style, "applyFont") ? cell_style["applyFont"] : "0"
        xroot = styles_xmlroot(wb)
        font_elements = find_all_nodes("/$SPREADSHEET_NAMESPACE_XPATH_ARG:styleSheet/$SPREADSHEET_NAMESPACE_XPATH_ARG:fonts", xroot)[begin]
        @assert parse(Int, font_elements["count"]) == length(XML.children(font_elements)) "Unexpected number of font definitions found : $(length(XML.children(font_elements))). Expected $(parse(Int, font_elements["count"]))"
        current_font = XML.children(font_elements)[parse(Int, fontid)+1] # Zero based!
        font_atts = Dict{String, Union{Dict{String, String}, Nothing}}()
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

# Only used in testing!
function styles_add_cell_font(wb::Workbook, attributes::Dict{String, Union{Dict{String, String}, Nothing}}) :: Int
    new_font = buildNode("font", attributes)
    return styles_add_cell_font(wb, new_font)
end
# Used by setFont()
function styles_add_cell_font(wb::Workbook, new_font::XML.Node) :: Int
    xroot = styles_xmlroot(wb)
    i, j = get_idces(xroot, "styleSheet", "fonts")
    existing_font_elements_count = length(XML.children(xroot[i][j]))
    @assert parse(Int, xroot[i][j]["count"]) == existing_font_elements_count "Wrong number of font elements found: $existing_font_elements_count. Expected $(parse(Int, xroot[i][j]["count"]))."

    # Check new_font doesn't duplicate any existing font. If yes, use that rather than create new.
    for (k, node) in enumerate(XML.children(xroot[i][j]))
        if XML.nodetype(node) == XML.nodetype(new_font) && XML.parse(XML.Node, XML.write(node)) == XML.parse(XML.Node, XML.write(new_font)) # XML.jl defines `Base.:(==)`
            #            if node == new_font # XML.jl defines `Base.:(==)`
            return k - 1 # CellDataFormat is zero-indexed
        end
    end

    push!(xroot[i][j], new_font)
    xroot[i][j]["count"] = string(existing_font_elements_count + 1)

    return existing_font_elements_count # turns out this is the new index (because it's zero-based)
end
