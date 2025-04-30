const highlights::Dict{String,Dict{String,Dict{String, String}}} = Dict(
    "redfilltext" => Dict(
        "font" => Dict("color"=>"FF9C0006"),
        "fill" => Dict("pattern" => "solid", "bgColor"=>"FFFFC7CE")
    ),
    "yellowfilltext" => Dict(
        "font" => Dict("color"=>"FFA51E00"),
        "fill" => Dict("pattern" => "solid", "bgColor"=>"FF9C5700")
    ),
    "greenfilltext" => Dict(
        "font" => Dict("color"=>"FF006100"),
        "fill" => Dict("pattern" => "solid", "bgColor"=>"FFC6EFCE")
    ),
    "redfill" => Dict(
        "fill" => Dict("pattern" => "solid", "bgColor"=>"FFFFC7CE")
    ),
    "redtext" => Dict(
        "font" => Dict("color"=>"FF9C0006"),
    ),
    "redborder" => Dict(
        "border" => Dict("color"=>"FF9C0006", "style"=>"thin")
    )
) # for type = :Cell

const colorscales::Dict{String,XML.Node} = Dict(    # Defines the 12 standard, built-in Excel color scales for conditional formatting.
    "greenyellowred" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFFEB84"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "redyellowgreengreenyellowred" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFFEB84"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "greenwhitered" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "redwhitegreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "bluewhitered" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FF5A8AC6")
        )
    ),
    "redwhiteblue" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF5A8AC6"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "whitered" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFCFCFF")
        )
    ),
    "redwhite" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "greenwhite" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "whitegreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFCFCFF")
        )
    ),
    "greenyellow" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFFFEF9C"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "yellowgreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFFEF9C")
        )
    )
)
function add_cf_to_XML(ws, new_cf) # Add a new conditional formatting to the worksheet XML.
    sheetdoc = xmlroot(ws.package, "xl/worksheets/sheet" * string(ws.sheetId) * ".xml") # The <conditionalFormatting> blocks come after the <sheetData>
    k, l = get_idces(sheetdoc, "worksheet", "sheetData")
    len = length(sheetdoc[k])
    if l != len
        push!(sheetdoc[k], sheetdoc[k][end])
        if l + 1 < len
            for pos = len-1:-1:l+1
                sheetdoc[k][pos+1] = sheetdoc[k][pos]
            end
        end
        sheetdoc[k][l+1] = new_cf
    else
        push!(sheetdoc[k], new_cf)
    end
end

function Add_Cf_Dx(wb::Workbook, new_dx::XML.Node)::DxFormat
    # Check if the workbook already has a dxfs element. If not, add one.
    xroot = styles_xmlroot(wb)
    i, j = get_idces(xroot, "styleSheet", "dxfs")
    
    if isnothing(j) # No existing conditional formats so need to add a block. Push everything lower down one.
        k, l = get_idces(xroot, "styleSheet", "cellStyles")
        l += 1 # The dxfs block comes after the cellXfs block.
        len = length(xroot[k])
        i != k && throw(XLSXError("Some problem here!"))
        push!(xroot[k], xroot[k][end]) # duplicate last element then move everything else down one
        if l < len
            for pos = len-1:-1:l
                xroot[k][pos+1] = xroot[k][pos]
            end
        end
        xroot[k][l] = XML.Element("dxsf", count="0")
        j = l
        println(XML.write(xroot[i][j]))
    else
        existing_dxf_elements_count = length(XML.children(xroot[i][j]))

        if parse(Int, xroot[i][j]["count"]) != existing_dxf_elements_count
            throw(XLSXError("Wrong number of xf elements found: $existing_cellxf_elements_count. Expected $(parse(Int, xroot[i][j]["count"]))."))
        end
    end
    # Check new_dx doesn't duplicate any existing dxf. If yes, use that rather than create new.
    # Need to work around XML.jl issue # 33
    for (k, node) in enumerate(XML.children(xroot[i][j]))
        if XML.parse(XML.Node, XML.write(node))[1] == XML.parse(XML.Node, XML.write(new_dx))[1] # XML.jl defines `Base.:(==)`
            return DxFormat(k - 1) # CellDataFormat is zero-indexed
        end
    end
    existingdx=XML.children(xroot[i][j])
    dxfs = unlink(xroot[i][j], ("dxfs", "dxf")) # Create the new <dxfs> Node
    if length(existingdx) > 0
        for c in existingdx
            push!(dxfs, c) # Copy each existing <dxf> into the new <dxfs> Node
        end
    end
    push!(dxfs, new_dx)
 
    xroot[i][j] = dxfs # Update the worksheet with the new cols.

    xroot[i][j]["count"] = string(existing_dxf_elements_count + 1)

    return DxFormat(existing_dxf_elements_count) # turns out this is the new index (because it's zero-based)

end

function convertref(c)
    if !isnothing(c)
        if is_valid_cellname(c)
            c = abscell(CellRef(c))
        elseif is_valid_sheet_cellname(c)
            c = mkabs(SheetCellRef(c))
        end
    end
    return c
end


"""
Get the conditional formats for a worksheet.

# Arguments
- `ws::Worksheet`: The worksheet to get the conditional formats for.

Return a vector of pairs: CellRange => Vector{String}, where String is the 
type of the conditional format applies.


"""
function allCfs(ws::Worksheet)
    sheetdoc = xmlroot(ws.package, "xl/worksheets/sheet" * string(ws.sheetId) * ".xml") # find all the <conditionalFormatting> blocks in the worksheet's xml file
    return find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":worksheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":conditionalFormatting", sheetdoc)
end

function getConditionalFormats(ws::Worksheet)::Vector{Pair{CellRange,Vector{String}}}
    allcfnodes = allCfs(ws::Worksheet)
    allcfs = Vector{Pair{CellRange,Vector{String}}}()
    for (i, cf) in enumerate(allcfnodes)
        cf_types = Vector{String}()
        for child in XML.children(cf)
            if XML.tag(child) == "cfRule"
                push!(cf_types, child["type"])
            end
        end
        push!(allcfs, CellRange(cf["sqref"]) => cf_types)
    end
    return allcfs
end

"""
    setConditionalFormat(ws::Worksheet, cr::String, type::Symbol; kw...) -> ::Int}
    setConditionalFormat(xf::XLSXFile,  cr::String, type::Symbol; kw...) -> ::Int

    setConditionalFormat(ws::Worksheet, row, col,   type::Symbol; kw...) -> ::Int}

Add a new conditional format to a worksheet.

!!! warning "In Develpment..."

    This function is still in development and may not work as expected.
    It is not yet implemented for all types of conditional formats.

Valid options for `type` are `:colorScale` (others in develpment) and these 
determine which type of conditional formatting is being defined.

Keyword options differ according to the `type` specified

# type = :colorScale

Define a 2-color or 3-color color scale conditional format.

Use the keyword `colorscale` to choose one of the 12 built-in Excel colorscales:

- `"redyellowgreen"`: Red, Yellow, Green color scale.
- `"greenyellowred"`: Green, Yellow, Red color scale.
- `"redwhitegreen"` : Red, White, Green color scale.
- `"greenwhitered"` : Green, White, Red color scale.
- `"redwhiteblue"`  : Red, White, Blue color scale.
- `"bluewhitered"`  : Blue, White, Red color scale.
- `"redwhite"`      : Red, White color scale.
- `"whitered"`      : White, Red color scale.
- `"whitegreen"`    : White, Green color scale.
- `"greenwhite"`    : Green, White color scale.
- `"yellowgreen"`   : Yellow, Green color scale.
- `"greenyellow"`   : Green, Yellow color scale. (default)

Alternatively, you can define a custom color scale by omitting the `colorscale` keyword and 
instead using the following keywords:

- `min_type`: The type of the minimum value. Valid values are: `min`, `percentile`, `percent` or `num`.
- `min_val` : The value of the minimum. Omit if `min_type="min"`.
- `min_col` : The color of the minimum value.
- `mid_type`: Valid values are: `percentile`, `percent` or `num`. Omit for a 2-color scale.
- `mid_val` : The value of the middle value. Omit for a 2-color scale.
- `mid_col` : The color of the middle value. Omit for a 2-color scale.
- `max_type`: The type of the maximum value. Valid values are: `max`, `percentile`, `percent` or `num`.
- `max_val` : The value of the maximum value. Omit if `max_type="max"`.
- `max_col` : The color of the maximum value.

The keywords `min_val`, `mid_val`, and `max_val` can be either a cell reference (e.g. `A1`) 
or a number. If a cell reference is used, it will be converted to an absolute cell reference 
when writing to an XLSXFile.

Colors can be specified using an 8-digit hex string (e.g. `FF0000FF` for blue) or any named 
color from Colors.jl ([here](https://juliagraphics.github.io/Colors.jl/stable/namedcolors/)).

# Examples

```julia
julia> XLSX.setConditionalFormat(f["Sheet1"], "A1:F12", :colorScale) # Defaults to the `greenyellow` built-in scale.
0

julia> XLSX.setConditionalFormat(f["Sheet1"], "A13:C18", :colorScale; colorscale="whitered")
0

julia> XLSX.setConditionalFormat(f["Sheet1"], "D13:F18", :colorScale; colorscale="bluewhitered")
0

julia> XLSX.setConditionalFormat(f["Sheet1"], "A13:F22", :colorScale;
            min_type="num", 
            min_val="2",
            min_col="tomato",
            mid_type="num",
            mid_val="6", 
            mid_col="lawngreen",
            max_type="num",
            max_val="10",
            max_col="cadetblue"
        )
0

```

# type = :cell

Defines a conditional format based on the value of a cell.

Valid keywords are:
- `operator` : Defines the comparison to make.
- `value1`   : defines the first value to compare against. This can be a cell reference (e.g. `A1`) or a number.
- `value2`   : defines the second value to compare against. This can be a cell reference (e.g. `A1`) or a number.
- `dxStyle`  : Used to select one of the built-in Excel formats to apply
- `format`   : defines the numFmt to apply if opting for a custom format.
- `font`     : defines the font to apply if opting for a custom format.
- `border`   : defines the border to apply if opting for a custom format.
- `fill`     : defines the fill to apply if opting for a custom format.

The keyword `operator` defines the comparison to use in the conditiopnal formatting. 
If the condition is met, the format is applied. Valid options are:
- `greaterThan`  (>)
- `greaterEqual` (>=)
- `lessThan`     (<)
- `lessEqual`    (<=)
- `between`      (requires `value2`)
- `notBetween`   (requires `value2`)
- `equal`        (==)
- `notEqual`     (!=)

The comparison is made against the value in `value1` and, if `operator` is either 
`between` or `notBetween`, `value2` sets the other bound on the condition. If not specified, 
`value1` will be the arithmetic average of the (non-missing) cell values in the range if 
values are numeric. If the cell values are non-numeric, an error is thrown.

Formatting to be applied if the condition is met can be defined in two ways. Use the keyword
`dxStyle` to select one of the built-in Excel formats. Valid options are:
- `redfilltext`     (light red fill, dark red text) (default)
- `yellowfilltext`  (light yellow fill, dark yellow text)
- `greenfilltext`   (light green fill, dark green text)
- `redfill`         (light red fill)
- `redtext`         (dark red text)
- `redborder`       (dark red cell borders)

Alternatively, you can define a custom format by using the keywords `format`, `font`,
`border`, and `fill` which each take a vector of pairs of strings. The first string 
is the name of the attribute to set and the second is the value to set it to.  
Valid attributes for each keyword are:
- `format` : `format``
- `font`   : `color`, `bold`, `italic`, `under`, `strike`
- `fill`   : `pattern`, `bgColor`, `fgColor`
- `border` : `style`, `color`

Refer to [`setFormat()`](@ref), [`setFont)`](@ref), [`setFill`](@ref) and [`setBorder@ref) for
more details on the valid attributes and values.

!!! Note

    Excel limits the formatting attributes that can be set in a conditional format.
    It is not possible to set the size or name of a font and nor is it possible to set 
    any of the cell alignment attributes. Diagonal borders cannot be set either.

    Although it is not a limitation of Excel, this function sets all the border attributes 
    for each side of a cell to be the same.

If both `dxStyle` and custom formatting keywords are specified, `dxStyle` will be used 
and the custom formatting will be ignored.
If neither `dxStyle` nor custom formatting keywords are specified, the default 
is `dxStyle="redfilltext"`.

# Examples

```julia
julia> XLSX.setConditionalFormat(s, "B1:B5", :cell) # Defaults to `operator="greaterThan"`, `dxStyle`="redfilltext"` and `value1` set to the arithmetic agverage of cell values in `rng`.

julia> XLSX.setConditionalFormat(s, "B1:B5", :cell;
            operator="between",
            value1="2",
            value2="3",
            fill = ["pattern" => "none", "bgColor"=>"FFFFC7CE"],
            format = ["format"=>"0.00%"],
            font = ["color"=>"blue", "bold"=>"true"]
        )

julia> XLSX.setConditionalFormat(s, "B1:B5", :cell; 
            operator="greaterThan",
            value1="4",
            fill = ["pattern" => "none", "bgColor"=>"green"],
            format = ["format"=>"0.0"],
            font = ["color"=>"red", "italic"=>"true"]
        )

julia> XLSX.setConditionalFormat(s, "B1:B5", :cell;
            operator="lessThan",
            value1="2",
            fill = ["pattern" => "none", "bgColor"=>"yellow"],
            format = ["format"=>"0.0"],
            font = ["color"=>"green"],
            border = ["style"=>"thick", "color"=>"coral"]
        )

```

"""
function setConditionalFormat(f, r, type::Symbol; kw...)
    if type == :colorScale
        setCfColorScale(f, r; kw...)
    elseif type == :cell
        setCfCell(f, r; kw...)
#    elseif type == :iconSet
#        throw(XLSXError("Icon sets are not yet implemented."))
#    elseif type == :Cell
#        throw(XLSXError("Cell conditional formats are not yet implemented."))
else
        throw(XLSXError("Invalid conditional format type: $type. Valid options are: :colorScale, :cell"))
    end
end
function setConditionalFormat(f, r, c, type::Symbol; kw...)
    if type == :colorScale
        setCfColorScale(f, r, c; kw...)
    elseif type == :cell
        setCfCell(f, r, c; kw...)
#    elseif type == :iconSet
#        throw(XLSXError("Icon sets are not yet implemented."))
#    elseif type == :Cell
#        throw(XLSXError("Cell conditional formats are not yet implemented."))
    else
        throw(XLSXError("Invalid conditional format type: $type. Valid options are: :colorScale, :cell."))
    end
end
setCfColorScale(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfColorScale, ws, row, nothing; kw...)
setCfColorScale(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfColorScale, ws, nothing, col; kw...)
setCfColorScale(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfColorScale, ws, nothing, nothing; kw...)
setCfColorScale(ws::Worksheet, ::Colon; kw...) = process_colon(setCfColorScale, ws, nothing, nothing; kw...)
setCfColorScale(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfColorScale(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfColorScale(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfColorScale(ws, rng.rng; kw...)
setCfColorScale(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfColorScale(ws, rng.colrng; kw...)
setCfColorScale(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfColorScale(ws, rng.rowrng; kw...)
setCfColorScale(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfColorScale, ws, rng; kw...)
setCfColorScale(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfColorScale, ws, rng; kw...)
setCfColorScale(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfColorScale, xl, sheetcell; kw...)
setCfColorScale(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfColorScale, ws, ref_or_rng; kw...)
function setCfColorScale(ws::Worksheet, rng::CellRange;
    colorscale::Union{Nothing,String}=nothing,
    min_type::Union{Nothing,String}="min",
    min_val::Union{Nothing,String}=nothing,
    min_col::Union{Nothing,String}="FFF8696B",
    mid_type::Union{Nothing,String}=nothing,
    mid_val::Union{Nothing,String}=nothing,
    mid_col::Union{Nothing,String}=nothing,
    max_type::Union{Nothing,String}="max",
    max_val::Union{Nothing,String}=nothing,
    max_col::Union{Nothing,String}="FFFFEB84",
)::Int

    !issubset(rng, get_dimension(ws)) && throw(XLSXError("Range `$rng` goes outside worksheet dimension."))
    length(rng) <=1 && throw(XLSXError("Range `$rng` must have more than one cell."))

    old_cf = getConditionalFormats(ws)
    for cf in old_cf
        if intersects(cf.first, rng)
            throw(XLSXError("Range `$rng` intersects with existing conditional format range `$(cf.first)`."))
        end
    end

    new_cf = XML.Element("conditionalFormatting"; sqref=rng)
    if isnothing(colorscale)

        min_type in ["min", "percentile", "percent", "num"] || throw(XLSXError("Invalid min_type: $min_type. Valid options are: min, percentile, percent, num."))
        isnothing(min_val) || is_valid_cellname(min_val) || !isnothing(tryparse(Float64,min_val)) || throw(XLSXError("Invalid mid_type: $min_val. Valid options are a CellRef (e.g. `A1`) or a number."))
        isnothing(mid_type) || mid_type in ["percentile", "percent", "num"] || throw(XLSXError("Invalid mid_type: $mid_type. Valid options are: percentile, percent, num."))
        isnothing(mid_val) || is_valid_cellname(mid_val) || !isnothing(tryparse(Float64,mid_val)) || throw(XLSXError("Invalid mid_type: $mid_val. Valid options are a CellRef (e.g. `A1`) or a number."))
        max_type in ["max", "percentile", "percent", "num"] || throw(XLSXError("Invalid max_type: $max_type. Valid options are: max, percentile, percent, num."))
        isnothing(max_val) || is_valid_cellname(max_val) || !isnothing(tryparse(Float64,max_val)) || throw(XLSXError("Invalid mid_type: $max_val. Valid options are a CellRef (e.g. `A1`) or a number."))

        
        min_val = convertref(min_val)
        mid_val = convertref(mid_val)
        max_val = convertref(max_val)

        push!(new_cf, XML.h.cfRule(type="colorScale", priority="1",
            XML.h.colorScale(
                isnothing(min_val) ? XML.h.cfvo(type=min_type) : XML.h.cfvo(type=min_type, val=min_val),
                isnothing(mid_type) ? nothing : XML.h.cfvo(type=mid_type, val=mid_val),
                isnothing(max_val) ? XML.h.cfvo(type=max_type) : XML.h.cfvo(type=max_type, val=max_val),
                XML.h.color(rgb=get_color(min_col)),
                isnothing(mid_type) ? nothing : XML.h.color(rgb=get_color(mid_col)),
                XML.h.color(rgb=get_color(max_col))
            )
        ))

    else
        if !haskey(colorscales, colorscale)
            throw(XLSXError("Invalid color scale: $colorScale. Valid options are: $(keys(colorscales))."))
        end
        new_cf = XML.Element("conditionalFormatting"; sqref=rng)
        push!(new_cf, colorscales[colorscale])
    end

    add_cf_to_XML(ws, new_cf)    # Insert the new conditional formatting into the worksheet XML

    update_worksheets_xml!(get_xlsxfile(ws))

    return 0
end

setCfCell(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfCell, ws, row, nothing; kw...)
setCfCell(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfCell, ws, nothing, col; kw...)
setCfCell(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfCell, ws, nothing, nothing; kw...)
setCfCell(ws::Worksheet, ::Colon; kw...) = process_colon(setCfCell, ws, nothing, nothing; kw...)
setCfCell(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfCell(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfCell(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfCell(ws, rng.rng; kw...)
setCfCell(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfCell(ws, rng.colrng; kw...)
setCfCell(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfCell(ws, rng.rowrng; kw...)
setCfCell(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfCell, ws, rng; kw...)
setCfCell(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfCell, ws, rng; kw...)
setCfCell(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfCell, xl, sheetcell; kw...)
setCfCell(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfCell, ws, ref_or_rng; kw...)
function setCfCell(ws::Worksheet, rng::CellRange;
    operator::Union{Nothing,String}="greaterThan",
    value1::Union{Nothing,String}=nothing,
    value2::Union{Nothing,String}=nothing,
    dxStyle::Union{Nothing,String}=nothing,
    format::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    font::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    border::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    fill::Union{Nothing,Vector{Pair{String,String}}}=nothing
)::Int

    !issubset(rng, get_dimension(ws)) && throw(XLSXError("Range `$rng` goes outside worksheet dimension."))
    length(rng) <=1 && throw(XLSXError("Range `$rng` must have more than one cell."))

    old_cf = getConditionalFormats(ws)
    for cf in old_cf
        if cf.first != rng && intersects(cf.first, rng)
            throw(XLSXError("Range `$rng` intersects with existing conditional format range `$(cf.first)` but is not the same. Must be the same as the existing range or entirely separate."))
        end
    end
    wb=get_workbook(ws)
    if isnothing(dxStyle)
        if all(isnothing.([border, fill, font, format]))
            dx=highlights["redfilltext"]
        else
            dx = Dict{String,Dict{String, String}}()
            for att in ["font" => font, "fill" => fill, "border" => border, "format" => format]
                if !isnothing(last(att))
                    dxx = Dict{String, String}()
                    for i in last(att)
                        push!(dxx, first(i) => last(i))
                    end
                    push!(dx, first(att) => dxx)
                end
            end
        end
    elseif haskey(highlights, dxStyle)
        dx = highlights[dxStyle]
    else
        throw(XLSXError("Invalid dxStyle: $dxStyle. Valid options are: $(keys(highlights))."))
    end
    new_dx = XML.Element("dxf")
    for k in ["font", "format", "fill", "border"] # Order is important to Excel.
        if haskey(dx, k)
            v = dx[k]
            if k=="fill"
                if !isnothing(v)
                    filldx=XML.Element("fill")
                    patterndx=XML.Element("patternFill")
                    for (y, z) in v
                        if y in ["fgColor", "bgColor"]
                            push!(patterndx, XML.Element(y, rgb=get_color(z)))
                        elseif y == "pattern" && z != "none"
                            patterndx["patternType"] = z
                        end
                    end
                    push!(filldx, patterndx)
                end
                push!(new_dx, filldx)
            elseif k=="font"
                if !isnothing(v)
                    fontdx=XML.Element("font")
                    for (y, z) in v
                        if y=="color"
                            push!(fontdx, XML.Element(y, rgb=get_color(z)))
                        elseif y == "bold"
                            z=="true" && push!(fontdx, XML.Element("b", val="0"))
                        elseif y == "italic"
                            z=="true" && push!(fontdx, XML.Element("i", val="0"))
                        elseif y == "under"
                            z != "none" && push!(fontdx, XML.Element("u"; val="v"))
                        elseif y == "strike"
                            strike=="true" && push!(fontdx, XML.Element(y; val="0"))
                        end
                    end
                end
                push!(new_dx, fontdx)
            elseif k=="border"
                if !isnothing(v)
                    borderdx=XML.Element("border")
                    cdx = haskey(v, "color") ? XML.Element("color", rgb=get_color(v["color"])) : nothing
                    sdx = haskey(v, "style") ? v["style"] : nothing
                    leftdx = XML.Element("left")
                    rightdx = XML.Element("right")
                    topdx = XML.Element("top")
                    bottomdx = XML.Element("bottom")
                    if !isnothing(sdx)
                        leftdx["style"]=sdx
                        rightdx["style"]=sdx
                        topdx["style"]=sdx
                        bottomdx["style"]=sdx
                    end
                    if !isnothing(cdx)
                        push!(leftdx, cdx)
                        push!(rightdx, cdx)
                        push!(topdx, cdx)
                        push!(bottomdx, cdx)
                    end
                end
                push!(borderdx, leftdx)
                push!(borderdx, rightdx)
                push!(borderdx, topdx)
                push!(borderdx, bottomdx)
                push!(new_dx, borderdx)
            elseif k=="format"
                if !isnothing(v)
                    if haskey(v, "format")
                        fmtCode = v["format"]
                        new_formatId = get_new_formatId(wb, fmtCode)
                        new_fmtCode = styles_numFmt_formatCode(wb, new_formatId)
                        fmtdx=XML.Element("numFmt"; numFmtId=string(new_formatId), formatCode=new_fmtCode)
                        push!(new_dx, fmtdx)
                    end
                end
            end
        end
        
    end

    dxid = Add_Cf_Dx(get_workbook(ws), new_dx)
    if isnothing(value1)
        value1 = all(ismissing.(ws[rng])) ? nothing : string(sum(skipmissing(ws[rng]))/count(!ismissing, ws[rng]))
    end
    cfx = XML.Element("cfRule"; type="cellIs", dxfId=Int(dxid.id), priority="1", operator=operator)
    if !isnothing(value1)
        push!(cfx, XML.Element("formula", XML.Text(value1)))
    end
    if !isnothing(value2)
        push!(cfx, XML.Element("formula", XML.Text(value2)))
    end

    allcfs = filter(x->x["sqref"]==string(rng), allCfs(ws)) # Match range with existing conditional formatting blocks.
    if length(allcfs) == 0                                  # No existing conditional formatting blocks for this range so create a new one.
        new_cf = XML.Element("conditionalFormatting"; sqref=rng)
    else                             # Existing conditional formatting block found for this range so add new rule to that.
        children=XML.children(allcfs[1])
        cfx["priority"] = string(maximum([parse(Int, c["priority"]) for c in children])+1)
        new_cf = allcfs[1]
    end


    push!(new_cf, cfx)

    add_cf_to_XML(ws, new_cf) # Add the new conditional formatting to the worksheet XML.

    update_worksheets_xml!(get_xlsxfile(ws))

    return 0
end