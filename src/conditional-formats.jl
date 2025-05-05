const needsValue2::Vector{String} = ["between", "notBetween"]
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

const timeperiods::Dict{String,String} = Dict(
    "last7Days" => "AND(TODAY()-FLOOR(__CR__,1)<=6,FLOOR(__CR__,1)<=TODAY())",
    "yesterday" => "FLOOR(__CR__,1)=TODAY()-1",
    "today"     => "FLOOR(__CR__,1)=TODAY()",
    "tomorrow"  => "FLOOR(__CR__,1)=TODAY()+1",
    "lastWeek"  => "AND(TODAY()-ROUNDDOWN(__CR__,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(__CR__,0)<(WEEKDAY(TODAY())+7))",
    "thisWeek"  => "AND(TODAY()-ROUNDDOWN(__CR__,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(__CR__,0)-TODAY()<=7-WEEKDAY(TODAY()))",
    "nextWeek"  => "AND(ROUNDDOWN(__CR__,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(__CR__,0)-TODAY()<(15-WEEKDAY(TODAY())))",
    "lastMonth" => "AND(MONTH(__CR__)=MONTH(EDATE(TODAY(),0-1)),YEAR(__CR__)=YEAR(EDATE(TODAY(),0-1)))",
    "thisMonth" => "AND(MONTH(__CR__)=MONTH(TODAY()),YEAR(__CR__)=YEAR(TODAY()))",
    "nextMonth" => "AND(MONTH(__CR__)=MONTH(EDATE(TODAY(),0+1)),YEAR(__CR__)=YEAR(EDATE(TODAY(),0+1)))"
)

function get_dx(dxStyle::Union{Nothing, String}, border::Union{Nothing, Vector{Pair{String, String}}}, fill::Union{Nothing, Vector{Pair{String, String}}}, font::Union{Nothing, Vector{Pair{String, String}}}, format::Union{Nothing, Vector{Pair{String, String}}})::Dict{String,Dict{String, String}}
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
    return dx
end
function get_new_dx(wb::Workbook, dx::Dict{String,Dict{String, String}})::XML.Node
    new_dx = XML.Element("dxf")
    for k in ["font", "format", "fill", "border"] # Order is important to Excel.
        if haskey(dx, k)
            v = dx[k]
            if k=="fill"
                if !isnothing(v)
                    filldx=XML.Element("fill")
                    patterndx=XML.Element("patternFill")
                    for (y, z) in v
                        y in ["pattern", "bgColor", "fgColor"] || throw(XLSXError("Invalid fill attribute: $k. Valid options are: `pattern`, `bgColor`, `fgColor`."))
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
                        y in ["color", "bold", "italic", "under", "strike"] || throw(XLSXError("Invalid font attribute: $y. Valid options are: `color`, `bold`, `italic`, `under`, `strike`.")) 
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
                    all([y in ["color", "style"] for y in keys(v)]) || throw(XLSXError("Invalid border attribute. Valid options are: `color`, `style`."))
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
    return new_dx
end

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

function update_worksheet_cfx!(allcfs, cfx, ws, rng)
    matchcfs = filter(x->x["sqref"]==string(rng), allcfs)   # Match range with existing conditional formatting blocks.
    l = length(matchcfs)
    if l == 0                                               # No existing conditional formatting blocks for this range so create a new one.
        new_cf = XML.Element("conditionalFormatting"; sqref=rng)
        push!(new_cf, cfx)
        add_cf_to_XML(ws, new_cf)                           # Add the new conditional formatting block to the worksheet XML.
    elseif l==1                                             # Existing conditional formatting block found for this range so add new rule to that block.
        push!(matchcfs[1], cfx)
    else
        throw(XLSXError("Too many conditional formatting blocks for range `$rng`. Must be one or none, found `$l`."))
    end
    update_worksheets_xml!(get_xlsxfile(ws))
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

function allCfs(ws::Worksheet)
    sheetdoc = xmlroot(ws.package, "xl/worksheets/sheet" * string(ws.sheetId) * ".xml") # find all the <conditionalFormatting> blocks in the worksheet's xml file
    return find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":worksheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":conditionalFormatting", sheetdoc)
end


"""
Get the conditional formats for a worksheet.

# Arguments
- `ws::Worksheet`: The worksheet to get the conditional formats for.

Return a vector of pairs: CellRange => NamedTuple{type::String, priority::Int}}.


"""
getConditionalFormats(ws::Worksheet) = getConditionalFormats(allCfs(ws))
function getConditionalFormats(allcfnodes::Vector{XML.Node})::Vector{Pair{CellRange,NamedTuple}}
    allcfs = Vector{Pair{CellRange,NamedTuple}}()
    for cf in allcfnodes
        for child in XML.children(cf)
            if XML.tag(child) == "cfRule"
                push!(allcfs, CellRange(cf["sqref"]) => (; type=child["type"], priority=parse(Int, child["priority"])))
            end
        end
    end
    return allcfs
end

"""
    setConditionalFormat(ws::Worksheet, cr::String, type::Symbol; kw...) -> ::Int
    setConditionalFormat(xf::XLSXFile,  cr::String, type::Symbol; kw...) -> ::Int

    setConditionalFormat(ws::Worksheet, rows, cols,   type::Symbol; kw...) -> ::Int

Add a new conditional format to a cell range, row range or column range in a 
worksheet or `XLSXFile`.  Alternatively, ranges can be specified by giving rows 
and columns can be specified separately.

!!! warning "In Develpment..."

    This function is still in development and may not work as expected.
    It is not yet implemented for all types of conditional formats.

Valid options for `type` are (others in develpment):
- `:colorScale`
- `:cellIs`
- `:top10`
- `:aboveAverage`
- `:containsText`
- `:notContainsText`
- `:beginsWith`
- `:endsWith`
- `:timePeriod`
- `:containsErrors`
- `:notContainsErrors`
- `:containsBlanks`
- `:notContainsBlanks`
- `:uniqueValues`
- `:duplicateValues`

The `type` argument determines which type of conditional formatting is being defined.
Keyword options differ according to the `type` specified, as set out below.

!!! note "Ovrlaying conditional formats"
    
    Conditional formats are applied to a cell range and it is possible to apply multiple 
    conditional formats to the same range or to overlapping ranges. Each format is applied 
    in turn to each cell in priority order which, here, is the order in which they are 
    created. Different format options may complement or override each other and the 
    finished appearance will be the resuilt of all formats overlaying each other.

    It is possible to terminate the sequential application of conditional formats to a 
    cell if the condition related to any format is met. This is achieved by setting the 
    keyword option `stopIfTrue=true` in the relevant conditional format.

    While the `stopIfTrue` keyword is available for most conditional formats, it is not 
    available for `:colorScale` conditional formats.

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

The keywords `min_val`, `mid_val`, and `max_val` can be either a cell reference (e.g. `"A1"`) 
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

# type = :cellIs

Defines a conditional format based on the value of each cell in a range.

Valid keywords are:
- `operator`   : Defines the comparison to make.
- `value`     : defines the first value to compare against. This can be a cell reference (e.g. `"A1"`) or a number.
- `value2`     : defines the second value to compare against. This can be a cell reference (e.g. `"A1"`) or a number.
- `stopIfTrue` : Stops evaluating the conditional formats for this cell if this one is true.
- `dxStyle`    : Used optionally to select one of the built-in Excel formats to apply
- `format`     : defines the numFmt to apply if opting for a custom format.
- `font`       : defines the font to apply if opting for a custom format.
- `border`     : defines the border to apply if opting for a custom format.
- `fill`       : defines the fill to apply if opting for a custom format.

All keywords are defined using Strings (e.g. `value = "2"` or `value = "A2"`).

The keyword `operator` defines the comparison to use in the conditional formatting. 
If the condition is met, the format is applied.
Valid options are:

- `greaterThan`     (cell >  `value`)
- `greaterEqual`    (cell >= `value`)
- `lessThan`        (cell <  `value`)
- `lessEqual`       (cell <= `value`)
- `equal`           (cell == `value`)
- `notEqual`        (cell != `value`)
- `between`         (cell between `value` and `value2`)
- `notBetween`      (cell not between `value` and `value2`)

If not specified (when required), `value` will be the arithmetic average of the 
(non-missing) cell values in the range if values are numeric. If the cell values 
are non-numeric, an error is thrown.

Formatting to be applied if the condition is met can be defined in two ways. Use the keyword
`dxStyle` to select one of the built-in Excel formats. 
Valid options are:

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

Refer to [`setFormat()`](@ref), [`setFont()`](@ref), [`setFill()`](@ref) and [`setBorder()](@ref) for
more details on the valid attributes and values.

!!! note

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
julia> XLSX.setConditionalFormat(s, "B1:B5", :cell) # Defaults to `operator="greaterThan"`, `dxStyle`="redfilltext"` and `value` set to the arithmetic agverage of cell values in `rng`.

julia> XLSX.setConditionalFormat(s, "B1:B5", :cell;
            operator="between",
            value="2",
            value2="3",
            fill = ["pattern" => "none", "bgColor"=>"FFFFC7CE"],
            format = ["format"=>"0.00%"],
            font = ["color"=>"blue", "bold"=>"true"]
        )

julia> XLSX.setConditionalFormat(s, "B1:B5", :cell; 
            operator="greaterThan",
            value="4",
            fill = ["pattern" => "none", "bgColor"=>"green"],
            format = ["format"=>"0.0"],
            font = ["color"=>"red", "italic"=>"true"]
        )

julia> XLSX.setConditionalFormat(s, "B1:B5", :cell;
            operator="lessThan",
            value="2",
            fill = ["pattern" => "none", "bgColor"=>"yellow"],
            format = ["format"=>"0.0"],
            font = ["color"=>"green"],
            border = ["style"=>"thick", "color"=>"coral"]
        )

```

# type = :top10

This conditional format can be used to highlight cells in the top (bottom) n within the 
range or in the top (bottom) n% (ie in the top 5 or in the top 5% of values in the range). 

The available keywords are:

- `operator`   : Defines the comparison to make.
- `value`     : Gives the for comparison or a cell reference (e.g. `"A1"`).
- `stopIfTrue` : Stops evaluating the conditional formats if this one is true.
- `dxStyle`    : Used optionally to select one of the built-in Excel formats to apply.
- `format`     : defines the numFmt to apply if opting for a custom format.
- `font`       : defines the font to apply if opting for a custom format.
- `border`     : defines the border to apply if opting for a custom format.
- `fill`       : defines the fill to apply if opting for a custom format.

Valid values for the `operator` keyword are the following:

- `topN`            (cell is in the top n (= `value`) values of the range)
- `bottomN`         (cell is in the bottom n (= `value`) values of the range)
- `topN%`           (cell is in the top n% (= `value`) values of the range)
- `bottomN%`        (cell is in the bottom n% (= `value`) values of the range)

The remaining keywords are defined as above for the `:cellIs` conditional format type.

# Examples

```julia
```

# type = :aboveAverage

This conditional format can be used to compare cell values in the range with the 
average value for the range. 

The available keywords are:

- `operator`   : Defines the comparison to make.
- `stopIfTrue` : Stops evaluating the conditional formats if this one is true.
- `dxStyle`    : Used optionally to select one of the built-in Excel formats to apply.
- `format`     : defines the numFmt to apply if opting for a custom format.
- `font`       : defines the font to apply if opting for a custom format.
- `border`     : defines the border to apply if opting for a custom format.
- `fill`       : defines the fill to apply if opting for a custom format.

Valid values for the `operator` keyword are the following:

- `aboveAverage`    (cell is above the average of the range)
- `aboveEqAverage`  (cell is above or equal to the average of the range)
- `plus1StdDev`     (cell is above the average of the range + 1 standard deviation)
- `plus2StdDev`     (cell is above the average of the range + 2 standard deviations)
- `plus3StdDev`     (cell is above the average of the range + 3 standard deviations)
- `belowAverage`    (cell is below the average of the range)
- `belowEqAverage`  (cell is below or equal to the average of the range)
- `minus1StdDev`    (cell is below the average of the range - 1 standard deviation)
- `minus2StdDev`    (cell is below the average of the range - 2 standard deviations)
- `minus3StdDev`    (cell is below the average of the range - 3 standard deviations)

The remaining keywords are defined as above for the `:cellIs` conditional format type.

# Examples

```julia
```

# type = :containsText
# type = :notContainsText
# type = :beginsWith
# type = :endsWith

Highlight cells in the range that contain (or do not contain), begin or end with 
a specific text string.

Valid keywords are:

- `value`     : Gives the literal text to match or provides a cell reference (e.g. `"A1"`).
- `stopIfTrue` : Stops evaluating the conditional formats if this one is true.
- `dxStyle`    : Used optionally to select one of the built-in Excel formats to apply.
- `format`     : defines the numFmt to apply if opting for a custom format.
- `font`       : defines the font to apply if opting for a custom format.
- `border`     : defines the border to apply if opting for a custom format.
- `fill`       : defines the fill to apply if opting for a custom format.

`value` gives the literal text to compare (eg. "Hello World") or provides a cell reference (e.g. `"A1"`).

The remaining keywords are optional and are defined as above for the `:cellIs` conditional format type.

# Examples

```julia
```

# type = :timePeriod

When cells contain dates, this conditional format can be used to highlight cells.
The available keywords are:

- `operator`   : Defines the comparison to make.
- `stopIfTrue` : Stops evaluating the conditional formats if this one is true.
- `dxStyle`    : Used optionally to select one of the built-in Excel formats to apply
- `format`     : defines the numFmt to apply if opting for a custom format.
- `font`       : defines the font to apply if opting for a custom format.
- `border`     : defines the border to apply if opting for a custom format.
- `fill`       : defines the fill to apply if opting for a custom format.

Valid values for the keyword `operator` are the following:

- `yesterday`
- `today`
- `tomorrow`
- `last7Days`
- `lastWeek`
- `thisWeek`
- `nextWeek`
- `lastMonth`
- `thisMonth`
- `nextMonth`

The remaining keywords are defined as above for the `:cellIs` conditional format type.

# Examples

```julia
```

# type = :containsErrors
# type = :notContainsErrors
# type = :containsBlanks
# type = :notContainsBlanks
# type = :uniqueValues
# type = :duplicateValues

These conditional formattimg options highlight cells that contain or don't contain errors, 
are blank or not blank, are unique in the range or duplicates within the range. 
The available keywords are: 

- `stopIfTrue` : Stops evaluating the conditional formats if this one is true.
- `dxStyle`    : Used optionally to select one of the built-in Excel formats to apply
- `format`     : defines the numFmt to apply if opting for a custom format.
- `font`       : defines the font to apply if opting for a custom format.
- `border`     : defines the border to apply if opting for a custom format.
- `fill`       : defines the fill to apply if opting for a custom format.

These keywords are defined as above for the `:cellIs` conditional format type.

# Examples

```julia
```


"""
function setConditionalFormat(f, r, type::Symbol; kw...)
    if type == :colorScale
        setCfColorScale(f, r; kw...)
    elseif type == :cellIs
        setCfCellIs(f, r; kw...)
    elseif type == :top10
        setCfTop10(f, r; kw...)
    elseif type == :aboveAverage
        setCfAboveAverage(f, r; kw...)
    elseif type == :timePeriod
        setCfTimePeriod(f, r; kw...)
    elseif type ∈ [:containsText, :notContainsText, :beginsWith, :endsWith]
        setCfContainsText(f, r; operator=String(type), kw...)
    elseif type ∈ [:containsBlanks, :notContainsBlanks, :containsErrors, :notContainsErrors, duplicateValues, uniqueValues]
        setCfContainsBlankErrorUniqDup(f, r; operator=String(type), kw...)
#    elseif type == :iconSet
#        throw(XLSXError("Icon sets are not yet implemented."))
#    elseif type == :dataBar
#        throw(XLSXError("Data bars are not yet implemented."))
else
        throw(XLSXError("Invalid conditional format type: $type. Valid options are: `:colorScale`, `:cellIs`"))
    end
end

function setConditionalFormat(f, r, c, type::Symbol; kw...)
    if type == :colorScale
        setCfColorScale(f, r, c; kw...)
    elseif type == :cellIs
        setCfCellIs(f, r, c; kw...)
    elseif type == :top10
        setCfTop10(f, r, c; kw...)
    elseif type == :aboveAverage
        setCfAboveAverage(f, r; kw...)
    elseif type == :timePeriod
        setCfTimePeriod(f, r, c; kw...)
    elseif type ∈ [:containsText, :notContainsText, :beginsWith, :endsWith]
        setCfContainsText(f, r, c; operator=String(type), kw...)
    elseif type ∈ [:containsBlanks, :notContainsBlanks, :containsErrors, :notContainsErrors, duplicateValues, uniqueValues]
        setCfContainsBlankErrorUniqDup(f, r, c; operator=String(type), kw...)
#    elseif type == :iconSet
#        throw(XLSXError("Icon sets are not yet implemented."))
#    elseif type == :dataBar
#        throw(XLSXError("Data bars are not yet implemented."))
    else
        throw(XLSXError("Invalid conditional format type: $type. Valid options are: `:colorScale`, `:cellIs`."))
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

    allcfs = allCfs(ws)                    # get all conditional format blocks
    old_cf = getConditionalFormats(allcfs) # extract conditional format info

    let new_pr, new_cf

        new_pr = length(old_cf) > 0 ? string(maximum([last(x).priority for x in values(old_cf)])+1) : 1

        if isnothing(colorscale)

            min_type in ["min", "percentile", "percent", "num"] || throw(XLSXError("Invalid min_type: $min_type. Valid options are: min, percentile, percent, num."))
            isnothing(min_val) || is_valid_cellname(min_val) || !is_valid_sheet_cellname(min_val) || !isnothing(tryparse(Float64,min_val)) || throw(XLSXError("Invalid mid_type: $min_val. Valid options are a CellRef (e.g. `A1`) or a number."))
            isnothing(mid_type) || mid_type in ["percentile", "percent", "num"] || throw(XLSXError("Invalid mid_type: $mid_type. Valid options are: percentile, percent, num."))
            isnothing(mid_val) || is_valid_cellname(mid_val) || !is_valid_sheet_cellname(mid_val) || !isnothing(tryparse(Float64,mid_val)) || throw(XLSXError("Invalid mid_type: $mid_val. Valid options are a CellRef (e.g. `A1`) or a number."))
            max_type in ["max", "percentile", "percent", "num"] || throw(XLSXError("Invalid max_type: $max_type. Valid options are: max, percentile, percent, num."))
            isnothing(max_val) || is_valid_cellname(max_val) || !is_valid_sheet_cellname(max_val) || !isnothing(tryparse(Float64,max_val)) || throw(XLSXError("Invalid mid_type: $max_val. Valid options are a CellRef (e.g. `A1`) or a number."))

            min_val = convertref(min_val)
            mid_val = convertref(mid_val)
            max_val = convertref(max_val)

            cfx = XML.h.cfRule(type="colorScale", priority=new_pr,
                XML.h.colorScale(
                    isnothing(min_val) ? XML.h.cfvo(type=min_type) : XML.h.cfvo(type=min_type, val=min_val),
                    isnothing(mid_type) ? nothing : XML.h.cfvo(type=mid_type, val=mid_val),
                    isnothing(max_val) ? XML.h.cfvo(type=max_type) : XML.h.cfvo(type=max_type, val=max_val),
                    XML.h.color(rgb=get_color(min_col)),
                    isnothing(mid_type) ? nothing : XML.h.color(rgb=get_color(mid_col)),
                    XML.h.color(rgb=get_color(max_col))
                )
            )

        else
            if !haskey(colorscales, colorscale)
                throw(XLSXError("Invalid color scale: $colorScale. Valid options are: $(keys(colorscales))."))
            end
            cfx=colorscales[colorscale]
            cfx["priority"] = new_pr
        end

        update_worksheet_cfx!(allcfs, cfx, ws, rng)

    end

    return 0
end

setCfCellIs(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfCellIs, ws, row, nothing; kw...)
setCfCellIs(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfCellIs, ws, nothing, col; kw...)
setCfCellIs(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfCellIs, ws, nothing, nothing; kw...)
setCfCellIs(ws::Worksheet, ::Colon; kw...) = process_colon(setCfCellIs, ws, nothing, nothing; kw...)
setCfCellIs(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfCellIs(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfCellIs(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfCellIs(ws, rng.rng; kw...)
setCfCellIs(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfCellIs(ws, rng.colrng; kw...)
setCfCellIs(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfCellIs(ws, rng.rowrng; kw...)
setCfCellIs(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfCellIs, ws, rng; kw...)
setCfCellIs(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfCellIs, ws, rng; kw...)
setCfCellIs(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfCellIs, xl, sheetcell; kw...)
setCfCellIs(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfCellIs, ws, ref_or_rng; kw...)
function setCfCellIs(ws::Worksheet, rng::CellRange;
    operator::Union{Nothing,String}="greaterThan",
    value::Union{Nothing,String}=nothing,
    value2::Union{Nothing,String}=nothing,
    stopIfTrue::Union{Nothing,String}=nothing,
    dxStyle::Union{Nothing,String}=nothing,
    format::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    font::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    border::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    fill::Union{Nothing,Vector{Pair{String,String}}}=nothing
)::Int

    !issubset(rng, get_dimension(ws)) && throw(XLSXError("Range `$rng` goes outside worksheet dimension."))

    allcfs = allCfs(ws)                    # get all conditional format blocks
    old_cf = getConditionalFormats(allcfs) # extract conditional format info

    !isnothing(value) && !is_valid_cellname(value) && !is_valid_sheet_cellname(value) && isnothing(tryparse(Float64, value)) && throw(XLSXError("Invalid `value`: $value. Must be a number or a CellRef."))
    !isnothing(value2) && !is_valid_cellname(value2) && !is_valid_sheet_cellname(value2) && isnothing(tryparse(Float64, value2)) && throw(XLSXError("Invalid `value2`: $value2. Must be a number or a CellRef."))

    wb=get_workbook(ws)
    dx = get_dx(dxStyle, format, font, border, fill)
    new_dx= get_new_dx(wb, dx)
    dxid = Add_Cf_Dx(wb, new_dx)

    if isnothing(value)
        value = all(ismissing.(ws[rng])) ? nothing : string(sum(skipmissing(ws[rng]))/count(!ismissing, ws[rng]))
    end
    cfx = XML.Element("cfRule"; type="cellIs", dxfId=Int(dxid.id), operator=operator)
    if !isnothing(stopIfTrue) && stopIfTrue == "true"
        cfx["stopIfTrue"] = "1"
    end
    cfx["priority"] = length(old_cf) > 0 ? string(maximum([last(x).priority for x in values(old_cf)])+1) : 1

    push!(cfx, XML.Element("formula", XML.Text(value)))
    if !isnothing(value2) && operator ∈ needsValue2
        push!(cfx, XML.Element("formula", XML.Text(value2)))
    end

    update_worksheet_cfx!(allcfs, cfx, ws, rng)

    return 0
end

setCfContainsText(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfContainsText, ws, row, nothing; kw...)
setCfContainsText(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfContainsText, ws, nothing, col; kw...)
setCfContainsText(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfContainsText, ws, nothing, nothing; kw...)
setCfContainsText(ws::Worksheet, ::Colon; kw...) = process_colon(setCfContainsText, ws, nothing, nothing; kw...)
setCfContainsText(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfContainsText(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfContainsText(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfContainsText(ws, rng.rng; kw...)
setCfContainsText(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfContainsText(ws, rng.colrng; kw...)
setCfContainsText(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfContainsText(ws, rng.rowrng; kw...)
setCfContainsText(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfContainsText, ws, rng; kw...)
setCfContainsText(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfContainsText, ws, rng; kw...)
setCfContainsText(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfContainsText, xl, sheetcell; kw...)
setCfContainsText(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfContainsText, ws, ref_or_rng; kw...)
function setCfContainsText(ws::Worksheet, rng::CellRange;
    operator::Union{Nothing,String}="containsText",
    value::Union{Nothing,String}="",
    stopIfTrue::Union{Nothing,String}=nothing,
    dxStyle::Union{Nothing,String}=nothing,
    format::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    font::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    border::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    fill::Union{Nothing,Vector{Pair{String,String}}}=nothing
)::Int

    !issubset(rng, get_dimension(ws)) && throw(XLSXError("Range `$rng` goes outside worksheet dimension."))

    allcfs = allCfs(ws)                    # get all conditional format blocks
    old_cf = getConditionalFormats(allcfs) # extract conditional format info

    !isnothing(value) && !is_valid_cellname(value) && !is_valid_sheet_cellname(value) && throw(XLSXError("Invalid `value`: $value. Must be a number or a CellRef."))

    wb=get_workbook(ws)
    dx = get_dx(dxStyle, format, font, border, fill)
    new_dx= get_new_dx(wb, dx)
    dxid = Add_Cf_Dx(wb, new_dx)

    if operator == "containsText"
        formula = "NOT(ISERROR(SEARCH(\"__txt__\",__CR__)))"
    elseif operator == "notContainsText"
        operator = "notContains"
        formula = "ISERROR(SEARCH(\"__txt__\",__CR__))"
    elseif operator == "beginsWith"
        operator = "beginsWith"
        formula = "LEFT(__CR__,LEN(\"__txt__\"))=\"__txt__\""
    elseif operator == "endsWith"
        operator = "endsWith"
        formula = "RIGHT(__CR__,LEN(\"__txt__\"))=\"__txt__\""
    else
        throw(XLSXError("Invalid operator: $operator. Valid options are: `containsText`, `notContainsText`, `beginsWith`, `endsWith`."))
    end
    formula = replace(formula, "__txt__" => value, "__CR__" => string(first(rng)))

    cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1", operator=operator, text=value)
    if !isnothing(stopIfTrue) && stopIfTrue == "true"
        cfx["stopIfTrue"] = "1"
    end
    cfx["priority"] = length(old_cf) > 0 ? string(maximum([last(x).priority for x in values(old_cf)])+1) : 1
    push!(cfx, XML.Element("formula", XML.Text(formula)))

    update_worksheet_cfx!(allcfs, cfx, ws, rng)

    return 0
end

setCfTop10(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfTop10, ws, row, nothing; kw...)
setCfTop10(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfTop10, ws, nothing, col; kw...)
setCfTop10(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfTop10, ws, nothing, nothing; kw...)
setCfTop10(ws::Worksheet, ::Colon; kw...) = process_colon(setCfTop10, ws, nothing, nothing; kw...)
setCfTop10(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfTop10(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfTop10(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfTop10(ws, rng.rng; kw...)
setCfTop10(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfTop10(ws, rng.colrng; kw...)
setCfTop10(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfTop10(ws, rng.rowrng; kw...)
setCfTop10(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfTop10, ws, rng; kw...)
setCfTop10(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfTop10, ws, rng; kw...)
setCfTop10(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfTop10, xl, sheetcell; kw...)
setCfTop10(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfTop10, ws, ref_or_rng; kw...)
function setCfTop10(ws::Worksheet, rng::CellRange;
    operator::Union{Nothing,String}="topN",
    value::Union{Nothing,String}="10",
    stopIfTrue::Union{Nothing,String}=nothing,
    dxStyle::Union{Nothing,String}=nothing,
    format::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    font::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    border::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    fill::Union{Nothing,Vector{Pair{String,String}}}=nothing
)::Int

    !issubset(rng, get_dimension(ws)) && throw(XLSXError("Range `$rng` goes outside worksheet dimension."))

    allcfs = allCfs(ws)                    # get all conditional format blocks
    old_cf = getConditionalFormats(allcfs) # extract conditional format info

    !isnothing(value) && !is_valid_cellname(value) && !is_valid_sheet_cellname(value) && isnothing(tryparse(Float64, value)) && throw(XLSXError("Invalid `value`: $value. Must be a number or a CellRef."))

    wb=get_workbook(ws)
    dx = get_dx(dxStyle, format, font, border, fill)
    new_dx= get_new_dx(wb, dx)
    dxid = Add_Cf_Dx(wb, new_dx)
 
    if operator == "topN"
        cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1", rank=value)
    elseif operator == "topN%"
        cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1", percent="1", rank=value)
    elseif operator == "bottomN"
        cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1", bottom="1", rank=value)
    elseif operator == "bottomN%"
        cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1", percent="1", bottom="1", rank=value)
    else
        throw(XLSXError("Invalid operator: $operator. Valid options are: `topN`, `topN%`, `bottomN`, `bottomN%`."))
    end

    if !isnothing(stopIfTrue) && stopIfTrue == "true"
        cfx["stopIfTrue"] = "1"
    end
    cfx["priority"] = length(old_cf) > 0 ? string(maximum([last(x).priority for x in values(old_cf)])+1) : 1

    update_worksheet_cfx!(allcfs, cfx, ws, rng)

    return 0
end

setCfAboveAverage(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfAboveAverage, ws, row, nothing; kw...)
setCfAboveAverage(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfAboveAverage, ws, nothing, col; kw...)
setCfAboveAverage(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfAboveAverage, ws, nothing, nothing; kw...)
setCfAboveAverage(ws::Worksheet, ::Colon; kw...) = process_colon(setCfAboveAverage, ws, nothing, nothing; kw...)
setCfAboveAverage(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfAboveAverage(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfAboveAverage(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfAboveAverage(ws, rng.rng; kw...)
setCfAboveAverage(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfAboveAverage(ws, rng.colrng; kw...)
setCfAboveAverage(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfAboveAverage(ws, rng.rowrng; kw...)
setCfAboveAverage(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfAboveAverage, ws, rng; kw...)
setCfAboveAverage(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfAboveAverage, ws, rng; kw...)
setCfAboveAverage(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfAboveAverage, xl, sheetcell; kw...)
setCfAboveAverage(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfAboveAverage, ws, ref_or_rng; kw...)
function setCfAboveAverage(ws::Worksheet, rng::CellRange;
    operator::Union{Nothing,String}="aboveAverage",
    stopIfTrue::Union{Nothing,String}=nothing,
    dxStyle::Union{Nothing,String}=nothing,
    format::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    font::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    border::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    fill::Union{Nothing,Vector{Pair{String,String}}}=nothing
)::Int

    !issubset(rng, get_dimension(ws)) && throw(XLSXError("Range `$rng` goes outside worksheet dimension."))

    allcfs = allCfs(ws)                    # get all conditional format blocks
    old_cf = getConditionalFormats(allcfs) # extract conditional format info

    isnothing(tryparse(Float64, value)) && throw(XLSXError("Invalid `value`: $value. Must be a number."))

    wb=get_workbook(ws)
    dx = get_dx(dxStyle, format, font, border, fill)
    new_dx= get_new_dx(wb, dx)
    dxid = Add_Cf_Dx(wb, new_dx)

     if operator == "aboveAverage"
        cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1")
    elseif operator == "aboveEqAverage"
        cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1", equalAverage="1")
    elseif operator == "plus1StdDev"
        cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1", bottom="1", stdDev="1")
    elseif operator == "plus2StdDev"
        cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1", percent="1", stdDev="2")
    elseif operator == "plus3StdDev"
        cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1", percent="1", stdDev="3")
    elseif operator == "belowAverage"
        cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1", aboveAverage="0", )
    elseif operator == "belowEqAverage"
        cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1", aboveAverage="0", equalAverage="1")
    elseif operator == "minus1StdDev"
        cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1", aboveAverage="0", stdDev="1")
    elseif operator == "minus2StdDev"
        cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1", aboveAverage="0", stdDev="2")
    elseif operator == "minus3StdDev"
        cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id), priority="1", aboveAverage="0", stdDev="3")
    else
        throw(XLSXError("Invalid operator: $operator. Valid options are: `aboveAverage`, `aboveEqAverage`, `plus1sStdDev`, `plus2StdDev`, `plus3StdDev`, `belowAverage`, `belowEqAverage`, `minus1StdDev`, `minus2StdDev`, `minus3StdDev`."))
    end

    if !isnothing(stopIfTrue) && stopIfTrue == "true"
        cfx["stopIfTrue"] = "1"
    end
    cfx["priority"] = length(old_cf) > 0 ? string(maximum([last(x).priority for x in values(old_cf)])+1) : 1

    update_worksheet_cfx!(allcfs, cfx, ws, rng)

    return 0
end

setCfTimePeriod(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfTimePeriod, ws, row, nothing; kw...)
setCfTimePeriod(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfTimePeriod, ws, nothing, col; kw...)
setCfTimePeriod(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfTimePeriod, ws, nothing, nothing; kw...)
setCfTimePeriod(ws::Worksheet, ::Colon; kw...) = process_colon(setCfTimePeriod, ws, nothing, nothing; kw...)
setCfTimePeriod(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfTimePeriod(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfTimePeriod(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfTimePeriod(ws, rng.rng; kw...)
setCfTimePeriod(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfTimePeriod(ws, rng.colrng; kw...)
setCfTimePeriod(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfTimePeriod(ws, rng.rowrng; kw...)
setCfTimePeriod(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfTimePeriod, ws, rng; kw...)
setCfTimePeriod(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfTimePeriod, ws, rng; kw...)
setCfTimePeriod(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfTimePeriod, xl, sheetcell; kw...)
setCfTimePeriod(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfTimePeriod, ws, ref_or_rng; kw...)
function setCfTimePeriod(ws::Worksheet, rng::CellRange;
    operator::Union{Nothing,String}="last7Days",
    stopIfTrue::Union{Nothing,String}=nothing,
    dxStyle::Union{Nothing,String}=nothing,
    format::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    font::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    border::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    fill::Union{Nothing,Vector{Pair{String,String}}}=nothing
)::Int

    !issubset(rng, get_dimension(ws)) && throw(XLSXError("Range `$rng` goes outside worksheet dimension."))

allcfs = allCfs(ws)                    # get all conditional format blocks
old_cf = getConditionalFormats(allcfs) # extract conditional format info

    if operator == "yesterday"
        formula = "FLOOR(__CR__,1)=TODAY()-1"
    elseif operator == "today"
        formula = "FLOOR(__CR__,1)=TODAY()"
    elseif operator == "tomorrow"
        formula = "FLOOR(__CR__,1)=TODAY()+1"
    elseif operator == "last7Days"
        formula = "AND(TODAY()-FLOOR(__CR__,1)<=6,FLOOR(__CR__,1)<=TODAY())"
    elseif operator == "lastWeek"
        formula = "AND(TODAY()-ROUNDDOWN(__CR__,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(__CR__,0)<(WEEKDAY(TODAY())+7))"
    elseif operator == "thisWeek"
        formula = "AND(TODAY()-ROUNDDOWN(__CR__,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(__CR__,0)-TODAY()<=7-WEEKDAY(TODAY()))"
    elseif operator == "nextWeek"
        formula = "AND(ROUNDDOWN(__CR__,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(__CR__,0)-TODAY()<(15-WEEKDAY(TODAY())))"
    elseif operator == "lastMonth"
        formula = "AND(MONTH(__CR__)=MONTH(EDATE(TODAY(),0-1)),YEAR(__CR__)=YEAR(EDATE(TODAY(),0-1)))"
    elseif operator == "thisMonth"
        formula = "AND(MONTH(__CR__)=MONTH(TODAY()),YEAR(__CR__)=YEAR(TODAY()))"
    elseif operator == "nextMonth"
        formula = "AND(MONTH(__CR__)=MONTH(EDATE(TODAY(),0+1)),YEAR(__CR__)=YEAR(EDATE(TODAY(),0+1)))"
    else
        throw(XLSXError("Invalid operator: $operator. Valid options are: `yesterday`, `today`, `tomorrow`, `last7Days`, `lastWeek`, `thisWeek`, `nextWeek`, `lastMonth`, `thisMonth`, `nextMonth`."))
    end
    formula = replace(formula, "__CR__" => string(first(rng)))

    wb=get_workbook(ws)
    dx = get_dx(dxStyle, format, font, border, fill)
    new_dx = get_new_dx(wb, dx)
    dxid = Add_Cf_Dx(wb, new_dx)

    cfx = XML.Element("cfRule"; type="timePeriod", dxfId=Int(dxid.id), operator=operator)
    if !isnothing(stopIfTrue) && stopIfTrue == "true"
        cfx["stopIfTrue"] = "1"
    end
    cfx["priority"] = length(old_cf) > 0 ? string(maximum([last(x).priority for x in values(old_cf)])+1) : 1

    push!(cfx, XML.Element("formula", XML.Text(formula)))

    update_worksheet_cfx!(allcfs, cfx, ws, rng)

    return 0
end

setCfContainsBlankErrorUniqDup(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setCfContainsBlankErrorUniqDup, ws, row, nothing; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setCfContainsBlankErrorUniqDup, ws, nothing, col; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setCfContainsBlankErrorUniqDup, ws, nothing, nothing; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, ::Colon; kw...) = process_colon(setCfContainsBlankErrorUniqDup, ws, nothing, nothing; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setCfContainsBlankErrorUniqDup(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfContainsBlankErrorUniqDup(ws, rng.rng; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfContainsBlankErrorUniqDup(ws, rng.colrng; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfContainsBlankErrorUniqDup(ws, rng.rowrng; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfContainsBlankErrorUniqDup, ws, rng; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfContainsBlankErrorUniqDup, ws, rng; kw...)
setCfContainsBlankErrorUniqDup(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfContainsBlankErrorUniqDup, xl, sheetcell; kw...)
setCfContainsBlankErrorUniqDup(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfContainsBlankErrorUniqDup, ws, ref_or_rng; kw...)
function setCfContainsBlankErrorUniqDup(ws::Worksheet, rng::CellRange;
    operator::Union{Nothing,String}="containsBlank",
    stopIfTrue::Union{Nothing,String}=nothing,
    dxStyle::Union{Nothing,String}=nothing,
    format::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    font::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    border::Union{Nothing,Vector{Pair{String,String}}}=nothing,
    fill::Union{Nothing,Vector{Pair{String,String}}}=nothing
)::Int

    !issubset(rng, get_dimension(ws)) && throw(XLSXError("Range `$rng` goes outside worksheet dimension."))

    allcfs = allCfs(ws)                    # get all conditional format blocks
    old_cf = getConditionalFormats(allcfs) # extract conditional format info

    if operator == "containsBlanks"
        formula = "LEN(TRIM(__CR__))=0"
    elseif operator == "notContainsBlanks"
        formula = "LEN(TRIM(__CR__))>0"
    elseif operator == "containsErrors"
        formula = "ISERROR(__CR__)"
    elseif operator == "notContainsErrors"
        formula = "NOT(ISERROR(__CR__))"
    elseif operator == "uniqueValues"
        formula = ""
    elseif operator == "duplicateValues"
        formula = ""
    else
        throw(XLSXError("Invalid operator: $operator. Valid options are: `containsBlanks`, `notContainsBlanks`, `containsErrors`, `notContainsErrors`, `uniqueValues`, `duplicateValues`."))
    end
    formula = replace(formula, "__CR__" => string(first(rng)))

    wb=get_workbook(ws)
    dx = get_dx(dxStyle, format, font, border, fill)
    new_dx = get_new_dx(wb, dx)
    dxid = Add_Cf_Dx(wb, new_dx)

    cfx = XML.Element("cfRule"; type=operator, dxfId=Int(dxid.id))
    if !isnothing(stopIfTrue) && stopIfTrue == "true"
        cfx["stopIfTrue"] = "1"
    end
    cfx["priority"] = length(old_cf) > 0 ? string(maximum([last(x).priority for x in values(old_cf)])+1) : 1
    formula !="" && push!(cfx, XML.Element("formula", XML.Text(formula)))

    update_worksheet_cfx!(allcfs, cfx, ws, rng)

    return 0
end