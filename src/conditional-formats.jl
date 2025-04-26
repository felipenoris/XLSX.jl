const colorscales = Dict(    # Defines the 12 standard, built-in Excel color scales for conditional formatting.
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
function getConditionalFormats(ws::Worksheet)::Vector{Pair{CellRange,Vector{String}}}
    sheetdoc = xmlroot(ws.package, "xl/worksheets/sheet" * string(ws.sheetId) * ".xml") # find all the <conditionalFormatting> blocks in the worksheet's xml file
    allcfnodes = find_all_nodes("/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":worksheet/" * SPREADSHEET_NAMESPACE_XPATH_ARG * ":conditionalFormatting", sheetdoc)
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

# Example
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


"""
function setConditionalFormat(f, r, type::Symbol; kw...)
    if type == :colorScale
        setCfColorScale(f, r; kw...)
#    elseif type == :dataBar
#        throw(XLSXError("Data bars are not yet implemented."))
#    elseif type == :iconSet
#        throw(XLSXError("Icon sets are not yet implemented."))
#    elseif type == :Cell
#        throw(XLSXError("Cell conditional formats are not yet implemented."))
else
        throw(XLSXError("Invalid conditional format type: $type. Valid options are: colorScale, dataBar, iconSet, formula."))
    end
end
function setConditionalFormat(f, r, c, type::Symbol; kw...)
    if type == :colorScale
        setCfColorScale(f, r, c; kw...)
#    elseif type == :dataBar
#        throw(XLSXError("Data bars are not yet implemented."))
#    elseif type == :iconSet
#        throw(XLSXError("Icon sets are not yet implemented."))
#    elseif type == :Cell
#        throw(XLSXError("Cell conditional formats are not yet implemented."))
    else
        throw(XLSXError("Invalid conditional format type: $type. Valid options are: colorScale, dataBar, iconSet, formula."))
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

    # Insert the new conditional formatting into the worksheet XML
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

    update_worksheets_xml!(get_xlsxfile(ws))

    return 0
end