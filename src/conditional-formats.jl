const colorscales = Dict(    # Defines the 12 standard, built-in Excel color scales for conditional formatting.
    "redyellowgreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFFEB84"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "greenyellowred" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFFEB84"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "redwhitegreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "greenwhitered" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "redwhiteblue" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FF5A8AC6")
        )
    ),
    "bluewhitered" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="percentile", val="50"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF5A8AC6"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "redwhite" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFF8696B"),
            XML.h.color(rgb="FFFCFCFF")
        )
    ),
    "whitered" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FFF8696B")
        )
    ),
    "whitegreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFFCFCFF"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "greenwhite" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFCFCFF")
        )
    ),
    "yellowgreen" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FFFFEF9C"),
            XML.h.color(rgb="FF63BE7B")
        )
    ),
    "greenyellow" => XML.h.cfRule(type="colorScale", priority="1",
        XML.h.colorScale(
            XML.h.cfvo(type="min"),
            XML.h.cfvo(type="max"),
            XML.h.color(rgb="FF63BE7B"),
            XML.h.color(rgb="FFFFEF9C")
        )
    )
)
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
                #                if any(XML.tag(c) == "extLst" for c in XML.children(child))
                #                    println("  extras: ", true)
                #                end
            end
        end
        push!(allcfs, CellRange(cf["sqref"]) => cf_types)
    end
    return allcfs
end

"""
    addConditionalFormat!(ws::Worksheet, rng::CellRange; kw...) -> nothing

Add a new conditional format to a worksheet.

Keyword argumenst `colorScale`, `dataBar`, `iconSet`, and `formula` are mutually exclusive.

Valid values for `colorScale` are:

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
- `"greenyellow"`   : Green, Yellow color scale.

These are the 12 built-in color scales in Excel.

"""
function addConditionalFormat!(ws::Worksheet, rng::CellRange;
        colorScale::Union{Nothing,AbstractString}=nothing,
        dataBar::Union{Nothing,AbstractString}=nothing,
        iconSet::Union{Nothing,AbstractString}=nothing,
        formula::Union{Nothing,AbstractString}=nothing,
    )::Nothing

    if !isnothing(colorScale) && !isnothing(dataBar) && !isnothing(iconSet) && !isnothing(formula)
        throw(XLSXError("Only one of colorScale, dataBar, iconSet, or formula can be specified."))
    end

    if isnothing(colorScale) && isnothing(dataBar) && isnothing(iconSet) && isnothing(formula)
        throw(XLSXError("At least one of colorScale, dataBar, iconSet, or formula must be specified."))
    end

    !issubset(rng, get_dimension(ws)) && throw(XLSXError("Range `$rng` goes outside worksheet dimension."))

    old_cf = getConditionalFormats(ws)
    for cf in old_cf
        if intersects(cf.first, rng)
            throw(XLSXError("Range `$rng` intersects with existing conditional format range `$(cf.first)`."))
        end
    end

    if !isnothing(colorScale)
        if !haskey(colorscales, colorScale)
            throw(XLSXError("Invalid color scale: $colorScale. Valid options are: $(keys(colorscales))."))
        end
        new_cf = XML.Element("conditionalFormatting"; sqref=rng)
        push!(new_cf, colorscales[colorScale])
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

    return nothing
end