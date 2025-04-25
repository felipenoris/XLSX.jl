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
    addConditionalFormat!(ws::Worksheet, rng::CellRange, type::Symbol; kw...)

Add a new conditional format to a worksheet.

!!! warning "In Develpment

    This function is still in development and may not work as expected.
    It is not yet implemented for all types of conditional formats.

Valid options for `type` are `:colorScale` (others in develpment) and these 
determine which type of conditional formatting is being defined.

Keyword options differ according to the `type` specified

Valid values for `colorScale` are:

- `:redyellowgreen`: Red, Yellow, Green color scale.
- `:greenyellowred`: Green, Yellow, Red color scale.
- `:redwhitegreen` : Red, White, Green color scale.
- `:greenwhitered` : Green, White, Red color scale.
- `:redwhiteblue`  : Red, White, Blue color scale.
- `:bluewhitered`  : Blue, White, Red color scale.
- `:redwhite`      : Red, White color scale.
- `:whitered`      : White, Red color scale.
- `:whitegreen`    : White, Green color scale.
- `:greenwhite`    : Green, White color scale.
- `:yellowgreen`   : Yellow, Green color scale.
- `:greenyellow`   : Green, Yellow color scale.

These are the 12 built-in color scales in Excel.

"""
function setConditionalFormat(xf::XLSXFile, ref_or_rng, type::Symbol; kw...)
    if type == :colorScale
        process_sheetcell(setCfColorScale, xf, ref_or_rng; kw...)
#    elseif type == :dataBar
#        throw(XLSXError("Data bars are not yet implemented."))
#    elseif type == :iconSet
#        throw(XLSXError("Icon sets are not yet implemented."))
#    elseif type == :formula
#        throw(XLSXError("Formulas are not yet implemented."))
    else
        throw(XLSXError("Invalid conditional format type: $type. Valid options are: colorScale, dataBar, iconSet, formula."))
    end
end
function setConditionalFormat(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon, type::Symbol; kw...)
    if type == :colorScale
        process_colon(setCfColorScale, ws, row, nothing; kw...)
#    elseif type == :dataBar
#        throw(XLSXError("Data bars are not yet implemented."))
#    elseif type == :iconSet
#        throw(XLSXError("Icon sets are not yet implemented."))
#    elseif type == :formula
#        throw(XLSXError("Formulas are not yet implemented."))
    else
        throw(XLSXError("Invalid conditional format type: $type. Valid options are: colorScale, dataBar, iconSet, formula."))
    end
end
function setConditionalFormat(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}, type::Symbol; kw...)
    if type == :colorScale
        process_colon(setCfColorScale, ws, nothing, col; kw...)
#    elseif type == :dataBar
#        throw(XLSXError("Data bars are not yet implemented."))
#    elseif type == :iconSet
#        throw(XLSXError("Icon sets are not yet implemented."))
#    elseif type == :formula
#        throw(XLSXError("Formulas are not yet implemented."))
    else
        throw(XLSXError("Invalid conditional format type: $type. Valid options are: colorScale, dataBar, iconSet, formula."))
    end
end
function setConditionalFormat(ws::Worksheet, ::Colon, ::Colon, type::Symbol; kw...)
    if type == :colorScale
        process_colon(setCfColorScale, ws, nothing, nothing; kw...)
#    elseif type == :dataBar
#        throw(XLSXError("Data bars are not yet implemented."))
#    elseif type == :iconSet
#        throw(XLSXError("Icon sets are not yet implemented."))
#    elseif type == :formula
#        throw(XLSXError("Formulas are not yet implemented."))
    else
        throw(XLSXError("Invalid conditional format type: $type. Valid options are: colorScale, dataBar, iconSet, formula."))
    end
end
function setConditionalFormat(ws::Worksheet, ::Colon, type::Symbol; kw...)
    if type == :colorScale
        process_colon(setCfColorScale, ws, nothing, nothing; kw...)
#    elseif type == :dataBar
#        throw(XLSXError("Data bars are not yet implemented."))
#    elseif type == :iconSet
#        throw(XLSXError("Icon sets are not yet implemented."))
#    elseif type == :formula
#        throw(XLSXError("Formulas are not yet implemented."))
    else
        throw(XLSXError("Invalid conditional format type: $type. Valid options are: colorScale, dataBar, iconSet, formula."))

    end
end
function setConditionalFormat(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}, type::Symbol; kw...)
    if type == :colorScale
        setCfColorScale(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
#    elseif type == :dataBar
#        throw(XLSXError("Data bars are not yet implemented."))
#    elseif type == :iconSet
#        throw(XLSXError("Icon sets are not yet implemented."))
#    elseif type == :formula
#        throw(XLSXError("Formulas are not yet implemented."))
    else
        throw(XLSXError("Invalid conditional format type: $type. Valid options are: colorScale, dataBar, iconSet, formula."))
    end
end
function setConditionalFormat(ws::Worksheet, ref_or_rng::AbstractString, type::Symbol; kw...)
    if type == :colorScale
        process_ranges(setCfColorScale, ws, ref_or_rng; kw...)::Int
#    elseif type == :dataBar
#        throw(XLSXError("Data bars are not yet implemented."))
#    elseif type == :iconSet
#        throw(XLSXError("Icon sets are not yet implemented."))
#    elseif type == :formula
#        throw(XLSXError("Formulas are not yet implemented."))
    else
        throw(XLSXError("Invalid conditional format type: $type. Valid options are: colorScale, dataBar, iconSet, formula."))
    end
end
setCfColorScale(ws::Worksheet, ref::SheetCellRef; kw...) = do_sheet_names_match(ws, ref) && setCfColorScale(ws, ref.cellref; kw...)
setCfColorScale(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setCfColorScale(ws, rng.rng; kw...)
setCfColorScale(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setCfColorScale(ws, rng.colrng; kw...)
setCfColorScale(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setCfColorScale(ws, rng.rowrng; kw...)
setCfColorScale(ws::Worksheet, rng::RowRange; kw...) = process_rowranges(setCfColorScale, ws, rng; kw...)
setCfColorScale(ws::Worksheet, rng::ColumnRange; kw...) = process_columnranges(setCfColorScale, ws, rng; kw...)
setCfColorScale(xl::XLSXFile, sheetcell::AbstractString; kw...)::Int = process_sheetcell(setCfColorScale, xl, sheetcell; kw...)
setCfColorScale(ws::Worksheet, ref_or_rng::AbstractString; kw...)::Int = process_ranges(setCfColorScale, ws, ref_or_rng; kw...)
function setCfColorScale(ws::Worksheet, rng::CellRange;
    colorScale::Union{Nothing,String}=nothing,
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
    if isnothing(colorScale)
        push!(new_cf, XML.h.cfRule(type="colorScale", priority="1",
            XML.h.colorScale(
                isnothing(min_val) ? XML.h.cfvo(type=min_type) : XML.h.cfvo(type=min_type, val=min_val),
                isnothing(mid_type) ? nothing : XML.h.cfvo(type=mid_type, val=mid_val),
                isnothing(max_val) ? XML.h.cfvo(type=max_type) : XML.h.cfvo(type=max_type, val=max_val),
                XML.h.color(rgb=get_color(min_col)),
                isnothing(mid_type) ? nothing : XML.h.color(rgb=get_color(mid_col)),
                XML.h.color(rgb=get_color(max_col))
            )
        )
        )
    else
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

    return 0
end