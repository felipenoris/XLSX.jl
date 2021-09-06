
# References
#
# 18.8.30 numFmt (Number Format)
# lists a table with predefined numFmtId.
# In this case, a numFmtId value is written on the xf record,
# but no corresponding numFmt element is written.
# Some of these Ids can be interpreted differently,
# depending on the UI language of the implementing application.
#
# General formatCodes ranges from 0 to 81.
#
# 18.8.10 cellXfs (Cell Formats)
import Base: isempty

function CellValue(ws::Worksheet, val::CellValueType)
    CellValue(val, default_cell_format(ws, val))
end

id(format::CellDataFormat) = string(format.id)
id(::EmptyCellDataFormat) = ""
isempty(::CellDataFormat) = false
isempty(::EmptyCellDataFormat) = true

# The number of predefined number formats in XLSX
# Any custom number formats must have an id >= PREDEFINED_NUMFMT_COUNT
const PREDEFINED_NUMFMT_COUNT = 164

# these formats may appear differently in different editors
const DEFAULT_DATE_numFmtId = 14 # dd-mm-yyyy
const DEFAULT_TIME_numFmtId = 20 # h:mm
const DEFAULT_DATETIME_numFmtId = 22 # dd-mm-yyyy h:mm

# Returns the default `CellDataFormat` for a type
default_cell_format(::Worksheet, ::CellValueType) = EmptyCellDataFormat()
default_cell_format(ws::Worksheet, ::Dates.Date) = get_num_style_index(ws, DEFAULT_DATE_numFmtId)
default_cell_format(ws::Worksheet, ::Dates.Time) = get_num_style_index(ws, DEFAULT_TIME_numFmtId)
default_cell_format(ws::Worksheet, ::Dates.DateTime) = get_num_style_index(ws, DEFAULT_DATETIME_numFmtId)

# Attempts to get CellDataFormat associated with a numFmtId and sets a default style if it is not found
# Use for ensuring default formats exist
function get_num_style_index(ws::Worksheet, numformatid::Integer)
    @assert numformatid >= 0 "Invalid number format id"

    wb = get_workbook(ws)
    style_index = styles_get_cellXf_with_numFmtId(wb, numformatid)
    if isempty(style_index)
        # adds default style <xf applyNumberFormat="1" borderId="0" fillId="0" fontId="0" numFmtId=formatid xfId="0"/>
        style_index = styles_add_cell_xf(wb, Dict("applyNumberFormat"=>"1", "borderId"=>"0", "fillId"=>"0", "fontId"=>"0", "numFmtId"=>string(numformatid), "xfId"=>"0"))
    end

    return style_index
end

# get styles document for workbook
function styles_xmlroot(workbook::Workbook)
    if workbook.styles_xroot == nothing
        STYLES_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
        if has_relationship_by_type(workbook, STYLES_RELATIONSHIP_TYPE)
            styles_target = get_relationship_target_by_type("xl", workbook, STYLES_RELATIONSHIP_TYPE)
            styles_root = xmlroot(get_xlsxfile(workbook), styles_target)

            # check root node name for styles.xml
            @assert get_default_namespace(styles_root) == SPREADSHEET_NAMESPACE_XPATH_ARG[1][2] "Unsupported styles XML namespace $(get_default_namespace(styles_root))."
            @assert EzXML.nodename(styles_root) == "styleSheet" "Malformed package. Expected root node named `styleSheet` in `styles.xml`."
            workbook.styles_xroot = styles_root
        else
            error("Styles not found for this workbook.")
        end
    end

    return workbook.styles_xroot
end

# Returns the xf XML node element for style `index`.
# `index` is 0-based.
function styles_cell_xf(wb::Workbook, index::Int) :: EzXML.Node
    xroot = styles_xmlroot(wb)
    xf_elements = findall("/xpath:styleSheet/xpath:cellXfs/xpath:xf", xroot, SPREADSHEET_NAMESPACE_XPATH_ARG)
    return xf_elements[index+1]
end

# Queries numFmtId from cellXfs -> xf nodes."
function styles_cell_xf_numFmtId(wb::Workbook, index::Int) :: Int
    el = styles_cell_xf(wb, index)
    if !haskey(el, "numFmtId")
        return 0
    end
    return parse(Int, el["numFmtId"])
end

# Defines a custom number format to render numbers, dates or text.
# Returns the index to be used as the `numFmtId` in a cellXf definition.
function styles_add_numFmt(wb::Workbook, format_code::AbstractString) :: Integer
    xroot = styles_xmlroot(wb)

    numfmts = findall("/xpath:styleSheet/xpath:numFmts", xroot, SPREADSHEET_NAMESPACE_XPATH_ARG)
    if isempty(numfmts)
        stylesheet = findfirst("/xpath:styleSheet", xroot, SPREADSHEET_NAMESPACE_XPATH_ARG)

        # We need to add the numFmts node directly after the styleSheet node
        child = EzXML.firstelement(stylesheet)
        numfmts = EzXML.addelement!(stylesheet, "numFmts")
        EzXML.unlink!(numfmts)
        EzXML.linkprev!(child, numfmts)
    else
        numfmts = numfmts[1]
    end

    existing_numFmt_elements_count = EzXML.countelements(numfmts)
    new_fmt = EzXML.addelement!(numfmts, "numFmt")

    fmt_code = existing_numFmt_elements_count + PREDEFINED_NUMFMT_COUNT
    new_fmt["numFmtId"] = fmt_code
    new_fmt["formatCode"] = xlsx_escape(format_code)

    return fmt_code
end

const FontAttribute = Union{AbstractString, Pair{String, Pair{String, String}}}

# Defines a custom font. Returns the index to be used as the `fontId` in a cellXf definition.
function styles_add_font(wb::Workbook, attributes::Vector{FontAttribute})
    xroot = styles_xmlroot(wb)
    fonts_element = findfirst("/xpath:styleSheet/xpath:fonts", xroot, SPREADSHEET_NAMESPACE_XPATH_ARG)
    existing_font_elements_count = EzXML.countelements(fonts_element)

    new_font = EzXML.addelement!(fonts_element, "font")
    for a in attributes
        if a isa Pair
            name, val = last(a)
            attr = EzXML.addelement!(new_font, first(a))
            attr[name] = val
        else
            EzXML.addelement!(new_font, a)
        end
    end

    return existing_font_elements_count
end


# Queries numFmt formatCode field by numFmtId.
function styles_numFmt_formatCode(wb::Workbook, numFmtId::AbstractString) :: String
    xroot = styles_xmlroot(wb)
    elements_found = findall("/xpath:styleSheet/xpath:numFmts/xpath:numFmt[@numFmtId='$(numFmtId)']", xroot, SPREADSHEET_NAMESPACE_XPATH_ARG)
    @assert length(elements_found) == 1 "numFmtId $numFmtId not found."
    return elements_found[1]["formatCode"]
end

styles_numFmt_formatCode(wb::Workbook, numFmtId::Int) = styles_numFmt_formatCode(wb, string(numFmtId))

const DATETIME_CODES = ["d", "m", "yy", "h", "s", "a/p", "am/pm"]

function remove_formatting(code::AbstractString)
    # this regex should cover all the formatting cases found here(colors/conditionals/quotes/spacing):
    # https://support.office.com/en-us/article/create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4
    ignoredformatting = r"""
        \[.{2,}?\]|
        &quot;.+&quot;|
        _.|
        \\.|
        \*.
        """x
    replace(code, ignoredformatting => "")
end

function styles_is_datetime(wb::Workbook, index::Int) :: Bool
    if !haskey(wb.buffer_styles_is_datetime, index)
        isdatetime = false

        numFmtId = styles_cell_xf_numFmtId(wb, index)

        if (14 <= numFmtId && numFmtId <= 22) || (45 <= numFmtId && numFmtId <= 47)
            isdatetime = true
        elseif numFmtId > 81
            code = lowercase(styles_numFmt_formatCode(wb, numFmtId))
            code = remove_formatting(code)
            if any(map(x->occursin(x, code), DATETIME_CODES))
                isdatetime = true
            end
        end

        wb.buffer_styles_is_datetime[index] = isdatetime
    end

    return wb.buffer_styles_is_datetime[index]
end

styles_is_datetime(wb::Workbook, fmt::CellDataFormat) = styles_is_datetime(wb, Int(fmt.id))

function styles_is_datetime(wb::Workbook, index::AbstractString)
    @assert !isempty(index)
    styles_is_datetime(wb, parse(Int, index))
end

styles_is_datetime(ws::Worksheet, index) = styles_is_datetime(get_workbook(ws), index)

function styles_is_float(wb::Workbook, index::Int) :: Bool
    if !haskey(wb.buffer_styles_is_float, index)
        isfloat = false
        numFmtId = styles_cell_xf_numFmtId(wb, index)

        if numFmtId == 2 || numFmtId == 4 || (7 <= numFmtId && numFmtId <= 11) || numFmtId == 39 || numFmtId == 40 || numFmtId == 44 || numFmtId == 48
            isfloat = true
        elseif numFmtId > 81
            code = styles_numFmt_formatCode(wb, numFmtId)
            code = remove_formatting(code)

            floatformats = r"""
                \.[0#?]|
                [0#?]e[+-]?[0#?]|
                [0#?]/[0#?]|
                %
                """ix
            if occursin(floatformats, code)
                isfloat = true
            end
        end

        wb.buffer_styles_is_float[index] = isfloat
    end

    return wb.buffer_styles_is_float[index]
end

function styles_is_float(wb::Workbook, index::AbstractString)
    @assert !isempty(index)
    styles_is_float(wb, parse(Int, index))
end

styles_is_float(wb::Workbook, fmt::CellDataFormat) = styles_is_float(wb, Int(fmt.id))
styles_is_float(ws::Worksheet, index) = styles_is_float(get_workbook(ws), index)

#=
Cell Xf element follows the XML format below.
This function queries the 0-based index of the first xf element that has the provided numFmtId.
Returns -1 if not found.

```
<styleSheet ...
    <cellXfs count="5">
            <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
            <xf applyNumberFormat="1" borderId="0" fillId="0" fontId="0" numFmtId="14" xfId="0"/>
            <xf applyNumberFormat="1" borderId="0" fillId="0" fontId="0" numFmtId="20" xfId="0"/>
            <xf applyNumberFormat="1" borderId="0" fillId="0" fontId="0" numFmtId="22" xfId="0"/>
```
=#
function styles_get_cellXf_with_numFmtId(wb::Workbook, numFmtId::Int) :: AbstractCellDataFormat
    xroot = styles_xmlroot(wb)
    elements_found = findall("/xpath:styleSheet/xpath:cellXfs/xpath:xf", xroot, SPREADSHEET_NAMESPACE_XPATH_ARG)

    if isempty(elements_found)
        return EmptyCellDataFormat()
    else
        for i in 1:length(elements_found)
            el = elements_found[i]
            if haskey(el, "numFmtId")
                if parse(Int, el["numFmtId"]) == numFmtId
                    return CellDataFormat(i-1)
                end
            end
        end

        # not found
        return EmptyCellDataFormat()
    end
end

function styles_add_cell_xf(wb::Workbook, attributes::Dict{String, String}) :: CellDataFormat
    xroot = styles_xmlroot(wb)
    cellXfs_element = findfirst("/xpath:styleSheet/xpath:cellXfs", xroot, SPREADSHEET_NAMESPACE_XPATH_ARG)
    existing_cellxf_elements_count = EzXML.countelements(cellXfs_element)

    new_xf = EzXML.addelement!(cellXfs_element, "xf")
    for k in keys(attributes)
        new_xf[k] = attributes[k]
    end

    return CellDataFormat(existing_cellxf_elements_count) # turns out this is the new index
end
