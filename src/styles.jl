
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

# get styles document for workbook
function styles_xmlroot(workbook::Workbook)
    if isnull(workbook.styles_xroot)
        STYLES_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
        if has_relationship_by_type(workbook, STYLES_RELATIONSHIP_TYPE)
            styles_target = get_relationship_target_by_type(workbook, STYLES_RELATIONSHIP_TYPE)
            styles_root = xmlroot(get_xlsxfile(workbook), "xl/" * styles_target)

            # check root node name for styles.xml
            @assert get_default_namespace(styles_root) == SPREADSHEET_NAMESPACE_XPATH_ARG[1][2] "Unsupported styles XML namespace $(get_default_namespace(styles_root))."
            @assert EzXML.nodename(styles_root) == "styleSheet" "Malformed package. Expected root node named `styleSheet` in `styles.xml`."
            workbook.styles_xroot = Nullable(styles_root)
        else
            error("Styles not found for this workbook.")
        end
    end

    return get(workbook.styles_xroot)
end

"""
Returns the xf XML node element for style `index`.
`index` is 0-based.
"""
function styles_cell_xf(wb::Workbook, index::Int) :: EzXML.Node
    xroot = styles_xmlroot(wb)
    xf_elements = find(xroot, "/xpath:styleSheet/xpath:cellXfs/xpath:xf", SPREADSHEET_NAMESPACE_XPATH_ARG)
    return xf_elements[index+1]
end

"Queries numFmtId from cellXfs -> xf nodes."
function styles_cell_xf_numFmtId(wb::Workbook, index::Int) :: Int
    el = styles_cell_xf(wb, index)
    return parse(Int, el["numFmtId"])
end

"""
Queries numFmt formatCode field by numFmtId.
"""
function styles_numFmt_formatCode(wb::Workbook, numFmtId::AbstractString) :: String
    xroot = styles_xmlroot(wb)
    elements_found = find(xroot, "/xpath:styleSheet/xpath:numFmts/xpath:numFmt[@numFmtId='$(numFmtId)']", SPREADSHEET_NAMESPACE_XPATH_ARG)
    @assert length(elements_found) == 1 "numFmtId $numFmtId not found."
    return elements_found[1]["formatCode"]
end

styles_numFmt_formatCode(wb::Workbook, numFmtId::Int) = styles_numFmt_formatCode(wb, string(numFmtId))

function styles_is_datetime(wb::Workbook, index::Int) :: Bool
    if !haskey(wb.buffer_styles_is_datetime, index)
        isdatetime = false

        numFmtId = styles_cell_xf_numFmtId(wb, index)

        if (14 <= numFmtId && numFmtId <= 22) || (45 <= numFmtId && numFmtId <= 47)
            isdatetime = true
        elseif numFmtId > 81
            code = styles_numFmt_formatCode(wb, numFmtId)
            if contains(code, "dd") || contains(code, "mm") || contains(code, "yy") || contains(code, "hh") || contains(code, "ss")
                isdatetime = true
            end
        end

        wb.buffer_styles_is_datetime[index] = isdatetime
    end

    return wb.buffer_styles_is_datetime[index]
end

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
            if contains(code, ".0")
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

styles_is_float(ws::Worksheet, index) = styles_is_float(get_workbook(ws), index)

"""

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
"""
function styles_get_cellXf_with_numFmtId(wb::Workbook, numFmtId::Int) :: Int
    xroot = styles_xmlroot(wb)
    elements_found = find(xroot, "/xpath:styleSheet/xpath:cellXfs/xpath:xf", SPREADSHEET_NAMESPACE_XPATH_ARG)

    if isempty(elements_found)
        return -1
    else
        for i in 1:length(elements_found)
            el = elements_found[i]
            if haskey(el, "numFmtId")
                if parse(Int, el["numFmtId"]) == numFmtId
                    return i-1
                end
            end
        end

        # not found
        return -1
    end
end

function styles_add_cell_xf(wb::Workbook, attributes::Dict{String, String}) :: Int
    xroot = styles_xmlroot(wb)
    cellXfs_element = findfirst(xroot, "/xpath:styleSheet/xpath:cellXfs", SPREADSHEET_NAMESPACE_XPATH_ARG)
    existing_cellxf_elements_count = EzXML.countelements(cellXfs_element)

    new_xf = EzXML.addelement!(cellXfs_element, "xf")
    for k in keys(attributes)
        new_xf[k] = attributes[k]
    end

    return existing_cellxf_elements_count # turns out this is the new index
end
