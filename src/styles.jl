
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

const STYLES_NAMESPACE_XPATH_ARG = [ "xpath" => "http://schemas.openxmlformats.org/spreadsheetml/2006/main" ]

# get styles document for workbook
function styles_xmlroot(workbook::Workbook)
    if isnull(workbook.styles_xroot)
        STYLES_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
        if has_relationship_by_type(workbook, STYLES_RELATIONSHIP_TYPE)
            styles_target = get_relationship_target_by_type(workbook, STYLES_RELATIONSHIP_TYPE)
            styles_root = xmlroot(get_xlsxfile(workbook), "xl/" * styles_target)

            # check root node name for styles.xml
            @assert get_default_namespace(styles_root) == STYLES_NAMESPACE_XPATH_ARG[1][2] "Unsupported styles XML namespace $(get_default_namespace(styles_root))."
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
    xf_elements = find(xroot, "/xpath:styleSheet/xpath:cellXfs/xpath:xf", STYLES_NAMESPACE_XPATH_ARG)
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
    elements_found = find(xroot, "/xpath:styleSheet/xpath:numFmts/xpath:numFmt[@numFmtId='$(numFmtId)']", STYLES_NAMESPACE_XPATH_ARG)
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
