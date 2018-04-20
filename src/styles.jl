
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

_styles_root(wb::Workbook) = LightXML.root(wb.styles)

"""
Returns the xf XMLElement for style `index`.
`index` is 0-based.
"""
function styles_cell_xf(wb::Workbook, index::Int) :: LightXML.XMLElement
    xroot = _styles_root(wb)
    return xroot["cellXfs"][1]["xf"][index+1]
end

styles_cell_xf(wb::Workbook, index::AbstractString) = styles_cell_xf(wb, parse(Int, index))
styles_cell_xf(ws::Worksheet, cell::Cell) = styles_cell_xf(ws.package.workbook, cell)

function styles_cell_xf(wb::Workbook, cell::Cell)
    @assert !isempty(cell.style) "Cell $(cell.ref.name) has empty style."
    styles_cell_xf(wb, cell.style)
end

"Queries numFmtId from cellXfs -> xf nodes."
function styles_cell_xf_numFmtId(wb::Workbook, index::Int) :: Int
    el = styles_cell_xf(wb, index)
    return parse(Int, LightXML.attribute(el, "numFmtId"))
end

"""
Queries numFmt formatCode field by numFmtId.
"""
function styles_numFmt_formatCode(wb::Workbook, numFmtId::AbstractString) :: String
    xroot = _styles_root(wb)

    for el in xroot["numFmts"][1]["numFmt"]
        if numFmtId == LightXML.attribute(el, "numFmtId")
            return LightXML.attribute(el, "formatCode")
        end
    end

    error("numFmtId $numFmtId not found.")
end

styles_numFmt_formatCode(wb::Workbook, numFmtId::Int) = styles_numFmt_formatCode(wb, string(numFmtId))

function styles_is_datetime(wb::Workbook, index::Int) :: Bool
    numFmtId = styles_cell_xf_numFmtId(wb, index)

    if (14 <= numFmtId && numFmtId <= 22) || (45 <= numFmtId && numFmtId <= 47)
        return true
    end

    if numFmtId > 81
        code = styles_numFmt_formatCode(wb, numFmtId)
        if contains(code, "dd") || contains(code, "mm") || contains(code, "yy") || contains(code, "hh") || contains(code, "ss")
            return true
        end
    end

    return false
end

function styles_is_datetime(wb::Workbook, index::AbstractString)
    @assert !isempty(index)
    styles_is_datetime(wb, parse(Int, index))
end

styles_is_datetime(ws::Worksheet, index) = styles_is_datetime(ws.package.workbook, index)

function styles_is_float(wb::Workbook, index::Int) :: Bool
    numFmtId = styles_cell_xf_numFmtId(wb, index)

    if numFmtId == 2 || numFmtId == 4 || (7 <= numFmtId && numFmtId <= 11) || numFmtId == 39 || numFmtId == 40 || numFmtId == 44 || numFmtId == 48
        return true
    end

    if numFmtId > 81
        code = styles_numFmt_formatCode(wb, numFmtId)
        if contains(code, ".0")
            return true
        end
    end

    return false
end

function styles_is_float(wb::Workbook, index::AbstractString)
    @assert !isempty(index)
    styles_is_float(wb, parse(Int, index))
end

styles_is_float(ws::Worksheet, index) = styles_is_float(ws.package.workbook, index)
