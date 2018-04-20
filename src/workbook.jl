
EmptyWorkbook() = Workbook(NullPackage(), Vector{Worksheet}(), false, Vector{Relationship}(), Vector{LightXML.XMLElement}(), LightXML.XMLDocument())

"""
Lists internal files from the XLSX package.
"""
filenames(xl::XLSXFile) = keys(xl.data)

"""
Lists Worksheet names for this Workbook.
"""
sheetnames(wb::Workbook) = [ s.name for s in wb.sheets ]
sheetnames(xl::XLSXFile) = sheetnames(xl.workbook)

"""
Counts the number of sheets in the Workbook.
"""
sheetcount(wb::Workbook) = length(wb.sheets)
sheetcount(xl::XLSXFile) = sheetcount(xl.workbook)

"""
    isdate1904(wb) :: Bool

Returns true if workbook follows date1904 convention.
"""
isdate1904(wb::Workbook) :: Bool = wb.date1904
isdate1904(xf::XLSXFile) :: Bool = isdate1904(xf.workbook)

"""
    xmldocument(xl::XLSXFile, filename::String) :: LightXML.XMLDocument

Utility method to find the XMLDocument associated with a given package filename.
Returns xl.data[filename] if it exists. Throws an error if it doesn't.
"""
function xmldocument(xl::XLSXFile, filename::String) :: LightXML.XMLDocument
    @assert in(filename, filenames(xl)) "$filename not found in XLSX package."
    return xl.data[filename]
end

"""
    xmlroot(xl::XLSXFile, filename::String) :: LightXML.XMLElement

Utility method to return the root element of a given XMLDocument from the package.
Returns LightXML.root(xl.data[filename]) if it exists.
"""
xmlroot(xl::XLSXFile, filename::String) :: LightXML.XMLElement = LightXML.root(xmldocument(xl, filename))

"""
  parse_workbook!(xf::XLSXFile)

Updates xf.workbook from xf.data[\"xl/workbook.xml\"]
"""
function parse_workbook!(xf::XLSXFile)
    xroot = xmlroot(xf, "xl/workbook.xml")
    @assert LightXML.name(xroot) == "workbook" "Malformed xl/workbook.xml. Root node name should be 'workbook'. Got '$(LightXML.name(xroot))'."

    # workbook to be parsed
    workbook = xf.workbook

    # workbookPr
    vec_workbookPr = xroot["workbookPr"]
    if length(vec_workbookPr) > 0
        @assert length(vec_workbookPr) == 1 "Malformed workbook. $xf has more than 1 workbookPr nodes in xl/workbook.xml."

        workbookPr_element = vec_workbookPr[1]
        if LightXML.has_attribute(workbookPr_element, "date1904")
            attribute_value_date1904 = LightXML.attribute(workbookPr_element, "date1904")

            if attribute_value_date1904 == "1" || attribute_value_date1904 == "true"
                workbook.date1904 = true
            elseif attribute_value_date1904 == "0" || attribute_value_date1904 == "false"
                workbook.date1904 = false
            else
                error("Could not parse xl/workbook -> workbookPr -> date1904 = $(attribute_value_date1904).")
            end
        else
            # does not have attribute => is not date1904
            workbook.date1904 = false
        end
    end

    # shared string table
    SHARED_STRINGS_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    if has_relationship_by_type(workbook, SHARED_STRINGS_RELATIONSHIP_TYPE)
        sst_root = xmlroot(xf, "xl/" * get_relationship_target_by_type(workbook, SHARED_STRINGS_RELATIONSHIP_TYPE))
        @assert LightXML.name(sst_root) == "sst" "Malformed workbook. sst file should have sst root."
        workbook.sst = sst_root["si"]
    end

    # sheets
    vec_sheets = xroot["sheets"]
    if length(vec_sheets) > 0
        @assert length(vec_sheets) == 1 "Malformed workbook. $xf has more than 1 sheet node in xl/workbook.xml."

        sheets_element = vec_sheets[1]

        vec_sheet = sheets_element["sheet"]
        num_sheets = length(vec_sheet)
        workbook.sheets = Vector{Worksheet}(num_sheets)

        for (index, sheet_element) in enumerate(vec_sheet)
            worksheet = Worksheet(xf, sheet_element)
            workbook.sheets[index] = worksheet
        end
    end

    # styles
    STYLES_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    if has_relationship_by_type(workbook, STYLES_RELATIONSHIP_TYPE)
        styles_target = get_relationship_target_by_type(workbook, STYLES_RELATIONSHIP_TYPE)
        workbook.styles = xmldocument(xf, "xl/" * styles_target)

        # check root node name for styles.xml
        styles_root = LightXML.root(workbook.styles)
        @assert LightXML.name(styles_root) == "styleSheet" "Malformed package. Expected root node named `styleSheet` in `styles.xml`."
    end

    nothing
end

function Worksheet(xf::XLSXFile, sheet_element::LightXML.XMLElement)
    @assert LightXML.name(sheet_element) == "sheet"

    sheetId = parse(Int, LightXML.attribute(sheet_element, "sheetId"))
    relationship_id = LightXML.attribute(sheet_element, "id")
    name = LightXML.attribute(sheet_element, "name")

    target = "xl/" * get_relationship_target_by_id(xf.workbook, relationship_id)
    sheet_data = xf.data[target]

    return Worksheet(xf, sheetId, relationship_id, name, sheet_data)
end

function Base.getindex(xl::XLSXFile, sheetname::String) :: Worksheet
    for ws in xl.workbook.sheets
        if ws.name == sheetname
            return ws
        end
    end
    error("$(xl.filepath) does not have a Worksheet named $sheetname.")
end

Base.getindex(xl::XLSXFile, sheet_index::Int) :: Worksheet = xl.workbook.sheets[sheet_index]

Base.show(io::IO, xf::XLSXFile) = print(io, "XLSXFile(\"$(xf.filepath)\")")
