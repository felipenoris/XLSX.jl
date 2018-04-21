
EmptyWorkbook() = Workbook(EmptyMSOfficePackage(), Vector{Worksheet}(), false, Vector{Relationship}(), Vector{LightXML.XMLElement}(), LightXML.XMLDocument())

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
