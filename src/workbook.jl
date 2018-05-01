
EmptyWorkbook() = Workbook(EmptyMSOfficePackage(), Vector{Worksheet}(), false, Vector{Relationship}(), SharedStrings(), EzXML.XMLDocument(), Dict{Int, Bool}(), Dict{Int, Bool}())

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
    xmldocument(xl::XLSXFile, filename::String) :: EzXML.Document

Utility method to find the XMLDocument associated with a given package filename.
Returns xl.data[filename] if it exists. Throws an error if it doesn't.
"""
function xmldocument(xl::XLSXFile, filename::String) :: EzXML.Document
    @assert in(filename, filenames(xl)) "$filename not found in XLSX package."
    return xl.data[filename]
end

"""
    xmlroot(xl::XLSXFile, filename::String) :: EzXML.Node

Utility method to return the root element of a given XMLDocument from the package.
Returns EzXML.root(xl.data[filename]) if it exists.
"""
xmlroot(xl::XLSXFile, filename::String) :: EzXML.Node = EzXML.root(xmldocument(xl, filename))

function getsheet(xl::XLSXFile, sheetname::String) :: Worksheet
    for ws in xl.workbook.sheets
        if ws.name == sheetname
            return ws
        end
    end
    error("$(xl.filepath) does not have a Worksheet named $sheetname.")
end

getsheet(xl::XLSXFile, sheet_index::Int) :: Worksheet = xl.workbook.sheets[sheet_index]
getsheet(filepath::AbstractString, s) = getsheet(read(filepath), s)

Base.getindex(xl::XLSXFile, s) = getsheet(xl, s)

Base.show(io::IO, xf::XLSXFile) = print(io, "XLSXFile(\"$(xf.filepath)\")")
