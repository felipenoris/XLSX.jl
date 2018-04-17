
"""
A `CellRef` represents a cell location given by row and column identifiers.

`CellRef("A6")` indicates a cell located at column `1` and row `6`.

Example:

```julia
import XLSX
cn = XLSX.CellRef("AB1")
println( XLSX.row_number(cn) ) # will print 1
println( XLSX.column_number(cn) ) # will print 28
println( string(cn) ) # will print out AB1
```

As a convenience, `@ref_str` macro is provided.

```julia
import XLSX
cn = XLSX.ref"AB1"
println( XLSX.row_number(cn) ) # will print 1
println( XLSX.column_number(cn) ) # will print 28
println( string(cn) ) # will print out AB1
```
"""
struct CellRef
    name::String
    column_name::SubString{String}
    row_number::Int
    column_number::Int
end

struct Cell
    ref::CellRef
    datatype::String
    style::String
    value::String
    formula::String
end

"""
A `CellRange` represents a rectangular range of cells in a spreadsheet.

`CellRange("A1:C4")` denotes cells ranging from `A1` (upper left corner) to `C4` (bottom right corner).

As a convenience, `@range_str` macro is provided.

```julia
import XLSX
cr = XLSX.range"A1:C4"
```
"""
struct CellRange
    start::CellRef
    stop::CellRef
end

abstract type MSOfficePackage end

struct NullPackage <: MSOfficePackage end

#struct EmptyMSOfficePackage <: MSOfficePackage
#end

"""
Relationships are defined in ECMA-376-1 Section 9.2.
This struct matches the `Relationship` tag attribute names.

A `Relashipship` defines relations between the files inside a MSOffice package.
Regarding Spreadsheets, there are two kinds of relationships:

    * package level: defined in `_rels/.rels`.
    * workbook level: defined in `xl/_rels/workbook.xml.rels`.

The function `parse_relationships!(xf::XLSXFile)` is used to parse
package and workbook level relationships.
"""
struct Relationship
    Id::String
    Type::String
    Target::String
end

mutable struct Worksheet
    package::MSOfficePackage # parent XLSXFile
    sheetId::Int
    relationship_id::String # r:id="rId1"
    name::String
    data::LightXML.XMLDocument # a copy of the reference xf.data[worksheet_file], xf :: XLSFile
end

"""
Workbook is the result of parsing file `xl/workbook.xml`.
"""
mutable struct Workbook
    package::MSOfficePackage # parent XLSXFile
    sheets::Vector{Worksheet} # workbook -> sheets -> <sheet name="Sheet1" r:id="rId1" sheetId="1"/>. sheetId determines the index of the WorkSheet in this vector.
    date1904::Bool              # workbook -> workbookPr -> attribute date1904 = "1" or absent
    relationships::Vector{Relationship} # contains workbook level relationships
    sst::Vector{LightXML.XMLElement} # shared string table ("si" elements)
    styles::LightXML.XMLDocument # a copy of the reference xf.data[styles_file]
end

"""
`XLSXFile` stores all XML data from an Excel file.

`filepath` is the filepath of the source file for this XLSXFile.
`data` stored the raw XML data. It maps internal XLSX filenames to XMLDocuments.
`workbook` is the result of parsing `xl/workbook.xml`.
"""
mutable struct XLSXFile <: MSOfficePackage
    filepath::AbstractString
    data::Dict{String, LightXML.XMLDocument} # maps filename => XMLDocument
    workbook::Workbook
    relationships::Vector{Relationship} # contains package level relationships
end
