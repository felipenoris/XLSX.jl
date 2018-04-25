
"""
A `CellRef` represents a cell location given by row and column identifiers.

`CellRef("A6")` indicates a cell located at column `1` and row `6`.

Example:

```julia
cn = XLSX.CellRef("AB1")
println( XLSX.row_number(cn) ) # will print 1
println( XLSX.column_number(cn) ) # will print 28
println( string(cn) ) # will print out AB1
```

As a convenience, `@ref_str` macro is provided.

```julia
cn = XLSX.ref"AB1"
println( XLSX.row_number(cn) ) # will print 1
println( XLSX.column_number(cn) ) # will print 28
println( string(cn) ) # will print out AB1
```
"""
struct CellRef
    name::String
    row_number::Int
    column_number::Int
end

abstract type AbstractCell end

struct Cell <: AbstractCell
    ref::CellRef
    datatype::String
    style::String
    value::String
    formula::String
end

struct EmptyCell <: AbstractCell
    ref::CellRef
end

"""
A `CellRange` represents a rectangular range of cells in a spreadsheet.

`CellRange("A1:C4")` denotes cells ranging from `A1` (upper left corner) to `C4` (bottom right corner).

As a convenience, `@range_str` macro is provided.

```julia
cr = XLSX.range"A1:C4"
```
"""
struct CellRange
    start::CellRef
    stop::CellRef

    function CellRange(a::CellRef, b::CellRef)

        top = row_number(a)
        bottom = row_number(b)
        left = column_number(a)
        right = column_number(b)

        @assert left <= right && top <= bottom "Invalid CellRange. Start cell should be at the top left corner of the range."

        return new(a, b)
    end
end

struct ColumnRange
    start::Int # column number
    stop::Int  # column number

    function ColumnRange(a::Int, b::Int)
        @assert a <= b "Invalid ColumnRange. Start column must be located before end column."
        new(a, b)
    end
end

abstract type MSOfficePackage end

struct EmptyMSOfficePackage <: MSOfficePackage
end

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

# Iterators

"""
    SheetRowIterator(sheet)

Iterates over Worksheet cells. See `eachrow` method docs.
"""
struct SheetRowIterator
    sheet::Worksheet
    xml_rows_iterator::LightXML.XMLElementIter
end

mutable struct SheetRow
    sheet::Worksheet
    row::Int
    row_xml_element::LightXML.XMLElement
    rowcells::Dict{Int, Cell} # column -> value
    is_rowcells_populated::Bool # indicates wether row_xml_element has been decoded into rowcells
end

mutable struct Index # based on DataFrames.jl
    lookup::Dict{Symbol, Int} # name -> table column index
    column_labels::Vector{Symbol}
    column_map::Dict{Int, Int} # table column index (1-based) -> sheet column index (cellref based)

    function Index(column_range::ColumnRange, column_labels::Vector{Symbol})
        lookup = Dict{Symbol, Int}()
        for (i, n) in enumerate(column_labels)
            lookup[n] = i
        end

        column_map = Dict{Int, Int}()
        for (i, n) in enumerate(column_range)
            column_map[i] = decode_column_number(n)
        end
        return new(lookup, column_labels, column_map)
    end
end

struct TableRowIterator
    itr::SheetRowIterator
    index::Index
    first_data_row::Int
end

struct TableRow
    itr::TableRowIterator
    sheet_row::SheetRow
    table_row_index::Int # Index of the row in the table. This is not relative to the worksheet cell row.
end
