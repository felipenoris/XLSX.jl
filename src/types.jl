
struct CellPosition
    row::Int
    column::Int
end

#=
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
=#
struct CellRef
    name::String
    row_number::Int
    column_number::Int
end

abstract type AbstractFormula end

"""
A default formula simply storing the formula string.
"""
struct Formula <: AbstractFormula
    formula::String
end

"""
The formula in this cell was defined somewhere else; we simply reference its ID.
"""
struct FormulaReference <: AbstractFormula
    id::Int
end

"""
Formula that is defined once and referenced in all cells given by the cell range given in `ref`.
"""
struct ReferencedFormula <: AbstractFormula
    formula::String
    id::Int
    ref::String # actually a CellRange, but defined later --> change if at some point we want to actively change formulae
end

abstract type AbstractCell end

mutable struct Cell <: AbstractCell
    ref::CellRef
    datatype::String
    style::String
    value::String
    formula::AbstractFormula
end

struct EmptyCell <: AbstractCell
    ref::CellRef
end

abstract type AbstractCellDataFormat end

struct EmptyCellDataFormat <: AbstractCellDataFormat end

# Keeps track of formatting information.
struct CellDataFormat <: AbstractCellDataFormat
    id::UInt
end

"""
    CellValueType

Concrete supported data-types.

```julia
Union{String, Missing, Float64, Int, Bool, Dates.Date, Dates.Time, Dates.DateTime}
```
"""
const CellValueType = Union{String, Missing, Float64, Int, Bool, Dates.Date, Dates.Time, Dates.DateTime}

# CellValue is a Julia type of a value read from a Spreadsheet.
struct CellValue
    value::CellValueType
    styleid::AbstractCellDataFormat
end

#=
A `CellRange` represents a rectangular range of cells in a spreadsheet.

`CellRange("A1:C4")` denotes cells ranging from `A1` (upper left corner) to `C4` (bottom right corner).

As a convenience, `@range_str` macro is provided.

```julia
cr = XLSX.range"A1:C4"
```
=#
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
        return new(a, b)
    end
end

struct SheetCellRef
    sheet::String
    cellref::CellRef
end

struct SheetCellRange
   sheet::String
   rng::CellRange
end

struct SheetColumnRange
    sheet::String
    colrng::ColumnRange
end

abstract type MSOfficePackage end

struct EmptyMSOfficePackage <: MSOfficePackage
end

#=
Relationships are defined in ECMA-376-1 Section 9.2.
This struct matches the `Relationship` tag attribute names.

A `Relashipship` defines relations between the files inside a MSOffice package.
Regarding Spreadsheets, there are two kinds of relationships:

    * package level: defined in `_rels/.rels`.
    * workbook level: defined in `xl/_rels/workbook.xml.rels`.

The function `parse_relationships!(xf::XLSXFile)` is used to parse
package and workbook level relationships.
=#
struct Relationship
    Id::String
    Type::String
    Target::String
end

const CellCache = Dict{Int, Dict{Int, Cell}} # row -> ( column -> cell )

#=
Iterates over Worksheet cells. See `eachrow` method docs.
Each element is a `SheetRow`.

Implementations: SheetRowStreamIterator, WorksheetCache.
=#
abstract type SheetRowIterator end

mutable struct SheetRowStreamIteratorState
    zip_io::ZipFile.Reader
    xml_stream_reader::EzXML.StreamReader
    is_open::Bool # indicated if zip_io and xml_stream_reader are opened
    row::Int # number of current row. It´s set to 0 in the start state.
end

mutable struct WorksheetCache{I<:SheetRowIterator} <: SheetRowIterator
    cells::CellCache # SheetRowNumber -> Dict{column_number, Cell}
    rows_in_cache::Vector{Int} # ordered vector with row numbers that are stored in cache
    row_index::Dict{Int, Int} # maps a row number to the index of the row number in rows_in_cache
    stream_iterator::I
    stream_state::Union{Nothing, SheetRowStreamIteratorState}
    dirty::Bool #indicate that data are not sorted, avoid sorting if we dont use the iterator
end

"""
A `Worksheet` represents a reference to an Excel Worksheet.

From a `Worksheet` you can query for Cells, cell values and ranges.

# Example

```julia
xf = XLSX.readxlsx("myfile.xlsx")
sh = xf["mysheet"] # get a reference to a Worksheet
println( sh[2, 2] ) # access element "B2" (2nd row, 2nd column)
println( sh["B2"] ) # you can also use the cell name
println( sh["A2:B4"] ) # or a cell range
println( sh[:] ) # all data inside worksheet's dimension
```
"""
mutable struct Worksheet
    package::MSOfficePackage # parent XLSXFile
    sheetId::Int
    relationship_id::String # r:id="rId1"
    name::String
    dimension::Union{Nothing, CellRange}
    is_hidden::Bool
    cache::Union{WorksheetCache, Nothing}

    function Worksheet(package::MSOfficePackage, sheetId::Int, relationship_id::String, name::String, dimension::Union{Nothing, CellRange}, is_hidden::Bool)
        return new(package, sheetId, relationship_id, name, dimension, is_hidden, nothing)
    end
end

struct SheetRowStreamIterator <: SheetRowIterator
    sheet::Worksheet
end

mutable struct SharedStringTable
    unformatted_strings::Vector{String}
    formatted_strings::Vector{String}
    index::Dict{String, Int64} # for unformatted_strings search optimisation
    is_loaded::Bool # for lazy-loading of sst XML file (implies that this struct must be mutable)
end

const DefinedNameValueTypes = Union{SheetCellRef, SheetCellRange, Int, Float64, String, Missing}

# Workbook is the result of parsing file `xl/workbook.xml`.
mutable struct Workbook
    package::MSOfficePackage # parent XLSXFile
    sheets::Vector{Worksheet} # workbook -> sheets -> <sheet name="Sheet1" r:id="rId1" sheetId="1"/>. sheetId determines the index of the WorkSheet in this vector.
    date1904::Bool              # workbook -> workbookPr -> attribute date1904 = "1" or absent
    relationships::Vector{Relationship} # contains workbook level relationships
    sst::SharedStringTable # shared string table
    buffer_styles_is_float::Dict{Int, Bool}      # cell style -> true if is float
    buffer_styles_is_datetime::Dict{Int, Bool}   # cell style -> true if is datetime
    workbook_names::Dict{String, DefinedNameValueTypes} # definedName
    worksheet_names::Dict{Tuple{Int, String}, DefinedNameValueTypes} # definedName. (sheetId, name) -> value.
    styles_xroot::Union{EzXML.Node, Nothing}
end

"""
`XLSXFile` represents a reference to an Excel file.

It is created by using [`XLSX.readxlsx`](@ref) or [`XLSX.openxlsx`](@ref).

From a `XLSXFile` you can navigate to a `XLSX.Worksheet` reference
as shown in the example below.

# Example

```julia
xf = XLSX.readxlsx("myfile.xlsx")
sh = xf["mysheet"] # get a reference to a Worksheet
```
"""
mutable struct XLSXFile <: MSOfficePackage
    source::Union{AbstractString, IO}
    use_cache_for_sheet_data::Bool # indicates whether Worksheet.cache will be fed while reading worksheet cells.
    io::ZipFile.Reader
    io_is_open::Bool
    files::Dict{String, Bool} # maps filename => isread bool
    data::Dict{String, EzXML.Document} # maps filename => XMLDocument
    binary_data::Dict{String, Vector{UInt8}} # maps filename => file content in bytes
    workbook::Workbook
    relationships::Vector{Relationship} # contains package level relationships
    is_writable::Bool # indicates whether this XLSX file can be edited

    function XLSXFile(source::Union{AbstractString, IO}, use_cache::Bool, is_writable::Bool)
        check_for_xlsx_file_format(source)
        io = ZipFile.Reader(source)
        xl = new(source, use_cache, io, true, Dict{String, Bool}(), Dict{String, EzXML.Document}(), Dict{String, Vector{UInt8}}(), EmptyWorkbook(), Vector{Relationship}(), is_writable)
        xl.workbook.package = xl
        finalizer(close, xl)
        return xl
    end
end

#
# Iterators
#

struct SheetRow
    sheet::Worksheet
    row::Int
    rowcells::Dict{Int, Cell} # column -> value
end

struct Index # based on DataFrames.jl
    lookup::Dict{Symbol, Int} # column label -> table column index
    column_labels::Vector{Symbol}
    column_map::Dict{Int, Int} # table column index (1-based) -> sheet column index (cellref based)

    function Index(column_range::Union{ColumnRange, AbstractString}, column_labels)
        column_labels_as_syms = [ Symbol(i) for i in column_labels ]
        column_range = convert(ColumnRange, column_range)
        @assert length(unique(column_labels_as_syms)) == length(column_labels_as_syms) "Column labels must be unique."

        lookup = Dict{Symbol, Int}()
        for (i, n) in enumerate(column_labels_as_syms)
            lookup[n] = i
        end

        column_map = Dict{Int, Int}()
        for (i, n) in enumerate(column_range)
            column_map[i] = decode_column_number(n)
        end
        return new(lookup, column_labels_as_syms, column_map)
    end
end

struct TableRowIterator{I<:SheetRowIterator}
    itr::I
    index::Index
    first_data_row::Int
    stop_in_empty_row::Bool
    stop_in_row_function::Union{Nothing, Function}
    keep_empty_rows::Bool
end

struct TableRow
    row::Int # Index of the row in the table. This is not relative to the worksheet cell row.
    index::Index
    cell_values::Vector{CellValueType}
end

struct TableRowIteratorState{S}
    table_row_index::Int
    sheet_row_index::Int
    sheet_row_iterator_state::S
end

struct DataTable
    data::Vector{Any} # columns
    column_labels::Vector{Symbol}
    column_label_index::Dict{Symbol, Int} # column_label -> column_index

    function DataTable(
            data::Vector{Any}, # columns
            column_labels::Vector{Symbol},
        )

        @assert length(data) == length(column_labels) "data has $(length(data)) columns but $(length(column_labels)) column labels."

        column_label_index = Dict{Symbol, Int}()
        for (i, sym) in enumerate(column_labels)
            @assert !haskey(column_label_index, sym) "DataTable has repeated label for column `$sym`"
            column_label_index[sym] = i
        end

        return new(data, column_labels, column_label_index)
    end
end
