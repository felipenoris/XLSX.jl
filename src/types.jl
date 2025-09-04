
struct CellPosition
    row::Int
    column::Int
end

"""
    CellRef(n::AbstractString)
    CellRef(row::Int, col::Int)

A `CellRef` represents a cell location given by row and column identifiers.

`CellRef("B6")` indicates a cell located at column `2` and row `6`.

These row and column integers can also be passed directly to the `CellRef` constructor: `CellRef(6,2) == CellRef("B6")`.

Finally, a convenience macro `@ref_str` is provided: `ref"B6" == CellRef("B6")`.

# Examples

```julia
cn = XLSX.CellRef("AB1")
println( XLSX.row_number(cn) ) # will print 1
println( XLSX.column_number(cn) ) # will print 28
println( string(cn) ) # will print out AB1

cn = XLSX.CellRef(1, 28)
println( XLSX.row_number(cn) ) # will print 1
println( XLSX.column_number(cn) ) # will print 28
println( string(cn) ) # will print out AB1

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

abstract type AbstractCellDataFormat end

struct EmptyCellDataFormat <: AbstractCellDataFormat end

# Keeps track of formatting information.
struct CellDataFormat <: AbstractCellDataFormat
    id::UInt
end

abstract type AbstractFormula end

"""
A default formula simply storing the formula string.
"""
mutable struct Formula <: AbstractFormula
    formula::String
    unhandled::Union{Dict{String,String},Nothing}
end
function Formula()
    return Formula("", nothing)
end
function Formula(s::String)
    return Formula(s, nothing)
end


"""
The formula in this cell was defined somewhere else; we simply reference its ID.
"""
mutable struct FormulaReference <: AbstractFormula
    id::Int
    unhandled::Union{Dict{String,String},Nothing}
end

"""
Formula that is defined once and referenced in all cells given by the cell range given in `ref` and using the same `id`.
"""
mutable struct ReferencedFormula <: AbstractFormula
    formula::String
    id::Int
    ref::String # actually a CellRange, but defined later --> change if at some point we want to actively change formulae
    unhandled::Union{Dict{String,String},Nothing}
end

struct CellFormula <: AbstractFormula
    value::T where T<:AbstractFormula
    styleid::AbstractCellDataFormat
end


mutable struct CellFont
    fontId::Int
    font::Dict{String, Union{Dict{String, String}, Nothing}} # fontAttribute -> (attribute -> value)
    applyFont::String

    function CellFont(fontid::Int, font::Dict{String, Union{Dict{String, String}, Nothing}}, applyFont::String)
        return new(fontid, font, applyFont)
    end
end

# A border postion element (e.g. `top` or `left`) has a style attribute, but `color` is a child element.
# The `color` element has an attribute (e.g. `rgb`) that defines the color of the border.
# These are both stored in the `border` field of `CellBorder`. The key for the color element
# will vary depending on how the color is defined (e.g. `rgb`, `indexed`, `auto`, etc.).
# Thus, for example, `"top" => Dict("style" => "thin", "rgb" => "FF000000")`
mutable struct CellBorder
    borderId::Int
    border::Dict{String, Union{Dict{String, String}, Nothing}} # borderAttribute -> (attribute -> value)
    applyBorder::String

    function CellBorder(borderid::Int, border::Dict{String, Union{Dict{String, String}, Nothing}}, applyBorder::String)
        return new(borderid, border, applyBorder)
    end
end

# A fill has a pattern type attribute and two children fgColor and bgColor, each with 
# one or two attributes of their own. These color attributes are pushed in to the Dict 
# of attributes with either `fg` or `bg` prepended to their name to support later 
# reconstruction of the xml element.
mutable struct CellFill
    fillId::Int
    fill::Dict{String, Union{Dict{String, String}, Nothing}} # fillAttribute -> (attribute -> value)
    applyFill::String

    function CellFill(fillid::Int, fill::Dict{String, Union{Dict{String, String}, Nothing}}, applyfill::String)
        return new(fillid, fill, applyfill)
    end
end
mutable struct CellFormat
    numFmtId::Int
    format::Dict{String, Union{Dict{String, String}, Nothing}} # fillAttribute -> (attribute -> value)
    applyNumberFormat::String

    function CellFormat(formatid::Int, format::Dict{String, Union{Dict{String, String}, Nothing}}, applynumberformat::String)
        return new(formatid, format, applynumberformat)
    end
end

mutable struct CellAlignment # Alignment is part of the cell style `xf` so doesn't need an Id
    alignment::Dict{String, Union{Dict{String, String}, Nothing}} # alignmentAttribute -> (attribute -> value)
    applyAlignment::String

    function CellAlignment(alignment::Dict{String, Union{Dict{String, String}, Nothing}}, applyalignment::String)
        return new(alignment, applyalignment)
    end
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

# Keeps track of conditional formatting information.
struct DxFormat <: AbstractCellDataFormat
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

abstract type AbstractCellRange end
abstract type ContiguousCellRange <: AbstractCellRange end
abstract type AbstractSheetCellRange <: AbstractCellRange end
abstract type ContiguousSheetCellRange <: AbstractSheetCellRange end

struct CellRange <: ContiguousCellRange
    start::CellRef
    stop::CellRef

    function CellRange(a::CellRef, b::CellRef)

        top = row_number(a)
        bottom = row_number(b)
        left = column_number(a)
        right = column_number(b)

        if left > right || top > bottom
            throw(XLSXError("Invalid CellRange. Start cell should be at the top left corner of the range."))
        end

        return new(a, b)
    end
end

struct ColumnRange <: ContiguousCellRange
    start::Int # column number
    stop::Int  # column number

    function ColumnRange(a::Int, b::Int)
        if a > b 
            throw(XLSXError("Invalid ColumnRange. Start column must be located before end column."))
        end
        return new(a, b)
    end
end
struct RowRange <: ContiguousCellRange
    start::Int # row number
    stop::Int  # row number

    function RowRange(a::Int, b::Int)
        if a > b
            throw(XLSXError("Invalid RowRange. Start row must be located before end row."))
        end
        return new(a, b)
    end
end

struct SheetCellRef
    sheet::String
    cellref::CellRef
end

struct SheetCellRange <: ContiguousSheetCellRange
   sheet::String
   rng::CellRange
end

struct NonContiguousRange <: AbstractSheetCellRange
    sheet::String
    rng::Vector{Union{CellRef, CellRange}}
end

struct SheetColumnRange <: ContiguousSheetCellRange
    sheet::String
    colrng::ColumnRange
end
struct SheetRowRange <: ContiguousSheetCellRange
    sheet::String
    rowrng::RowRange
end

abstract type MSOfficePackage end

struct EmptyMSOfficePackage <: MSOfficePackage
end

#=
Relationships are defined in ECMA-376-1 Section 9.2.
This struct matches the `Relationship` tag attribute names.

A `Relationship` defines relations between the files inside a MSOffice package.
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
    itr::XML.LazyNode # Worksheet being processed
    itr_state::Union{Nothing, XML.LazyNode} # Worksheet state
    row::Int # number of current row in the worksheet. It´s set to 0 in the start state.
    ht::Union{Float64, Nothing} # row height
end

mutable struct WorksheetCacheIteratorState
    row_from_last_iteration::Int
    full_cache::Bool # is the cache full (true) or does it need filling (false)
end

mutable struct WorksheetCache{I<:SheetRowIterator} <: SheetRowIterator
    is_full::Bool # false until iterator runs to completion
    cells::CellCache # SheetRowNumber -> Dict{column_number, Cell}
    rows_in_cache::Vector{Int} # ordered vector with row numbers that are stored in cache
    row_ht::Dict{Int, Union{Float64, Nothing}} # Maps a row number to a row height
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
    unhandled_attributes::Union{Nothing,Dict{Int,Dict{String,String}}}
    sst_count::Int # number of cells containing a shared string

    function Worksheet(package::MSOfficePackage, sheetId::Int, relationship_id::String, name::String, dimension::Union{Nothing, CellRange}, is_hidden::Bool)
        return new(package, sheetId, relationship_id, name, dimension, is_hidden, nothing, nothing, 0)
    end
end

struct SheetRowStreamIterator <: SheetRowIterator
    sheet::Worksheet
end

#------------------------------------------------------------------------------ sharedStrings
mutable struct SharedStringTable
    unformatted_strings::Vector{String}
    formatted_strings::Vector{String}
    index::Dict{String, Int64} # for unformatted_strings search optimisation
    is_loaded::Bool # for lazy-loading of sst XML file (implies that this struct must be mutable)
end
struct SstToken
    n::XML.LazyNode
    idx::Int
end
struct Sst
    unformatted::String
    formatted::String
    idx::Int
end
const DefinedNameValueTypes = Union{SheetCellRef, SheetCellRange, NonContiguousRange, Int, Float64, String, Missing}
const DefinedNameRangeTypes = Union{SheetCellRef, SheetCellRange, NonContiguousRange}

struct DefinedNameValue
    value::DefinedNameValueTypes
    isabs::Union{Bool, Vector{Bool}}
end

# Workbook is the result of parsing file `xl/workbook.xml`.
# The `xl/workbook.xml` will need to be updated using the Workbook_names and 
# worksheet_names from here when a workbook is saved in case any new defined 
# names have been created.
mutable struct Workbook
    package::MSOfficePackage # parent XLSXFile
    sheets::Vector{Worksheet} # workbook -> sheets -> <sheet name="Sheet1" r:id="rId1" sheetId="1"/>. sheetId determines the index of the WorkSheet in this vector.
    date1904::Bool              # workbook -> workbookPr -> attribute date1904 = "1" or absent
    relationships::Vector{Relationship} # contains workbook level relationships
    sst::SharedStringTable # shared string table
    buffer_styles_is_float::Dict{Int, Bool}      # cell style -> true if is float
    buffer_styles_is_datetime::Dict{Int, Bool}   # cell style -> true if is datetime
    workbook_names::Dict{String, DefinedNameValue} # definedName
    worksheet_names::Dict{Tuple{Int, String}, DefinedNameValue} # definedName. (sheetId, name) -> value.
    styles_xroot::Union{XML.Node, Nothing}
end

"""
`XLSXFile` represents a reference to an Excel file.

It is created by using [`XLSX.readxlsx`](@ref) or [`XLSX.openxlsx`](@ref) 
or [`XLSX.opentemplate`](@ref) or [`XLSX.newxlsx`](@ref).

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
    io::ZipArchives.ZipReader
    files::Dict{String, Bool} # maps filename => isread bool
    data::Dict{String, XML.Node} # maps filename => XMLDocument (with row/sst elements removed)
    binary_data::Dict{String, Vector{UInt8}} # maps filename => file content in bytes
    workbook::Workbook
    relationships::Vector{Relationship} # contains package level relationships
    is_writable::Bool # indicates whether this XLSX file can be edited

    function XLSXFile(source::Union{AbstractString, IO}, use_cache::Bool, is_writable::Bool)
        check_for_xlsx_file_format(source)
        if use_cache || (source isa IO)
            io = ZipArchives.ZipReader(read(source))
        else
            io = ZipArchives.ZipReader(FileArray(abspath(source)))
        end
        xl = new(source, use_cache, io, Dict{String, Bool}(), Dict{String, XML.Node}(), Dict{String, Vector{UInt8}}(), EmptyWorkbook(), Vector{Relationship}(), is_writable)
        xl.workbook.package = xl
        return xl
    end
end

struct ReadFile
    node::Union{Nothing,XML.Node}
    raw::Union{Nothing,XML.Raw}
    bin::Union{Nothing,Vector{UInt8}}
    name::String
end

#
# Iterators
#

struct SheetRow
    sheet::Worksheet
    row::Int                  # index of the row in the worksheet
    ht::Union{Float64, Nothing}   # row height
    rowcells::Dict{Int, Cell} # column -> value
end

struct Index # based on DataFrames.jl
    lookup::Dict{Symbol, Int} # column label -> table column index
    column_labels::Vector{Symbol}
    column_map::Dict{Int, Int} # table column index (1-based) -> sheet column index (cellref based)

    function Index(column_range::Union{ColumnRange, AbstractString}, column_labels)
        column_labels_as_syms = [ Symbol(i) for i in column_labels ]
        column_range = convert(ColumnRange, column_range)
        if length(unique(column_labels_as_syms)) != length(column_labels_as_syms)
            throw(XLSXError("Column labels must be unique."))
        end

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
    missing_rows::Int # number of completely empty rows between the last row and the current row
    row_pending::Union{Nothing, SheetRow} # if the last row was empty, this is the row that was pending to be returned
end

struct DataTable
    data::Vector{Any} # columns
    column_labels::Vector{Symbol}
    column_label_index::Dict{Symbol, Int} # column_label -> column_index

    function DataTable(
            data::Vector{Any}, # columns
            column_labels::Vector{Symbol},
        )

        if length(data) != length(column_labels)
            throw(XLSXError("Data has $(length(data)) columns but $(length(column_labels)) column labels."))
        end

        column_label_index = Dict{Symbol, Int}()
        for (i, sym) in enumerate(column_labels)
            if haskey(column_label_index, sym)
                throw(XLSXError("DataTable has repeated label for column `$sym`"))
            end
            column_label_index[sym] = i
        end

        return new(data, column_labels, column_label_index)
    end
end

struct xpath
    node::XML.Node
    path::String

    function xpath(node::XML.Node, path::String)
        new(node, path)
    end
end

struct XLSXError <: Exception
    msg::String
end
Base.showerror(io::IO, e::XLSXError) = print(io, "XLSXError: ",e.msg)

struct FileArray <: AbstractVector{UInt8}
    filename::String
    offset::Int64
    len::Int64
end
