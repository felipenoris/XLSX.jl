
EmptyWorkbook() = Workbook(EmptyMSOfficePackage(), Vector{Worksheet}(), false, Vector{Relationship}(), SharedStrings(), EzXML.XMLDocument(), Dict{Int, Bool}(), Dict{Int, Bool}(), Dict{String, Union{SheetCellRef, SheetCellRange}}())

"""
Lists internal files from the XLSX package.
"""
@inline filenames(xl::XLSXFile) = keys(xl.data)

"""
Lists Worksheet names for this Workbook.
"""
sheetnames(wb::Workbook) = [ s.name for s in wb.sheets ]
@inline sheetnames(xl::XLSXFile) = sheetnames(xl.workbook)

function hassheet(wb::Workbook, sheetname::AbstractString) :: Bool
    for s in wb.sheets
        if s.name == sheetname
            return true
        end
    end
    return false
end

@inline hassheet(xl::XLSXFile, sheetname::AbstractString) = hassheet(xl.workbook, sheetname)

"""
Counts the number of sheets in the Workbook.
"""
@inline sheetcount(wb::Workbook) = length(wb.sheets)
@inline sheetcount(xl::XLSXFile) = sheetcount(xl.workbook)

"""
    isdate1904(wb) :: Bool

Returns true if workbook follows date1904 convention.
"""
@inline isdate1904(wb::Workbook) :: Bool = wb.date1904
@inline isdate1904(xf::XLSXFile) :: Bool = isdate1904(xf.workbook)

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
@inline xmlroot(xl::XLSXFile, filename::String) :: EzXML.Node = EzXML.root(xmldocument(xl, filename))

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

Base.show(io::IO, xf::XLSXFile) = print(io, "XLSXFile(\"$(xf.filepath)\")")

Base.getindex(xl::XLSXFile, i::Integer) = getsheet(xl, i)

function Base.getindex(xl::XLSXFile, s::AbstractString)
    if hassheet(xl, s)
        return getsheet(xl, s)
    else
        return getdata(xl, s)
    end
end

function getdata(xl::XLSXFile, ref::SheetCellRef)
    @assert hassheet(xl, ref.sheet) "Sheet $(ref.sheet) not found."
    return getdata(getsheet(xl, ref.sheet), ref.cellref)
end

function getdata(xl::XLSXFile, rng::SheetCellRange)
    @assert hassheet(xl, rng.sheet) "Sheet $(rng.sheet) not found."
    return getdata(getsheet(xl, rng.sheet), rng.rng)
end

function getdata(xl::XLSXFile, rng::SheetColumnRange)
    @assert hassheet(xl, rng.sheet) "Sheet $(rng.sheet) not found."
    return getdata(getsheet(xl, rng.sheet), rng.colrng)
end

function getdata(xl::XLSXFile, s::AbstractString)
    if is_valid_sheet_cellname(s)
        return getdata(xl, SheetCellRef(s))
    elseif is_valid_sheet_cellrange(s)
        return getdata(xl, SheetCellRange(s))
    elseif is_valid_sheet_column_range(s)
        return getdata(xl, SheetColumnRange(s))
    elseif is_defined_name(xl, s)
        return getdata(xl, get_defined_name_reference(xl, s))
    end

    error("$s is not a valid sheetname or cell/range reference.")
end

function getcell(xl::XLSXFile, ref::SheetCellRef)
    @assert hassheet(xl, ref.sheet) "Sheet $(ref.sheet) not found."
    return getcell(getsheet(xl, ref.sheet), ref.cellref)
end

getcell(xl::XLSXFile, ref_str::AbstractString) = getcell(xl, SheetCellRef(ref_str))

function getcellrange(xl::XLSXFile, rng::SheetCellRange)
    @assert hassheet(xl, rng.sheet) "Sheet $(rng.sheet) not found."
    return getcellrange(getsheet(xl, rng.sheet), rng.rng)
end

function getcellrange(xl::XLSXFile, rng::SheetColumnRange)
    @assert hassheet(xl, rng.sheet) "Sheet $(rng.sheet) not found."
    return getcellrange(getsheet(xl, rng.sheet), rng.colrng)
end

function getcellrange(xl::XLSXFile, rng_str::AbstractString)
    if is_valid_sheet_cellrange(rng_str)
        return getcellrange(xl, SheetCellRange(rng_str))
    elseif is_valid_sheet_column_range(rng_str)
        return getcellrange(xl, SheetColumnRange(rng_str))
    end

    error("$rng_str is not a valid range reference.")
end

@inline is_defined_name(wb::Workbook, name::AbstractString) :: Bool = haskey(wb.defined_names, name)
@inline is_defined_name(xl::XLSXFile, name::AbstractString) :: Bool = is_defined_name(xl.workbook, name)
@inline get_defined_name_reference(wb::Workbook, name::AbstractString) = wb.defined_names[name]
@inline get_defined_name_reference(xl::XLSXFile, name::AbstractString) = get_defined_name_reference(xl.workbook, name)
