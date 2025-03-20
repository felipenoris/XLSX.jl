
EmptyWorkbook() = Workbook(EmptyMSOfficePackage(), Vector{Worksheet}(), false,
    Vector{Relationship}(), SharedStringTable(), Dict{Int, Bool}(), Dict{Int, Bool}(),
    Dict{String, DefinedNameValueTypes}(), Dict{Tuple{Int, String}, DefinedNameValueTypes}(), nothing)

#=
Indicates whether this XLSX file can be edited.
This controls if assignment to worksheet cells is allowed.
Writable XLSXFile instances are opened with `XLSX.open_xlsx_template` method.
=#
is_writable(xl::XLSXFile) = xl.is_writable

"""
    sheetnames(xl::XLSXFile)
    sheetnames(wb::Workbook)

Return a vector with Worksheet names for this Workbook.
"""
sheetnames(wb::Workbook) = [ s.name for s in wb.sheets ]
@inline sheetnames(xl::XLSXFile) = sheetnames(xl.workbook)

"""
    hassheet(wb::Workbook, sheetname::AbstractString)
    hassheet(xl::XLSXFile, sheetname::AbstractString)

Return `true` if `wb` contains a sheet named `sheetname`.
"""
function hassheet(wb::Workbook, sheetname::AbstractString) :: Bool
    for s in wb.sheets
        if s.name == unquoteit(sheetname)
            return true
        end
    end
    return false
end

@inline hassheet(xl::XLSXFile, sheetname::AbstractString) = hassheet(xl.workbook, sheetname)

"""
    sheetcount(xlsfile) :: Int

Count the number of sheets in the Workbook.
"""
@inline sheetcount(wb::Workbook) = length(wb.sheets)
@inline sheetcount(xl::XLSXFile) = sheetcount(xl.workbook)

# Returns true if workbook follows date1904 convention.
@inline isdate1904(wb::Workbook) :: Bool = wb.date1904
@inline isdate1904(xf::XLSXFile) :: Bool = isdate1904(get_workbook(xf))

function getsheet(wb::Workbook, sheetname::String) :: Worksheet
    for ws in wb.sheets
        if ws.name == unquoteit(sheetname)
            return ws
        end
    end
    error("$(get_xlsxfile(wb).source) does not have a Worksheet named $sheetname.")
end

@inline getsheet(wb::Workbook, sheet_index::Int) :: Worksheet = wb.sheets[sheet_index]
@inline getsheet(xl::XLSXFile, sheetname::String) :: Worksheet = getsheet(xl.workbook, sheetname)
@inline getsheet(xl::XLSXFile, sheet_index::Int) :: Worksheet = getsheet(xl.workbook, sheet_index)

function Base.show(io::IO, xf::XLSXFile)

    function sheetcountstr(workbook)
        sc = sheetcount(workbook)
        if sc == 1
            return "1 Worksheet"
        else
            return "$sc Worksheets"
        end
    end

    wb = xf.workbook
    print(io, "XLSXFile(\"$(xf.source)\") ",
              "containing $(sheetcountstr(wb))\n")
    @printf(io, "%21s %-13s %-13s\n", "sheetname", "size", "range")
    println(io, "-"^(21+1+13+1+13))

    for s in wb.sheets
        sheetname = s.name 
        if textwidth(sheetname) > 20 
            sheetname = sheetname[collect(eachindex(s.name))[1:20]] * "…"
        end

        if s.dimension !== nothing
            rg = s.dimension
            _size = size(rg) |> x -> string(x[1], "x", x[2])
            @printf(io, "%21s %-13s %-13s\n", sheetname, _size, rg)
        else
            @printf(io, "%21s size unknown\n", sheetname)
        end
    end
end

@inline Base.getindex(xl::XLSXFile, i::Integer) = getsheet(xl, i)

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

function getdata(xl::XLSXFile, rng::SheetRowRange)
    @assert hassheet(xl, rng.sheet) "Sheet $(rng.sheet) not found."
    return getdata(getsheet(xl, rng.sheet), rng.rowrng)
end

function getdata(xl::XLSXFile, rng::NonContiguousRange)
    @assert hassheet(xl, rng.sheet) "Sheet $(rng.sheet) not found."
    return getdata(getsheet(xl, rng.sheet), rng)
end

function getdata(xl::XLSXFile, s::AbstractString)
    if is_workbook_defined_name(xl, s)
        v = get_defined_name_value(xl.workbook, s)
        if is_defined_name_value_a_constant(v)
            return v
        elseif is_defined_name_value_a_reference(v)
            return getdata(xl, v)
        else
            error("Unexpected defined name value: $v.")
        end
    elseif is_valid_sheet_cellname(s)
        return getdata(xl, SheetCellRef(s))
    elseif is_valid_sheet_cellrange(s)
        return getdata(xl, SheetCellRange(s))
    elseif is_valid_sheet_column_range(s)
        return getdata(xl, SheetColumnRange(s))
    elseif is_valid_sheet_row_range(s)
        return getdata(xl, SheetRowRange(s))
    elseif is_valid_non_contiguous_range(s)
        return getdata(xl, nonContiguousRange(s))
    end

    error("$s is not a valid definedName or cell/range reference.")
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

function getcellrange(xl::XLSXFile, rng::SheetRowRange)
    @assert hassheet(xl, rng.sheet) "Sheet $(rng.sheet) not found."
    return getcellrange(getsheet(xl, rng.sheet), rng.rowrng)
end

function getcellrange(xl::XLSXFile, rng::NonContiguousRange)
    @assert hassheet(xl, rng.sheet) "Sheet $(rng.sheet) not found."
    return getcellrange(getsheet(xl, rng.sheet), rng)
end

function getcellrange(xl::XLSXFile, rng_str::AbstractString)
    if is_valid_sheet_cellrange(rng_str)
        return getcellrange(xl, SheetCellRange(rng_str))
    elseif is_valid_sheet_column_range(rng_str)
        return getcellrange(xl, SheetColumnRange(rng_str))
    elseif is_valid_sheet_row_range(rng_str)
        return getcellrange(xl, SheetRowRange(rng_str))
    elseif is_valid_non_contiguous_range(rng_str)
        return getcellrange(xl, nonContiguousRange(rng_str))
    end

    error("$rng_str is not a valid range reference.")
end

@inline is_workbook_defined_name(wb::Workbook, name::AbstractString) :: Bool = haskey(wb.workbook_names, name)
@inline is_workbook_defined_name(xl::XLSXFile, name::AbstractString) :: Bool = is_workbook_defined_name(get_workbook(xl), name)
@inline is_worksheet_defined_name(ws::Worksheet, name::AbstractString) :: Bool = is_worksheet_defined_name(get_workbook(ws), ws.sheetId, name)
@inline is_worksheet_defined_name(wb::Workbook, sheetId::Int, name::AbstractString) :: Bool = haskey(wb.worksheet_names, (sheetId, name))
@inline is_worksheet_defined_name(wb::Workbook, sheet_name::AbstractString, name::AbstractString) :: Bool = is_worksheet_defined_name(wb, getsheet(wb, sheet_name).sheetId, name)

@inline get_defined_name_value(wb::Workbook, name::AbstractString) :: DefinedNameValueTypes = wb.workbook_names[name].value

function get_defined_name_value(ws::Worksheet, name::AbstractString) :: DefinedNameValueTypes
    wb = get_workbook(ws)
    sheetId = ws.sheetId
    dn = wb.worksheet_names[(sheetId, name)]
    return dn.value
end

@inline is_defined_name_value_a_reference(v::DefinedNameValueTypes) = isa(v, SheetCellRef) || isa(v, SheetCellRange) || isa(v, NonContiguousRange)
@inline is_defined_name_value_a_constant(v::DefinedNameValueTypes) = !is_defined_name_value_a_reference(v)

function is_valid_defined_name(name::AbstractString) :: Bool
    if isempty(name)
        return false
    end
    if !isletter(name[1]) && name[1] != '_'
        return false
    end
    for c in name
        if !isletter(c) && !isdigit(c) && c != '_' && c != '\\'
            return false
        end
    end
    return true
end

function addDefName(xf::XLSXFile, name::AbstractString, value::DefinedNameValueTypes; absolute=true)
    if !is_valid_defined_name(name)
        error("Invalid defined name: $name.")
    end
    if is_workbook_defined_name(xf, name)
        error("Workbook already has a defined name called $name.")
    end
    xf.workbook.workbook_names[name] = DefinedNameValue(value, absolute)
end
function addDefName(ws::Worksheet, name::AbstractString, value::DefinedNameValueTypes; absolute=true)
    wb = get_workbook(ws)
    if !is_valid_defined_name(name)
        error("Invalid defined name: $name.")
    end
    if is_worksheet_defined_name(ws, name)
        error("Worksheet $(ws.name) already has a defined name called $name.")
    end
#    local abs::Union{Bool, Vector{Bool}}
    if value isa NonContiguousRange
        @assert replace(value.sheet, "'" => "") == ws.name "Non-contiguous range must be in the same worksheet."
        abs = absolute ? fill(true, length(value.rng)) : fill(false, length(value.rng))
    else
        abs = absolute ? true : false
    end
    wb.worksheet_names[(ws.sheetId, name)] = DefinedNameValue(value, abs)
end

quoteit(x::AbstractString) = occursin(r"^[0-9]|[\s,:!&#@*]", x) ? "'$x'" : x
unquoteit(x::AbstractString) = replace(x, "'" => "")

"""
    addDefinedName(xf::XLSXFile,  name::AbstractString, value::Union{Int, Float64, String}; absolute=true)
    addDefinedName(xf::XLSXFile,  name::AbstractString, value::AbstractString; absolute=true)
    addDefinedName(sh::Worksheet, name::AbstractString, value::Union{Int, Float64, String}; absolute=true)
    addDefinedName(sh::Worksheet, name::AbstractString, value::AbstractString; absolute=true)

Add a defined name to the Workbook or Worksheet. If an `XLSXFile` is passed, the defined name 
is added to the Workbook. If a `Worksheet` is passed, the defined name is added to the Worksheet.

If the new `definedName` is a cell reference or range, by default, it will be an absolute 
reference (e.g. \$A\$1:\$C\$6). If `absolute=false` is specified, the new `definedName will be 
a relative reference(e.g. A1:C6). The `absolute` argument is ignored if the `definedName` is
not a cell reference or range.

In the context of `XLSX.jl` there is no difference between an absolute reference and a relative 
reference. However, Excel treats them differently. When `definedNames` are read in as part of 
an XLSXFile, we keep track of whether they are absolute or not. If the XLSXFile is subsequently 
written out again, the status of the `definedNames` is preserved. 

# Examples
```julia
julia> XLSX.addDefinedName(sh, "ID", "C21")

julia> XLSX.addDefinedName(sh, "NEW", "'Mock-up'!A1:B2")

julia> XLSX.addDefinedName(sh, "my_name", "A1,B2,C3")

julia> XLSX.addDefinedName(xf, "Life_the_universe_and_everything", 42)

julia> XLSX.addDefinedName(xf, "first_name", "Hello World")

```
"""
function addDefinedName end
addDefinedName(xf::XLSXFile, name::AbstractString, value::Union{Int, Float64}; absolute=true) = addDefName(xf, name, value)
addDefinedName(ws::Worksheet, name::AbstractString, value::Union{Int, Float64}; absolute=true) = addDefName(ws, name, value)
function addDefinedName(xf::XLSXFile, name::AbstractString, value::AbstractString; absolute=true)
    if value == ""
        error("Defined name value cannot be an empty string.")
    end
    if is_valid_sheet_cellname(value)
        return addDefName(xf, name, SheetCellRef(value); absolute)
    elseif is_valid_sheet_cellrange(value)
        return addDefName(xf, name, SheetCellRange(value); absolute)
    elseif is_valid_non_contiguous_sheetcellrange(value)
        return addDefName(xf, name, nonContiguousRange(value); absolute)
    else
        return addDefName(xf, name, value isa String ? "\"$value\"" : value)
    end
end
function addDefinedName(ws::Worksheet, name::AbstractString, value::AbstractString; absolute=true)
    if value == ""
        error("Defined name value cannot be an empty string.")
    end
    if is_valid_cellname(value)
        return addDefName(ws, name, SheetCellRef(ws.name, CellRef(value)); absolute)
    elseif is_valid_cellrange(value)
        return addDefName(ws, name, SheetCellRange(ws.name, CellRange(value)); absolute)
    elseif is_valid_non_contiguous_sheetcellrange(value)
        return addDefName(ws, name, nonContiguousRange(value); absolute)
    elseif is_valid_non_contiguous_cellrange(value)
        return addDefName(ws, name, nonContiguousRange(ws, value); absolute)
    else
        return addDefName(ws, name, value isa String ? "\"$value\"" : value)
    end
end
