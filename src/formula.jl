const EXCEL_FUNCTION_PREFIX = Dict( # Prefixes needed for newer Excel functions
    # Dynamic array core
    "UNIQUE"      => "_xlfn.",
    "FILTER"      => "_xlfn.",
    "SEQUENCE"    => "_xlfn.",
    "RANDARRAY"   => "_xlfn.",

    # Sorting / reshaping
    "SORT"        => "_xlfn._xlws.",
    "SORTBY"      => "_xlfn._xlws.",   # some builds use _xlfn., safest to force _xlws

    # Lookup
    "XLOOKUP"     => "_xlfn.",
    "XMATCH"      => "_xlfn.",

    # Financial / data
    "STOCKHISTORY"=> "_xlfn.",

    # Text split/join
    "TEXTSPLIT"   => "_xlfn.",
    "TEXTBEFORE"  => "_xlfn.",
    "TEXTAFTER"   => "_xlfn.",

    # Stacking / reshaping
    "VSTACK"      => "_xlfn.",
    "HSTACK"      => "_xlfn.",
    "TAKE"        => "_xlfn.",
    "DROP"        => "_xlfn.",
    "TOROW"       => "_xlfn.",
    "TOCOL"       => "_xlfn.",
    "WRAPROWS"    => "_xlfn.",
    "WRAPCOLS"    => "_xlfn.",
    "EXPAND"      => "_xlfn.",
    "CHOOSECOLS"  => "_xlfn.",
    "CHOOSEROWS"  => "_xlfn.",

    # Lambda / Let
    "LAMBDA"      => "_xlfn.",
    "LET"         => "_xlfn.",

    # Parameter markers (appear only inside LAMBDA/LET bodies)
    "_xlpm"       => "_xlpm.",

    # User‑defined functions
    "_xludf"      => "_xludf."
)

Base.isempty(f::Formula) = f.formula == ""
Base.isempty(f::ReferencedFormula) = f.formula == ""
Base.isempty(f::FormulaReference) = false # always links to another formula
Base.hash(f::Formula) = hash(f.formula) + hash(f.unhandled)
Base.hash(f::FormulaReference) = hash(f.id) + hash(f.unhandled)
Base.hash(f::ReferencedFormula) = hash(f.formula) + hash(f.id) + hash(f.ref) + hash(f.unhandled)

function new_ReferencedFormula_Id(ws::Worksheet)
    # get all referenceFormaula Ids currently in use
    all_f=([z.formula.id for x in [values(last(x)) for x in (ws.cache.cells)] for z in x if isa(z.formula, ReferencedFormula)])

    # find the lowest integer not currently used as an Id
    id = 0
    while id ∈ all_f
        id +=1
    end

    return id

end

# If overwriting a cell containing a referencedFormula, need to re-reference all referring cells.
# The referencedFormula will be in the top right cell of the referenced block. Need to rereference 
# the rest of the block on this top row (without the first, overwritten cell) and then the rest of 
# the block without this top row. Need to do this as two new, separate rectangular blocks with the 
# referencedFormula in the first cell of each and the other cells set to formulaReferences.
# Note that a block of referencedFormulas can have a separate referencedFormula block set within it! 
# 
function rereference_formulae(ws::Worksheet, cell::Cell)
    old_range = CellRange(cell.formula.ref)
    if size(old_range) == (1, 2) || size(old_range) == (2, 1)
        getcell(ws, old_range.stop).formula = Formula(cell.formula.formula)
        return
    end
    ranges=CellRange[]
    if old_range.stop.column_number > old_range.start.column_number
        push!(ranges, CellRange(CellRef(old_range.start.row_number, old_range.start.column_number+1), CellRef(old_range.start.row_number, old_range.stop.column_number)))
    end
    if old_range.stop.row_number > old_range.start.row_number
        push!(ranges, CellRange(CellRef(old_range.start.row_number+1, old_range.start.column_number), CellRef(old_range.stop.row_number, old_range.stop.column_number)))
    end

    for newrng in ranges
        if size(newrng) == (1, 1)# || size(newrng) == (2, 1)
            getcell(ws, newrng.stop).formula = Formula(cell.formula.formula)
        else
            newid = new_ReferencedFormula_Id(ws)
            rereference_formulae(ws, cell, newrng, newid)
        end
    end
end

function rereference_formulae(ws::Worksheet, oldcell::Cell, newrng::CellRange, newid::Int64)
    oldform = oldcell.formula.formula
    oldunhandled = oldcell.formula.unhandled
    offset = (newrng.start.row_number - oldcell.ref.row_number, newrng.start.column_number - oldcell.ref.column_number)
    newform = ReferencedFormula(shift_excel_references(oldform, offset), newid, string(newrng), oldunhandled)
    for fr in newrng
        newfr = getcell(ws, fr)
        if fr != newrng.start
            if newfr.formula isa FormulaReference && newfr.formula.id == oldcell.formula.id
                setdata!(ws, Cell(fr, newfr.datatype, newfr.style, "", "", FormulaReference(newid, oldunhandled)))
            end
        else
            setdata!(ws, Cell(fr, oldcell.datatype, newfr.style, "", "", newform))
        end
    end
    return nothing
end

# Replace formula references to a sheet that has been deleted
function update_formulas_missing_sheet!(wb::Workbook, name::String)
    pattern = (name * "!" => "#REF!", r"\$?[A-Z]{1,3}\$?[1-9][0-9]*" => "")
    for i = 1:sheetcount(wb)
        s = getsheet(wb, i)
        for r in eachrow(s)
            for (_, cell) in r.rowcells
                cell.formula isa FormulaReference && continue
                oldform = cell.formula.formula
                if occursin(name * "!", cell.formula.formula)
                    for (pat, r) in pattern
                        cell.formula.formula = replace(cell.formula.formula, pat => r)
                    end
                    if oldform != cell.formula.formula
                        cell.datatype = "e"
                        cell.value = "#REF!"
                    end
                end
            end
        end
    end
end

"""
    setFormula(ws::Worksheet, RefOrRange::AbstractString, formula::String)
    setFormula(xf::XLSXFile,  RefOrRange::AbstractString, formula::String)

    setFormula(sh::Worksheet, row, col, formula::String)

Set the Excel formula to be used in the given cell or cell range.

Formulae must be valid Excel formulae and written in US english with comma
separators. Cell references may be absolute or relative references in either 
the row or the column or both (e.g. `\$A\$2`). No validation of the specified 
formulae is made by `XLSX.jl` and formulae are stored verbatim, as given.

If a contiguous range is specfied, `setFormula` will usually create a 
referencedFormula. This is the same as Excel would use if using drag fill to 
copy a formula into a range of cells.

Since XLSX.jl does not and cannot replicate all the functions built in to Excel, 
setting a formula in a cell does not permit the cell's value to be re-calculated.
Instead, the value is set to missing so that Excel will re-calculate it on opening. 

Unlike with conventional functions, dynamic array functions (e.g. `SORT` or `UNIQUE`)
should not be put into ReferencedFormulae. When functions of this type are identified,
they will instead be duplicated individually in each cell in the given range (but 
with relative cell references appropriately adjusted).

Note that dynamic array functions will return values into a spill range the extent of 
which depends on the data on which the functions is operating. If any of the cells in the 
spill range already contains a value, Excel will show an `@SPILL` error.

# Examples:

```julia

julia> using XLSX

julia> f=newxlsx("setting formulas")
XLSXFile("blank.xlsx") containing 1 Worksheet
            sheetname size          range        
-------------------------------------------------
     setting formulas 1x1           A1:A1        


julia> s=f[1]
1×1 Worksheet: ["setting formulas"](A1:A1) 

julia> s["A1:G1"]=1
1

julia> s["A2:A10"]=1
1

julia> s["A1:J1"]=1
1

julia> setFormula(s, "B2:J10", "=A2+B1") # adds formulae but cannot update calculated values

julia> s[:]
10×10 Matrix{Any}:
 1  1         1         1         1         1         1         1         1         1
 1   missing   missing   missing   missing   missing   missing   missing   missing   missing
 1   missing   missing   missing   missing   missing   missing   missing   missing   missing
 1   missing   missing   missing   missing   missing   missing   missing   missing   missing
 1   missing   missing   missing   missing   missing   missing   missing   missing   missing
 1   missing   missing   missing   missing   missing   missing   missing   missing   missing
 1   missing   missing   missing   missing   missing   missing   missing   missing   missing
 1   missing   missing   missing   missing   missing   missing   missing   missing   missing
 1   missing   missing   missing   missing   missing   missing   missing   missing   missing
 1   missing   missing   missing   missing   missing   missing   missing   missing   missing

# formulae are there for when the saved file is opened in Excel.
julia> XLSX.getcell(s, "B2")
XLSX.Cell(B2, "", "", "", XLSX.ReferencedFormula("=A2+B1", 0, "B2:J10", nothing))

julia> XLSX.getcell(s, "J10")
XLSX.Cell(J10, "", "", "", XLSX.FormulaReference(0, nothing))

julia> addsheet!(f, "trig functions")
1×1 Worksheet: ["trig functions"](A1:A1) 

julia> f
XLSXFile("mytest.xlsx") containing 2 Worksheets
            sheetname size          range        
-------------------------------------------------
     setting formulas 10x10         A1:J10
       trig functions 1x1           A1:A1


julia> s2=f[2]
1×1 Worksheet: ["trig functions"](A1:A1)

julia> for i=1:100, s2[i, 1] = 2.0*pi*i/100.0; end

julia> setFormula(s2, "B1:B100", "=sin(A1)")

julia> setFormula(s2, "C1:C100", "=cos(A1)")

julia> setFormula(s2, "D1:D100", "=sin(A1)^2 + cos(A1)^2")

julia> f=newxlsx("mysheet")
XLSXFile("blank.xlsx") containing 1 Worksheet
            sheetname size          range
-------------------------------------------------
              mysheet 1x1           A1:A1

julia> s=f[1]
1×1 Worksheet: ["mysheet"](A1:A1)

julia> s["A1"]=["Header1" "Header2" "Header3"; 1 2 3; 4 5 6; 7 8 9; 1 2 3; 4 5 6; 7 8 9]
7×3 Matrix{Any}:
  "Header1"   "Header2"   "Header3"
 1           2           3
 4           5           6
 7           8           9
 1           2           3
 4           5           6
 7           8           9

julia> setFormula(s, "E1:G1", "=sort(unique(A2:A7),,-1)") # using dynamic array functions
```
![image|320x500](../images/SortUnique.png)

"""
setFormula(w, r, f::AbstractString) = setFormula(w, r; val=f)
setFormula(w, r, c, f::AbstractString) = setFormula(w, r, c; val=f)
setFormula(ws::Worksheet, ref::SheetCellRef; kw...) = do_sheet_names_match(ws, ref) && setFormula(ws, ref.cellref; kw...)
setFormula(ws::Worksheet, rng::SheetCellRange; kw...) = do_sheet_names_match(ws, rng) && setFormula(ws, rng.rng; kw...)
setFormula(ws::Worksheet, rng::SheetColumnRange; kw...) = do_sheet_names_match(ws, rng) && setFormula(ws, rng.colrng; kw...)
setFormula(ws::Worksheet, rng::SheetRowRange; kw...) = do_sheet_names_match(ws, rng) && setFormula(ws, rng.rowrng; kw...)
#setFormula(ws::Worksheet, rng::CellRange; kw...) = process_cellranges(setFormula, ws, rng; kw...) # do this explicitly, below
setFormula(ws::Worksheet, colrng::ColumnRange; kw...) = process_columnranges(setFormula, ws, colrng; kw...)
setFormula(ws::Worksheet, rowrng::RowRange; kw...) = process_rowranges(setFormula, ws, rowrng; kw...)
setFormula(ws::Worksheet, ncrng::NonContiguousRange; kw...) = process_ncranges(setFormula, ws, ncrng; kw...)
setFormula(ws::Worksheet, ref_or_rng::AbstractString; kw...) = process_ranges(setFormula, ws, ref_or_rng; kw...)
setFormula(xl::XLSXFile, sheetcell::String; kw...) = process_sheetcell(setFormula, xl, sheetcell; kw...)
setFormula(ws::Worksheet, row::Integer, col::Integer; kw...) = setFormula(ws, CellRef(row, col); kw...)
setFormula(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setFormula, ws, row, nothing; kw...)
setFormula(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setFormula, ws, nothing, col; kw...)
setFormula(ws::Worksheet, ::Colon, ::Colon; kw...) = process_colon(setFormula, ws, nothing, nothing; kw...)
setFormula(ws::Worksheet, ::Colon; kw...) = process_colon(setFormula, ws, nothing, nothing; kw...)
setFormula(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_veccolon(setFormula, ws, row, nothing; kw...)
setFormula(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon(setFormula, ws, nothing, col; kw...)
setFormula(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setFormula, ws, row, col; kw...)
setFormula(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_vecint(setFormula, ws, row, col; kw...)
setFormula(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setFormula, ws, row, col; kw...)
setFormula(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setFormula(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
function setFormula(ws::Worksheet, rng::CellRange; val::AbstractString)

    is_array=false
    for k in keys(EXCEL_FUNCTION_PREFIX) # Identify dynamic arrays
        r = Regex(k, "i")
        is_array |= occursin(r, val)
    end

    if is_array # Don't use ReferencedFormulas for dynamic arrays. set each cell individually
        start = rng.start
        for c in rng
            offset = (c.row_number - start.row_number, c.column_number - start.column_number)
            newval=shift_excel_references(val, offset)
            setFormula(ws, c, newval)
        end
        return
    end

    first_cell = getcell(ws, rng.start)
    if !isa(first_cell, EmptyCell) && first_cell.formula isa ReferencedFormula
        if CellRange(first_cell.formula.ref) == rng # range matches, so just need to change the referenced formula
            first_cell.formula.formula = val
            return
        end
    end
    
    newid = new_ReferencedFormula_Id(ws)
    for c in rng
        if c == rng.start
            newform = ReferencedFormula(val, newid, string(rng), nothing)
        else
            newform = FormulaReference(newid, nothing)
        end
        cell = getcell(ws, c)
        if cell isa EmptyCell || cell.style==""
            setdata!(ws, c, CellFormula(ws, newform))
        else
            setdata!(ws, c, CellFormula(newform, CellDataFormat(parse(Int,cell.style))))
        end
    end
    return
end
function setFormula(ws::Worksheet, cellref::CellRef; val::AbstractString)
    xf=get_xlsxfile(ws)
    c=getcell(ws, cellref)
    t   = ""
    ref = ""
    cm  = ""

    formula=val
    for (k, v) in EXCEL_FUNCTION_PREFIX # add prefixes to any array functions
        r = Regex(k, "i")
        formula = replace(formula, r => v*k) # replace any dynamic array function name with its prefixed name
    end

    if formula != val # contains a dynamic array function (now with prefix(es))
        t = "array"
        ref = cellref.name*":"*cellref.name
        cm = "1"
        if !haskey(xf.files, "xl/metadata.xml") # add metadata.xml on first use of a dynamicArray formula
#            xf.data["xl/metadata.xml"] = XML.Node(XML.Raw(read(joinpath(_relocatable_data_path(), "metadata.xml"))))
            xf.data["xl/metadata.xml"] = XML.Node(XML.Raw(read(raw"C:\Users\tim\OneDrive\Documents\Julia\XLSX\XLSX.jl\data\metadata.xml")))
            xf.files["xl/metadata.xml"] = true # set file as read
#            wbdoc = xmlroot(xf, "xl/workbook.xml")
            add_override!(xf, "/xl/metadata.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml")
            rId = add_relationship!(get_workbook(ws), "metadata.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata")
        end
    end

    if c isa EmptyCell || c.style==""
        setdata!(ws, cellref, CellFormula(ws, Formula(formula, t, ref, nothing)))
    else
        setdata!(ws, cellref, CellFormula(Formula(formula, t, ref, nothing), CellDataFormat(parse(Int,c.style))))
    end
    c=getcell(ws, cellref)
    c.meta = cm
end
