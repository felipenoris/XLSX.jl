#----------------------------------------------------------------------------------------------------
# metadata.xml should perhaps better be a package artifact. Put it here in the meantime.
const metadata = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>"""
#-----------------------------------------------------------------------------------------------------

const RGX_FORMULA_SHEET_CELL = r"!\$?[A-Z]+\$?[0-9]" # to recognise sheetcell references like "otherSheet!A1"
const EXCEL_FUNCTION_PREFIX = Dict( # Prefixes needed for newer Excel functions - previously two different prefixes (hence Dict) but now only one.
    # Core dynamic array + LAMBDA family
    "MAKEARRAY"   => "_xlfn.",
    "MAP"         => "_xlfn.",
    "REDUCE"      => "_xlfn.",
    "SCAN"        => "_xlfn.",
    "BYROW"       => "_xlfn.",
    "BYCOL"       => "_xlfn.",
    "LAMBDA"      => "_xlfn.",
    "ANCHORARRAY" => "_xlfn.",

    # Generators
    "SEQUENCE"    => "_xlfn.",
    "RANDARRAY"   => "_xlfn.",

    # Array shaping/stacking
    "VSTACK"      => "_xlfn.",
    "HSTACK"      => "_xlfn.",
    "TOCOL"       => "_xlfn.",
    "TOROW"       => "_xlfn.",
    "WRAPROWS"    => "_xlfn.",
    "WRAPCOLS"    => "_xlfn.",
    "TAKE"        => "_xlfn.",
    "DROP"        => "_xlfn.",
    "CHOOSECOLS"  => "_xlfn.",
    "CHOOSEROWS"  => "_xlfn.",

    # Sort/filter/distinct (historically also seen with "_xlws.")
    "SORT"        => "_xlfn.",
    "SORTBY"      => "_xlfn.",
    "FILTER"      => "_xlfn.",
    "UNIQUE"      => "_xlfn.",

    # Lookup
    "XLOOKUP"     => "_xlfn.",
    "XMATCH"      => "_xlfn.",

    # Text functions (dynamic-array aware)
    "TEXTSPLIT"   => "_xlfn.",
    "TEXTBEFORE"  => "_xlfn.",
    "TEXTAFTER"   => "_xlfn."
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
# The referencedFormula will be in the top left cell of the referenced block. Need to rereference 
# the rest of the block on this top row (without the first, overwritten cell) and then the rest of 
# the block without this top row. Need to do this as two new, separate rectangular blocks with the 
# referencedFormula in the first cell of each and the other cells set to formulaReferences.
# 
# overwritten newRF1    FR1       FR1       FR1
# newRF2      FR2       FR2       FR2       FR2
# FR2         FR2       FR2       FR2       FR2
#
# Note that a block of referencedFormulas can have a separate referencedFormula block set within it! 
# 
function rereference_formulae(ws::Worksheet, cell::Cell)
    old_range = CellRange(cell.formula.ref)
    ranges=CellRange[]
    if old_range.stop.column_number > old_range.start.column_number
        push!(ranges, CellRange(CellRef(old_range.start.row_number, old_range.start.column_number+1), CellRef(old_range.start.row_number, old_range.stop.column_number)))
    end
    if old_range.stop.row_number > old_range.start.row_number
        push!(ranges, CellRange(CellRef(old_range.start.row_number+1, old_range.start.column_number), CellRef(old_range.stop.row_number, old_range.stop.column_number)))
    end

    for newrng in ranges
        if size(newrng) == (1, 1)
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

# shift the relative cell references in a formula when shifting a ReferencedFormula
function shift_excel_references(formula::String, offset::Tuple{Int64,Int64})
    # Regex to match Excel-style cell references (e.g., A1, $A$1, A$1, $A1)
    pattern = r"\$?[A-Z]{1,3}\$?[1-9][0-9]*"
    row_shift, col_shift = offset

    initial = [string(x.match) for x in eachmatch(pattern, formula)]
    result = Vector{String}()

    for ref in eachmatch(pattern, formula)
        # Extract parts using regex
        m = match(r"(\$?)([A-Z]{1,3})(\$?)([1-9][0-9]*)", ref.match)
        col_abs, col_letters, row_abs, row_digits = m.captures

        col_num = decode_column_number(col_letters)
        row_num = parse(Int, row_digits)

        # Apply shifts only if not absolute
        new_col = col_abs == "\$" ? col_letters : encode_column_number(col_num + col_shift)
        new_row = row_abs == "\$" ? row_digits : string(row_num + row_shift)

        push!(result, col_abs * new_col * row_abs * new_row)
    end

    pairs = Dict(zip(initial, result))
    if !isempty(pairs)
        formula = replace(formula, pairs...)
    end

    return formula
end

# Replace formula references to a sheet that has been deleted with #REF errors
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
formula is made by `XLSX.jl` and formulae are stored verbatim, as given.

If a contiguous range is specfied, `setFormula` will usually create a 
referencedFormula. This is the same as Excel would use if using drag fill to 
copy a formula into a range of cells.

Non-contiguous ranges are not supported by `setFormula`. Set the formula in 
each cell or contiguous range separately.

An `XLSXFile` must be open in write mode to use `setFormula`.

Since XLSX.jl does not and cannot replicate all the functions built in to Excel, 
setting a formula in a cell does not permit the cell's value to be re-calculated 
within XLSX.jl. Instead, although the formula is properly added to the cell, the 
value is set to missing so that Excel will re-calculate it on opening. 

Unlike with conventional functions, dynamic array functions (e.g. `SORT` or `UNIQUE`)
should not be put into ReferencedFormulae and neither should functions that refer to 
cells in other sheets. When functions of this type are identified, they will instead be 
duplicated individually in each cell in the given range (but with relative cell references 
appropriately adjusted). This is an internal process that should be transparent 
to the user.

Note that dynamic array functions will return values into a spill range the extent of 
which depends on the data on which the function is operating. If any of the cells in the 
spill range already contains a value, Excel will show an `#SPILL` error.

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

!!! note

    Excel is often very fussy about the structure of the internal structure of an xlsx file but 
    often the resulting error messages (when Excel tries to open a file it considers mal-formed) 
    are somewhat cryptic. If there is an error in the formula you enter, it may not be clear what 
    it is from the error Excel produces. A safe fall back may be to test the formula in Excel itself 
    and copy/paste it into julia.

    For example:

    ```julia

    julia> f=newxlsx()
    XLSXFile("blank.xlsx") containing 1 Worksheet
                sheetname size          range        
    -------------------------------------------------
                Sheet1 1x1           A1:A1        

    julia> f[1][1:3, 1]=collect(1:3)
    3-element Vector{Int64}:
    1
    2
    3

    julia> setFormula(f[1], "B1", "=max(A1:A3") # missing closing parenthesis in the formula
    ""

    julia> XLSX.getcell(f[1], "B1")
    XLSX.Cell(B1, "", "", "", "", XLSX.Formula("=max(A1:A3", "", "", nothing))

    julia> writexlsx("mytest.xlsx", f, overwrite=true)
    "C:\\Users\\tim\\OneDrive\\Documents\\Julia\\XLSX\\mytest.xlsx"

    ```

    ![image|320x500](../images/problemContent.png)
    ![image|320x500](../images/ExcelRepairs.png)

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
#setFormula(ws::Worksheet, ncrng::NonContiguousRange; kw...) = process_ncranges(setFormula, ws, ncrng; kw...)
setFormula(ws::Worksheet, ref_or_rng::AbstractString; kw...) = process_ranges(setFormula, ws, ref_or_rng; kw...)
setFormula(xl::XLSXFile, sheetcell::String; kw...) = process_sheetcell(setFormula, xl, sheetcell; kw...)
setFormula(ws::Worksheet, row::Integer, col::Integer; kw...) = setFormula(ws, CellRef(row, col); kw...)
setFormula(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon; kw...) = process_colon(setFormula, ws, row, nothing; kw...)
setFormula(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_colon(setFormula, ws, nothing, col; kw...)
setFormula(ws::Worksheet, ::Colon, ::Colon; kw...) = setFormula(ws, :; kw...)
setFormula(ws::Worksheet, ::Colon; kw...) = process_colon(setFormula, ws, nothing, nothing; kw...)
#setFormula(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon; kw...) = process_veccolon(setFormula, ws, row, nothing; kw...)
#setFormula(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_veccolon(setFormula, ws, nothing, col; kw...)
#setFormula(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setFormula, ws, row, col; kw...)
#setFormula(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = process_vecint(setFormula, ws, row, col; kw...)
#setFormula(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}}; kw...) = process_vecint(setFormula, ws, row, col; kw...)
setFormula(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}; kw...) = setFormula(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))); kw...)
function setFormula(ws::Worksheet, rng::CellRange; val::AbstractString)

    xf=get_xlsxfile(ws)

    if xf.is_writable == false
        throw(XLSXError("Cannot set formula because because XLSXFile is not writable."))
    end

    is_array=false
    for k in keys(EXCEL_FUNCTION_PREFIX) # Identify formulas containing dynamic array functions
        r = Regex(k, "i")
        is_array |= occursin(r, val)
    end

    is_sheetcell = occursin(RGX_FORMULA_SHEET_CELL, val)
    
    if is_array || is_sheetcell || occursin("#", val) # Don't use ReferencedFormulas for sheetcell formulas or dynamic array functions. Set each cell individually.
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

    if xf.is_writable == false
        throw(XLSXError("Cannot set formula because because XLSXFile is not writable."))
    end

    c=getcell(ws, cellref)
    t   = ""
    ref = ""
    cm  = ""


    formula=val
    if occursin("#", formula) # handle spill references like A1# or myName#
        formula=replace(formula, r"(\$?[A-Za-z]{1,3}\$?\d+|[A-Za-z_][A-Za-z0-9_.]*)#" => s"ANCHORARRAY(\1)")
    end
    for (k, v) in EXCEL_FUNCTION_PREFIX # add prefixes to any array functions
        r = Regex(k, "i")
        formula = replace(formula, r => v*k) # replace any dynamic array function name with its prefixed name
    end
    if formula != val # contains a dynamic array function (now with prefix(es))
        if occursin(r"^ *=? *_xlfn", formula)
            t = "array"
            ref = cellref.name*":"*cellref.name
            cm = "1"
        end
        if !haskey(xf.files, "xl/metadata.xml") # add metadata.xml on first use of a dynamicArray formula
#            xf.data["xl/metadata.xml"] = XML.Node(XML.Raw(read(joinpath(_relocatable_data_path(), "metadata.xml"))))
            xf.data["xl/metadata.xml"] = parse(metadata, XML.Node)
            xf.files["xl/metadata.xml"] = true # set file as read
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
