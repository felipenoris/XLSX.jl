Base.isempty(f::Formula) = f.formula == ""
Base.isempty(f::ReferencedFormula) = f.formula == ""
Base.isempty(f::FormulaReference) = false # always links to another formula
Base.hash(f::Formula) = hash(f.formula) + hash(f.unhandled)
Base.hash(f::FormulaReference) = hash(f.id) + hash(f.unhandled)
Base.hash(f::ReferencedFormula) = hash(f.formula) + hash(f.id) + hash(f.ref) + hash(f.unhandled)

# if overwriting a cell containing a referenced formula, need to re-reference all referring cells
function rereference_formulae(ws::Worksheet, cell::Cell)
    process_range = CellRange(cell.formula.ref)
    done = CellRange("A1:A1")
    first_range = true
    while done != process_range
        for c in process_range
            newcell = getcell(ws, c)
            if newcell.formula isa FormulaReference
                newid = first_range ? cell.formula.id : length(unique([z.formula.id for x in [values(last(x)) for x in (ws.cache.cells)] for z in x if z.formula isa ReferencedFormula]))
                newref = CellRange(CellRef(newcell.ref.row_number, process_range.stop.column_number), process_range.stop)
                offset = (newcell.ref.row_number - cell.ref.row_number, newcell.ref.column_number - cell.ref.column_number)
                done = rereference_formulae(ws, cell, newref, offset, newid)
                if done.start.row_number + 1 > process_range.stop.row_number || done.start.column_number - 1 < process_range.start.column_number
                    process_range = done
                    break
                end
                process_range = CellRange(CellRef(done.start.row_number + 1, process_range.start.column_number), CellRef(process_range.stop.row_number, done.start.column_number - 1))
                first_range = false
                break
            end
        end
    end
end
function rereference_formulae(ws::Worksheet, cell::Cell, newref::CellRange, offset::Tuple{Int64,Int64}, newid::Int64)::CellRange
    oldform = cell.formula.formula
    oldunhandled = cell.formula.unhandled
    newform = ReferencedFormula(shift_excel_references(oldform, offset), newid, string(newref), oldunhandled)
    for fr in newref
        if fr != newref.start
            newfr = getcell(ws, fr)
            setdata!(ws, Cell(fr, newfr.datatype, newfr.style, newfr.value, FormulaReference(newid, oldunhandled)))
        end
    end
    setdata!(ws, Cell(newref.start, cell.datatype, cell.style, cell.value, newform))
    return newref
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
