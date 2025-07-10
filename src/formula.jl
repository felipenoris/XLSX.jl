Base.isempty(f::Formula) = f.formula == ""
Base.isempty(f::ReferencedFormula) = f.formula == ""
Base.isempty(f::FormulaReference) = false # always links to another formula
Base.hash(f::Formula) = hash(f.formula) + hash(f.unhandled)
Base.hash(f::FormulaReference) = hash(f.id) + hash(f.unhandled)
Base.hash(f::ReferencedFormula) = hash(f.formula) + hash(f.id) + hash(f.ref) + hash(f.unhandled)