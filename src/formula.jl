Base.:(==)(f1::AbstractFormula, f2::AbstractFormula) = false

Base.:(==)(f1::Formula, f2::Formula) = f1.formula == f2.formula

Base.:(==)(f1::ReferencedFormula, f2::Formula) = f1.formula == f2.formula
Base.:(==)(f1::Formula, f2::ReferencedFormula) = f1.formula == f2.formula

# WARNING: This is problematic because indexing is unique per sheet
Base.:(==)(f1::Union{FormulaReference, ReferencedFormula}, f2::FormulaReference) = f1.id == f2.id 
Base.:(==)(f1::FormulaReference, f2::ReferencedFormula) = f1.id == f2.id 

Base.:(==)(f1::ReferencedFormula, f2::ReferencedFormula) = f1.id == f2.id && f1.formula == f2.formula

Base.hash(f::Formula) = hash(f.formula)
Base.hash(f::ReferencedFormula) = hash(f.formula) + hash(f.id) + hash(f.ref)
Base.hash(f::FormulaReference) = hash(f.id)

Base.isempty(f::Formula) = f.formula == ""
Base.isempty(f::ReferencedFormula) = f.formula == ""
Base.isempty(f::FormulaReference) = false # always links to another formula