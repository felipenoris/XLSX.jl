
# Tables.jl interface

Tables.isrowtable(::Type{<:TableRowIterator}) = true
Tables.columnnames(tr::TableRow) = Tuple(tr.index.column_labels)
Tables.getcolumn(tr::TableRow, nm::Symbol) = getdata(tr, nm)

function writetable(filename::AbstractString, x; kw...)

	_as_vector(y::AbstractVector) = y
	_as_vector(y) = collect(y)

    writetable(filename, Any[_as_vector(c) for c in Tables.Columns(x)], collect(Symbol, Tables.columnnames(x)); kw...)
end
