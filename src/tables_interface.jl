
# Tables.jl interface

Tables.istable(::Type{<:TableRowIterator}) = true
Tables.rowaccess(::Type{<:TableRowIterator}) = true
Tables.rows(itr::TableRowIterator) = itr
Tables.schema(itr::TableRowIterator) = Tables.Schema(itr.index.column_labels, fill(Any, length(itr.index.column_labels)))
Tables.columnnames(tr::TableRow) = tr.index.column_labels
Tables.getcolumn(tr::TableRow, nm::Symbol) = getdata(tr, nm)
Tables.getcolumn(tr::TableRow, i::Integer) = getdata(tr, i)

function writetable(filename::AbstractString, x; kw...)

	_as_vector(y::AbstractVector) = y
	_as_vector(y) = collect(y)

    writetable(filename, Any[_as_vector(c) for c in Tables.Columns(x)], collect(Symbol, Tables.columnnames(x)); kw...)
end
