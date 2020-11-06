
# Tables.jl interface

Tables.isrowtable(::Type{<:TableRowIterator}) = true
Tables.columnnames(tr::TableRow) = Tuple(tr.index.column_labels)
Tables.getcolumn(tr::TableRow, nm::Symbol) = getdata(tr, nm)

function writetable(filename::AbstractString, x; kw...)

	_as_vector(y::AbstractVector) = y
	_as_vector(y) = collect(y)

    writetable(filename, Any[_as_vector(c) for c in Tables.Columns(x)], collect(Symbol, Tables.columnnames(x)); kw...)
end

function writetable(filename::AbstractString, data_vector::Vector{Tuple{S, T}}; kw...) where {S<:AbstractString, T}
	_as_vector(y::AbstractVector) = y
	_as_vector(y) = collect(y)

	tables = Tuple{AbstractString, Vector{Any}, Vector{Any}}[]
	for sheet in data_vector
		push!(tables, (sheet[1], Any[_as_vector(c) for c in Tables.Columns(sheet[2])], collect(Symbol, Tables.columnnames(sheet[2]))))
	end
	writetable(filename, tables; kw...)
end

function writetable(filename::AbstractString, data_dict::Dict{S, T}; kw...) where {S<:AbstractString, T}
	writetable(filename, [(sheetname, data) for (sheetname, data) in data_dict]; kw...)
end