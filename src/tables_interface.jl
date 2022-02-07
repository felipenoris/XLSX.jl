
# Tables.jl interface

Tables.istable(::Type{<:TableRowIterator}) = true
Tables.rowaccess(::Type{<:TableRowIterator}) = true
Tables.rows(itr::TableRowIterator) = itr
Tables.schema(itr::TableRowIterator) = Tables.Schema(itr.index.column_labels, fill(Any, length(itr.index.column_labels)))
Tables.columnnames(tr::TableRow) = tr.index.column_labels
Tables.getcolumn(tr::TableRow, nm::Symbol) = getdata(tr, nm)
Tables.getcolumn(tr::TableRow, i::Integer) = getdata(tr, i)

function table_to_arrays(x)
    _as_vector(y::AbstractVector) = y
    _as_vector(y) = collect(y)
    columns = Any[_as_vector(c) for c in Tables.Columns(x)]
    names = collect(Symbol, Tables.columnnames(x))
    return columns, names
end

"""
    writetable(filename, table; [overwrite], [sheetname])

Write Tables.jl `table` to the specified filename
"""
writetable(filename::AbstractString, x; kw...) = writetable(filename, table_to_arrays(x)...; kw...)


"""
    writetable(filename::AbstractString, tables::Vector{Pair{String, T}}; overwrite::Bool=false)
    writetable(filename::AbstractString, tables::Pair{String, Any}...; overwrite::Bool=false)
"""
function writetable(filename::AbstractString, tables::Vector{<:Pair}; kw...)
    data = [(name, table_to_arrays(x)...) for (name, x) in tables]
    return writetable(filename, data; kw...)
end
writetable(filename::AbstractString, tables::Pair{<:String, <:Any}...; kw...) = writetable(filename, collect(tables); kw...)

"""
    writetable!(sheet::Worksheet, table; anchor_cell::CellRef=CellRef("A1")))

Write Tables.jl `table` to the specified sheet
"""
writetable!(sheet::Worksheet, x, kw...) = writetable!(sheet, table_to_arrays(x)...; kw...)