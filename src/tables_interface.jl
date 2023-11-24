
# Tables.jl interface

Tables.istable(::Type{<:TableRowIterator}) = true
Tables.rowaccess(::Type{<:TableRowIterator}) = true
Tables.rows(itr::TableRowIterator) = itr
Tables.schema(itr::TableRowIterator) = Tables.Schema(itr.index.column_labels, fill(Any, length(itr.index.column_labels)))
Tables.columnnames(tr::TableRow) = tr.index.column_labels
Tables.getcolumn(tr::TableRow, nm::Symbol) = getdata(tr, nm)
Tables.getcolumn(tr::TableRow, i::Integer) = getdata(tr, i)

_as_vector(y::AbstractVector) = y
_as_vector(y) = collect(y)

function _table_to_arrays(x)
    if Tables.istable(x)
            columns = Any[_as_vector(c) for c in Tables.Columns(x)]
            colnames = collect(Symbol, Tables.columnnames(x))
            return columns, colnames
    else
        error("$(typeof(x)) does not implement Tables.jl interface.")
    end
end

"""
    writetable(filename, table; [overwrite], [sheetname])

Write Tables.jl `table` to the specified filename.
"""
writetable(filename::Union{AbstractString, IO}, x; kw...) = writetable(filename, _table_to_arrays(x)...; kw...)

"""
    writetable(filename::Union{AbstractString, IO}, tables::Vector{Pair{String, T}}; overwrite::Bool=false)
    writetable(filename::Union{AbstractString, IO}, tables::Pair{String, Any}...; overwrite::Bool=false)
"""
function writetable(filename::Union{AbstractString, IO}, tables::Vector{<:Pair}; kw...)
    data = [(name, _table_to_arrays(x)...) for (name, x) in tables]
    return writetable(filename, data; kw...)
end

writetable(filename::Union{AbstractString, IO}, tables::Pair{<:String, <:Any}...; kw...) = writetable(filename, collect(tables); kw...)

"""
    writetable!(sheet::Worksheet, table; anchor_cell::CellRef=CellRef("A1")))

Write Tables.jl `table` to the specified sheet.
"""
writetable!(sheet::Worksheet, x; kw...) = writetable!(sheet, _table_to_arrays(x)...; kw...)

#
# DataTable
#

Tables.istable(::Type{DataTable}) = true
Tables.columnaccess(::Type{DataTable}) = true
Tables.columns(dt::DataTable) = dt # DataTable implements Tables.AbstractColumns interface
Tables.columnnames(dt::DataTable) = dt.column_labels
Tables.getcolumn(dt::DataTable, i::Int) = dt.data[i]

function Tables.getcolumn(dt::DataTable, column_label::Symbol)
    if !haskey(dt.column_label_index, column_label)
        error("Column `$column_label` not found.")
    end

    column_index = dt.column_label_index[column_label]
    return Tables.getcolumn(dt, column_index)
end
