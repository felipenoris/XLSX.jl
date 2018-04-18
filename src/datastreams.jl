struct Table
    ws::Worksheet
    schema::Data.Schema
end

function Table(ws::Worksheet)
    dim = dimension(ws)
    rows, cols = size(dim)
    return Table(ws, Data.Schema(collect(Any for i = 1:cols), rows))
end

Data.schema(source::Table) = source.schema
Data.accesspattern(::Type{Table}) = Data.RandomAccess
@inline Data.isdone(::Table, row, col, rows, cols) = row > rows || col > cols
@inline Data.isdone(table::Table, row, col) = Data.isdone(table, row, col, size(table.schema)...)
Data.streamtype(::Type{<:Table}, ::Type{Data.Field}) = true
@inline Data.streamfrom(source::Table, ::Type{Data.Field}, ::T, row, col::Int) = source.ws[CellRef(row, col)]

