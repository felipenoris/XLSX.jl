#=
for i = I   # or  "for i in I"
    # body
end

is translated into:

state = start(I)
while !done(I, state)
    (i, state) = next(I, state)
    # body
end
=#

"""
    SheetRowIterator(sheet, [cellrange]; [skip_empty_rows])

Iterates over Worksheet cells.
A `cellrange` can be specified to query for a rectangular subset of the worksheet data. It defaults to `dimension(sheet)`.
If `skip_empty_rows == true`, the iterator will skip empty rows.

See also `SheetRow`, `eachrow`.
"""
struct SheetRowIterator
    sheet::Worksheet
    xml_rows_iterator::LightXML.XMLElementIter
end

mutable struct SheetRow
    sheet::Worksheet
    row::Int
    row_xml_element::LightXML.XMLElement
    rowcells::Dict{Int, Cell} # column -> value
    is_rowcells_populated::Bool # indicates wether row_xml_element has been decoded into rowcells
end

# creates SheetRow with unpopulated rowcells
SheetRow(ws::Worksheet, row::Int, xml_element::LightXML.XMLElement) = SheetRow(ws, row, xml_element, Dict{Int, Cell}(), false)

function populate_row_cells!(r::SheetRow)
    if !r.is_rowcells_populated
        for c in r.row_xml_element["c"]
            ref = CellRef(LightXML.attribute(c, "r"))
            cell = Cell(c)
            @assert row_number(cell.ref) == r.row "Malformed Excel file. range_row = $(r.row), cell.ref = $(cell.ref)"
            r.rowcells[column_number(ref)] = cell
        end
        r.is_rowcells_populated = true
    end
    nothing
end

Base.start(itr::SheetRowIterator) = start(itr.xml_rows_iterator)
Base.done(itr::SheetRowIterator, state) = done(itr.xml_rows_iterator, state)

#(i, state) = next(I, state)
function Base.next(itr::SheetRowIterator, state)
    xml_element, next_state = next(itr.xml_rows_iterator, state)
    row = parse(Int, LightXML.attribute(xml_element, "r"))
    return SheetRow(itr.sheet, row, xml_element), next_state
end

function SheetRowIterator(ws::Worksheet)
    xroot = LightXML.root(ws.data)
    @assert LightXML.name(xroot) == "worksheet" "Malformed sheet $(ws.name)."
    vec_sheetdata = xroot["sheetData"]
    @assert length(vec_sheetdata) <= 1 "Malformed sheet $(ws.name)."
    return SheetRowIterator(ws, LightXML.child_elements(vec_sheetdata[1]))
end

row_number(sr::SheetRow) = sr.row

function getcell(r::SheetRow, column_index::Int) :: AbstractCell
    populate_row_cells!(r)

    if haskey(r.rowcells, column_index)
        return r.rowcells[column_index]
    else
        return EmptyCell()
    end
end

function getcell(r::SheetRow, column_name::AbstractString)
    @assert is_valid_column_name(column_name) "$column_name is not a valid column name."
    return getcell(r, decode_column_number(column_name))
end

eachrow(ws::Worksheet) = SheetRowIterator(ws)
