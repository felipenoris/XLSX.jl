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

#struct SheetRowIteratorState
#    xml_rows_iterator_state::Ptr{LightXML.xmlBuffer}
#end

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

#=
function getrowcells(itr::SheetRowIterator, xml_vector_index::Int) :: Dict{Int, Cell}
    @assert !isempty(itr.xml_rows) && xml_vector_index >= 1 && xml_vector_index <= length(itr.xml_rows)
    rowcells = Dict{Int, Cell}()

    left = column_number(itr.rng.start)
    right = column_number(itr.rng.stop)

    for c in itr.xml_rows[xml_vector_index]["c"]
        ref = CellRef(LightXML.attribute(c, "r"))
        
        if column_number(ref) < left || right < column_number(ref)
            # filters out cells outside requested CellRange
            continue
        else
            cell = Cell(c)
            range_row = parse(Int, LightXML.attribute(itr.xml_rows[xml_vector_index], "r"))
            @assert row_number(cell.ref) == range_row "Malformed Excel file. range_row = $range_row, cell.ref = $(cell.ref)"
            rowcells[column_number(ref)] = cell
        end
    end

    return rowcells
end
=#

#=
"""
    Base.isempty(itr::SheetRowIterator, xml_vector_index::Int) :: Bool

Checks if `itr.xml_rows[xml_vector_index]` has at least one cell inside `itr.rng`.
"""
function Base.isempty(itr::SheetRowIterator, xml_vector_index::Int) :: Bool
    @assert !isempty(itr.xml_rows) && xml_vector_index >= 1 && xml_vector_index <= length(itr.xml_rows)

    left = column_number(itr.rng.start)
    right = column_number(itr.rng.stop)

    for c in itr.xml_rows[xml_vector_index]["c"]
        ref = CellRef(LightXML.attribute(c, "r"))
        
        if left >= column_number(ref) && column_number(ref) <= right
            return false
        end
    end

    return true
end

function Base.next(itr::SheetRowIterator, state::SheetRowIteratorState)
    top = row_number(itr.rng.start)
    bottom = row_number(itr.rng.stop)

    xml_vector_index = 1

    # looks for the next row of data inside required range
    xml_data_row = parse(Int, LightXML.attribute(itr.xml_rows[xml_vector_index], "r"))
    last_xml_data_row = xml_data_row - 1 # initial value. Will be used to check if rows are ordered inside the XML file

    while xml_vector_index <= length(itr.xml_rows)
        @assert last_xml_data_row < xml_data_row "Malformed Excel file: Worksheet rows are not sorted."
        
        xml_data_row = parse(Int, LightXML.attribute(itr.xml_rows[xml_vector_index], "r"))

        if xml_data_row < range_row
            # xml data is behind required row. Let's try the next one.
            last_xml_data_row = xml_data_row
            xml_vector_index += 1
        else
            if xml_data_row > range_row
                # xml data is ahead required data

                if itr.skip_empty_rows
                    # will point tho the xml row if it's stil inside the required range
                    if row_number(itr.rng.start) <= xml_data_row && xml_data_row <= row_number(itr.rng.stop)
                        range_row = xml_data_row
                    else
                        # next xml row outside required range. Will return empty data.
                        return SheetRowIteratorState(range_row, xml_vector_index, xml_data_row, true, true, true)
                    end
                else
                    # xml data is ahead required row. Will return missing values if we can't skip empty rows.
                    return SheetRowIteratorState(range_row, xml_vector_index, xml_data_row, true, false, false)
                end
            else
                @assert xml_data_row == range_row "Unicorn!"

                # checks if we should skip this row
                if itr.skip_empty_rows && isempty(itr, xml_vector_index)
                    range_row += 1
                    
                    if range_row > bottom
                        # there's no data to read
                        return SheetRowIteratorState(range_row, xml_vector_index, xml_data_row, true, true, true) 
                    else
                        last_xml_data_row = xml_data_row
                        xml_vector_index += 1
                    end
                else
                    return SheetRowIteratorState(range_row, xml_vector_index, xml_data_row, true, false, false)
                end
            end
        end
    end

    error("Unicorn!")
end

function Base.start(itr::SheetRowIterator) :: SheetRowIteratorState
    
    top = row_number(itr.rng.start)
    bottom = row_number(itr.rng.stop)

    range_row = top

    if isempty(itr.xml_rows)
        if itr.skip_empty_rows
            # will signal this is the last value, so done() will return true
            return SheetRowIteratorState(range_row, 0, 0, true, true, true)
        else
            # there's no data to be read, but will return missing if we can't skip empty rows
            return SheetRowIteratorState(range_row, 0, 0, true, false, true)
        end
    else
        ###
        # CUT
        ###

        return SheetRowIteratorState(range_row, xml_vector_index, xml_data_row, true, true, true)
    end

    error("Unicorn!")
end

Base.done(itr::SheetRowIterator, state::SheetRowIteratorState) = state.is_last

function Base.next(itr::SheetRowIterator, state::SheetRowIteratorState)

    local next_range_row::Int
    local next_xml_vector_index::Int

    if state.is_first
        next_range_row = state.range_row
        next_xml_vector_index = state.xml_vector_index
        next_xml_data_row = state.xml_data_row
    else
        # TODO
    end

    next_sheet_row = SheetRow(itr.sheet, next_range_row, getrowcells(itr, next_xml_vector_index)
    next_state = SheetRowIteratorState(next_range_row, next_xml_vector_index, next_xml_data_row, false, is_last, is_empty)
end
=#