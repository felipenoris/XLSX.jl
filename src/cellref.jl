
function CellRef(n::AbstractString)
    @assert is_valid_cellname(n) "$n is not a valid CellRef."

    m_row = match(r"[0-9]+$", n)
    row_number = parse(Int, m_row.match)
    m_column_name = match(r"^[A-X]?[A-Z]?[A-Z]", n)

    return CellRef(n, m_column_name.match, row_number, decode_column_number(m_column_name.match))
end

CellRef(row::Int, col::Int) = CellRef(encode_column_number(col) * string(row))

"""
Converts column name to a column number.
"""
function decode_column_number(column_name::AbstractString) :: Int
    local result::Int = 0

    num_characters = length(column_name)

    iteration = 1
    for i in num_characters:-1:1
        column_char_as_int = Int(column_name[i])
        result += (26^(iteration-1)) * (column_char_as_int - 64) # From A to Z we have 26 values. 'A' Char is ASCII 65.
        iteration += 1
    end

    return result
end

"""
Converts column number to a column name.
"""
function encode_column_number(column_number::Int) :: String
    @assert column_number > 0 && column_number <= 16384 "Column number should be in the range from 1 to 16384."

    third_letter_sequence = div(column_number, 26^2)
    column_number = column_number - third_letter_sequence*(26^2)

    second_letter_sequence = div(column_number, 26) # 26^1
    column_number = column_number - second_letter_sequence*(26)

    first_letter_sequence = column_number # 26^0

    if third_letter_sequence > 0
        # result will have 3 letters
        return String([ Char(third_letter_sequence+64), Char(second_letter_sequence+64), Char(first_letter_sequence+64) ])

    elseif second_letter_sequence > 0
        # result will have 2 letters
        return String([ Char(second_letter_sequence+64), Char(first_letter_sequence+64) ])

    else
        # result will have 1 letter
        return String([ Char(first_letter_sequence+64) ])
    end
end

Base.string(c::CellRef) = c.name
Base.show(io::IO, c::CellRef) = show(io, string(c))

Base.:(==)(c1::CellRef, c2::CellRef) = c1.name == c2.name
Base.hash(c::CellRef) = hash(c.name)

const RGX_CELLNAME = r"^[A-Z]+[0-9]+$"
const RGX_CELLRANGE = r"^[A-Z]+[0-9]+:[A-Z]+[0-9]+$"

# Cellname is bounded by A1 : XFD1048576
function is_valid_cellname(n::AbstractString) :: Bool

    if !ismatch(RGX_CELLNAME, n)
        return false
    end

    m_row = match(r"[0-9]+$", n)
    row = parse(Int, m_row.match)

    if row < 1 || row > 1048576
        return false
    end

    m_column = match(r"^[A-Z]+", n)
    column_name = m_column.match
    column_number = decode_column_number(column_name)

    if column_number < 1 || column_number > 16384
        return false
    end

    return true
end

function is_valid_cellrange(n::AbstractString) :: Bool
    if !ismatch(RGX_CELLRANGE, n)
        return false
    end
    
    m_start = match(r"^[A-Z]+[0-9]+", n)
    start_name = m_start.match
    if !is_valid_cellname(start_name)
        return false
    end

    m_stop = match(r"[A-Z]+[0-9]+$", n)
    stop_name = m_stop.match
    if !is_valid_cellname(stop_name)
        return false
    end

    return true
end

macro ref_str(ref)
    CellRef(ref)
end

function CellRange(r::AbstractString)
    @assert ismatch(RGX_CELLRANGE, r) "Invalid cell range: $r."
    
    m_start = match(r"^[A-Z]+[0-9]+", r)
    start_name = CellRef(m_start.match)

    m_stop = match(r"[A-Z]+[0-9]+$", r)
    stop_name = CellRef(m_stop.match)

    return CellRange(start_name, stop_name)
end

Base.string(cr::CellRange) = "$(string(cr.start)):$(string(cr.stop))"
Base.show(io::IO, cr::CellRange) = show(io, string(cr))

Base.:(==)(cr1::CellRange, cr2::CellRange) = cr1.start == cr2.start && cr2.stop == cr2.stop
Base.hash(cr::CellRange) = hash(cr.start) + hash(cr.stop)

macro range_str(cellrange)
    CellRange(cellrange)
end

"""
Checks wether `c` is a cell name inside a range given by `r`.
"""
function Base.in(cell::CellRef, rng::CellRange)
    top = row_number(rng.start)
    bottom = row_number(rng.stop)
    r = row_number(cell)

    if top <= r && r <= bottom
        left = column_number(rng.start)
        right = column_number(rng.stop)
        c = column_number(cell)

        if left <= c && c <= right
            return true
        end
    end

    return false
end

Base.issubset(subrng::CellRange, rng::CellRange) = in(subrng.start, rng) && in(subrng.stop, rng)

function Base.size(rng::CellRange)
    top = row_number(rng.start)
    bottom = row_number(rng.stop)
    left = column_number(rng.start)
    right = column_number(rng.stop)

    return ( bottom - top + 1, right - left + 1 )
end

row_number(c::CellRef) = c.row_number
column_number(c::CellRef) = c.column_number

"""
Returns (row, column) representing a `ref` position relative to `rng`.

For example, for a range "B2:D4", we have:

* "C3" relative position is (2, 2)

* "B2" relative position is (1, 1)

* "C4" relative position is (3, 2)

* "D4" relative position is (3, 3)

"""
function relative_cell_position(ref::CellRef, rng::CellRange)
    @assert ref ∈ rng "$ref is outside range $rng."

    top = row_number(rng.start)
    left = column_number(rng.start)

    r, c = row_number(ref), column_number(ref)

    return ( r - top + 1 , c - left + 1 )
end
