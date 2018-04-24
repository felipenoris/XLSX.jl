
function CellRef(n::AbstractString)
    @assert is_valid_cellname(n) "$n is not a valid CellRef."

    m_row = match(r"[0-9]+$", n)
    row_number = parse(Int, m_row.match)
    m_column_name = match(r"^[A-X]?[A-Z]?[A-Z]", n)

    return CellRef(n, row_number, decode_column_number(m_column_name.match))
end

CellRef(row::Int, col::Int) = CellRef(encode_column_number(col) * string(row))

"""
    decode_column_number(column_name::AbstractString) :: Int

Converts column name to a column number.

```julia
julia> XLSX.decode_column_number("D")
4
```

See also: `encode_column_number`.
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
    encode_column_number(column_number::Int) :: String

Converts column number to a column name.

```julia
julia> XLSX.encode_column_number(4)
"D"
```

See also: `decode_column_number`.
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
Base.show(io::IO, c::CellRef) = print(io, string(c))

Base.:(==)(c1::CellRef, c2::CellRef) = c1.name == c2.name
Base.hash(c::CellRef) = hash(c.name)

const RGX_COLUMN_NAME = r"^[A-Z]?[A-Z]?[A-Z]$"
const RGX_CELLNAME = r"^[A-Z]+[0-9]+$"
const RGX_CELLRANGE = r"^[A-Z]+[0-9]+:[A-Z]+[0-9]+$"

function is_valid_column_name(n::AbstractString) :: Bool
    if !ismatch(RGX_COLUMN_NAME, n)
        return false
    end

    column_number = decode_column_number(n)
    if column_number < 1 || column_number > 16384
        return false
    end

    return true
end

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
    if !is_valid_column_name(m_column.match)
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
Base.show(io::IO, cr::CellRange) = print(io, string(cr))

Base.:(==)(cr1::CellRange, cr2::CellRange) = cr1.start == cr2.start && cr2.stop == cr2.stop
Base.hash(cr::CellRange) = hash(cr.start) + hash(cr.stop)

macro range_str(cellrange)
    CellRange(cellrange)
end

"""
    Base.in(ref::CellRef, rng::CellRange) :: Bool

Checks wether `ref` is a cell reference inside a range given by `rng`.
"""
function Base.in(ref::CellRef, rng::CellRange) :: Bool
    top = row_number(rng.start)
    bottom = row_number(rng.stop)
    r = row_number(ref)

    if top <= r && r <= bottom
        left = column_number(rng.start)
        right = column_number(rng.stop)
        c = column_number(ref)

        if left <= c && c <= right
            return true
        end
    end

    return false
end

"""
    Base.issubset(subrng::CellRange, rng::CellRange)

Checks wether `subrng` is a cell range contained in `rng`.
"""
Base.issubset(subrng::CellRange, rng::CellRange) :: Bool = in(subrng.start, rng) && in(subrng.stop, rng)

function Base.size(rng::CellRange)
    top = row_number(rng.start)
    bottom = row_number(rng.stop)
    left = column_number(rng.start)
    right = column_number(rng.stop)

    return ( bottom - top + 1, right - left + 1 )
end

"""
    row_number(c::CellRef) :: Int

Returns the row number of a given cell reference.
"""
row_number(c::CellRef) :: Int = c.row_number

"""
    column_number(c::CellRef) :: Int

Returns the column number of a given cell reference.
"""
column_number(c::CellRef) :: Int = c.column_number

column_name(c::CellRef) :: String = encode_column_number(column_number(c))

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

#
# ColumnRange
#

const RGX_COLUMN_RANGE = r"^[A-Z]?[A-Z]?[A-Z]:[A-Z]?[A-Z]?[A-Z]$"

function is_valid_column_range(r::AbstractString) :: Bool
    if !ismatch(RGX_COLUMN_RANGE, r)
        return false
    end

    start_name = match(r"^[A-Z]+", r).match
    stop_name = match(r"[A-Z]+$", r).match

    if !is_valid_column_name(start_name)
        return false
    end

    if !is_valid_column_name(stop_name)
        return false
    end

    return true
end

function ColumnRange(r::AbstractString)
    @assert is_valid_column_range(r) "Invalid column range: $r."

    start_name = match(r"^[A-Z]+", r).match
    stop_name = match(r"[A-Z]+$", r).match

    return ColumnRange(decode_column_number(start_name), decode_column_number(stop_name))
end

convert(::Type{ColumnRange}, str::AbstractString) = ColumnRange(str)
convert(::Type{ColumnRange}, column_range::ColumnRange) = column_range

column_bounds(r::ColumnRange) = (r.start, r.stop)
Base.length(r::ColumnRange) = r.stop - r.start + 1

# ColumnRange iterator
Base.start(itr::ColumnRange) = itr.start
Base.done(itr::ColumnRange, column_index::Int) = column_index > itr.stop
Base.next(itr::ColumnRange, column_index::Int) = (encode_column_number(column_index), column_index + 1)

# CellRange iterator
struct CellRefIteratorState
    row::Int
    col::Int
end

Base.start(rng::CellRange) = CellRefIteratorState(row_number(rng.start), column_number(rng.start))
Base.done(rng::CellRange, state::CellRefIteratorState) = state.row > row_number(rng.stop)

function Base.length(rng::CellRange)
    (r, c) = size(rng)
    return r * c
end

# (i, state) = next(I, state)
function Base.next(rng::CellRange, state::CellRefIteratorState)
    local next_state::CellRefIteratorState
    if state.col == column_number(rng.stop)
        # reached last column. Go to the next row.
        next_state = CellRefIteratorState(state.row + 1, column_number(rng.start))
    else
        # go to the next column
        next_state = CellRefIteratorState(state.row, state.col + 1)
    end

    return CellRef(state.row, state.col), next_state
end
