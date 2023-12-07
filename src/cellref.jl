
function CellRef(n::AbstractString)
    @assert is_valid_cellname(n) "$n is not a valid CellRef."
    column_name, row_number = split_cellname(n)
    return CellRef(n, row_number, decode_column_number(column_name))
end

@inline CellRef(row::Int, col::Int) = CellRef(encode_column_number(col) * string(row))
@inline CellPosition(ref::CellRef) = CellPosition(row_number(ref), column_number(ref))
@inline row_number(p::CellPosition) = p.row
@inline column_number(p::CellPosition) = p.column
@inline CellRef(p::CellPosition) = CellRef(row_number(p), column_number(p))

# Converts column name to a column number. See also XLSX.encode_column_number.
function decode_column_number(column_name::AbstractString) :: Int
    local result::Int = 0

    @assert isascii(column_name) "$column_name is not a valid column name."
    num_characters = length(column_name) # this is safe, since `column_name` is encoded as ASCII

    iteration = 1
    for i in num_characters:-1:1
        column_char_as_int = Int(column_name[i])
        result += (26^(iteration-1)) * (column_char_as_int - 64) # From A to Z we have 26 values. 'A' Char is ASCII 65.
        iteration += 1
    end

    return result
end

# Converts column number to a column name. See also XLSX.decode_column_number.
function encode_column_number(column_number::Int) :: String
    @assert column_number > 0 && column_number <= EXCEL_MAX_COLS "Column number should be in the range from 1 to $EXCEL_MAX_COLS."

    third_letter_sequence = div(column_number - 26 - 1, 26*26)
    column_number = column_number - third_letter_sequence*(26*26)

    second_letter_sequence = div(column_number - 1, 26) # 26^1
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
    if !occursin(RGX_COLUMN_NAME, n)
        return false
    end

    column_number = decode_column_number(n)
    if column_number < 1 || column_number > EXCEL_MAX_COLS
        return false
    end

    return true
end

const RGX_CELLNAME_LEFT = r"^[A-Z]+"
const RGX_CELLNAME_RIGHT = r"[0-9]+$"

# Splits a string representing a cell name to its column name and row number.
@inline function split_cellname(n::AbstractString)
    @assert isascii(n) "$n is not a valid cell name."
    for (i, c) in enumerate(n)
        if isdigit(c) # this block is safe since n is encoded as ASCII
            column_name = SubString(n, 1, i-1)
            row = parse(Int, SubString(n, i, length(n)))

            return column_name, row
        end
    end

    error("Couldn't split (column_name, row) for cellname $n.")
end

# Checks whether `n` is a valid name for a cell.
function is_valid_cellname(n::AbstractString) :: Bool

    if !occursin(RGX_CELLNAME, n)
        return false
    end

    column_name, row = split_cellname(n)

    if row < 1 || row > EXCEL_MAX_ROWS
        return false
    end

    if !is_valid_column_name(column_name)
        return false
    end

    return true
end

const RGX_CELLRANGE_START = r"^[A-Z]+[0-9]+"
const RGX_CELLRANGE_STOP = r"[A-Z]+[0-9]+$"

#=
    split_cellrange(n::AbstractString) -> start_name, stop_name

Splits a string representing a cell range into its cell names.

# Example

```julia
julia> XLSX.split_cellrange("AB12:CD24")
("AB12", "CD24")
```
=#
@inline function split_cellrange(n::AbstractString)
    s = split(n, ":")
    @assert length(s) == 2 "$n is not a valid cell range."
    return s[1], s[2]
end

function is_valid_cellrange(n::AbstractString) :: Bool

    if !occursin(RGX_CELLRANGE, n)
        return false
    end

    start_name, stop_name = split_cellrange(n)

    if !is_valid_cellname(start_name)
        return false
    end

    if !is_valid_cellname(stop_name)
        return false
    end

    return true
end

macro ref_str(ref)
    CellRef(ref)
end

function CellRange(r::AbstractString)
    @assert occursin(RGX_CELLRANGE, r) "Invalid cell range: $r."
    start_name, stop_name = split_cellrange(r)
    return CellRange(CellRef(start_name), CellRef(stop_name))
end

CellRange(start_row::Integer, start_column::Integer, stop_row::Integer, stop_column::Integer) = CellRange(CellRef(start_row, start_column), CellRef(stop_row, stop_column))

Base.string(cr::CellRange) = "$(string(cr.start)):$(string(cr.stop))"
Base.show(io::IO, cr::CellRange) = print(io, string(cr))
Base.:(==)(cr1::CellRange, cr2::CellRange) = cr1.start == cr2.start && cr2.stop == cr2.stop
Base.hash(cr::CellRange) = hash(cr.start) + hash(cr.stop)

macro range_str(cellrange)
    CellRange(cellrange)
end

# Checks whether `ref` is a cell reference inside a range given by `rng`.
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

# Checks whether `subrng` is a cell range contained in `rng`.
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

#=
Returns (row, column) representing a `ref` position relative to `rng`.

For example, for a range "B2:D4", we have:

* "C3" relative position is (2, 2)

* "B2" relative position is (1, 1)

* "C4" relative position is (3, 2)

* "D4" relative position is (3, 3)
=#
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

Base.string(cr::ColumnRange) = "$(encode_column_number(cr.start)):$(encode_column_number(cr.stop))"
Base.show(io::IO, cr::ColumnRange) = print(io, string(cr))
Base.:(==)(cr1::ColumnRange, cr2::ColumnRange) = cr1.start == cr2.start && cr2.stop == cr2.stop
Base.hash(cr::ColumnRange) = hash(cr.start) + hash(cr.stop)
Base.in(column_number::Integer, rng::ColumnRange) = rng.start <= column_number && column_number <= rng.stop

function relative_column_position(column_number::Integer, rng::ColumnRange)
    @assert column_number ∈ rng "Column $column_number is outside range $rng."
    return column_number - rng.start + 1
end

@inline relative_column_position(ref::CellRef, rng::ColumnRange) = relative_column_position(column_number(ref), rng)

const RGX_COLUMN_RANGE = r"^[A-Z]?[A-Z]?[A-Z]:[A-Z]?[A-Z]?[A-Z]$"
const RGX_COLUMN_RANGE_START = r"^[A-Z]+"
const RGX_COLUMN_RANGE_STOP = r"[A-Z]+$"
const RGX_SINGLE_COLUMN = r"^[A-Z]+$"

# Returns tuple (column_name_start, column_name_stop).
@inline function split_column_range(n::AbstractString)
    if !occursin(":", n)
        return n, n
    else
        s = split(n, ":")
        return s[1], s[2]
    end
end

function is_valid_column_range(r::AbstractString) :: Bool

    if occursin(RGX_SINGLE_COLUMN, r)
        return true
    end

    if !occursin(RGX_COLUMN_RANGE, r)
        return false
    end

    start_name, stop_name = split_column_range(r)

    if !is_valid_column_name(start_name) || !is_valid_column_name(stop_name)
        return false
    end

    return true
end

function ColumnRange(r::AbstractString)
    @assert is_valid_column_range(r) "Invalid column range: $r."
    start_name, stop_name = split_column_range(r)
    return ColumnRange(decode_column_number(start_name), decode_column_number(stop_name))
end

convert(::Type{ColumnRange}, str::AbstractString) = ColumnRange(str)
convert(::Type{ColumnRange}, column_range::ColumnRange) = column_range

column_bounds(r::ColumnRange) = (r.start, r.stop)
Base.length(r::ColumnRange) = r.stop - r.start + 1

# ColumnRange iterator: element is a String with the column name, the state is the column number.
function Base.iterate(itr::ColumnRange, state::Int=itr.start)
    if state > itr.stop
        return nothing
    end

    return encode_column_number(state), state + 1
end

# CellRange iterator: element is a CellRef, the state is a CellPosition.
function Base.iterate(rng::CellRange, state::CellPosition=CellPosition(rng.start))

    if row_number(state) > row_number(rng.stop)
        return nothing
    elseif column_number(state) == column_number(rng.stop)
        # reached last column. Go to the next row.
        next_state = CellPosition(row_number(state) + 1, column_number(rng.start))
    else
        # go to the next column
        next_state = CellPosition(row_number(state), column_number(state) + 1)
    end

    return CellRef(state), next_state
end

function Base.length(rng::CellRange)
    (r, c) = size(rng)
    return r * c
end

#
# SheetCellRef, SheetCellRange, SheetColumnRange
#

Base.string(cr::SheetCellRef) = string(cr.sheet, "!", cr.cellref)
Base.show(io::IO, cr::SheetCellRef) = print(io, string(cr))
Base.:(==)(cr1::SheetCellRef, cr2::SheetCellRef) = cr1.sheet == cr2.sheet && cr2.cellref == cr2.cellref
Base.hash(cr::SheetCellRef) = hash(cr.sheet) + hash(cr.cellref)

Base.string(cr::SheetCellRange) = string(cr.sheet, "!", cr.rng)
Base.show(io::IO, cr::SheetCellRange) = print(io, string(cr))
Base.:(==)(cr1::SheetCellRange, cr2::SheetCellRange) = cr1.sheet == cr2.sheet && cr2.rng == cr2.rng
Base.hash(cr::SheetCellRange) = hash(cr.sheet) + hash(cr.rng)

Base.string(cr::SheetColumnRange) = string(cr.sheet, "!", cr.colrng)
Base.show(io::IO, cr::SheetColumnRange) = print(io, string(cr))
Base.:(==)(cr1::SheetColumnRange, cr2::SheetColumnRange) = cr1.sheet == cr2.sheet && cr2.colrng == cr2.colrng
Base.hash(cr::SheetColumnRange) = hash(cr.sheet) + hash(cr.colrng)

const RGX_SHEET_CELLNAME = r"^.+![A-Z]+[0-9]+$"
const RGX_SHEET_CELLRANGE = r"^.+![A-Z]+[0-9]+:[A-Z]+[0-9]+$"
const RGX_SHEET_COLUMN_RANGE = r"^.+![A-Z]?[A-Z]?[A-Z]:[A-Z]?[A-Z]?[A-Z]$"

const RGX_SHEET_CELLNAME_RIGHT = r"[A-Z]+[0-9]+$"
const RGX_SHEET_CELLRANGE_RIGHT = r"[A-Z]+[0-9]+:[A-Z]+[0-9]+$"
const RGX_SHEET_COLUMN_RANGE_RIGHT = r"[A-Z]?[A-Z]?[A-Z]:[A-Z]?[A-Z]?[A-Z]$"

function is_valid_sheet_cellname(n::AbstractString) :: Bool
    if !occursin(RGX_SHEET_CELLNAME, n)
        return false
    end

    cellname = match(RGX_SHEET_CELLNAME_RIGHT, n).match
    if !is_valid_cellname(cellname)
        return false
    end

    return true
end

function is_valid_sheet_cellrange(n::AbstractString) :: Bool
    if !occursin(RGX_SHEET_CELLRANGE, n)
        return false
    end

    cellrange = match(RGX_SHEET_CELLRANGE_RIGHT, n).match
    if !is_valid_cellrange(cellrange)
        return false
    end

    return true
end

function is_valid_sheet_column_range(n::AbstractString) :: Bool
    if !occursin(RGX_SHEET_COLUMN_RANGE, n)
        return false
    end

    column_range = match(RGX_SHEET_COLUMN_RANGE_RIGHT, n).match
    if !is_valid_column_range(column_range)
        return false
    end

    return true
end

const RGX_SHEET_PREFIX = r"^.+!"
const RGX_CELLNAME_RIGHT_FIXED = r"\$[A-Z]+\$[0-9]+$"
const RGX_SHEET_CELNAME_RIGHT_FIXED = r"\$[A-Z]+\$[0-9]+:\$[A-Z]+\$[0-9]+$"

function parse_sheetname_from_sheetcell_name(n::AbstractString) :: SubString
    @assert occursin(RGX_SHEET_PREFIX, n) "$n is not a SheetCell reference."
    sheetname = match(RGX_SHEET_PREFIX, n).match
    sheetname = SubString(sheetname, firstindex(sheetname), prevind(sheetname, lastindex(sheetname)))
    return sheetname
end

function SheetCellRef(n::AbstractString)
    local cellref::CellRef

    if is_valid_fixed_sheet_cellname(n)
        fixed_cellname = match(RGX_CELLNAME_RIGHT_FIXED, n).match
        cellref = CellRef(replace(fixed_cellname, "\$" => ""))
    else
        @assert is_valid_sheet_cellname(n) "$n is not a valid SheetCellRef."
        cellref = CellRef(match(RGX_SHEET_CELLNAME_RIGHT, n).match)
    end
    sheetname = parse_sheetname_from_sheetcell_name(n)
    return SheetCellRef(sheetname, cellref)
end

function SheetCellRange(n::AbstractString)
    local cellrange::CellRange

    if is_valid_fixed_sheet_cellrange(n)
        fixed_cellrange = match(RGX_SHEET_CELNAME_RIGHT_FIXED, n).match
        cellrange = CellRange(replace(fixed_cellrange, "\$" => ""))
    else
        @assert is_valid_sheet_cellrange(n) "$n is not a valid SheetCellRange."
        cellrange = CellRange(match(RGX_SHEET_CELLRANGE_RIGHT, n).match)
    end

    sheetname = parse_sheetname_from_sheetcell_name(n)
    return SheetCellRange(sheetname, cellrange)
end

function SheetColumnRange(n::AbstractString)
    @assert is_valid_sheet_column_range(n) "$n is not a valid SheetColumnRange."
    column_range = match(RGX_SHEET_COLUMN_RANGE_RIGHT, n).match
    sheetname = parse_sheetname_from_sheetcell_name(n)
    return SheetColumnRange(sheetname, ColumnRange(column_range))
end

# Named ranges
const RGX_FIXED_SHEET_CELLNAME = r"^.+!\$[A-Z]+\$[0-9]+$"
const RGX_FIXED_SHEET_CELLRANGE = r"^.+!\$[A-Z]+\$[0-9]+:\$[A-Z]+\$[0-9]+$"

is_valid_fixed_sheet_cellname(s::AbstractString) = occursin(RGX_FIXED_SHEET_CELLNAME, s)
is_valid_fixed_sheet_cellrange(s::AbstractString) = occursin(RGX_FIXED_SHEET_CELLRANGE, s)
