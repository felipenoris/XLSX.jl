
@inline Base.isempty(::EmptyCell) = true
@inline Base.isempty(::AbstractCell) = false
@inline iserror(c::Cell) = c.datatype == "e"
@inline iserror(::AbstractCell) = false
@inline row_number(c::EmptyCell) = row_number(c.ref)
@inline column_number(c::EmptyCell) = column_number(c.ref)
@inline row_number(c::Cell) = row_number(c.ref)
@inline column_number(c::Cell) = column_number(c.ref)
@inline relative_cell_position(c::Cell, rng::CellRange) = relative_cell_position(c.ref, rng)
@inline relative_cell_position(c::EmptyCell, rng::CellRange) = relative_cell_position(c.ref, rng)
@inline relative_column_position(c::Cell, rng::ColumnRange) = relative_column_position(c.ref, rng)
@inline relative_column_position(c::EmptyCell, rng::ColumnRange) = relative_column_position(c.ref, rng)

Base.:(==)(c1::Cell, c2::Cell) = c1.ref == c2.ref && c1.datatype == c2.datatype && c1.style == c2.style && c1.value == c2.value && c1.formula == c2.formula
Base.hash(c::Cell) = hash(c.ref) + hash(c.datatype) + hash(c.style) + hash(c.value) + hash(c.formula)

Base.:(==)(c1::EmptyCell, c2::EmptyCell) = c1.ref == c2.ref
Base.hash(c::EmptyCell) = hash(c.ref) + 10

#=
function find_t_node_recursively(n::XML.LazyNode) :: Union{Nothing, XML.LazyNode}
    if XML.tag(n) == "t"
        return n
    else
        for child in XML.children(n)
            result = find_t_node_recursively(child)
            if result !== nothing
                return result
            end
        end
    end

    return nothing
end
=#
function Cell(c::XML.LazyNode, ws::Worksheet; mylock::Union{ReentrantLock,Nothing}=nothing) :: Union{Cell, EmptyCell}
    # c (Cell) element is defined at section 18.3.1.4
    # t (Cell Data Type) is an enumeration representing the cell's data type. The possible values for this attribute are defined by the ST_CellType simple type (§18.18.11).
    # s (Style Index) is the index of this cell's style. Style records are stored in the Styles Part.

    wb=get_workbook(ws)

    if XML.tag(c) != "c"
        throw(XLSXError("`Cell` Expects a `c` (cell) XML node."))
    end

    a = XML.attributes(c) # Dict of cell attributes

    ref = CellRef(a["r"])

    # type
    if haskey(a, "t")
        t = a["t"]
    else
        t = ""
    end

    # style
    if haskey(a, "s")
        s = a["s"]
    else
        s = ""
    end

    # Cell metadata flag (for dynamicArrays)
    if haskey(a, "cm")
        m = a["cm"]
    else
        m = ""
    end

    # iterate v and f elements
    local v::String = ""
    local f::AbstractFormula = Formula()
    local found_v::Bool = false
    local found_f::Bool = false

    for c_child_element in XML.children(c)

        if t == "inlineStr" # Convert to sharedString
            if XML.tag(c_child_element) == "is"
                uft = unformatted_text(c_child_element)
                if uft=="" # Convert empty inlineStrings to missing. Can't have empty sharedStrings
                    v=""
                    t=""
                else
                    ft=("<si>\n  "*join(XML.write.(XML.children(c_child_element)), "\n")*"\n</si>")
                    t = "s"
                    v = string(add_shared_string!(wb, uft, ft; mylock))
                end
            end
        else
            if XML.tag(c_child_element) == "v"
                if found_v # we should have only one v element
                    throw(XLSXError("Unsupported: cell $(ref) has more than 1 `v` elements."))
                else
                    found_v = true
                end              
                # v = length(c_child_element)==0 ? "" : XML.unescape(XML.simple_value(c_child_element))
                ch=XML.children(c_child_element)
                v = length(ch)==0 ? "" : XML.unescape(XML.value(ch[1])) # saves a little time!
            elseif XML.tag(c_child_element) == "f"
                if found_f # we should have only one f element
                    throw(XLSXError("Unsupported: cell $(ref) has more than 1 `f` elements."))
                else
                    found_f = true
                end
                f = parse_formula_from_element(c_child_element)
            end
        end
    end
    return Cell(ref, t, s, v, m, f)
end

function parse_formula_from_element(c_child_element) :: AbstractFormula

    if XML.tag(c_child_element) != "f"
        throw(XLSXError("Expected nodename `f`. Found: `$(XML.tag(c_child_element))`"))
    end

    if XML.is_simple(c_child_element)
        formula_string = XML.unescape(XML.simple_value(c_child_element))
    else
        fs = [x for x in XML.children(c_child_element) if XML.nodetype(x) == XML.Text]
        if length(fs)==0
            formula_string=""
        else
            formula_string=XML.unescape(XML.value(fs[1]))
        end
    end

    a = XML.attributes(c_child_element)
    unhandled_attributes=Dict{String,String}()
    if !isnothing(a)
        for (k, v) in a
            if k ∉ ["t", "si", "ref"]
                push!(unhandled_attributes, k => v)
            end
        end
    end
    is_array=false
    let ref = nothing
        if !isnothing(a)
            if haskey(a, "t")
                if a["t"] == "shared"
                    haskey(a, "si") || throw(XLSXError("Expected shared formula to have an index. `si` attribute is missing: $c_child_element"))
                    if haskey(a, "ref")
                        return ReferencedFormula(
                            formula_string,
                            parse(Int, a["si"]),
                            a["ref"],
                            length(unhandled_attributes) > 0 ? unhandled_attributes : nothing,
                        )
                    else
                        return FormulaReference(
                            parse(Int, a["si"]),
                            length(unhandled_attributes) > 0 ? unhandled_attributes : nothing,
                        )
                    end
                elseif a["t"] == "array"
                    is_array=true
                    ref = haskey(a,"ref") ? a["ref"] : nothing
                end
            end
        end
        return Formula(
            formula_string,
            is_array ? "array" : nothing,
            ref,
            length(unhandled_attributes) > 0 ? unhandled_attributes : nothing)
    end
end

# Constructor with simple formula string for backward compatibility
function Cell(ref::CellRef, datatype::String, style::String, value::String, meta::String, formula::String)
    return Cell(ref, datatype, style, value, meta, Formula(formula))
end

@inline getdata(ws::Worksheet, empty::EmptyCell) = missing

const RGX_INTEGER = r"^\-?[0-9]+$"

"""
    getdata(ws::Worksheet, cell::Cell) :: CellValue

Returns a Julia representation of a given cell value.
The result data type is chosen based on the value of the cell as well as its style.

For example, date is stored as integers inside the spreadsheet, and the style is the
information that is taken into account to chose `Date` as the result type.

For numbers, if the style implies that the number is visualized with decimals,
the method will return a float, even if the underlying number is stored
as an integer inside the spreadsheet XML.

If `cell` has empty value or empty `String`, this function will return `missing`.
"""
function getdata(ws::Worksheet, cell::Cell) :: CellValueType

    if iserror(cell)
        return missing
    end

    ecv=isempty(cell.value)
    ecd=isempty(cell.datatype)
    ecs=isempty(cell.style)

#=
    if cell.datatype == "inlineStr" # Now converted to shared strings on read
        if ecv
            return missing
        else
            return cell.value
        end
    end
=#

    if cell.datatype == "s"

        if ecv
            return missing
        end

        # use sst
        str = sst_unformatted_string(ws, cell.value)

        if isempty(str)
            return missing
        else
            return str
        end

    elseif (ecd || cell.datatype == "n")

        if ecv
            return missing
        end

        if !ecs && styles_is_datetime(ws, cell.style)
            # datetime
            return _celldata_datetime(cell.value, isdate1904(ws))

        elseif !ecs && styles_is_float(ws, cell.style)
            # float
            return parse(Float64, cell.value)

        else
            # fallback to unformatted number
            if occursin(RGX_INTEGER, cell.value)  # if contains only numbers
                v_num = parse(Int, cell.value)
            else
                v_num = parse(Float64, cell.value)
            end

            return v_num
        end
    elseif cell.datatype == "b"
        if cell.value == "0"
            return false
        elseif cell.value == "1"
            return true
        else
            throw(XLSXError("Unknown boolean value: $(cell.value)."))
        end
    elseif cell.datatype == "str"
        # plain string
        if ecv
            return missing
        else
            return cell.value
        end
    end

    throw(XLSXError("Couldn't parse data for $cell."))
end

function _celldata_datetime(v::AbstractString, _is_date_1904::Bool) :: Union{Dates.DateTime, Dates.Date, Dates.Time}

    # does not allow empty string
    if isempty(v) 
        throw(XLSXError("Cannot convert an empty string into a datetime value."))
    end

    if occursin(".", v) || v == "0"
        time_value = parse(Float64, v)
        if time_value < 0
            throw(XLSXError("Cannot have a datetime value < 0. Got $time_value"))
        end

        if time_value <= 1
            # Time
            return excel_value_to_time(time_value)
        else
            # DateTime
            return excel_value_to_datetime(time_value, _is_date_1904)
        end
    else
        # Date
        return excel_value_to_date(parse(Int, v), _is_date_1904)
    end
end

# Converts Excel number to Time.
# `x` must be between 0 and 1.
# To represent Time, Excel uses the decimal part
# of a floating point number. `1` equals one day.
function excel_value_to_time(x::Float64) :: Dates.Time
    if x >= 0 && x <= 1
        return Dates.Time(Dates.Nanosecond(round(Int, x * 86400) * 1E9 ))
    else
        throw(XLSXError("A value must be between 0 and 1 to be converted to time. Got $x"))
    end
end

time_to_excel_value(x::Dates.Time) :: Float64 = Dates.value(x) / ( 86400 * 1E9 )

# Converts Excel number to Date. See also XLSX.isdate1904.
function excel_value_to_date(x::Int, _is_date_1904::Bool) :: Dates.Date
    if _is_date_1904
        return Dates.Date(Dates.rata2datetime(x + 695056))
    else
        return Dates.Date(Dates.rata2datetime(x + 693594))
    end
end

function date_to_excel_value(date::Dates.Date, _is_date_1904::Bool) :: Int
    if _is_date_1904
        return Dates.datetime2rata(date) - 695056
    else
        return Dates.datetime2rata(date) - 693594
    end
end

# Converts Excel number to DateTime.
# The decimal part represents the Time (see `_time` function).
# The integer part represents the Date.
# See also XLSX.isdate1904.
function excel_value_to_datetime(x::Float64, _is_date_1904::Bool) :: Dates.DateTime
    if x < 0
        throw(XLSXError("Cannot have a datetime value < 0. Got $x"))
    end

    local dt::Dates.Date
    local hr::Dates.Time

    dt_part = trunc(Int, x)
    hr_part = x - dt_part

    dt = excel_value_to_date(dt_part, _is_date_1904)
    hr = excel_value_to_time(hr_part)

    return dt + hr
end

function datetime_to_excel_value(dt::Dates.DateTime, _is_date_1904::Bool) :: Float64
    return date_to_excel_value(Dates.Date(dt), _is_date_1904) + time_to_excel_value(Dates.Time(dt))
end

# Extract cells from a <row> LazyNode and push them (in place) into a Dict(column -> Cell)
function get_rowcells!(rowcells::Dict{Int, Cell}, row::XML.LazyNode, ws::Worksheet; mylock::Union{ReentrantLock,Nothing}=nothing)

#=
    # threaded cell extraction causes hugely more lock conflicts for low cellchunk size.
    # may be worthwhile if many columns (hundreds+), with a cellchunk size > ~10 or ~20, but this is unverified.

    # debug
    # @assert row.tag == "row" "Not a row node"
    cellchunk=8 # bigger chunks, fewer lock conflicts but columns are generally relatively few.
    sst_count=0
    d=row.depth

    row_cellnodes = Channel{Vector{XML.LazyNode}}(1 << 10)
    row_cells = Channel{Vector{XLSX.Cell}}(1 << 10)

    # consumer task
    consumer = @async begin
        for cells in row_cells  
            for cell in cells      
                sst_count += cell.datatype == "s" ? 1 : 0
                rowcells[column_number(cell)] = cell
            end
        end
    end

    # Feed row_cellnodes
    cellnodes = Vector{XML.LazyNode}(undef, cellchunk)
    pos=0
    cellnode=XML.next(row)
    while !isnothing(cellnode) && cellnode.depth > d
        if cellnode.tag == "c" # This is a cell
            pos += 1
            cellnodes[pos] = cellnode
        end
        if pos >= cellchunk
            put!(row_cellnodes, copy(cellnodes))
            pos=0
        end
        cellnode = XML.next(cellnode)
    end
    if pos>0 # handle last incomplete chunk
        put!(row_cellnodes, cellnodes[1:pos])
    end
    close(row_cellnodes)

    # Producer tasks
    mylock = ReentrantLock() # lock for thread-safe access to shared string table in case of inlineStrings
    @sync for _ in 1:Threads.nthreads()
        Threads.@spawn begin
            chunk = Vector{XLSX.Cell}(undef, cellchunk)
            for cns in row_cellnodes
                cell_count=0
                for cn in cns
                    cell_count += 1
                    chunk[cell_count] = Cell(cn, ws; mylock)
                    if cell_count >= cellchunk
                        put!(row_cells, copy(chunk))
                        cell_count=0
                    end
                end
                if cell_count > 0 # handle last incomplete chunk
                    put!(row_cells, chunk[1:cell_count])
                end
            end
        end
    end
    close(row_cells)

    wait(consumer)  # ensure consumer is done

    if !isnothing(cellnode) && cellnode.tag == "row" # have reached the end of last row, beginning of next
        return cellnode, sst_count
    else                                             # no more rows
        return nothing, sst_count
    end
=#
    # unthreaded cell extraction is (exceedingly marginally) slower but no lock conflicts introduced.

    # debug
    # @assert row.tag == "row" "Not a row node"

    sst_count=0

    d=row.depth

    cellnode=XML.next(row)

    while !isnothing(cellnode) && cellnode.depth > d
        if cellnode.tag == "c" # This is a cell
            cell = Cell(cellnode, ws; mylock) # construct an XLSX.Cell from an XML.LazyNode
            sst_count += cell.datatype == "s" ? 1 : 0
            rowcells[column_number(cell)] = cell
        end
        cellnode = XML.next(cellnode)
    end
    if !isnothing(cellnode) && cellnode.tag == "row" # have reached the end of last row, beginning of next
        return cellnode, sst_count
    else                                             # no more rows
        return nothing, sst_count
    end

end
