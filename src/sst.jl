
SharedStrings() = SharedStrings(Vector{String}(), Vector{String}(), Dict{UInt64, Int}(), false)

@inline get_sst(wb::Workbook) = wb.sst
@inline get_sst(xl::XLSXFile) = get_sst(get_workbook(xl))
@inline Base.length(sst::SharedStrings) = length(sst.formatted_strings)
@inline Base.isempty(sst::SharedStrings) = isempty(sst.formatted_strings)

"""
Checks if string is inside shared string table.
Returns -1 if it's not in the shared string table.
Returns the index of the string in the shared string table. The index is 0-based.
"""
function get_shared_string_index(sst::SharedStrings, str_formatted::AbstractString) :: Int
    @assert sst.is_loaded "Can't query shared string table because it's not loaded into memory."

    h = hash(str_formatted)
    if haskey(sst.hashmap, h)
        i = sst.hashmap[h]
        @assert sst.formatted_strings[i+1] == str_formatted "\"Congratulations. You've just discovered the secret message. Please send your answer to Old Pink, care of the Funny Farm, Chalfont…\"\nPlease, file an issue at https://github.com/felipenoris/XLSX.jl..."
        return i
    else
        return -1
    end
end

"""
    add_shared_string!(sheet, str_unformatted, [str_formatted]) :: Int

Add string to shared string table. Returns the 0-based index of the shared string in the shared string table.
"""
function add_shared_string!(wb::Workbook, str_unformatted::AbstractString, str_formatted::AbstractString) :: Int
    @assert !(isempty(str_unformatted) || isempty(str_formatted)) "Can't add empty string to Shared String Table."
    sst = get_sst(wb)
    @assert sst.is_loaded "Can't query shared string table because it's not loaded into memory."

    i = get_shared_string_index(sst, str_formatted)
    if i != -1
        # it's already in the table
        return i
    end

    push!(sst.unformatted_strings, str_unformatted)
    push!(sst.formatted_strings, str_formatted)
    new_index = length(sst.formatted_strings) - 1 # 0-based
    sst.hashmap[hash(str_formatted)] = new_index
    return new_index
end

function add_shared_string!(wb::Workbook, str_unformatted::AbstractString) :: Int
    str_formatted = string("<si><t>", str_unformatted, "</t></si>")
    return add_shared_string!(wb, str_unformatted, str_formatted)
end

function sst_load!(workbook::Workbook)
    if !workbook.sst.is_loaded

        relationship_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
        if has_relationship_by_type(workbook, relationship_type)
            sst_root = xmlroot(get_xlsxfile(workbook), "xl/" * get_relationship_target_by_type(workbook, relationship_type))

            @assert EzXML.nodename(sst_root) == "sst"

            formatted_string_buffer = IOBuffer()
            for el in EzXML.eachelement(sst_root)
                @assert EzXML.nodename(el) == "si" "Unsupported node $(EzXML.nodename(el)) in sst table."
                push!(get_sst(workbook).unformatted_strings, unformatted_text(el))

                print(formatted_string_buffer, el)
                push!(get_sst(workbook).formatted_strings, String(take!(formatted_string_buffer)))
            end

            workbook.sst.is_loaded=true
            return
        end

        error("Shared Strings Table not found for this workbook.")
    end
end

"""
    unformatted_text(el::EzXML.Node) :: String

Helper function to gather unformatted text from Excel data files.
It looks at all childs of `el` for tag name `t` and returns
a join of all the strings found.
"""
function unformatted_text(el::EzXML.Node) :: String

    function gather_strings!(v::Vector{String}, e::EzXML.Node)
        if EzXML.nodename(e) == "t"
            push!(v, EzXML.nodecontent(e))
        end

        for ch in EzXML.eachelement(e)
            gather_strings!(v, ch)
        end

        nothing
    end

    v_string = Vector{String}()
    gather_strings!(v_string, el)
    return join(v_string)
end

"""
    sst_unformatted_string(wb, index) :: String

Looks for a string inside the Shared Strings Table (sst).
`index` starts at 0.
"""
@inline function sst_unformatted_string(wb::Workbook, index::Int)
    sst_load!(wb)
    return get_sst(wb).unformatted_strings[index+1]
end

"""
    sst_formatted_string(wb, index) :: String

Looks for a formatted string inside the Shared Strings Table (sst).
`index` starts at 0.
"""
@inline function sst_formatted_string(wb::Workbook, index::Int)
    sst_load!(wb)
    return get_sst(wb).formatted_strings[index+1]
end

@inline sst_unformatted_string(xl::XLSXFile, index::Int) :: String = sst_unformatted_string(get_workbook(xl), index)
@inline sst_unformatted_string(ws::Worksheet, index::Int) :: String = sst_unformatted_string(get_xlsxfile(ws), index)
@inline sst_unformatted_string(target::Union{Workbook, XLSXFile, Worksheet}, index_str::String) :: String = sst_unformatted_string(target, parse(Int, index_str))
