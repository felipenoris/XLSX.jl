
SharedStringTable() = SharedStringTable(Vector{String}(), Vector{String}(), Dict{String, Int64}(), false)

@inline get_sst(wb::Workbook) = wb.sst
@inline get_sst(xl::XLSXFile) = get_sst(get_workbook(xl))
@inline Base.length(sst::SharedStringTable) = length(sst.formatted_strings)
@inline Base.isempty(sst::SharedStringTable) = isempty(sst.formatted_strings)

# Checks if string is inside shared string table.
# Returns `nothing` if it's not in the shared string table.
# Returns the index of the string in the shared string table. The index is 0-based.
function get_shared_string_index(sst::SharedStringTable, str_formatted::AbstractString) :: Union{Nothing, Int}
    @assert sst.is_loaded "Can't query shared string table because it's not loaded into memory."

    #using a Dict is much more efficient than the findfirst approach especially on large datasets
    if haskey(sst.index, str_formatted)
        return sst.index[str_formatted] - 1
    else
        return nothing
    end

end

function add_shared_string!(sst::SharedStringTable, str_unformatted::AbstractString, str_formatted::AbstractString) :: Int
    i = get_shared_string_index(sst, str_formatted)
    if i != nothing
        # it's already in the table
        return i
    else
        push!(sst.unformatted_strings, str_unformatted)
        push!(sst.formatted_strings, str_formatted)
        sst.index[str_formatted] = length(sst.formatted_strings)
        new_index = length(sst.formatted_strings) - 1 # 0-based
        @assert new_index == get_shared_string_index(sst, str_formatted) "Inconsistent state after adding a string to the Shared String Table."
        return new_index
    end
end

# Adds a string to shared string table. Returns the 0-based index of the shared string in the shared string table.
function add_shared_string!(wb::Workbook, str_unformatted::AbstractString, str_formatted::AbstractString) :: Int
    @assert is_writable(get_xlsxfile(wb)) "XLSXFile instance is not writable."
    @assert !(isempty(str_unformatted) || isempty(str_formatted)) "Can't add empty string to Shared String Table."
    sst = get_sst(wb)

    if !sst.is_loaded
        # if got to this point, the file was opened as template but doesn't have a Shared String Table.
        # Will create a new one.
        sst.is_loaded = true

        # add relationship
        #<Relationship Id="rId16" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"/>
        add_relationship!(wb, "sharedStrings.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")

        # add Content Type <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" PartName="/xl/sharedStrings.xml"/>
        ctype_root = xmlroot(get_xlsxfile(wb), "[Content_Types].xml")
        @assert EzXML.nodename(ctype_root) == "Types"
        override_node = EzXML.addelement!(ctype_root, "Override")
        override_node["ContentType"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
        override_node["PartName"] = "/xl/sharedStrings.xml"
        init_sst_index(sst)
    end

    return add_shared_string!(sst, str_unformatted, str_formatted)
end

function add_shared_string!(wb::Workbook, str_unformatted::AbstractString) :: Int
    str_formatted = string("<si><t>", str_unformatted, "</t></si>")
    return add_shared_string!(wb, str_unformatted, str_formatted)
end

function sst_load!(workbook::Workbook)
    sst = get_sst(workbook)

    if !sst.is_loaded

        relationship_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
        if has_relationship_by_type(workbook, relationship_type)
            sst_root = xmlroot(get_xlsxfile(workbook), get_relationship_target_by_type("xl", workbook, relationship_type))

            @assert EzXML.nodename(sst_root) == "sst"

            formatted_string_buffer = IOBuffer()
            for el in EzXML.eachelement(sst_root)
                @assert EzXML.nodename(el) == "si" "Unsupported node $(EzXML.nodename(el)) in sst table."
                push!(sst.unformatted_strings, unformatted_text(el))

                print(formatted_string_buffer, el)
                push!(sst.formatted_strings, String(take!(formatted_string_buffer)))
            end
            init_sst_index(sst)
            sst.is_loaded=true
            return
        end

        error("Shared Strings Table not found for this workbook.")
    end
end

# Checks whether this workbook has a Shared String Table.
function has_sst(workbook::Workbook) :: Bool
    relationship_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    return has_relationship_by_type(workbook, relationship_type)
end

# Helper function to gather unformatted text from Excel data files.
# It looks at all children of `el` for tag name `t` and returns
# a join of all the strings found.
function unformatted_text(el::EzXML.Node) :: String

    function gather_strings!(v::Vector{String}, e::EzXML.Node)
        if EzXML.nodename(e) == "t"
            push!(v, EzXML.nodecontent(e))
        end

        for ch in EzXML.eachelement(e)
            if EzXML.nodename(e) != "rPh"
                gather_strings!(v, ch)
            end 
        end

        nothing
    end

    v_string = Vector{String}()
    gather_strings!(v_string, el)
    return join(v_string)
end

# Looks for a string inside the Shared Strings Table (sst).
# `index` starts at 0.
@inline function sst_unformatted_string(wb::Workbook, index::Int)
    sst_load!(wb)
    return get_sst(wb).unformatted_strings[index+1]
end

# Looks for a formatted string inside the Shared Strings Table (sst).
# `index` starts at 0.
@inline function sst_formatted_string(wb::Workbook, index::Int)
    sst_load!(wb)
    return get_sst(wb).formatted_strings[index+1]
end

@inline sst_unformatted_string(xl::XLSXFile, index::Int) :: String = sst_unformatted_string(get_workbook(xl), index)
@inline sst_unformatted_string(ws::Worksheet, index::Int) :: String = sst_unformatted_string(get_xlsxfile(ws), index)
@inline sst_unformatted_string(target::Union{Workbook, XLSXFile, Worksheet}, index_str::String) :: String = sst_unformatted_string(target, parse(Int, index_str))


# init the index table
function init_sst_index(sst::SharedStringTable)
    empty!(sst.index)
    for i in 1:length(sst.formatted_strings)
        sst.index[sst.formatted_strings[i]] = i
    end
end
