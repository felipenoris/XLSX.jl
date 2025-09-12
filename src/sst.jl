
SharedStringTable() = SharedStringTable(Vector{String}(), Vector{String}(), Dict{String, Int64}(), false)

@inline get_sst(wb::Workbook) = wb.sst
@inline get_sst(xl::XLSXFile) = get_sst(get_workbook(xl))
@inline Base.length(sst::SharedStringTable) = length(sst.formatted_strings)
@inline Base.isempty(sst::SharedStringTable) = isempty(sst.formatted_strings)

# Checks if string is inside shared string table.
# Returns `nothing` if it's not in the shared string table.
# Returns the index of the string in the shared string table. The index is 0-based.
function get_shared_string_index(sst::SharedStringTable, str_formatted::AbstractString) :: Union{Nothing, Int}
    !sst.is_loaded && throw(XLSXError("Can't query shared string table because it's not loaded into memory."))

    #using a Dict is much more efficient than the findfirst approach especially on large datasets
    if haskey(sst.index, str_formatted)
        return sst.index[str_formatted] - 1
    else
        return nothing
    end

end

function add_shared_string!(sst::SharedStringTable, str_unformatted::AbstractString, str_formatted::AbstractString) :: Int
    i = get_shared_string_index(sst, str_formatted)
    if i !== nothing
        # it's already in the table
        return i
    else
        push!(sst.unformatted_strings, str_unformatted)
        push!(sst.formatted_strings, str_formatted)
        sst.index[str_formatted] = length(sst.formatted_strings)
        new_index = length(sst.formatted_strings) - 1 # 0-based
        if new_index != get_shared_string_index(sst, str_formatted)
            throw(XLSXError("Inconsistent state after adding a string to the Shared String Table."))
        end
        return new_index
    end
end

# Adds a string to shared string table. Returns the 0-based index of the shared string in the shared string table.
function add_shared_string!(wb::Workbook, str_unformatted::AbstractString, str_formatted::AbstractString) :: Int
    !is_writable(get_xlsxfile(wb)) && throw(XLSXError("XLSXFile instance is not writable."))
    (isempty(str_unformatted) || isempty(str_formatted)) && throw(XLSXError("Can't add empty string to Shared String Table."))
    sst = get_sst(wb)

    if !sst.is_loaded
        # if got to this point, the file was opened as template but doesn't have a Shared String Table.
        # Will create a new one.
        sst.is_loaded = true

        # add relationship
        #<Relationship Id="rId16" Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"/>
        add_relationship!(wb, "sharedStrings.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")

        # add Content Type <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" PartName="/xl/sharedStrings.xml"/>
        ctype_root = xmlroot(get_xlsxfile(wb), "[Content_Types].xml")[end]
        XML.tag(ctype_root) != "Types" && throw(XLSXError("Something wrong here!"))
        override_node = XML.Element("Override";
            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            PartName = "/xl/sharedStrings.xml"
        )
        push!(ctype_root, override_node)
        init_sst_index(sst)
    end

    return add_shared_string!(sst, str_unformatted, str_formatted)
end

# allow to write cells containing only whitespace characters or with leading or trailing whitespace.
function add_shared_string!(wb::Workbook, str_unformatted::AbstractString) :: Int
    if startswith(str_unformatted, " ") || endswith(str_unformatted, " ")
        str_formatted = string("<si><t xml:space=\"preserve\">", XML.escape(str_unformatted), "</t></si>")
    else
        str_formatted = string("<si><t>", XML.escape(str_unformatted), "</t></si>")
    end
    return add_shared_string!(wb, str_unformatted, str_formatted)
end

function sst_load!(workbook::Workbook)
    sst = get_sst(workbook)
    if !sst.is_loaded

        relationship_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
        if has_relationship_by_type(workbook, relationship_type)
            sst_chan = stream_ssts(open_internal_file_stream(get_xlsxfile(workbook), "xl/sharedStrings.xml")[end])
            load_sst_table!(workbook, sst_chan, Threads.nthreads())
            init_sst_index(sst)
            sst.is_loaded=true
            
            return
        end

        throw(XLSXError("Shared Strings Table not found for this workbook."))
    end
end
function stream_ssts(io::XML.LazyNode; channel_size::Int=1 << 20)
    n = XML.next(io)
    i=0
    Channel{SstToken}(channel_size) do out
        while !isnothing(n)
            if n.tag == "si"
                i += 1
                put!(out, SstToken(n, i))
            end
            n = XML.next(n)
        end
    end
end

function process_sst(sst::SstToken)
    el = sst.n
    i = sst.idx

    if XML.nodetype(el) != XML.Text
        XML.tag(el) != "si" && throw(XLSXError("Unsupported node $(XML.tag(el)) in sst table."))
        sst = Sst(unformatted_text(el), XML.write(el), i)
        return sst

    end

end

function load_sst_table!(wb::Workbook, chan::Channel, nthreads::Int)
    sst_table = get_sst(wb)

    sst_results = Channel{Sst}(1 << 24)

    all_ssts = Vector{Tuple{Int,Sst}}()
    consumer = @async begin        
        for sst in sst_results
            push!(all_ssts, (sst.idx, sst))
        end    

        sort!(all_ssts, by = x -> x[1])
    
        empty!(sst_table.index)
        for sst in all_ssts
            push!(sst_table.unformatted_strings, sst[end].unformatted)
            push!(sst_table.formatted_strings, sst[end].formatted)
            sst_table.index[sst[end].formatted] = sst[begin]
        end
    
    end

    # Producer tasks
    @sync begin
        for _ in 1:nthreads
            Threads.@spawn begin
                for tok in chan
                    result = process_sst(tok)
                    put!(sst_results, result)
                end
            end
        end
    end
    close(sst_results)

    wait(consumer)  # ensure consumer is done

    sst_table.is_loaded=true
    
end

# Checks whether this workbook has a Shared String Table.
function has_sst(workbook::Workbook) :: Bool
    relationship_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    return has_relationship_by_type(workbook, relationship_type)
end

# Helper function to gather unformatted text from Excel data files.
# It looks at all children of `el` for tag name `t` and returns
# a join of all the strings found.
function unformatted_text(el::XML.LazyNode) :: String

    function gather_strings!(v::Vector{String}, e::XML.LazyNode)
        if XML.tag(e) == "t"
            c=XML.children(e)
            if length(c) == 1
                push!(v, XML.is_simple(c[1]) ? XML.simple_value(c[1]) : XML.value(c[1]))
            elseif length(c) == 0
                push!(v, isnothing(XML.value(e)) ? "" : XML.is_simple(e) ? XML.simple_value(e) : XML.value(e))
            else
                throw(XLSXError("Unexpected number of children in <t> node: $(length(c)). Expected 0 or 1."))
            end
        end

        if XML.tag(e) != "rPh"
            for ch in XML.children(e)
                # recursively gather strings from children
                gather_strings!(v, ch)
            end 
        end

        nothing
    end

    v_string = Vector{String}()
    gather_strings!(v_string, el)

    return XML.unescape(join(v_string))
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
