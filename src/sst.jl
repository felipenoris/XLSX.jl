
SharedStrings() = SharedStrings(Vector{String}(), false)

function sst_load!(workbook::Workbook)
    if !workbook.sst.is_loaded

        relationship_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
        if has_relationship_by_type(workbook, relationship_type)
            sst_root = xmlroot(get_xlsxfile(workbook), "xl/" * get_relationship_target_by_type(workbook, relationship_type))

            @assert EzXML.nodename(sst_root) == "sst"

            for el in EzXML.eachelement(sst_root)
                @assert EzXML.nodename(el) == "si" "Unsupported node $(EzXML.nodename(el)) in sst table."
                push!(workbook.sst.unformatted_strings, unformatted_text(el))
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
    return wb.sst.unformatted_strings[index+1]
end

@inline sst_unformatted_string(xl::XLSXFile, index::Int) :: String = sst_unformatted_string(get_workbook(xl), index)
@inline sst_unformatted_string(ws::Worksheet, index::Int) :: String = sst_unformatted_string(get_xlsxfile(ws), index)
@inline sst_unformatted_string(target::Union{Workbook, XLSXFile, Worksheet}, index_str::String) :: String = sst_unformatted_string(target, parse(Int, index_str))
