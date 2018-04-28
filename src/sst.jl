
SharedStrings() = SharedStrings(EzXML.ElementNode("sst"))

function SharedStrings(root::EzXML.Node)
    @assert EzXML.nodename(root) == "sst"

    unformatted_strings = Vector{String}()
    for el in EzXML.elements(root)
        @assert EzXML.nodename(el) == "si" "Unsupported node $(EzXML.nodename(el)) in sst table."
        push!(unformatted_strings, unformatted_text(el))
    end

    return SharedStrings(root, unformatted_strings)
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

        for ch in EzXML.elements(e)
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
@inline sst_unformatted_string(sst::SharedStrings, index::Int) = sst.unformatted_strings[index+1]
@inline sst_unformatted_string(wb::Workbook, index::Int) :: String = sst_unformatted_string(wb.sst, index)
@inline sst_unformatted_string(xl::XLSXFile, index::Int) :: String = sst_unformatted_string(xl.workbook, index)
@inline sst_unformatted_string(ws::Worksheet, index::Int) :: String = sst_unformatted_string(ws.package, index)
sst_unformatted_string(target::Union{Workbook, XLSXFile, Worksheet}, index_str::String) :: String = sst_unformatted_string(target, parse(Int, index_str))
