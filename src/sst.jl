
"""
    unformatted_text(el::LightXML.XMLElement) :: String

Helper function to gather unformatted text from Excel data files.
It looks at all childs of `el` for tag name `t` and returns
a join of all the strings found.
"""
function unformatted_text(el::LightXML.XMLElement) :: String

    function gather_strings!(v::Vector{String}, e::LightXML.XMLElement)
        if LightXML.name(e) == "t"
            push!(v, LightXML.content(e))
        end

        for ch in LightXML.child_elements(e)
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
sst_unformatted_string(wb::Workbook, index::Int) :: String = unformatted_text(wb.sst[index+1])
sst_unformatted_string(xl::XLSXFile, index::Int) :: String = sst_unformatted_string(xl.workbook, index)
sst_unformatted_string(ws::Worksheet, index::Int) :: String = sst_unformatted_string(ws.package, index)
sst_unformatted_string(target::Union{Workbook, XLSXFile, Worksheet}, index_str::String) = sst_unformatted_string(target, parse(Int, index_str))
