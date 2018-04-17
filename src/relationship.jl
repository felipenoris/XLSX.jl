
function Relationship(e::LightXML.XMLElement) :: Relationship
    @assert LightXML.name(e) == "Relationship" "Unexpected XMLElement: $(LightXML.name(e)). Expected: \"Relationship\"."

    return Relationship(
        LightXML.attribute(e, "Id"),
        LightXML.attribute(e, "Type"),
        LightXML.attribute(e, "Target")
    )
end

"""
Parses package level relationships defined in `_rels/.rels`.
Prases workbook level relationships defined in `xl/_rels/workbook.xml.rels`.
"""
function parse_relationships!(xf::XLSXFile)
    xroot = xmlroot(xf, "_rels/.rels")
    @assert LightXML.name(xroot) == "Relationships" "Malformed XLSX file $(xf.filepath). _rels/.rels root node name should be `Relationships`. Found $(LightXML.name(xroot))."
    #@assert LightXML.attribute(xroot, "xmlns") == "http://schemas.openxmlformats.org/package/2006/relationships" "Unsupported schema for Relationships: $(LightXML.attribute(xroot, "xmlns"))."

    for el in xroot["Relationship"]
        push!(xf.relationships, Relationship(el))
    end

    xroot = xmlroot(xf, "xl/_rels/workbook.xml.rels")
    @assert LightXML.name(xroot) == "Relationships" "Malformed XLSX file $(xf.filepath). xl/_rels/workbook.xml.rels root node name should be `Relationships`. Found $(LightXML.name(xroot))."
    #@assert LightXML.attribute(xroot, "xmlns") == "http://schemas.openxmlformats.org/package/2006/relationships" "Unsupported schema for Relationships: $(LightXML.attribute(xroot, "xmlns"))."

    for el in xroot["Relationship"]
        push!(xf.workbook.relationships, Relationship(el))
    end

    nothing
end

function get_relationship_target_by_id(wb::Workbook, Id::String) :: String
    for r in wb.relationships
        if Id == r.Id
            return r.Target
        end
    end
    error("Relationship Id=$(Id) not found")
end

function get_relationship_target_by_type(wb::Workbook, _type_::String) :: String
    for r in wb.relationships
        if _type_ == r.Type
            return r.Target
        end
    end
    error("Relationship Type=$(_type_) not found")
end

function has_relationship_by_type(wb::Workbook, _type_::String) :: Bool
    for r in wb.relationships
        if _type_ == r.Type
            return true
        end
    end
    false
end
