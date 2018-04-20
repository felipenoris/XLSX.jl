
function Relationship(e::LightXML.XMLElement) :: Relationship
    @assert LightXML.name(e) == "Relationship" "Unexpected XMLElement: $(LightXML.name(e)). Expected: \"Relationship\"."

    return Relationship(
        LightXML.attribute(e, "Id"),
        LightXML.attribute(e, "Type"),
        LightXML.attribute(e, "Target")
    )
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
