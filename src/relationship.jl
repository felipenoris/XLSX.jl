
function Relationship(e::EzXML.Node) :: Relationship
    @assert EzXML.nodename(e) == "Relationship" "Unexpected XMLElement: $(EzXML.nodename(e)). Expected: \"Relationship\"."

    return Relationship(
        e["Id"],
        e["Type"],
        e["Target"]
    )
end

function parse_relationship_target(prefix::String, target::String) :: String
    @assert !isempty(prefix) && !isempty(target)

    if target[1] == '/'
        @assert sizeof(target) > 1 "Incomplete target path $target."
        return target[2:end]
    else
        return prefix * '/' * target
    end
end

function get_relationship_target_by_id(prefix::String, wb::Workbook, Id::String) :: String
    for r in wb.relationships
        if Id == r.Id
            return parse_relationship_target(prefix, r.Target)
        end
    end
    error("Relationship Id=$(Id) not found")
end

function get_relationship_target_by_type(prefix::String, wb::Workbook, _type_::String) :: String
    for r in wb.relationships
        if _type_ == r.Type
            return parse_relationship_target(prefix, r.Target)
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

function get_package_relationship_root(xf::XLSXFile) :: EzXML.Node
    xroot = xmlroot(xf, "_rels/.rels")
    @assert EzXML.nodename(xroot) == "Relationships" "Malformed XLSX file $(xf.source). _rels/.rels root node name should be `Relationships`. Found $(EzXML.nodename(xroot))."
    @assert (""=>"http://schemas.openxmlformats.org/package/2006/relationships") ∈ EzXML.namespaces(xroot) "Unexpected namespace at workbook relationship file: `$(EzXML.namespaces(xroot))`."
    return xroot
end

function get_workbook_relationship_root(xf::XLSXFile) :: EzXML.Node
    xroot = xmlroot(xf, "xl/_rels/workbook.xml.rels")
    @assert EzXML.nodename(xroot) == "Relationships" "Malformed XLSX file $(xf.source). xl/_rels/workbook.xml.rels root node name should be `Relationships`. Found $(EzXML.nodename(xroot))."
    @assert (""=>"http://schemas.openxmlformats.org/package/2006/relationships") ∈ EzXML.namespaces(xroot) "Unexpected namespace at workbook relationship file: `$(EzXML.namespaces(xroot))`."
    return xroot
end

# Adds new relationship. Returns new generated rId.
function add_relationship!(wb::Workbook, target::String, _type::String) :: String
    xf = get_xlsxfile(wb)
    @assert is_writable(xf) "XLSXFile instance is not writable."
    local rId :: String

    let
        got_unique_id = false
        id = 1

        while !got_unique_id
            got_unique_id = true
            rId = string("rId", id)
            for r in wb.relationships
                if r.Id == rId
                    got_unique_id = false
                    id += 1
                    break
                end
            end
        end
    end

    # adds to relationship vector
    new_relationship = Relationship(rId, _type, target)
    push!(wb.relationships, new_relationship)

    # adds to XML tree
    xroot = get_workbook_relationship_root(xf)
    el = EzXML.addelement!(xroot, "Relationship")
    el["Id"] = rId
    el["Target"] = target
    el["Type"] = _type

    return rId
end
