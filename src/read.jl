
"""
    XLSXFile(filepath)

Creates an empty instance of XLSXFile.
"""
function XLSXFile(filepath::AbstractString)
    @assert isfile(filepath) "File $filepath not found."
    io = ZipFile.Reader(filepath)

    xl = XLSXFile(filepath, io, true, Dict{String, Bool}(), Dict{String, EzXML.Document}(), EmptyWorkbook(), Vector{Relationship}())
    xl.workbook.package = xl
    return xl
end

"""
    readxlsx(filepath) :: XLSXFile

Main function for reading an Excel file.
"""
function readxlsx(filepath::AbstractString) :: XLSXFile

    xf = XLSXFile(filepath)

    try
        for f in xf.io.files

            # parse only XML files
            if !ismatch(r".xml", f.name) && !ismatch(r".rels", f.name)
                #warn("Ignoring non-XML file $(f.name).") # debug
                continue
            end

            # ignore xl/calcChain.xml for now
            if f.name == "xl/calcChain.xml"
                continue
            end

            internal_file_add!(xf, f.name)
            internal_file_read(xf, f.name)
        end

        check_minimum_requirements(xf)
        parse_relationships!(xf)
        parse_workbook!(xf)

    finally
        close(xf)
    end

    return xf
end

"""
    openxlsx(filepath) :: XLSXFile

Open a XLSX file for reading. The user must close this file after using it.
XML data will be fetched from disk as needed.
"""
function openxlsx(filepath::AbstractString) :: XLSXFile

    xf = XLSXFile(filepath)

    for f in xf.io.files

        # parse only XML files
        if !ismatch(r".xml", f.name) && !ismatch(r".rels", f.name)
            #warn("Ignoring non-XML file $(f.name).") # debug
            continue
        end

        internal_file_add!(xf, f.name)
    end

    check_minimum_requirements(xf)
    parse_relationships!(xf)
    parse_workbook!(xf)

    return xf
end

function get_default_namespace(r::EzXML.Node) :: String
    for (prefix, ns) in EzXML.namespaces(r)
        if prefix == ""
            return ns
        end
    end

    error("No default namespace found.")
end

# See section 12.2 - Package Structure
function check_minimum_requirements(xf::XLSXFile)
    mandatory_files = ["_rels/.rels",
                       "xl/workbook.xml",
                       "[Content_Types].xml",
                       "xl/_rels/workbook.xml.rels"
                       ]
 
    for f in mandatory_files
        @assert in(f, filenames(xf)) "Malformed XLSX File. Couldn't find file $f in the package."
    end

    nothing
end

"""
Parses package level relationships defined in `_rels/.rels`.
Prases workbook level relationships defined in `xl/_rels/workbook.xml.rels`.
"""
function parse_relationships!(xf::XLSXFile)
    xroot = xmlroot(xf, "_rels/.rels")
    @assert EzXML.nodename(xroot) == "Relationships" "Malformed XLSX file $(xf.filepath). _rels/.rels root node name should be `Relationships`. Found $(EzXML.nodename(xroot))."
    @assert EzXML.namespaces(xroot) == Pair{String,String}[""=>"http://schemas.openxmlformats.org/package/2006/relationships"]

    for el in EzXML.eachelement(xroot)
        push!(xf.relationships, Relationship(el))
    end
    @assert !isempty(xf.relationships) "Relationships not found in _rels/.rels!"

    xroot = xmlroot(xf, "xl/_rels/workbook.xml.rels")
    @assert EzXML.nodename(xroot) == "Relationships" "Malformed XLSX file $(xf.filepath). xl/_rels/workbook.xml.rels root node name should be `Relationships`. Found $(EzXML.nodename(xroot))."
    @assert EzXML.namespaces(xroot) == Pair{String,String}[""=>"http://schemas.openxmlformats.org/package/2006/relationships"]

    for el in EzXML.eachelement(xroot)
        push!(xf.workbook.relationships, Relationship(el))
    end
    @assert !isempty(xf.workbook.relationships) "Relationships not found in xl/_rels/workbook.xml.rels"

    nothing
end

"""
  parse_workbook!(xf::XLSXFile)

Updates xf.workbook from xf.data[\"xl/workbook.xml\"]
"""
function parse_workbook!(xf::XLSXFile)
    xroot = xmlroot(xf, "xl/workbook.xml")
    @assert EzXML.nodename(xroot) == "workbook" "Malformed xl/workbook.xml. Root node name should be 'workbook'. Got '$(EzXML.nodename(xroot))'."

    # workbook to be parsed
    workbook = xf.workbook

    # workbookPr
    local foundworkbookPr::Bool = false
    for node in EzXML.eachelement(xroot)

        if EzXML.nodename(node) == "workbookPr"
            foundworkbookPr = true

            # read date1904 attribute
            if haskey(node, "date1904")
                attribute_value_date1904 = node["date1904"]

                if attribute_value_date1904 == "1" || attribute_value_date1904 == "true"
                    workbook.date1904 = true
                elseif attribute_value_date1904 == "0" || attribute_value_date1904 == "false"
                    workbook.date1904 = false
                else
                    error("Could not parse xl/workbook -> workbookPr -> date1904 = $(attribute_value_date1904).")
                end
            else
                # does not have attribute => is not date1904
                workbook.date1904 = false
            end

            break
        end
    end
    @assert foundworkbookPr "Malformed: couldn't find workbookPr node element in 'xl/workbook.xml'."

    # sheets
    sheets = Vector{Worksheet}()
    for node in EzXML.eachelement(xroot)
        if EzXML.nodename(node) == "sheets"

            for sheet_node in EzXML.eachelement(node)
                @assert EzXML.nodename(sheet_node) == "sheet" "Unsupported node $(EzXML.nodename(sheet_node)) in 'xl/workbook.xml'."
                worksheet = Worksheet(xf, sheet_node)
                push!(sheets, worksheet)
            end

            break
        end
    end
    workbook.sheets = sheets

    # named ranges
    for node in EzXML.eachelement(xroot)
        if EzXML.nodename(node) == "definedNames"
            for defined_name_node in EzXML.eachelement(node)
                @assert EzXML.nodename(defined_name_node) == "definedName"
                defined_value_string = EzXML.nodecontent(defined_name_node)
                name = defined_name_node["name"]

                local defined_value::DefinedNameValueTypes

                if is_valid_fixed_sheet_cellname(defined_value_string) || is_valid_sheet_cellname(defined_value_string)
                    defined_value = SheetCellRef(defined_value_string)
                elseif is_valid_fixed_sheet_cellrange(defined_value_string) || is_valid_sheet_cellrange(defined_value_string)
                    defined_value = SheetCellRange(defined_value_string)
                elseif ismatch(r"^\".*\"$", defined_value_string) # is enclosed by quotes
                    defined_value = defined_value_string[2:end-1] # remove enclosing quotes
                    if isempty(defined_value)
                        defined_value = Missings.missing
                    end
                elseif tryparse(Int, defined_value_string)
                    defined_value = parse(Int, defined_value_string)
                elseif tryparse(Float64, defined_value_string)
                    defined_value = parse(Float64, defined_value_string)
                elseif isempty(defined_value_string)
                    defined_value = Missings.missing
                else
                    error("Could not parse value $(defined_value_string) for definedName node $name.")
                end

                if haskey(defined_name_node, "localSheetId")
                    # is a Worksheet level name

                    # localSheetId is the 0-based index of the Worksheet in the order
                    # that it is displayed on screen.
                    # Which is the order of the elements under <sheets> element in workbook.xml .
                    localSheetId = parse(Int, defined_name_node["localSheetId"]) + 1
                    sheetId = workbook.sheets[localSheetId].sheetId
                    workbook.worksheet_names[(sheetId, name)] = defined_value
                else
                    # is a Workbook level name
                    workbook.workbook_names[name] = defined_value
                end
            end

            break
        end
    end

    nothing
end

function tryparse(t::Type, s::String)
    try
        parse(t, s)
        return true
    catch
        return false
    end
end

#EzXML.Document

# Lazy loading of XML files

"""
Lists internal files from the XLSX package.
"""
@inline filenames(xl::XLSXFile) = keys(xl.files)

"""
Returns true if the file data was read into xl.data.
"""
internal_file_isread(xl::XLSXFile, filename::String) :: Bool = xl.files[filename]
internal_file_exists(xl::XLSXFile, filename::String) :: Bool = haskey(xl.files, filename)

function internal_file_add!(xl::XLSXFile, filename::String)
    xl.files[filename] = false
    nothing
end

function internal_file_read(xf::XLSXFile, filename::String) :: EzXML.Document
    @assert internal_file_exists(xf, filename) "Couldn't find $filename in $(xf.filepath)."

    if !internal_file_isread(xf, filename)
        @assert isopen(xf) "Can't read from a closed XLSXFile."
        file_not_found = true
        for f in xf.io.files
            if f.name == filename
                xf.files[filename] = true # set file as read
                xf.data[filename] = EzXML.readxml(f)
                file_not_found = false
                break
            end
        end

        if file_not_found
            # shouldn't happen
            error("$filename not found in XLSX package.")
        end
    end

    return xf.data[filename]
end

function Base.close(xl::XLSXFile)
    xl.io_is_open = false
    close(xl.io)
end

Base.isopen(xl::XLSXFile) = xl.io_is_open

"""
    xmldocument(xl::XLSXFile, filename::String) :: EzXML.Document

Utility method to find the XMLDocument associated with a given package filename.
Returns xl.data[filename] if it exists. Throws an error if it doesn't.
"""
@inline xmldocument(xl::XLSXFile, filename::String) :: EzXML.Document = internal_file_read(xl, filename)

"""
    xmlroot(xl::XLSXFile, filename::String) :: EzXML.Node

Utility method to return the root element of a given XMLDocument from the package.
Returns EzXML.root(xl.data[filename]) if it exists.
"""
@inline xmlroot(xl::XLSXFile, filename::String) :: EzXML.Node = EzXML.root(xmldocument(xl, filename))

@inline function xmldocument(ws::Worksheet)
    wb = ws.package.workbook
    target_file = "xl/" * get_relationship_target_by_id(wb, ws.relationship_id)
    return xmldocument(ws.package, target_file)
end

@inline function xmlroot(ws::Worksheet)
    xroot = EzXML.root(xmldocument(ws))
    @assert EzXML.nodename(xroot) == "worksheet" "Malformed sheet $(ws.name)."
    return xroot
end

# get styles document for workbook
function styles_xmlroot(workbook::Workbook)
    # styles
    STYLES_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    if has_relationship_by_type(workbook, STYLES_RELATIONSHIP_TYPE)
        styles_target = get_relationship_target_by_type(workbook, STYLES_RELATIONSHIP_TYPE)
        styles_root = xmlroot(workbook.package, "xl/" * styles_target)

        # check root node name for styles.xml
        @assert get_default_namespace(styles_root) == STYLES_NAMESPACE_XPATH_ARG[1][2] "Unsupported styles XML namespace $(get_default_namespace(styles_root))."
        @assert EzXML.nodename(styles_root) == "styleSheet" "Malformed package. Expected root node named `styleSheet` in `styles.xml`."
        return styles_root
    end

    error("Styles not found for this workbook.")
end
