
function XLSXFile(filepath::AbstractString)
    if isfile(filepath)
        return read(filepath)
    else
        error("XLSXFile write not supported yet...")
    end
end

"""
    read(filepath) :: XLSXFile

Main function for reading an Excel file.
"""
function read(filepath::AbstractString) :: XLSXFile
    @assert isfile(filepath) "File $filepath not found."
    xf = XLSXFile(filepath, Dict{String, LightXML.XMLDocument}(), EmptyWorkbook(), Vector{Relationship}())

    xlfile = ZipFile.Reader(filepath)
    try
        for f in xlfile.files

            # parse only XML files
            if !ismatch(r".xml", f.name) && !ismatch(r".rels", f.name)
                #warn("Ignoring non-XML file $(f.name).") # debug
                continue
            end

            xf.data[f.name] = LightXML.parse_string(readstring(f))
        end

        # Check for minimum package requirements
        check_minimum_requirements(xf)

        parse_relationships!(xf)
        parse_workbook!(xf)

        return xf
    catch e
        error("Error parsing $filepath. $e. $(catch_stacktrace()).")
    finally
        close(xlfile)
    end
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

"""
  parse_workbook!(xf::XLSXFile)

Updates xf.workbook from xf.data[\"xl/workbook.xml\"]
"""
function parse_workbook!(xf::XLSXFile)
    xroot = xmlroot(xf, "xl/workbook.xml")
    @assert LightXML.name(xroot) == "workbook" "Malformed xl/workbook.xml. Root node name should be 'workbook'. Got '$(LightXML.name(xroot))'."

    # workbook to be parsed
    workbook = xf.workbook

    # workbookPr
    vec_workbookPr = xroot["workbookPr"]
    if length(vec_workbookPr) > 0
        @assert length(vec_workbookPr) == 1 "Malformed workbook. $xf has more than 1 workbookPr nodes in xl/workbook.xml."

        workbookPr_element = vec_workbookPr[1]
        if LightXML.has_attribute(workbookPr_element, "date1904")
            attribute_value_date1904 = LightXML.attribute(workbookPr_element, "date1904")

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
    end

    # shared string table
    SHARED_STRINGS_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    if has_relationship_by_type(workbook, SHARED_STRINGS_RELATIONSHIP_TYPE)
        sst_root = xmlroot(xf, "xl/" * get_relationship_target_by_type(workbook, SHARED_STRINGS_RELATIONSHIP_TYPE))
        @assert LightXML.name(sst_root) == "sst" "Malformed workbook. sst file should have sst root."
        workbook.sst = sst_root["si"]
    end

    # sheets
    vec_sheets = xroot["sheets"]
    if length(vec_sheets) > 0
        @assert length(vec_sheets) == 1 "Malformed workbook. $xf has more than 1 sheet node in xl/workbook.xml."

        sheets_element = vec_sheets[1]

        vec_sheet = sheets_element["sheet"]
        num_sheets = length(vec_sheet)
        workbook.sheets = Vector{Worksheet}(num_sheets)

        for (index, sheet_element) in enumerate(vec_sheet)
            worksheet = Worksheet(xf, sheet_element)
            workbook.sheets[index] = worksheet
        end
    end

    # styles
    STYLES_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    if has_relationship_by_type(workbook, STYLES_RELATIONSHIP_TYPE)
        styles_target = get_relationship_target_by_type(workbook, STYLES_RELATIONSHIP_TYPE)
        workbook.styles = xmldocument(xf, "xl/" * styles_target)

        # check root node name for styles.xml
        styles_root = LightXML.root(workbook.styles)
        @assert LightXML.name(styles_root) == "styleSheet" "Malformed package. Expected root node named `styleSheet` in `styles.xml`."
    end

    nothing
end
