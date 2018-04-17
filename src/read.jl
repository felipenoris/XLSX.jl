
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

            try
                xf.data[f.name] = LightXML.parse_string(readstring(f))
            catch ee
                error("Error parsing $(f.name): $ee.")
            end
        end

        # Check for minimum package requirements
        check_minimum_requirements(xf)

        parse_relationships!(xf)
        parse_workbook!(xf)

        return xf
    catch e
        error("Error parsing $filepath: $e.")
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
