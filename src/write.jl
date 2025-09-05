
"""
    savexlsx(f::XLSXFile)

Save an `XLSXFile` instance back to the file from which it was opened (given in `f.source`), 
overwriting original content.

A new `XLSXFile` created with `XLSX.newxlsx` (or using `openxlsx` without specifying a filename) will 
have `source` set to `"blank.xlsx"` and cannot be saved with this function. Use [`writexlsx`](@ref) instead 
to specify a file name for the saved file.

Returns the filepath of the written file if a filename is supplied, or `nothing` if writing to an `IO`.

"""
function savexlsx(f::XLSXFile)
    f.source == "blank.xlsx" && throw(XLSXError("Can't save to a blank `XLSXFile` instance. Use `writexlsx` instead to specify a file name."))
    return writexlsx(f.source, f; overwrite=true)
end


"""
    writexlsx(output_source::Union{AbstractString,IO}, xf::XLSXFile; [overwrite=false])

Write an XLSXFile given by `xf` to the IO or filepath `output_source`.

The source attribute of the `XLSXFile` will be updated to the `output_source` if it is a filepath.

Returns the filepath of the written file if a filename is supplied, or `nothing` if writing to an `IO`.

If `overwrite=true`, `output_source` (when a filepath) will be overwritten if it exists.

See also [`savexlsx`](@ref).
"""
function writexlsx(output_source::Union{AbstractString,IO}, xf::XLSXFile; overwrite::Bool=false)

    !is_writable(xf) && throw(XLSXError("XLSXFile instance is not writable."))
    if !all(values(xf.files))
        throw(XLSXError("Some internal files were not loaded into memory for unknown reasons."))
    end
    if output_source isa AbstractString && !overwrite
        isfile(output_source) && throw(XLSXError("Output file $output_source already exists."))
    end

    update_workbook_xml!(xf)
    
    wb=get_workbook(xf)

    ZipArchives.ZipWriter(output_source) do xlsx

        # write XML files not in cache
        for f in keys(xf.files)
            if !occursin(r"xl/worksheets/sheet\d+\.xml|xl/sharedStrings\.xml", f) # will be generated from cache below
                ZipArchives.zip_newfile(xlsx, f; compress=true)
                write(xlsx, XML.write(xf.data[f]))
            end
        end

        # write worksheet files from cache (cach must be enabled in write mode)
        for sheet_no in 1:sheetcount(wb)
            doc= update_single_sheet!(wb, sheet_no, true)
            f = get_relationship_target_by_id("xl", wb, getsheet(wb, sheet_no).relationship_id)
            ZipArchives.zip_newfile(xlsx, f; compress=true)
            write(xlsx, XML.write(doc))
        end

        # write binary files
        for f in keys(xf.binary_data)
            ZipArchives.zip_newfile(xlsx, f; compress=true)
            write(xlsx, xf.binary_data[f])
        end
        
        # write sharedString table
        if !isempty(get_sst(wb))
            ZipArchives.zip_newfile(xlsx, "xl/sharedStrings.xml"; compress=true)
            print(xlsx, generate_sst_xml_string(wb))
        end        
    end

    if !(output_source isa IO)
        (xf.source = output_source) # update source if output_source is a file path
        return abspath(xf.source)
    else
        return nothing
    end

end

get_worksheet_internal_file(ws::Worksheet) = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
get_worksheet_xml_document(ws::Worksheet) = get_xlsxfile(ws).data[get_worksheet_internal_file(ws)]

function set_worksheet_xml_document!(ws::Worksheet, xdoc::XML.Node)
    XML.nodetype(xdoc) != XML.Document && throw(XLSXError("Expected an XML Document node, got $(XML.nodetype(xdoc))."))
    xf = get_xlsxfile(ws)
    filename = get_worksheet_internal_file(ws)
    !haskey(xf.data, filename) && throw(XLSXError("Internal file not found for $(ws.name)."))
    xf.data[filename] = xdoc
end

function generate_sst_xml_string(wb::Workbook)::String
    sst=wb.sst
    !sst.is_loaded && throw(XLSXError("Can't generate XML string from a Shared String Table that is not loaded."))
    buff = IOBuffer()

    # TODO - Done! : <sst count="89" (TimG: I don't know what this means! UPDATE: Aha! Got it!)
    sst_total = 0
    for sheet in wb.sheets
        sst_total += sheet.sst_count
    end
    print(
        buff,
        """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst count="$sst_total" uniqueCount="$(length(sst))" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
"""
    )

    for s in sst.formatted_strings
        print(buff, s)
    end

    print(buff, "</sst>")
    return String(take!(buff))
end

function add_node_formula!(node, f::Formula)
    f_node = XML.Element("f", XML.Text(XML.escape(f.formula)))
    if !isnothing(f.unhandled)
        for (k, v) in f.unhandled
            f_node[k] = v
        end
    end
    push!(node, f_node)
end

function add_node_formula!(node, f::FormulaReference)
    f_node = XML.Element("f"; t="shared")
    if !isnothing(f.unhandled)
        for (k, v) in f.unhandled
            f_node[k] = v
        end
    end
    f_node["si"] = string(f.id)
    push!(node, f_node)
end

function add_node_formula!(node, f::ReferencedFormula)
    f_node = XML.Element("f", XML.Text(XML.escape(f.formula)); t="shared", ref=f.ref)
    if !isnothing(f.unhandled)
        for (k, v) in f.unhandled
            f_node[k] = v
        end
    end
    f_node["si"] = string(f.id)
    push!(node, f_node)
end

function find_all_nodes(givenpath::String, doc::XML.Node)::Vector{XML.Node}
    XML.nodetype(doc) != XML.Document && throw(XLSXError("Something wrong here!"))
    found_nodes = Vector{XML.Node}()
    for xp in get_node_paths(doc)
        if xp.path == givenpath
            push!(found_nodes, xp.node)
        end
    end
    return found_nodes
end
function get_node_paths(node::XML.Node)
    XML.nodetype(node) != XML.Document && throw(XLSXError("Something wrong here!"))
    default_ns = get_default_namespace(node[end])
    xpaths = Vector{xpath}()
    get_node_paths!(xpaths, node, default_ns, "")
    return xpaths
end

function get_node_paths!(xpaths::Vector{xpath}, node::XML.Node, default_ns, path)
    for c in XML.children(node)
        if XML.nodetype(c) ∉ [XML.Declaration, XML.Comment, XML.Text]
            node_tag = XML.tag(c)
            if !occursin(":", node_tag)
                node_tag = default_ns * ":" * node_tag
            end
            npath = path * "/" * node_tag
            push!(xpaths, xpath(c, npath))
            if length(XML.children(c)) > 0
                get_node_paths!(xpaths, c, default_ns, npath)
            end
        end
    end
    return nothing
end

# Remove all children with tag given by att[2] from a parent XML node with a tag given by att[1].
function unlink(node::XML.Node, att::Tuple{String,String})
    new_node = XML.Element(first(att))
    atts = XML.attributes(node)
    if !isnothing(atts) # Copy attributes across to new node
        for (k, v) in atts
            new_node[k] = v
        end
    end
    for child in XML.children(node) # Copy any child nodes with tags that are not att[2] across to new node
        if XML.tag(child) != last(att)
            push!(new_node, child)
        end
    end
    return new_node
end
function get_idces(doc::XML.LazyNode, t, b)
    i = 1
    j = 1
    n=XML.next(doc)
    while XML.tag(n) != t
        n=XML.next(n)
        if isnothing(n)
            return nothing, nothing
        end
        i += 1
    end
    n=XML.next(n)
    d=n.depth
    while XML.tag(n) != b
        n=XML.next(n)
        if isnothing(n)
            return i, nothing
        end
        if n.depth == d
            j += 1
        end
    end
    return i, j
end
function get_idces(doc::XML.Node, t, b)
    i = 1
    j = 1
    chn=XML.children(doc)
    l=length(chn)
    while XML.tag(chn[i]) != t
        i += 1
        if i > l
            return nothing, nothing
        end
    end
    chn=XML.children(chn[i])
    l=length(chn)
    while XML.tag(chn[j]) != b
        j += 1
        if j > l
            return i, nothing
        end
    end
    return i, j
end

"""
    update_worksheets_xml!(xl::XLSXFile; full=false)

Update worksheet files held in `xf.data` whenever needed. These files had row and cell data 
stripped out when they were read because, when needed (ie in write mode), these data are 
stored in the cache. Worksheet xml files are fully reconstructed to include rows and cells 
from the cache only as a transient precursor to writing an xlsx file and these reconstructed 
worksheet xml files are not stored.
"""
function update_worksheets_xml!(xl::XLSXFile; full=false)
    wb = get_workbook(xl)
    for sheet_no in 1:sheetcount(wb)
        update_single_sheet!(wb, sheet_no, full)
    end
    return nothing
end
function update_single_sheet!(wb::Workbook, sheet_no::Int, full::Bool)::XML.Node
    sheet = getsheet(wb, sheet_no)
    doc = copynode(get_worksheet_xml_document(sheet))
    xroot = doc[end]

    # check namespace and root node name
    get_default_namespace(xroot) != SPREADSHEET_NAMESPACE_XPATH_ARG && throw(XLSXError("Unsupported Spreadsheet XML namespace $(get_default_namespace(xroot))."))
    XML.tag(xroot) != "worksheet" && throw(XLSXError("Malformed Excel file. Expected root node named `worksheet` in worksheet XML file."))

    if full # need to reconstruct row and cell data from cache

        # update sheetData from cache
        i, j = get_idces(doc, "worksheet", "sheetData")
            sheetData_node = XML.Element(XML.tag(doc[i][j]))
            a = XML.attributes(doc[i][j])
            if !isnothing(a)
                for (k, v) in a
                    sheetData_node[k] = v
                end
            end

        # iterates over WorksheetCache cells and writes the XML
        rows = get_cache_rows(sheet)
        sort!(rows, by = x -> x[1])
        for r in rows
            push!(sheetData_node, r[2])
        end
        doc[i][j] = sheetData_node

        # updates worksheet dimension
        i, j = get_idces(doc, "worksheet", "dimension")
        if !isnothing(j)
            dimension_node = doc[i][j]
            dimension_node["ref"] = string(get_dimension(sheet))
            doc[i][j] = dimension_node
        end
    end

    !full && set_worksheet_xml_document!(sheet, doc) # no need to save full reconstructed data

    return doc
end

function stream_cache_rows(sheet::Worksheet)

    Channel{Tuple{CellRange, SheetRow, Dict{String,String}}}(1 << 20) do out
        d = get_dimension(sheet)
        for r in eachrow(sheet)
            rn=row_number(r)
            uh = (!isnothing(sheet.unhandled_attributes) && haskey(sheet.unhandled_attributes, rn)) ? sheet.unhandled_attributes[rn] : Dict{String,String}()
            put!(out, (d, r, uh))
        end
    end
end

function get_cache_rows(sheet::Worksheet)
    read_cache_rows = Channel{Tuple{Int64,XML.Node}}(1 << 24)
    cache_rows = stream_cache_rows(sheet)

    @sync for _ in 1:Threads.nthreads()
        Threads.@spawn begin
            for row in cache_rows
                put!(read_cache_rows, process_cache_row(row))
            end
        end
    end

    close(read_cache_rows) # after all workers finish

    return collect(read_cache_rows)
end

function process_cache_row(cacherow::Tuple{CellRange, XLSX.SheetRow, Dict{String, String}})
    d, r, unhandled_attributes = cacherow
    spans_str = string(column_number(d.start), ":", column_number(d.stop))

    row_nr = row_number(r)
    ordered_column_indexes = sort(collect(keys(r.rowcells)))

    row_node = XML.Element("row"; r=string(row_nr))
    if spans_str != ""
        row_node["spans"] = spans_str
    end
    if !isnothing(r.ht)
        row_node["ht"] = string(r.ht)
        row_node["customHeight"] = "1"
    end

    if !isempty(unhandled_attributes)
        for (attribute, value) in unhandled_attributes
            row_node[attribute] = value
        end
    end
    # add cells to row
    for c in ordered_column_indexes
        cell = getcell(r, c)
        c_element = XML.Element("c"; r=cell.ref.name)

        if cell.datatype != ""
            c_element["t"] = cell.datatype
        end

        if cell.style != ""
            c_element["s"] = cell.style
        end

        if !isempty(cell.formula)
            add_node_formula!(c_element, cell.formula)
        end

        if cell.value != ""
            v_node = XML.Element("v", XML.Text(XML.escape(cell.value)))
            push!(c_element, v_node)
        end
        push!(row_node, c_element)
    end
    return (row_nr, row_node)
end

function abscell(c::CellRef)
    col, row = split_cellname(c.name)
    return "\$" * col * "\$" * string(row)
end

mkabs(c::SheetCellRef) = abscell(c.cellref)
mkabs(c::SheetCellRange) = abscell(c.rng.start) * ":" * abscell(c.rng.stop)
function make_absolute(dn::DefinedNameValue)
    if dn.value isa NonContiguousRange
        v = ""
        for (i, r) in enumerate(dn.value.rng)
            cr = r isa CellRange ? SheetCellRange(dn.value.sheet, r) : SheetCellRef(dn.value.sheet, r) # need to separate and handle separately
            if dn.isabs[i]
                v *= quoteit(cr.sheet) * "!" * mkabs(cr) * ","
            else
                v *= string(cr) * ","
            end
        end
        return v[begin:prevind(v, end)]
    else
        return dn.isabs ? quoteit(dn.value.sheet) * "!" * mkabs(dn.value) : string(dn.value)
    end
end

function update_workbook_xml!(xl::XLSXFile) # Need to update <sheets> and <definedNames>. 
    wb = get_workbook(xl)

    #update defined names
    if length(wb.workbook_names) > 0 || length(wb.worksheet_names) > 0 # skip if no defined names present
        wbdoc = xmlroot(xl, "xl/workbook.xml") # find the <definedNames> block in the workbook's xml file
        i, j = get_idces(wbdoc, "workbook", "definedNames")
        if isnothing(j)
            # there is no <definedNames> block in the workbook's xml file, so we'll need to create one
            # The <definedNames> block goes after the <sheets> block. Need to move everything down one to make room.    
            m, n = get_idces(wbdoc, "workbook", "sheets")
            nchildren = length(XML.children(wbdoc[m]))
            push!(wbdoc[m], wbdoc[m][end])
            for c in nchildren-1:-1:n+1
                wbdoc[m][c+1] = wbdoc[m][c]
            end
            definedNames = XML.Element("definedNames")
            j = n + 1
        else
            definedNames = unlink(wbdoc[i][j], ("definedNames", "definedName")) # Remove old defined names
        end
        for (k, v) in wb.workbook_names
            if typeof(v.value) <: DefinedNameRangeTypes
                v = make_absolute(v)
            else
                v = string(v.value)
            end
            dn_node = XML.Element("definedName", name=k, XML.Text(v))
            push!(definedNames, dn_node)
        end
        for (k, v) in wb.worksheet_names
            if typeof(v.value) <: DefinedNameRangeTypes
                v = make_absolute(v)
            else
                v = string(v.value)
            end
            dn_node = XML.Element("definedName", name=last(k), localSheetId=first(k) - 1, XML.Text(v))
            push!(definedNames, dn_node)
        end
        wbdoc[i][j] = definedNames # Add the new definedNames block to the workbook's xml file
    end

    #update sheets
    doc = xmlroot(xl, "xl/workbook.xml")
    i, j = get_idces(doc, "workbook", "sheets")
    unlink(doc[i][j], ("sheets", "sheet"))
    sheets_element = XML.Element("sheets")
    for s in wb.sheets
        sheet_element = XML.Element("sheet"; name=XML.escape(s.name))
        sheet_element["sheetId"] = s.sheetId
        sheet_element["r:id"] = s.relationship_id
        push!(sheets_element, sheet_element)
    end
    doc[i][j] = sheets_element

    return nothing
end

function add_cell_to_worksheet_dimension!(ws::Worksheet, cell::Cell)
    # update worksheet dimension
    ws_dimension = get_dimension(ws)

    top = row_number(ws_dimension.start)
    left = column_number(ws_dimension.start)

    bottom = row_number(ws_dimension.stop)
    right = column_number(ws_dimension.stop)

    r = row_number(cell)
    c = column_number(cell)

    if r < top || c < left
        top = min(r, top)
        left = min(c, left)
        set_dimension!(ws, CellRange(top, left, bottom, right))
    elseif r > bottom || c > right
        bottom = max(r, bottom)
        right = max(c, right)
        set_dimension!(ws, CellRange(top, left, bottom, right))
    end

    nothing
end

function setdata!(ws::Worksheet, cell::Cell)
    !is_writable(get_xlsxfile(ws)) && throw(XLSXError("XLSXFile instance is not writable. Open Excel file with `mode=\"rw\"` instead"))
    ws.cache === nothing && throw(XLSXError("Can't write data to a Worksheet with empty cache."))
    cache = ws.cache

    r = row_number(cell)
    c = column_number(cell)

    if !haskey(cache.cells, r)
        push!(cache.rows_in_cache, r)
        cache.cells[r] = Dict{Int,Cell}()
        cache.row_ht[r] = nothing
        cache.dirty = true
    end
    cache.cells[r][c] = cell
    add_cell_to_worksheet_dimension!(ws, cell)

    nothing
end

# This set of characters works in my use case. I don't know:
# - if the set is sufficient, or if other charachers may be needed in other use cases
# - if all of these characters are necessary or if one or two coulld be dropped
# - What the optimum replacement character should be.
const ILLEGAL_CHARS = [
    Char(0x00) => "",
    Char(0x01) => "",
    Char(0x02) => "",
    Char(0x03) => "",
    Char(0x04) => "",
    Char(0x05) => "",
    Char(0x06) => "",
    Char(0x07) => "",
    Char(0x08) => "",
    Char(0x12) => "&apos;",
    Char(0x16) => ""
]
function strip_illegal_chars(x::String) # No-longer needed. Issue arose in XML.jl (https://github.com/JuliaComputing/XML.jl/issues/48)
    result = x
    for (pat, r) in ILLEGAL_CHARS
        result = replace(result, pat => r)
    end
    return result
end

# Returns the datatype and value for `val` to be inserted into `ws`.
function xlsx_encode(ws::Worksheet, val::AbstractString)
    if isempty(val)
        return ("", "")
    end
#    sst_ind = add_shared_string!(get_workbook(ws), strip_illegal_chars(val))
    sst_ind = add_shared_string!(get_workbook(ws), val)
    ws.sst_count+=1

    return ("s", string(sst_ind))
end

xlsx_encode(::Worksheet, val::Missing) = ("", "")
xlsx_encode(::Worksheet, val::Bool) = ("b", val ? "1" : "0")
xlsx_encode(::Worksheet, val::Union{Int,Float64}) = ("", string(val))
xlsx_encode(ws::Worksheet, val::Dates.Date) = ("", string(date_to_excel_value(val, isdate1904(get_xlsxfile(ws)))))
xlsx_encode(ws::Worksheet, val::Dates.DateTime) = ("", string(datetime_to_excel_value(val, isdate1904(get_xlsxfile(ws)))))
xlsx_encode(::Worksheet, val::Dates.Time) = ("", string(time_to_excel_value(val)))

Base.setindex!(ws::Worksheet, v, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}) = setdata!(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))), v)
Base.setindex!(ws::Worksheet, v::AbstractVector, r::Union{Integer,UnitRange{<:Integer}}, c::UnitRange{T}) where {T<:Integer} = setdata!(ws, r, c, v)
Base.setindex!(ws::Worksheet, v::AbstractVector, r::UnitRange{T}, c::Union{Integer,UnitRange{<:Integer}}) where {T<:Integer} = setdata!(ws, r, c, v)
Base.setindex!(ws::Worksheet, v::AbstractVector, ref; dim::Integer=2) = setdata!(ws, ref, v, dim)
Base.setindex!(ws::Worksheet, v::AbstractVector, r, c; dim::Integer=2) = setdata!(ws, r, c, v, dim)
Base.setindex!(ws::Worksheet, v, ref) = setdata!(ws, ref, v)
Base.setindex!(ws::Worksheet, v, r, c) = setdata!(ws, r, c, v)
function Base.setindex!(ws::Worksheet, v, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}})
    for a in row, b in col
        setdata!(ws, CellRef(a, b), v)
    end
end
function Base.setindex!(ws::Worksheet, v, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}})
    for a in row, b in col
        setdata!(ws, CellRef(a, b), v)
    end
end
function Base.setindex!(ws::Worksheet, v, row::Union{Vector{Int},StepRange{<:Integer}}, col::Union{Vector{Int},StepRange{<:Integer}})
    for a in row, b in col
        setdata!(ws, CellRef(a, b), v)
    end
end

function setdata!(ws::Worksheet, ref::CellRef, val::CellFormula)
    v = ""
    t = ""
    cell = Cell(ref, t, id(val.styleid), v, Formula(val.value.formula))
    setdata!(ws, cell)
end
function setdata!(ws::Worksheet, ref::CellRef, val::CellValue)
    t, v = xlsx_encode(ws, val.value)
    cell = Cell(ref, t, id(val.styleid), v, Formula())
    setdata!(ws, cell)
end

# convert AbstractTypes to concrete
setdata!(ws::Worksheet, ref::CellRef, val::AbstractString) = setdata!(ws, ref, CellValue(ws, convert(String, val)))
setdata!(ws::Worksheet, ref::CellRef, val::Integer) = setdata!(ws, ref, convert(Int, val))
setdata!(ws::Worksheet, ref::CellRef, val::Real) = setdata!(ws, ref, convert(Float64, val))

# convert nothing to missing when writing
setdata!(ws::Worksheet, ref::CellRef, ::Nothing) = setdata!(ws, ref, CellValue(ws, missing))

setdata!(ws::Worksheet, row::Integer, col::Integer, val::CellValue) = setdata!(ws, CellRef(row, col), val)

setdata!(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, col::Union{Integer,UnitRange{<:Integer}}, v) = setdata!(ws, CellRange(CellRef(first(row), first(col)), CellRef(last(row), last(col))), v)

# shift the relative cell references ina formula when shifting a ReferencedFormula
function shift_excel_references(formula::String, offset::Tuple{Int64,Int64})
    # Regex to match Excel-style cell references (e.g., A1, $A$1, A$1, $A1)
    pattern = r"\$?[A-Z]{1,3}\$?[1-9][0-9]*"
    row_shift, col_shift = offset

    initial = [string(x.match) for x in eachmatch(pattern, formula)]
    result = Vector{String}()

    for ref in eachmatch(pattern, formula)
        # Extract parts using regex
        m = match(r"(\$?)([A-Z]{1,3})(\$?)([1-9][0-9]*)", ref.match)
        col_abs, col_letters, row_abs, row_digits = m.captures

        col_num = decode_column_number(col_letters)
        row_num = parse(Int, row_digits)

        # Apply shifts only if not absolute
        new_col = col_abs == "\$" ? col_letters : encode_column_number(col_num + col_shift)
        new_row = row_abs == "\$" ? row_digits : string(row_num + row_shift)

        push!(result, col_abs * new_col * row_abs * new_row)
    end

    pairs = Dict(zip(initial, result))
    if !isempty(pairs)
        for (from, to) in pairs
            formula = replace(formula, from => to)
        end
    end
    return formula
end

function setdata!(ws::Worksheet, ref::CellRef, val::Union{AbstractFormula,CellValueType}) # use existing cell format if it exists
    c = getcell(ws, ref)
    if !(c isa EmptyCell) && c.formula isa ReferencedFormula
        rereference_formulae(ws, c)
    end
    if c isa EmptyCell || c.style == ""
        if val isa AbstractFormula
            return setdata!(ws, ref, CellFormula(ws, val))
        else
            return setdata!(ws, ref, CellValue(ws, val))
        end
    else
        existing_style = CellDataFormat(parse(Int, c.style))
        isa_dt = styles_is_datetime(ws.package.workbook, existing_style)
        if val isa Dates.Date
            if isa_dt == false
                c.style = string(update_template_xf(ws, existing_style, ["numFmtId", "applyNumberFormat"], [string(DEFAULT_DATE_numFmtId), "1"]).id)
            end
        elseif val isa Dates.Time
            if isa_dt == false
                c.style = string(update_template_xf(ws, existing_style, ["numFmtId", "applyNumberFormat"], [string(DEFAULT_TIME_numFmtId), "1"]).id)
            end
        elseif val isa Dates.DateTime
            if isa_dt == false
                c.style = string(update_template_xf(ws, existing_style, ["numFmtId", "applyNumberFormat"], [string(DEFAULT_DATETIME_numFmtId), "1"]).id)
            end
        elseif val isa Float64 || val isa Int
            if styles_is_float(ws.package.workbook, existing_style) == false && Int(existing_style.id) ∉ [0, 1]
                c.style = string(update_template_xf(ws, existing_style, ["numFmtId"], [string(DEFAULT_NUMBER_numFmtId)]).id)
            end
        elseif val isa Bool # Now rerouted here rather than assigning an EmptyCellDataFormat.
            # Change any style to General (0) and retain other formatting.
            c.style = string(update_template_xf(ws, existing_style, ["numFmtId"], [string(DEFAULT_BOOL_numFmtId)]).id)
        end
        if val isa AbstractFormula
            return setdata!(ws, ref, CellFormula(val, CellDataFormat(parse(Int, c.style))))
        else
            return setdata!(ws, ref, CellValue(val, CellDataFormat(parse(Int, c.style))))
        end
    end
end
function setdata!(ws::Worksheet, ref::AbstractString, value)
    if value isa String
        i = findfirst(!isspace, value)
        if !isnothing(i) && length(value[i:end]) > 1 && value[i] == '=' # it's a formula!
            value = Formula(last(split(value, '=')))
        end
    end
    if is_worksheet_defined_name(ws, ref)
        v = get_defined_name_value(ws, ref)
        if is_defined_name_value_a_reference(v)
            return setdata!(ws, v, value)
        else
            throw(XLSXError("`$ref` is not a valid cell or range reference."))
        end
    elseif is_workbook_defined_name(get_workbook(ws), ref)
        wb = get_workbook(ws)
        v = get_defined_name_value(wb, ref)
        if is_defined_name_value_a_reference(v)
            return setdata!(ws, v, value)
        else
            throw(XLSXError("`$ref` is not a valid cell or range reference."))
        end
    elseif is_valid_cellname(ref)
        return setdata!(ws, CellRef(ref), value)
    elseif is_valid_sheet_cellname(ref)
        return setdata!(ws, SheetCellRef(ref), value)
    elseif is_valid_cellrange(ref)
        return setdata!(ws, CellRange(ref), value)
    elseif is_valid_column_range(ref)
        return setdata!(ws, ColumnRange(ref), value)
    elseif is_valid_row_range(ref)
        return setdata!(ws, RowRange(ref), value)
    elseif is_valid_non_contiguous_range(ref)
        return setdata!(ws, NonContiguousRange(ws, ref), value)
    elseif is_valid_sheet_cellrange(ref)
        return setdata!(ws, SheetCellRange(ref), value)
    elseif is_valid_sheet_column_range(ref)
        return setdata!(ws, SheetColumnRange(ref), value)
    elseif is_valid_sheet_row_range(ref)
        return setdata!(ws, SheetRowRange(ref), value)
    elseif is_valid_non_contiguous_cellrange(ref)
        return setdata!(ws, NonContiguousRange(ws, ref), value)
    elseif is_valid_non_contiguous_sheetcellrange(ref)
        nc = NonContiguousRange(ref)
        return do_sheet_names_match(ws, nc) && setdata!(ws, nc, value)
    end
    throw(XLSXError("`$ref` is not a valid cell or range reference."))
end
function setdata!(ws::Worksheet, rng::CellRange, value)
    for row in rng.start.row_number:rng.stop.row_number
        for col in rng.start.column_number:rng.stop.column_number
            setdata!(ws, row, col, value)
        end
    end
end
function setdata!(ws::Worksheet, rng::RowRange, value)
    dim = get_dimension(ws)
    start = CellRef(rng.start, dim.start.column_number,)
    stop = CellRef(rng.stop, dim.stop.column_number)
    setdata!(ws, CellRange(start, stop), value)
end
function setdata!(ws::Worksheet, rng::ColumnRange, value)
    dim = get_dimension(ws)
    start = CellRef(dim.start.row_number, rng.start)
    stop = CellRef(dim.stop.row_number, rng.stop)
    setdata!(ws, CellRange(start, stop), value)
end
function setdata!(ws::Worksheet, rng::NonContiguousRange, value)
    for r in rng.rng
        if r isa CellRef
            setdata!(ws, r, value)
        else
            for cell in r
                setdata!(ws, cell, value)
            end
        end
    end
end
setdata!(ws::Worksheet, ::Colon, ::Colon, v) = setdata!(ws::Worksheet, :, v)
function setdata!(ws::Worksheet, ::Colon, v)
    dim = get_dimension(ws)
    setdata!(ws, dim, v)
end
function setdata!(ws::Worksheet, row::Union{Integer,UnitRange{<:Integer}}, ::Colon, v)
    dim = get_dimension(ws)
    setdata!(ws, CellRange(CellRef(first(row), dim.start.column_number), CellRef(last(row), dim.stop.column_number)), v)
end
function setdata!(ws::Worksheet, ::Colon, col::Union{Integer,UnitRange{<:Integer}}, v)
    dim = get_dimension(ws)
    setdata!(ws, CellRange(CellRef(dim.start.row_number, first(col)), CellRef(dim.stop.row_number, last(col))), v)
end
function setdata!(ws::Worksheet, row::Union{Vector{Int},StepRange{<:Integer}}, ::Colon, v)
    dim = get_dimension(ws)
    for a in row
        for b in dim.start.column_number:dim.stop.column_number
            setdata!(ws, CellRef(a, b), v)
        end
    end
end
function setdata!(ws::Worksheet, ::Colon, col::Union{Vector{Int},StepRange{<:Integer}}, v)
    dim = get_dimension(ws)
    for b in col
        for a in dim.start.row_number:dim.stop.row_number
            setdata!(ws, CellRef(a, b), v)
        end
    end
end
setdata!(ws::Worksheet, ref::SheetCellRef, value) = do_sheet_names_match(ws, ref) && setdata!(ws, ref.cellref, value)
setdata!(ws::Worksheet, rng::SheetCellRange, value) = do_sheet_names_match(ws, rng) && setdata!(ws, rng.rng, value)
setdata!(ws::Worksheet, rng::SheetColumnRange, value) = do_sheet_names_match(ws, rng) && setdata!(ws, rng.colrng, value)
setdata!(ws::Worksheet, rng::SheetRowRange, value) = do_sheet_names_match(ws, rng) && setdata!(ws, rng.rowrng, value)
setdata!(ws::Worksheet, ref_str::AbstractString, value::Vector, dim::Integer) = setdata!(ws, CellRef(ref_str), value, dim)
setdata!(ws::Worksheet, row::Integer, col::Integer, data) = setdata!(ws, CellRef(row, col), data)
setdata!(ws::Worksheet, ref::CellRef, value) = throw(XLSXError("Unsupported datatype $(typeof(value)) for writing data to Excel file. Supported data types are $(CellValueType) or $(CellValue)."))
setdata!(ws::Worksheet, row::Integer, col::Integer, data::AbstractVector, dim::Integer) = setdata!(ws, CellRef(row, col), data, dim)

function setdata!(sheet::Worksheet, ref::CellRef, data::AbstractVector, dim::Integer)
    for (i, val) in enumerate(data)
        target_cell_ref = target_cell_ref_from_offset(ref, i - 1, dim)
        setdata!(sheet, target_cell_ref, val)
    end
end

Base.setindex!(ws::Worksheet, v::AbstractVector, r::Union{Colon,UnitRange{T}}, c) where {T<:Integer} = setdata!(ws, r, c, v)
Base.setindex!(ws::Worksheet, v::AbstractVector, r, c::Union{Colon,UnitRange{T}}) where {T<:Integer} = setdata!(ws, r, c, v)
setdata!(sheet::Worksheet, ::Colon, col::Integer, data::AbstractVector) = setdata!(sheet, 1, col, data, 1)
setdata!(sheet::Worksheet, row::Integer, ::Colon, data::AbstractVector) = setdata!(sheet, row, 1, data, 2)

function setdata!(sheet::Worksheet, row::Integer, cols::UnitRange{T}, data::AbstractVector) where {T<:Integer}
    length(data) != length(cols) && throw(XLSXError("Column count mismatch between `data` ($(length(data)) columns) and column range $cols ($(length(cols)) columns)."))
    anchor_cell_ref = CellRef(row, first(cols))

    # since cols is the unit range, this is a column-based operation
    setdata!(sheet, anchor_cell_ref, data, 2)
end

function setdata!(sheet::Worksheet, rows::UnitRange{T}, col::Integer, data::AbstractVector) where {T<:Integer}
    length(data) != length(rows) && throw(XLSXError("Row count mismatch between `data` ($(length(data)) rows) and row range $rows ($(length(rows)) rows)."))
    anchor_cell_ref = CellRef(first(rows), col)

    # since rows is the unit range, this is a row-based operation
    setdata!(sheet, anchor_cell_ref, data, 1)
end

function setdata!(sheet::Worksheet, ref_or_rng::AbstractString, matrix::Array{T,2}) where {T}
    if is_valid_cellrange(ref_or_rng)
        setdata!(sheet, CellRange(ref_or_rng), matrix)
    elseif is_valid_cellname(ref_or_rng)
        setdata!(sheet, CellRef(ref_or_rng), matrix)
    else
        throw(XLSXError("Invalid cell reference or range: $ref_or_rng"))
    end
end

function setdata!(sheet::Worksheet, ref::CellRef, matrix::Array{T,2}) where {T}
    rows, cols = size(matrix)
    anchor_row = row_number(ref)
    anchor_col = column_number(ref)

    @inbounds for c in 1:cols, r in 1:rows
        setdata!(sheet, anchor_row + r - 1, anchor_col + c - 1, matrix[r, c])
    end
end

function setdata!(sheet::Worksheet, rng::CellRange, matrix::Array{T,2}) where {T}
    size(rng) != size(matrix) && throw(XLSXError("Target range $rng size ($(size(rng))) must be equal to the input matrix size ($(size(matrix)))"))
    setdata!(sheet, rng.start, matrix)
end

# Given an anchor cell at (anchor_row, anchor_col).
# Returns a CellRef at:
# - (anchor_row + offset, anchol_col) if dim = 1 (operates on rows)
# - (anchor_row, anchor_col + offset) if dim = 2 (operates on cols)
function target_cell_ref_from_offset(anchor_row::Integer, anchor_col::Integer, offset::Integer, dim::Integer)::CellRef
    if dim == 1
        return CellRef(anchor_row + offset, anchor_col)
    elseif dim == 2
        return CellRef(anchor_row, anchor_col + offset)
    else
        throw(XLSXError("Invalid dimension: $dim."))
    end
end

function target_cell_ref_from_offset(anchor_cell::CellRef, offset::Integer, dim::Integer)::CellRef
    return target_cell_ref_from_offset(row_number(anchor_cell), column_number(anchor_cell), offset, dim)
end

const ALLOWED_TYPES = Union{Number,String,Bool,Dates.Date,Dates.Time,Dates.DateTime,Missing,Nothing}
function process_vector(col) # Convert any disallowed types to strings. #239.
    if eltype(col) <: ALLOWED_TYPES
        # Case 1: All elements are of allowed types
        return col
    elseif eltype(col) <: Any && all(x -> !(typeof(x) <: ALLOWED_TYPES), col)
        # Case 2: All elements are of disallowed types
        return map(x -> string(x), col)
    else
        # Case 3: Mixed types, process each element
        return [typeof(x) <: ALLOWED_TYPES ? x : string(x) for x in col]
    end
end

"""
    writetable!(
        sheet::Worksheet,
        data,
        columnnames;
        anchor_cell::CellRef=CellRef("A1"),
        write_columnnames::Bool=true,
    )

Write tabular data `data` with labels given by `columnnames` to `sheet`,
starting at `anchor_cell`.

`data` must be a vector of columns.
`columnnames` must be a vector of column labels.

Column labels that are not of type `String` will be converted 
to strings before writing. Any data columns that are not of 
type `String`, `Float64`, `Int64`, `Bool`, `Date`, `Time`, 
`DateTime`, `Missing`, or `Nothing` will be converted to strings 
before writing.


See also: [`XLSX.writetable`](@ref).
"""
function writetable!(
    sheet::Worksheet,
    data,
    columnnames;
    anchor_cell::CellRef=CellRef("A1"),
    write_columnnames::Bool=true,
)

    # read dimensions
    col_count = length(data)
    col_count != length(columnnames) && throw(XLSXError("Column count mismatch between `data` ($col_count columns) and `columnnames` ($(length(columnnames)) columns)."))
    col_count <= 0 && throw(XLSXError("Can't write table with no columns."))
    col_count > EXCEL_MAX_COLS && throw(XLSXError("`data` contains $col_count columns, but Excel only supports up to $EXCEL_MAX_COLS; must reduce `data` size"))
    row_count = length(data[1])
    row_count > EXCEL_MAX_ROWS - 1 && throw(XLSXError("`data` contains $row_count rows, but Excel only supports up to $(EXCEL_MAX_ROWS-1); must reduce `data` size"))
    if col_count > 1
        for c in 2:col_count
            length(data[c]) != row_count && throw(XLSXError("Row count mismatch between column 1 ($row_count rows) and column $c ($(length(data[c])) rows)."))
        end
    end

    anchor_row = row_number(anchor_cell)
    anchor_col = column_number(anchor_cell)
    start_from_anchor = 1

    # write table header
    if write_columnnames
        for c in 1:col_count
            target_cell_ref = CellRef(anchor_row, c + anchor_col - 1)
#            sheet[target_cell_ref] = strip_illegal_chars(XML.escape(string(columnnames[c])))
#            sheet[target_cell_ref] = XML.escape(string(columnnames[c]))
            sheet[target_cell_ref] = string(columnnames[c])
        end
        start_from_anchor = 0
    end

    # write table data
    data = [process_vector(col) for col in data] # Address issue #239
    for c in 1:col_count
        for r in 1:row_count
            target_cell_ref = CellRef(r + anchor_row - start_from_anchor, c + anchor_col - 1)
            v = data[c][r]
#            sheet[target_cell_ref] = v isa String ? strip_illegal_chars(XML.escape(v)) : v
#            sheet[target_cell_ref] = v isa String ? XML.escape(v) : v
            sheet[target_cell_ref] = v 
        end
    end
end

"""
    rename!(ws::Worksheet, name::AbstractString)

Rename a `Worksheet` to `name`.

"""
function rename!(ws::Worksheet, name::AbstractString)

    # no-op if the name has not changed
    if ws.name == name
        return
    end

    xf = get_xlsxfile(ws)
    !is_writable(xf) && throw(XLSXError("XLSXFile instance is not writable."))
    name ∈ sheetnames(xf) && throw(XLSXError("Sheetname $name is already in use."))

    # updates XML
    xroot = xmlroot(xf, "xl/workbook.xml")[end]
    for node in XML.children(xroot)
        if XML.tag(node) == "sheets"

            for sheet_node in XML.children(node)
                if sheet_node["name"] == ws.name
                    # assign new name
                    sheet_node["name"] = name
                    break
                end
            end

            break
        end
    end

    # updates the new name in the worksheet instance
    ws.name = name
    nothing
end

"""
    addsheet!(wb::Workbook, [name::AbstractString=""]) --> ::Worksheet
    addsheet!(xf::XLSXFile, [name::AbstractString=""]) --> ::Worksheet

Create a new worksheet named `name`.
If `name` is not provided, a unique name is created.

See also [copysheet!](@ref), [deletesheet!](@ref)

"""
addsheet!(xl::XLSXFile, name::AbstractString="")::Worksheet = addsheet!(get_workbook(xl), name)::Worksheet
function addsheet!(wb::Workbook, name::AbstractString=""; relocatable_data_path::String=_relocatable_data_path())::Worksheet
    file_sheet_template = joinpath(relocatable_data_path, "sheet_template.xml")
    !isfile(file_sheet_template) && throw(XLSXError("Couldn't find template file $file_sheet_template."))
    bytes = read(file_sheet_template)
    f, _ = skipNode(XML.Raw(bytes), "sheetData")
    xdoc = XML.Node(XML.Raw(f))

    new_cache = XLSX.WorksheetCache(
        true,
        Dict{Int64,Dict{Int64,XLSX.Cell}}(),
        Int64[],
        Dict{Int,Union{Float64,Nothing}}(),
        Dict{Int64,Int64}(),
        SheetRowStreamIterator(get_xlsxfile(wb)[1]), # Dummy - not needed because using full cache.
        nothing,
        false
    )
    new_ws = insertsheet!(wb, xdoc, new_cache, 0, name)
    return new_ws
end

"""
    copysheet!(ws::Worksheet, [name::AbstractString=""]) --> ::Worksheet

Create a copy of the worksheet `ws` and add it to the end of the workbook with the 
specified worksheet name.
Return the new worksheet object.
If `name` is not provided, a new name is generated by appending " (copy)" to the original 
worksheet name, with a further numerical suffix to guarantee uniqueness if necessary.
To copy worksheets, the `XLSXFile` must be writable (opened with `mode="rw"` or as a template).
See also [`XLSX.openxlsx`](@ref) and [XLSX.opentemplate](@ref).

!!! warning "Experimental"
    This function is experimental is not guaranteed to work with all XLSX files, 
    especially those with complex features. However, cell formats, conditional formats 
    and worksheet defined names should all copy OK. Please report any issues.

See also [addsheet!](@ref), [deletesheet!](@ref)

# Examples
```julia
julia> f=XLSX.openxlsx("general.xlsx", mode="rw")
XLSXFile("C:\\...\\general.xlsx") containing 13 Worksheets
            sheetname size          range
-------------------------------------------------
              general 10x6          A1:F10
               table3 5x6           A2:F6
               table4 4x3           E12:G15
                table 12x8          A2:H13
               table2 5x3           A1:C5
                empty 1x1           A1:A1
               table5 6x1           C3:C8
               table6 8x2           B1:C8
               table7 7x2           B2:C8
               lookup 4x9           B2:J5
         header_error 3x4           B2:E4
       named_ranges_2 4x5           A1:E4
         named_ranges 14x6          A2:F15

julia> XLSX.copysheet!(f[4])
12×8 XLSX.Worksheet: ["table (copy)"](A2:H13)

julia> f
XLSXFile("C:\\...\\general.xlsx") containing 14 Worksheets
            sheetname size          range
-------------------------------------------------
              general 10x6          A1:F10
               table3 5x6           A2:F6
               table4 4x3           E12:G15
                table 12x8          A2:H13
               table2 5x3           A1:C5
                empty 1x1           A1:A1
               table5 6x1           C3:C8
               table6 8x2           B1:C8
               table7 7x2           B2:C8
               lookup 4x9           B2:J5
         header_error 3x4           B2:E4
       named_ranges_2 4x5           A1:E4
         named_ranges 14x6          A2:F15
         table (copy) 12x8          A2:H13

```

"""
function copysheet!(ws::Worksheet, name::AbstractString="")::Worksheet
    wb = get_workbook(ws)
    xl = get_xlsxfile(ws)
    !is_writable(get_xlsxfile(ws)) && throw(XLSXError("XLSXFile instance is not writable."))
    dim = get_dimension(ws)

    # make sure cache and XML are consistent
    update_worksheets_xml!(xl)

    # create a copy of the XML document
    xdoc = copynode(get_worksheet_xml_document(ws))

    # if copied sheet is the currently selected sheet, do not copy this attribute over.
    # The original sheet will remain the only selected sheet.
    for c in XML.children(xdoc[end])
        if c.tag=="sheetViews"
            for c2 in XML.children(c)
                if c2.tag=="sheetView"
                    atts=XML.attributes(c2)
                    if haskey(atts, "tabSelected")
                        atts["tabSelected"]="0"
                    end
                end
            end
        end
    end

    # generate a new name for the copied sheet
    name = name == "" ? ws.name * " (copy)" : name

    # cache of copied sheet must be full
    ws.cache.is_full == true || throw(XLSXError("Cannot copy worksheet that does not have a full cache"))

    # copy the original worksheet cache to the new worksheet
    new_cache = WorksheetCache(
        true,
        ws.cache.cells,
        ws.cache.rows_in_cache,
        ws.cache.row_ht,
        ws.cache.row_index,
        SheetRowStreamIterator(ws), # Dummy - not needed because using full cache.
        nothing,
        ws.cache.dirty,
    )

    # insert the copied sheet into the workbook
    new_ws = insertsheet!(wb, xdoc, new_cache, ws.sst_count, name; dim)

    # copy defined names from the original worksheet to the new worksheet
    ws_keys = [x for x in keys(wb.worksheet_names) if first(x) == ws.sheetId]
    for k in ws_keys
        val = wb.worksheet_names[k].value
        val = val isa CellRange ? string(val) :
              val isa SheetCellRange ? new_ws.name * "!" * string(val.rng) :
              val
        addDefinedName(new_ws, last(k), val; absolute=wb.worksheet_names[k].isabs)
    end

    return new_ws
end

function insertsheet!(wb::Workbook, xdoc::XML.Node, new_cache::WorksheetCache, sst_count::Int, name::AbstractString=""; dim=CellRange("A1:A1"))::Worksheet
    xf = get_xlsxfile(wb)
    !is_writable(xf) && throw(XLSXError("XLSXFile instance is not writable."))

    if name == ""
        new_name = ""
    else
        new_name = name
    end
    # ensure name is unique.
    i = 1
    current_sheet_names = sheetnames(wb)
    while new_name ∈ current_sheet_names || new_name == ""
        new_name = (name == "" ? "Sheet" : name * " ") * string(i)
        i += 1
    end

    new_name == "" && throw(XLSXError("Something wrong here!"))

    # checks if name is a unique sheet name
    function check_valid_sheetname(n::AbstractString)
        max_length = 31
        if length(n) > max_length
            throw(XLSXError("Invalid sheetname $n: must have at most $max_length characters. Found $(length(n))"))
        end

        if occursin(r"[:\\/\?\*\[\]]+", n)
            throw(XLSXError("Sheetname cannot contain characters: ':', '\\', '/', '?', '*', '[' or ']'."))
        end
    end

    check_valid_sheetname(new_name)

    # generate sheetId
    current_sheet_ids = [ws.sheetId for ws in wb.sheets]
    sheetId = max(current_sheet_ids...) + 1

    # generate a unique ID for the new sheet
    xdoc[2]["xr:uid"] = "{" * string(UUIDs.uuid4()) * "}"

    # generate a unique name for the XML
    local xml_filename::String
    i = 1
    while true
        xml_filename = "xl/worksheets/sheet" * string(i) * ".xml"
        #        if !in(xml_filename, keys(xf.files))
        if !haskey(xf.files, xml_filename)
            break
        end
        i += 1
    end

    # adds doc do XLSXFile
    xf.files[xml_filename] = true # is read
    xf.data[xml_filename] = xdoc

    # adds workbook-level relationship
    # <Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
    rId = add_relationship!(wb, xml_filename[4:end], "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")

    # creates Worksheet instance
    ws = Worksheet(xf, sheetId, rId, new_name, dim, false)
    ws.cache = new_cache
    ws.sst_count = sst_count

    # adds the new sheet to the list of sheets in the workbook
    push!(wb.sheets, ws)

    # update [Content_Types].xml (fix for issue #275)
    ctype_root = xmlroot(get_xlsxfile(wb), "[Content_Types].xml")[end]
    XML.tag(ctype_root) != "Types" && throw(XLSXError("Something wrong here!"))
    override_node = XML.Element("Override";
        PartName="/xl/worksheets/sheet" * string(sheetId) * ".xml",
        ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    )
    push!(ctype_root, override_node)

    update_workbook_xml!(xf)

    return ws
end

function renumber_files!(xf::XLSXFile, rId::String)
    wb = get_workbook(xf)
    id = parse(Int64, rId[4:end])
    #= UNNECESSARY!
        for s in wb.sheets
            r=parse(Int64, s.relationship_id[4:end])
            if r > id
                newId=string(r-1)
                oldId=string(r)
    #            s.relationship_id = "rId"*newId
                new_filename = "xl/worksheets/sheet" * newId * ".xml"
                old_filename = "xl/worksheets/sheet" * oldId * ".xml"
                if haskey(xf.files, old_filename)
                    xf.files[new_filename]=xf.files[old_filename]
                    delete!(xf.files, old_filename)
                end
                if haskey(xf.data, old_filename)
                    xf.data[new_filename]=xf.data[old_filename]
                    delete!(xf.data, old_filename)
                end
            end
        end
    =#
    # update active tab
    wbdoc = xmlroot(xf, "xl/workbook.xml")
    i, j = get_idces(wbdoc, "workbook", "bookViews")
    w = XML.children(wbdoc[i][j])
    if length(w) > 0
        for c in w
            if XML.tag(c) == "workbookView"
                a = XML.attributes(c)
                if haskey(a, "activeTab")
                    at = parse(Int64, a["activeTab"])
                    if at >= id
                        c["activeTab"] = string(at - 1)
                    end
                end
            end
        end
    end
end


"""
    deletesheet!(ws::Worksheet) -> ::XLSXFile
    deletesheet!(wb::Workbook, name::AbstractString) -> ::XLSXFile
    deletesheet!(xf::XLSXFile, name::AbstractString) -> ::XLSXFile
    deletesheet!(xf::XLSXFile, sheetId::Integer) -> ::XLSXFile

Delete the given worksheet, the worksheet with the given name or the worksheet with the given `sheetId` from its `XLSXFile` 
(`sheetId` is a 1-based integer representing the order in which worksheet tabs are displayed in Excel).

# note "Caution"
    Cells in the other sheets that have references to the deleted sheet will fail when the sheet is deleted.
    The formulae are updated to contain a `#Ref!` error in place of each sheetcell reference.
    

See also [addsheet!](@ref), [copysheet!](@ref)

# Examples

```julia
julia> f = XLSX.opentemplate("general.xlsx")
XLSXFile("C:\\...\\general.xlsx") containing 13 Worksheets
            sheetname size          range
-------------------------------------------------
              general 10x6          A1:F10
               table3 5x6           A2:F6
               table4 4x3           E12:G15
                table 12x8          A2:H13
               table2 5x3           A1:C5
                empty 1x1           A1:A1        
               table5 6x1           C3:C8
               table6 8x2           B1:C8
               table7 7x2           B2:C8
               lookup 4x9           B2:J5
         header_error 3x4           B2:E4
       named_ranges_2 4x5           A1:E4
         named_ranges 14x6          A2:F15


julia> XLSX.deletesheet!(f[4])
XLSXFile("C:\\...\\general.xlsx") containing 12 Worksheets
            sheetname size          range
-------------------------------------------------
              general 10x6          A1:F10
               table3 5x6           A2:F6
               table4 4x3           E12:G15
               table2 5x3           A1:C5
                empty 1x1           A1:A1
               table5 6x1           C3:C8
               table6 8x2           B1:C8
               table7 7x2           B2:C8
               lookup 4x9           B2:J5
         header_error 3x4           B2:E4
       named_ranges_2 4x5           A1:E4
         named_ranges 14x6          A2:F15


julia> XLSX.deletesheet!(f, "table5")
XLSXFile("C:\\...\\general.xlsx") containing 11 Worksheets
            sheetname size          range
-------------------------------------------------
              general 10x6          A1:F10
               table3 5x6           A2:F6
               table4 4x3           E12:G15
               table2 5x3           A1:C5
                empty 1x1           A1:A1
               table6 8x2           B1:C8
               table7 7x2           B2:C8
               lookup 4x9           B2:J5
         header_error 3x4           B2:E4
       named_ranges_2 4x5           A1:E4
         named_ranges 14x6          A2:F15


julia> XLSX.deletesheet!(f, 1)
XLSXFile("C:\\...\\general.xlsx") containing 10 Worksheets
            sheetname size          range
-------------------------------------------------
               table3 5x6           A2:F6
               table4 4x3           E12:G15
               table2 5x3           A1:C5
                empty 1x1           A1:A1
               table6 8x2           B1:C8
               table7 7x2           B2:C8
               lookup 4x9           B2:J5
         header_error 3x4           B2:E4
       named_ranges_2 4x5           A1:E4
         named_ranges 14x6          A2:F15
```

"""
deletesheet!(ws::Worksheet) = deletesheet!(get_workbook(ws), ws.name)
deletesheet!(xl::XLSXFile, sheetId::Integer) = deletesheet!(get_workbook(xl), xl[sheetId].name)
deletesheet!(xl::XLSXFile, name::AbstractString) = deletesheet!(get_workbook(xl), name)
function deletesheet!(wb::Workbook, name::AbstractString)::XLSXFile
    hassheet(wb, name) || throw(XLSXError("Worksheet `$name` not found in workbook."))
    sheetcount(wb) > 1 || throw(XLSXError("`$name` is this workbook's only sheet. Cannot delete the only sheet!"))

    xf = get_xlsxfile(wb)

    # Worksheets and relationships
    s = (findfirst(s -> s.name == name, wb.sheets))
    sId = wb.sheets[s].sheetId
    rId = wb.sheets[s].relationship_id
    r = findfirst(y -> occursin("worksheet", y.Type) && y.Id == rId, wb.relationships)
    delete_relationships!(xf, wb.relationships[r])
    deleteat!(wb.relationships, r)
    deleteat!(wb.sheets, s)

    # Defined Names
    found_wbnames = Vector{String}()
    for (k, v) in wb.workbook_names
        wbn = v.value
        if typeof(wbn) <: DefinedNameRangeTypes
            if wbn.sheet == name
                push!(found_wbnames, k)
            end
        end
    end
    found_wsnames = Vector{Tuple{Int64,String}}()
    for (k, v) in wb.worksheet_names
        if first(k) == sId
            push!(found_wsnames, k)
        end
    end
    for key in found_wbnames
        delete!(wb.workbook_names, key)
    end
    for key in found_wsnames
        delete!(wb.worksheet_names, key)
    end
    renumber_keys = Vector{Pair{Tuple{Int64,String},Tuple{Int64,String}}}()
    for (k, _) in wb.worksheet_names
        first(k) == sId && throw(XLSXError("Something wrong here!"))
        if first(k) > sId
            push!(renumber_keys, k => (sId, last(k)))
        end
    end
    for (oldkey, newkey) in renumber_keys
        wb.worksheet_names[newkey] = wb.worksheet_names[oldkey]
        delete!(wb.worksheet_names, oldkey)
    end

    # Files
    xml_filename = "xl/worksheets/sheet" * rId[4:end] * ".xml"
    if in(xml_filename, keys(xf.files))
        delete!(xf.files, xml_filename)
    end
    if in(xml_filename, keys(xf.data))
        delete!(xf.data, xml_filename)
    end
    if in(xml_filename, keys(xf.binary_data))
        delete!(xf.binary_data, xml_filename)
    end

    # update [Content_Types].xml
    ctype_root = xmlroot(get_xlsxfile(wb), "[Content_Types].xml")[end]
    XML.tag(ctype_root) != "Types" && throw(XLSXError("Something wrong here!"))
    cont = XML.children(ctype_root)
    let idx = 0
        for (i, c) in enumerate(cont)
            if haskey(c, "PartName") && c["PartName"] == "/xl/worksheets/sheet" * rId[4:end] * ".xml"
                idx = i
                break
            end
        end
        if idx > 0
            deleteat!(cont, idx)
        end
    end

    update_formulas_missing_sheet!(wb, name)
    renumber_files!(xf, rId)
    update_workbook_xml!(xf)

    return xf
end

#
# Helper Functions
#

"""
    writetable(filename, data, columnnames; [overwrite], [sheetname])

- `data` is a vector of columns.
- `columnames` is a vector of column labels.
- `overwrite` is a `Bool` to control if `filename` should be overwritten if already exists.
- `sheetname` is the name for the worksheet.

Returns the filepath of the written file if a filename is supplied, or `nothing` if writing to an `IO`.

# Example

```julia
import XLSX
columns = [ [1, 2, 3, 4], ["Hey", "You", "Out", "There"], [10.2, 20.3, 30.4, 40.5] ]
colnames = [ "integers", "strings", "floats" ]
XLSX.writetable("table.xlsx", columns, colnames)
```

See also: [`XLSX.writetable!`](@ref).
"""
function writetable(filename::Union{AbstractString,IO}, data, columnnames; overwrite::Bool=false, sheetname::AbstractString="", anchor_cell::Union{String,CellRef}=CellRef("A1"))

    if filename isa AbstractString && !overwrite
        isfile(filename) && throw(XLSXError("$filename already exists."))
    end

    xf = open_empty_template(sheetname)
    sheet = xf[1]

    if isa(anchor_cell, String)
        anchor_cell = CellRef(anchor_cell)
    end

    writetable!(sheet, data, columnnames; anchor_cell=anchor_cell)

    # write output file
    writexlsx(filename, xf, overwrite=overwrite)
end

"""
    writetable(filename::Union{AbstractString, IO}; overwrite::Bool=false, kw...)
    writetable(filename::Union{AbstractString, IO}, tables::Vector{Tuple{String, Vector{Any}, Vector{String}}}; overwrite::Bool=false)

Write multiple tables.

`kw` is a variable keyword argument list. Each element should be in this format: `sheetname=( data, column_names )`,
where `data` is a vector of columns and `column_names` is a vector of column labels.

Returns the filepath of the written file if a filename is supplied, or `nothing` if writing to an `IO`.

Example:

```julia
julia> import DataFrames, XLSX

julia> df1 = DataFrames.DataFrame(COL1=[10,20,30], COL2=["Fist", "Sec", "Third"])

julia> df2 = DataFrames.DataFrame(AA=["aa", "bb"], AB=[10.1, 10.2])

julia> XLSX.writetable("report.xlsx", "REPORT_A" => df1, "REPORT_B" => df2)
```
"""
function writetable(filename::Union{AbstractString,IO}; overwrite::Bool=false, kw...)

    if filename isa AbstractString && !overwrite
        isfile(filename) && throw(XLSXError("$filename already exists."))
    end

    xf = open_empty_template()
    is_first = true

    for (sheetname, (data, column_names)) in kw
        if is_first
            # first sheet already exists in template file
            sheet = xf[1]
            rename!(sheet, string(sheetname))
            writetable!(sheet, data, column_names)

            is_first = false
        else
            sheet = addsheet!(xf, string(sheetname))
            writetable!(sheet, data, column_names)
        end
    end

    # write output file
    writexlsx(filename, xf, overwrite=overwrite)

end

function writetable(filename::Union{AbstractString,IO}, tables::Vector{Tuple{String,S,Vector{T}}}; overwrite::Bool=false) where {S<:Vector{U} where {U},T<:Union{String,Symbol}}

    if filename isa AbstractString && !overwrite
        isfile(filename) && throw(XLSXError("$filename already exists."))
    end

    xf = open_empty_template()

    is_first = true
    for (sheetname, data, column_names) in tables
        if is_first
            # first sheet already exists in template file
            sheet = xf[1]
            rename!(sheet, string(sheetname))
            writetable!(sheet, data, column_names)

            is_first = false
        else
            sheet = addsheet!(xf, string(sheetname))
            writetable!(sheet, data, column_names)
        end
    end

    # write output file
    writexlsx(filename, xf, overwrite=overwrite)
end
