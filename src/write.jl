
#=
    open_xlsx_template(source::Union{AbstractString, IO}) :: XLSXFile

Open an Excel file as template for editing and saving to another file with `XLSX.writexlsx`.
=#
@inline open_xlsx_template(source::Union{AbstractString, IO}) :: XLSXFile = open_or_read_xlsx(source, true, true, true)

function _relocatable_data_path(; path::AbstractString=Artifacts.artifact"XLSX_relocatable_data")
    return path
end

#=
    open_empty_template(
        sheetname::AbstractString="";
        relocatable_data_path::AbstractString=_relocatable_data_path()
    ) :: XLSXFile

Returns an empty, writable `XLSXFile` with 1 worksheet.

# Arguments

* `sheetname` is the name of the worksheet. When provided with an empty string `""`, this routine selects the first sheet of the workbook.

* `relocatable_data_path` is the filepath for a blank workbook template. It defaults to the template provided by the package artifact.
=#
function open_empty_template(
            sheetname::AbstractString="";
            relocatable_data_path::AbstractString=_relocatable_data_path()
        ) :: XLSXFile

    empty_excel_template = joinpath(relocatable_data_path, "blank.xlsx")
    @assert isfile(empty_excel_template) "Couldn't find template file $empty_excel_template."
    xf = open_xlsx_template(empty_excel_template)

    if sheetname != ""
        rename!(xf[1], sheetname)
    end

    return xf
end

"""
    writexlsx(output_source, xlsx_file; [overwrite=false])

Writes an Excel file given by `xlsx_file::XLSXFile` to IO or filepath `output_source`.

If `overwrite=true`, `output_source` (when a filepath) will be overwritten if it exists.
"""
function writexlsx(output_source::Union{AbstractString, IO}, xf::XLSXFile; overwrite::Bool=false)

    @assert is_writable(xf) "XLSXFile instance is not writable."
    @assert all(values(xf.files)) "Some internal files were not loaded into memory. Did you use `XLSX.open_xlsx_template` to open this file?"
    if output_source isa AbstractString && !overwrite
        @assert !isfile(output_source) "Output file $output_source already exists."
    end

   update_worksheets_xml!(xf)

    ZipArchives.ZipWriter(output_source) do xlsx
        # write XML files
        for f in keys(xf.files)

            if f == "xl/sharedStrings.xml"
                # sst will be generated below
                continue
            end
            ZipArchives.zip_newfile(xlsx, f; compress=true)
            write(xlsx, replace(XML.write(xf.data[f]), r" *\n *" => "")) # Don't want pretty printing.
        end

        # write binary files
        for f in keys(xf.binary_data)
            ZipArchives.zip_newfile(xlsx, f; compress=true)
            write(xlsx, xf.binary_data[f])
        end

        if !isempty(get_sst(xf))
            ZipArchives.zip_newfile(xlsx, "xl/sharedStrings.xml"; compress=true)
            print(xlsx, generate_sst_xml_string(get_sst(xf)))
        end
    end
    nothing
end

get_worksheet_internal_file(ws::Worksheet) = get_relationship_target_by_id("xl", get_workbook(ws), ws.relationship_id)
get_worksheet_xml_document(ws::Worksheet) = get_xlsxfile(ws).data[ get_worksheet_internal_file(ws) ]

function set_worksheet_xml_document!(ws::Worksheet, xdoc::XML.Node)
    @assert XML.nodetype(xdoc) == XML.Document "Expected an XML Document node, got $(XML.nodetype(xdoc))."
    xf = get_xlsxfile(ws)
    filename = get_worksheet_internal_file(ws)
    @assert haskey(xf.data, filename) "Internal file not found for $(ws.name)."
    xf.data[filename] = xdoc
    
end

function generate_sst_xml_string(sst::SharedStringTable) :: String
    @assert sst.is_loaded "Can't generate XML string from a Shared String Table that is not loaded."
    buff = IOBuffer()

    # TODO: <sst count="89"
    print(buff, """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst uniqueCount="$(length(sst))" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
""")

    for s in sst.formatted_strings
        print(buff, s)
    end
 
    print(buff, "</sst>")
    return String(take!(buff))
end

function add_node_formula!(node, f::Formula)
    f_node = XML.Element("f", Text(f.formula))
    push!(node, f_node)
end

function add_node_formula!(node, f::FormulaReference)
    f_node = XML.Element("f"; t = "shared", si = string(f.id))
    push!(node, f_node)
end

function add_node_formula!(node, f::ReferencedFormula)
    f_node = XML.Element("f", Text(f.formula); t = "shared", si = string(f.id), ref = f.ref)
    push!(node, f_node)
end

function find_all_nodes(givenpath::String, doc::XML.Node)::Vector{XML.Node}
    @assert XML.nodetype(doc) == XML.Document
    found_nodes = Vector{XML.Node}()
    for xp in get_node_paths(doc)
        if xp.path == givenpath
            push!(found_nodes, xp.node)
        end
    end
    return found_nodes
end
function get_node_paths(node::XML.Node)
    @assert XML.nodetype(node) == XML.Document
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
            if length(XML.children(c))>0
                get_node_paths!(xpaths, c, default_ns, npath)
            end
        end
    end 
    return nothing
end
function unlink_rows(node::XML.Node) # removes all rows from a sheetData XML node.
    new_worksheet = XML.Element("sheetData")
    a = XML.attributes(node)
    if !isnothing(a) # Copy attributes across to new node
        for (k, v) in XML.attributes(node)
            new_worksheet[k] = v
        end
    end
    for child in XML.children(node) # Copy any child nodes that are not rows across to new node
        if XML.tag(child) != "row"
            push!(new_worksheet, child)
        end
    end
    return new_worksheet
end
function get_idces(doc, t, b)
    i=1
    j=1
    while XML.tag(doc[i]) != t
        i+=1
    end
    while XML.tag(doc[i][j]) != b
        j+=1
    end
    return i, j
end

function update_worksheets_xml!(xl::XLSXFile)
    wb = get_workbook(xl)

    for i in 1:sheetcount(wb)
        sheet = getsheet(wb, i)
        doc = get_worksheet_xml_document(sheet)
        xroot = doc[end]

        # check namespace and root node name
        @assert get_default_namespace(xroot) == SPREADSHEET_NAMESPACE_XPATH_ARG "Unsupported Spreadsheet XML namespace $(get_default_namespace(xroot))."
        @assert XML.tag(xroot) == "worksheet" "Malformed Excel file. Expected root node named `worksheet` in worksheet XML file."

        # Since we do not at the moment track changes, we need to delete all data and re-write it, but this could entail losses.
        # |- Column formatting is preserved in the <cols> subtree.
        # |- Row formatting would get lost if we simply remove the children of <sheetData> storing the rows
        unhandled_attributes = Dict{Int, Dict{String,String}}() # from row number to attribute and its value

        # The following attributes will be overwritten by us and need not be preserved
        handled_attributes = Set{String}([
            "r",     # the row number
            "spans", # the columns the row spans
        ])

        let
            child_nodes = find_all_nodes("/$SPREADSHEET_NAMESPACE_XPATH_ARG:worksheet/$SPREADSHEET_NAMESPACE_XPATH_ARG:sheetData/$SPREADSHEET_NAMESPACE_XPATH_ARG:row", doc)

            i, j = get_idces(doc, "worksheet", "sheetData")
            parent = doc[i][j]

            for c in child_nodes # all elements under sheetData should be <row> elements

                if XML.tag(c) == "row"

                    attributes = XML.attributes(c)
                    if !isnothing(attributes)
                        unhandled_attributes_ = filter(attribute -> !in(first(attribute), handled_attributes), attributes)
                        if length(unhandled_attributes_)>0
                            row_nr = parse(Int, c["r"])
                            unhandled_attributes[row_nr] = unhandled_attributes_
                        end
                    end
                else
                    @warn("Unexpected node under sheetData: $(XML.tag(c))")
                end            
            end

            doc[i][j] = unlink_rows(parent)
        end

        # updates sheetData
        i, j = get_idces(doc, "worksheet", "sheetData")
        sheetData_node = doc[i][j]
        if isnothing(sheetData_node.children)
            a = XML.attributes(sheetData_node)
            sheetData_node = XML.Element(XML.tag(sheetData_node))
            if !isnothing(a)
                for (k, v) in a
                    sheetData_node[k] = v
                end
            end
        end

        local spans_str::String = ""

        # Every row has the `spans=1:<n_cols>` property. Set it to the whole range of columns by default
        if !isnothing(get_dimension(sheet))
            spans_str = string(column_number(get_dimension(sheet).start), ":", column_number(get_dimension(sheet).stop))
        end

        # iterates over WorksheetCache cells and write the XML
        for r in eachrow(sheet)
            row_nr = row_number(r)
            ordered_column_indexes = sort(collect(keys(r.rowcells)))

            row_node = XML.Element("row"; r = string(row_nr))
            if spans_str != ""
                row_node["spans"] = spans_str
            end

            if haskey(unhandled_attributes, row_nr)
                for (attribute, value) in unhandled_attributes[row_nr]
                    row_node[attribute] = value
                end
            end

            # add cells to row
            for c in ordered_column_indexes
                cell = getcell(r, c)
                c_element = XML.Element("c"; r = cell.ref.name)

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
                    v_node = XML.Element("v", Text(cell.value))
                    push!(c_element, v_node)
                end
                push!(row_node, c_element)
            end

            push!(sheetData_node, row_node)
        end
        doc[i][j]=sheetData_node

        # updates worksheet dimension
        if get_dimension(sheet) !== nothing
            i, j = get_idces(doc, "worksheet", "dimension")
            dimension_node = doc[i][j]
            dimension_node["ref"] = string(get_dimension(sheet))
            doc[i][j] = dimension_node
        end

        set_worksheet_xml_document!(sheet, doc)
    end

    nothing
end

function add_cell_to_worksheet_dimension!(ws::Worksheet, cell::Cell)
    # update worksheet dimension
    ws_dimension = get_dimension(ws)

    if ws_dimension === nothing
        set_dimension!(ws, CellRange(cell.ref, cell.ref))
        return
    end

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
    @assert is_writable(get_xlsxfile(ws)) "XLSXFile instance is not writable."
    @assert ws.cache !== nothing "Can't write data to a Worksheet with empty cache."
    cache = ws.cache

    r = row_number(cell)
    c = column_number(cell)

    if !haskey(cache.cells, r)
        push!(cache.rows_in_cache, r)
        cache.cells[r] = Dict{Int, Cell}()
        cache.dirty = true
    end
    cache.cells[r][c] = cell
    add_cell_to_worksheet_dimension!(ws, cell)

    nothing
end

# Returns the datatype and value for `val` to be inserted into `ws`.
function xlsx_encode(ws::Worksheet, val::AbstractString)
    if isempty(val)
        return ("", "")
    end
    sst_ind = add_shared_string!(get_workbook(ws), XML.escape(val))
    return ("s", string(sst_ind))
end

xlsx_encode(::Worksheet, val::Missing) = ("", "")
xlsx_encode(::Worksheet, val::Bool) = ("b", val ? "1" : "0")
xlsx_encode(::Worksheet, val::Union{Int, Float64}) = ("", string(val))
xlsx_encode(ws::Worksheet, val::Dates.Date) = ("", string(date_to_excel_value(val, isdate1904(get_xlsxfile(ws)))))
xlsx_encode(ws::Worksheet, val::Dates.DateTime) = ("", string(datetime_to_excel_value(val, isdate1904(get_xlsxfile(ws)))))
xlsx_encode(::Worksheet, val::Dates.Time) = ("", string(time_to_excel_value(val)))

function setdata!(ws::Worksheet, ref::CellRef, val::CellValue)
    t, v = xlsx_encode(ws, val.value)
    cell = Cell(ref, t, id(val.styleid), v, Formula(""))
    setdata!(ws, cell)
end

# convert AbstractTypes to concrete
setdata!(ws::Worksheet, ref::CellRef, val::AbstractString) = setdata!(ws, ref, CellValue(ws, convert(String, val)))
setdata!(ws::Worksheet, ref::CellRef, val::Bool) = setdata!(ws, ref, CellValue(ws, val))
setdata!(ws::Worksheet, ref::CellRef, val::Integer) = setdata!(ws, ref, CellValue(ws, convert(Int, val)))
setdata!(ws::Worksheet, ref::CellRef, val::Real) = setdata!(ws, ref, CellValue(ws, convert(Float64, val)))

# convert nothing to missing when writing
setdata!(ws::Worksheet, ref::CellRef, ::Nothing) = setdata!(ws, ref, CellValue(ws, missing))

setdata!(ws::Worksheet, row::Integer, col::Integer, val::CellValue) = setdata!(ws, CellRef(row, col), val)

Base.setindex!(ws::Worksheet, v, ref) = setdata!(ws, ref, v)
Base.setindex!(ws::Worksheet, v, r, c) = setdata!(ws, r, c, v)

Base.setindex!(ws::Worksheet, v::AbstractVector, ref; dim::Integer=2) = setdata!(ws, ref, v, dim)
Base.setindex!(ws::Worksheet, v::AbstractVector, r, c; dim::Integer=2) = setdata!(ws, r, c, v, dim)

setdata!(ws::Worksheet, ref::CellRef, val::CellValueType) = setdata!(ws, ref, CellValue(ws, val))
setdata!(ws::Worksheet, ref_str::AbstractString, value) = setdata!(ws, CellRef(ref_str), value)
setdata!(ws::Worksheet, ref_str::AbstractString, value::Vector, dim::Integer) = setdata!(ws, CellRef(ref_str), value, dim)
setdata!(ws::Worksheet, row::Integer, col::Integer, data) = setdata!(ws, CellRef(row, col), data)
setdata!(ws::Worksheet, ref::CellRef, value) = error("Unsupported datatype $(typeof(value)) for writing data to Excel file. Supported data types are $(CellValueType) or $(CellValue).")
setdata!(ws::Worksheet, row::Integer, col::Integer, data::AbstractVector, dim::Integer) = setdata!(ws, CellRef(row, col), data, dim)

function setdata!(sheet::Worksheet, ref::CellRef, data::AbstractVector, dim::Integer)
    for (i, val) in enumerate(data)
        target_cell_ref = target_cell_ref_from_offset(ref, i-1, dim)
        setdata!(sheet, target_cell_ref, val)
    end
end

Base.setindex!(ws::Worksheet, v::AbstractVector, r::Union{Colon, UnitRange{T}}, c) where {T<:Integer} = setdata!(ws, r, c, v)
Base.setindex!(ws::Worksheet, v::AbstractVector, r, c::Union{Colon, UnitRange{T}}) where {T<:Integer} = setdata!(ws, r, c, v)
setdata!(sheet::Worksheet, ::Colon, col::Integer, data::AbstractVector) = setdata!(sheet, 1, col, data, 1)
setdata!(sheet::Worksheet, row::Integer, ::Colon, data::AbstractVector) = setdata!(sheet, row, 1, data, 2)

function setdata!(sheet::Worksheet, row::Integer, cols::UnitRange{T}, data::AbstractVector) where {T<:Integer}
    @assert length(data) == length(cols) "Column count mismatch between `data` ($(length(data)) columns) and column range $cols ($(length(cols)) columns)."
    anchor_cell_ref = CellRef(row, first(cols))

    # since cols is the unit range, this is a column-based operation
    setdata!(sheet, anchor_cell_ref, data, 2)
end

function setdata!(sheet::Worksheet, rows::UnitRange{T}, col::Integer, data::AbstractVector) where {T<:Integer}
    @assert length(data) == length(rows) "Row count mismatch between `data` ($(length(data)) rows) and row range $rows ($(length(rows)) rows)."
    anchor_cell_ref = CellRef(first(rows), col)

    # since rows is the unit range, this is a row-based operation
    setdata!(sheet, anchor_cell_ref, data, 1)
end

function setdata!(sheet::Worksheet, ref_or_rng::AbstractString, matrix::Array{T, 2}) where {T}
    if is_valid_cellrange(ref_or_rng)
        setdata!(sheet, CellRange(ref_or_rng), matrix)
    elseif is_valid_cellname(ref_or_rng)
        setdata!(sheet, CellRef(ref_or_rng), matrix)
    else
        error("Invalid cell reference or range: $ref_or_rng")
    end
end

function setdata!(sheet::Worksheet, ref::CellRef, matrix::Array{T, 2}) where {T}
    rows, cols = size(matrix)
    anchor_row = row_number(ref)
    anchor_col = column_number(ref)

    @inbounds for c in 1:cols, r in 1:rows
        setdata!(sheet, anchor_row + r - 1, anchor_col + c - 1, matrix[r, c])
    end
end

function setdata!(sheet::Worksheet, rng::CellRange, matrix::Array{T, 2}) where {T}
    @assert size(rng) == size(matrix) "Target range $rng size ($(size(rng))) must be equal to the input matrix size ($(size(matrix))) "
    setdata!(sheet, rng.start, matrix)
end

# Given an anchor cell at (anchor_row, anchor_col).
# Returns a CellRef at:
# - (anchor_row + offset, anchol_col) if dim = 1 (operates on rows)
# - (anchor_row, anchor_col + offset) if dim = 2 (operates on cols)
function target_cell_ref_from_offset(anchor_row::Integer, anchor_col::Integer, offset::Integer, dim::Integer) :: CellRef
    if dim == 1
        return CellRef(anchor_row + offset, anchor_col)
    elseif dim == 2
        return CellRef(anchor_row, anchor_col + offset)
    else
        error("Invalid dimension: $dim.")
    end
end

function target_cell_ref_from_offset(anchor_cell::CellRef, offset::Integer, dim::Integer) :: CellRef
    return target_cell_ref_from_offset(row_number(anchor_cell), column_number(anchor_cell), offset, dim)
end

"""
    writetable!(
        sheet::Worksheet,
        data,
        columnnames;
        anchor_cell::CellRef=CellRef("A1"),
        write_columnnames::Bool=true,
    )

Writes tabular data `data` with labels given by `columnnames` to `sheet`,
starting at `anchor_cell`.

`data` must be a vector of columns.
`columnnames` must be a vector of column labels.

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
    @assert col_count == length(columnnames) "Column count mismatch between `data` ($col_count columns) and `columnnames` ($(length(columnnames)) columns)."
    @assert col_count > 0 "Can't write table with no columns."
    @assert col_count <= EXCEL_MAX_COLS "`data` contains $col_count columns, but Excel only supports up to $EXCEL_MAX_COLS; must reduce `data` size"
    row_count = length(data[1])
    @assert row_count <= EXCEL_MAX_ROWS-1 "`data` contains $row_count rows, but Excel only supports up to $(EXCEL_MAX_ROWS-1); must reduce `data` size"
    if col_count > 1
        for c in 2:col_count
            @assert length(data[c]) == row_count "Row count mismatch between column 1 ($row_count rows) and column $c ($(length(data[c])) rows)."
        end
    end

    anchor_row = row_number(anchor_cell)
    anchor_col = column_number(anchor_cell)
    start_from_anchor = 1
    # write table header
    if write_columnnames
        for c in 1:col_count
            target_cell_ref = CellRef(anchor_row, c + anchor_col - 1)
            sheet[target_cell_ref] = XML.escape(string(columnnames[c]))
        end
        start_from_anchor = 0
    end

    # write table data
    for r in 1:row_count, c in 1:col_count
        target_cell_ref = CellRef(r + anchor_row - start_from_anchor, c + anchor_col - 1)
        sheet[target_cell_ref] = data[c][r] isa String ? XML.escape(data[c][r]) : data[c][r]
    end
end

"""
    rename!(ws::Worksheet, name::AbstractString)

Renames a `Worksheet`.
"""
function rename!(ws::Worksheet, name::AbstractString)

    # no-op if the name has not changed
    if ws.name == name
        return
    end

    xf = get_xlsxfile(ws)
    @assert is_writable(xf) "XLSXFile instance is not writable."
    @assert name ∉ sheetnames(xf) "Sheetname $name is already in use."

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

addsheet!(xl::XLSXFile, name::AbstractString="") :: Worksheet = addsheet!(get_workbook(xl), name)

"""
    addsheet!(workbook, [name]) :: Worksheet

Create a new worksheet with named `name`.
If `name` is not provided, a unique name is created.
"""
function addsheet!(wb::Workbook, name::AbstractString=""; relocatable_data_path::String = _relocatable_data_path()) :: Worksheet
    xf = get_xlsxfile(wb)
    @assert is_writable(xf) "XLSXFile instance is not writable."

    file_sheet_template = joinpath(relocatable_data_path, "sheet_template.xml")
    @assert isfile(file_sheet_template) "Couldn't find template file $file_sheet_template."

    if name == ""
        # name was not provided. Will find a unique name.
        i = 1
        current_sheet_names = sheetnames(wb)
        while true
            name = "Sheet$i"
            if !in(name, current_sheet_names)
                # found a unique name
                break
            end
            i += 1
        end
    else
    end

    @assert name != ""

    # checks if name is a unique sheet name
    @assert name ∉ sheetnames(wb) "A sheet named `$name` already exists in this workbook."

    function check_valid_sheetname(n::AbstractString)
        max_length = 31
        @assert(length(n) <= max_length,
                "Invalid sheetname $n: must have at most $max_length characters. Found $(length(n))"
               )

        @assert(!occursin(r"[:\\/\?\*\[\]]+", n),
                "Sheetname cannot contain characters: ':', '\\', '/', '?', '*', '[', ']'."
               )
    end

    check_valid_sheetname(name)

    # generate sheetId
    current_sheet_ids = [ ws.sheetId for ws in wb.sheets ]
    sheetId = max(current_sheet_ids...) + 1

    xdoc = XML.read(file_sheet_template, XML.Node)

    # generate a unique name for the XML
    local xml_filename::String
    i = 1
    while true
        xml_filename = "xl/worksheets/sheet$i.xml"
        if !in(xml_filename, keys(xf.files))
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
    ws = Worksheet(xf, sheetId, rId, name, CellRange("A1:A1"), false)
    # creates a mock WorksheetCache
    # because we can't write to sheet with empty cache (see setdata!(ws::Worksheet, cell::Cell))
    # and the stream should be closed
    # to indicate that no more rows will be fetched from SheetRowStreamIterator in Base.iterate(ws_cache::WorksheetCache, row_from_last_iteration::Int)
    reader = open_internal_file_stream(xf, "xl/worksheets/sheet1.xml") # could be any file
    state =  SheetRowStreamIteratorState(reader, nothing, 0)
    ws.cache = XLSX.WorksheetCache(
        Dict{Int64, Dict{Int64, XLSX.Cell}}(),
        Int64[],
        Dict{Int64, Int64}(),
        SheetRowStreamIterator(ws),
        state,
        false
     ) 

    # adds the new sheet to the list of sheets in the workbook
    push!(wb.sheets, ws)

    # updates workbook xml
    xroot = xmlroot(xf, "xl/workbook.xml")[end]
    for node in XML.children(xroot)
        if XML.tag(node) == "sheets"
            sheet_element = XML.Element("sheet"; name = name)
            sheet_element["r:id"] = rId
            sheet_element["sheetId"] = string(sheetId)
            push!(node, sheet_element)
            break
        end
    end

    return ws
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

# Example

```julia
import XLSX
columns = [ [1, 2, 3, 4], ["Hey", "You", "Out", "There"], [10.2, 20.3, 30.4, 40.5] ]
colnames = [ "integers", "strings", "floats" ]
XLSX.writetable("table.xlsx", columns, colnames)
```

See also: [`XLSX.writetable!`](@ref).
"""
function writetable(filename::Union{AbstractString, IO}, data, columnnames; overwrite::Bool=false, sheetname::AbstractString="", anchor_cell::Union{String, CellRef}=CellRef("A1"))

    if filename isa AbstractString && !overwrite
        @assert !isfile(filename) "$filename already exists."
    end

    xf = open_empty_template(sheetname)
    sheet = xf[1]

    if isa(anchor_cell, String)
        anchor_cell = CellRef(anchor_cell)
    end

    writetable!(sheet, data, columnnames; anchor_cell=anchor_cell)

    # write output file
    writexlsx(filename, xf, overwrite=overwrite)
    nothing
end

"""
    writetable(filename::Union{AbstractString, IO}; overwrite::Bool=false, kw...)
    writetable(filename::Union{AbstractString, IO}, tables::Vector{Tuple{String, Vector{Any}, Vector{String}}}; overwrite::Bool=false)

Write multiple tables.

`kw` is a variable keyword argument list. Each element should be in this format: `sheetname=( data, column_names )`,
where `data` is a vector of columns and `column_names` is a vector of column labels.

Example:

```julia
julia> import DataFrames, XLSX

julia> df1 = DataFrames.DataFrame(COL1=[10,20,30], COL2=["Fist", "Sec", "Third"])

julia> df2 = DataFrames.DataFrame(AA=["aa", "bb"], AB=[10.1, 10.2])

julia> XLSX.writetable("report.xlsx", "REPORT_A" => df1, "REPORT_B" => df2)
```
"""
function writetable(filename::Union{AbstractString, IO}; overwrite::Bool=false, kw...)

    if filename isa AbstractString && !overwrite
        @assert !isfile(filename) "$filename already exists."
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
    nothing
end

function writetable(filename::Union{AbstractString, IO}, tables::Vector{Tuple{String, S, Vector{T}}}; overwrite::Bool=false) where {S<:Vector{U} where U, T<:Union{String, Symbol}}

    if filename isa AbstractString && !overwrite
        @assert !isfile(filename) "$filename already exists."
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
