
@inline get_xlsxfile(wb::Workbook) :: XLSXFile = wb.package
@inline get_xlsxfile(ws::Worksheet) :: XLSXFile = ws.package
@inline get_workbook(ws::Worksheet) :: Workbook = get_xlsxfile(ws).workbook
@inline get_workbook(xl::XLSXFile) :: Workbook = xl.workbook

const ZIP_FILE_HEADER = [ 0x50, 0x4b, 0x03, 0x04 ]
const XLS_FILE_HEADER = [ 0xd0, 0xcf, 0x11, 0xe0 ]

function check_for_xlsx_file_format(source::IO, label::AbstractString="input")
    local header::Vector{UInt8}

    mark(source)
    header = Base.read(source, 4)
    reset(source)

    if header == ZIP_FILE_HEADER # valid Zip file header
        return
    elseif header == XLS_FILE_HEADER # old XLS file
        error("$label looks like an old XLS file (not XLSX). This package does not support XLS file format.")
    else
        error("$label is not a valid XLSX file.")
    end
end

function check_for_xlsx_file_format(filepath::AbstractString)
    @assert isfile(filepath) "File $filepath not found."

    open(filepath, "r") do io
        check_for_xlsx_file_format(io, filepath)
    end
end

"""
    readxlsx(source::Union{AbstractString, IO}) :: XLSXFile

Main function for reading an Excel file.
This function will read the whole Excel file into memory
and return a closed XLSXFile.

Consider using [`XLSX.openxlsx`](@ref) for lazy loading of Excel file contents.
"""
@inline readxlsx(source::Union{AbstractString, IO}) :: XLSXFile = open_or_read_xlsx(source, true, true, false)

"""
    openxlsx(f::F, source::Union{AbstractString, IO}; mode::AbstractString="r", enable_cache::Bool=true) where {F<:Function}

Open XLSX file for reading and/or writing. It returns an opened XLSXFile that will be automatically closed after applying `f` to the file.

# `Do` syntax

This function should be used with `do` syntax, like in:

```julia
XLSX.openxlsx("myfile.xlsx") do xf
    # read data from `xf`
end
```

# Filemodes

The `mode` argument controls how the file is opened. The following modes are allowed:

* `r` : read mode. The existing data in `source` will be accessible for reading. This is the **default** mode.

* `w` : write mode. Opens an empty file that will be written to `source`.

* `rw` : edit mode. Opens `source` for editing. The file will be saved to disk when the function ends.

!!! warning

    The `rw` mode is known to produce some data loss. See [#159](https://github.com/felipenoris/XLSX.jl/issues/159).

    Simple data should work fine. Users are advised to use this feature with caution when working with formulas and charts.

# Arguments

* `source` is IO or the complete path to the file.

* `mode` is the file mode, as explained in the last section.

* `enable_cache`:

If `enable_cache=true`, all read worksheet cells will be cached.
If you read a worksheet cell twice it will use the cached value instead of reading from disk
in the second time.

If `enable_cache=false`, worksheet cells will always be read from disk.
This is useful when you want to read a spreadsheet that doesn't fit into memory.

The default value is `enable_cache=true`.

# Examples

## Read from file

The following example shows how you would read worksheet cells, one row at a time,
where `myfile.xlsx` is a spreadsheet that doesn't fit into memory.

```julia
julia> XLSX.openxlsx("myfile.xlsx", enable_cache=false) do xf
          for r in XLSX.eachrow(xf["mysheet"])
              # read something from row `r`
          end
       end
```

## Write a new file

```julia
XLSX.openxlsx("new.xlsx", mode="w") do xf
    sheet = xf[1]
    sheet[1, :] = [1, Date(2018, 1, 1), "test"]
end
```

## Edit an existing file

```julia
XLSX.openxlsx("edit.xlsx", mode="rw") do xf
    sheet = xf[1]
    sheet[2, :] = [2, Date(2019, 1, 1), "add new line"]
end
```

See also [`XLSX.readxlsx`](@ref).
"""
function openxlsx(f::F, source::Union{AbstractString, IO};
                  mode::AbstractString="r", enable_cache::Bool=true) where {F<:Function}

    _read, _write = parse_file_mode(mode)

    if _read
        @assert source isa IO || isfile(source) "File $source not found."
        xf = open_or_read_xlsx(source, _write, enable_cache, _write)
    else
        xf = open_empty_template()
    end

    try
        f(xf)
    finally

        if _write
            writexlsx(source, xf, overwrite=true)
        else
            close(xf)
        end

        # fix libuv issue on windows (#42) and other systems (#173)
        GC.gc()
    end
end

"""
    openxlsx(source::Union{AbstractString, IO}; mode="r", enable_cache=true) :: XLSXFile

Supports opening a XLSX file without using do-syntax.
In this case, the user is responsible for closing the `XLSXFile`
using `close` or writing it to file using `XLSX.writexlsx`.

See also [`XLSX.writexlsx`](@ref).
"""
function openxlsx(source::Union{AbstractString, IO};
                  mode::AbstractString="r",
                  enable_cache::Bool=true) :: XLSXFile

    _read, _write = parse_file_mode(mode)

    if _read
        @assert source isa IO || isfile(source) "File $source not found."
        return open_or_read_xlsx(source, _write, enable_cache, _write)
    else
        return open_empty_template()
    end
end

function parse_file_mode(mode::AbstractString) :: Tuple{Bool, Bool}
    if mode == "r"
        return (true, false)
    elseif mode == "w"
        return (false, true)
    elseif mode == "rw" || mode == "wr"
        return (true, true)
    else
        error("Couldn't parse file mode $mode.")
    end
end

function open_or_read_xlsx(source::Union{IO, AbstractString}, read_files::Bool, enable_cache::Bool, read_as_template::Bool) :: XLSXFile
    # sanity check
    if read_as_template
        @assert read_files && enable_cache
    end

    xf = XLSXFile(source, enable_cache, read_as_template)

    try
        for f in xf.io.files

            # ignore xl/calcChain.xml in any case (#31)
            if f.name == "xl/calcChain.xml"
                continue
            end

            if endswith(f.name, ".xml") || endswith(f.name, ".rels")
                # XML file
                internal_xml_file_add!(xf, f.name)
                if read_files

                    # ignore worksheet files because they'll be read thru streaming
                    # If reading as template, it will be loaded in two places: here and WorksheetCache.
                    if !read_as_template && startswith(f.name, "xl/worksheets") && endswith(f.name, ".xml")
                        continue
                    end

                    # ignore custom XML internal files
                    if startswith(f.name, "customXml")
                        continue
                    end

                    internal_xml_file_read(xf, f.name)
                end
            elseif read_as_template

                # Binary file
                # we only read binary files to save the Excel file later
                bytes = ZipFile.read(f)
                @assert sizeof(bytes) == f.uncompressedsize
                xf.binary_data[f.name] = bytes
            end
        end

        check_minimum_requirements(xf)
        parse_relationships!(xf)
        parse_workbook!(xf)

        # read data from Worksheet streams
        if read_files
            for sheet_name in sheetnames(xf)
                sheet = getsheet(xf, sheet_name)

                # to read sheet content, we just need to iterate a SheetRowIterator and the data will be stored in cache
                for r in eachrow(sheet)
                    nothing
                end
            end
        end

        if read_as_template
            wb = get_workbook(xf)
            if has_sst(wb)
                sst_load!(wb)
            end
        end

    finally
        if read_files
            close(xf)
        end
    end

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

# Parses package level relationships defined in `_rels/.rels`.
# Parses workbook level relationships defined in `xl/_rels/workbook.xml.rels`.
function parse_relationships!(xf::XLSXFile)

    # package level relationships
    xroot = get_package_relationship_root(xf)
    for el in EzXML.eachelement(xroot)
        push!(xf.relationships, Relationship(el))
    end
    @assert !isempty(xf.relationships) "Relationships not found in _rels/.rels!"

    # workbook level relationships
    wb = get_workbook(xf)
    xroot = get_workbook_relationship_root(xf)
    for el in EzXML.eachelement(xroot)
        push!(wb.relationships, Relationship(el))
    end
    @assert !isempty(wb.relationships) "Relationships not found in xl/_rels/workbook.xml.rels"

    nothing
end

# Updates xf.workbook from xf.data[\"xl/workbook.xml\"]
function parse_workbook!(xf::XLSXFile)
    xroot = xmlroot(xf, "xl/workbook.xml")
    @assert EzXML.nodename(xroot) == "workbook" "Malformed xl/workbook.xml. Root node name should be 'workbook'. Got '$(EzXML.nodename(xroot))'."

    # workbook to be parsed
    workbook = get_workbook(xf)

    # workbookPr -> date1904
    # does not have attribute => is not date1904
    workbook.date1904 = false

    # changes workbook.date1904 if there is a setting in the workbookPr node
    for node in EzXML.eachelement(xroot)
        if EzXML.nodename(node) == "workbookPr"

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
            end

            break
        end
    end

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
                elseif occursin(r"^\".*\"$", defined_value_string) # is enclosed by quotes
                    defined_value = defined_value_string[2:end-1] # remove enclosing quotes
                    if isempty(defined_value)
                        defined_value = missing
                    end
                elseif tryparse(Int, defined_value_string) != nothing
                    defined_value = parse(Int, defined_value_string)
                elseif tryparse(Float64, defined_value_string) != nothing
                    defined_value = parse(Float64, defined_value_string)
                elseif isempty(defined_value_string)
                    defined_value = missing
                else

                    # Couldn't parse definedName. Will silently ignore it, since this is not a critical feature.
                    continue

                    # debug
                    #error("Could not parse value $(defined_value_string) for definedName $name.")
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

# Lazy loading of XML files

# Lists internal files from the XLSX package.
@inline filenames(xl::XLSXFile) = keys(xl.files)

# Returns true if the file data was read into xl.data.
@inline internal_xml_file_isread(xl::XLSXFile, filename::String) :: Bool = xl.files[filename]
@inline internal_xml_file_exists(xl::XLSXFile, filename::String) :: Bool = haskey(xl.files, filename)

function internal_xml_file_add!(xl::XLSXFile, filename::String)
    @assert endswith(filename, ".xml") || endswith(filename, ".rels")
    xl.files[filename] = false
    nothing
end

function internal_xml_file_read(xf::XLSXFile, filename::String) :: EzXML.Document
    @assert internal_xml_file_exists(xf, filename) "Couldn't find $filename in $(xf.source)."

    if !internal_xml_file_isread(xf, filename)
        @assert isopen(xf) "Can't read from a closed XLSXFile."
        file_not_found = true
        for f in xf.io.files
            if f.name == filename
                xf.files[filename] = true # set file as read

                try
                    xf.data[filename] = EzXML.readxml(f)
                catch err
                    @error("Failed to parse internal XML file `$filename`")
                    rethrow()
                end

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

    # close all internal file streams from worksheet caches
    for sheet in xl.workbook.sheets
        if sheet.cache != nothing && sheet.cache.stream_state != nothing
            close(sheet.cache.stream_state)
        end
    end
end

Base.isopen(xl::XLSXFile) = xl.io_is_open

# Utility method to find the XMLDocument associated with a given package filename.
# Returns xl.data[filename] if it exists. Throws an error if it doesn't.
@inline xmldocument(xl::XLSXFile, filename::String) :: EzXML.Document = internal_xml_file_read(xl, filename)

# Utility method to return the root element of a given XMLDocument from the package.
# Returns EzXML.root(xl.data[filename]) if it exists.
@inline xmlroot(xl::XLSXFile, filename::String) :: EzXML.Node = EzXML.root(xmldocument(xl, filename))

#
# Helper Functions
#

"""
    readdata(source, sheet, ref)
    readdata(source, sheetref)

Returns a scalar or matrix with values from a spreadsheet.

See also [`XLSX.getdata`](@ref).

# Examples

These function calls are equivalent.

```julia
julia> XLSX.readdata("myfile.xlsx", "mysheet", "A2:B4")
3×2 Array{Any,2}:
 1  "first"
 2  "second"
 3  "third"

julia> XLSX.readdata("myfile.xlsx", 1, "A2:B4")
3×2 Array{Any,2}:
 1  "first"
 2  "second"
 3  "third"

julia> XLSX.readdata("myfile.xlsx", "mysheet!A2:B4")
3×2 Array{Any,2}:
 1  "first"
 2  "second"
 3  "third"
```
"""
function readdata(source::Union{AbstractString, IO}, sheet::Union{AbstractString, Int}, ref)
    c = openxlsx(source, enable_cache=false) do xf
        getdata(getsheet(xf, sheet), ref)
    end
    return c
end

function readdata(source::Union{AbstractString, IO}, sheetref::AbstractString)
    c = openxlsx(source, enable_cache=false) do xf
        getdata(xf, sheetref)
    end
    return c
end

"""
    readtable(
        source,
        sheet,
        [columns];
        [first_row],
        [column_labels],
        [header],
        [infer_eltypes],
        [stop_in_empty_row],
        [stop_in_row_function],
        [keep_empty_rows]
    ) -> DataTable

Returns tabular data from a spreadsheet as a struct `XLSX.DataTable`.
Use this function to create a `DataFrame` from package `DataFrames.jl`.

Use `columns` argument to specify which columns to get.
For example, `"B:D"` will select columns `B`, `C` and `D`.
If `columns` is not given, the algorithm will find the first sequence
of consecutive non-empty cells.

Use `first_row` to indicate the first row from the table.
`first_row=5` will look for a table starting at sheet row `5`.
If `first_row` is not given, the algorithm will look for the first
non-empty row in the spreadsheet.

`header` is a `Bool` indicating if the first row is a header.
If `header=true` and `column_labels` is not specified, the column labels
for the table will be read from the first row of the table.
If `header=false` and `column_labels` is not specified, the algorithm
will generate column labels. The default value is `header=true`.

Use `column_labels` to specify names for the header of the table.

Use `infer_eltypes=true` to get `data` as a `Vector{Any}` of typed vectors.
The default value is `infer_eltypes=false`.

`stop_in_empty_row` is a boolean indicating whether an empty row marks the end of the table.
If `stop_in_empty_row=false`, the `TableRowIterator` will continue to fetch rows until there's no more rows in the Worksheet.
The default behavior is `stop_in_empty_row=true`.

`stop_in_row_function` is a Function that receives a `TableRow` and returns a `Bool` indicating if the end of the table was reached.

Example for `stop_in_row_function`:

```
function stop_function(r)
    v = r[:col_label]
    return !ismissing(v) && v == "unwanted value"
end
```

`keep_empty_rows` determines whether rows where all column values are equal to `missing` are kept (`true`) or dropped (`false`) from the resulting table. 
`keep_empty_rows` never affects the *bounds* of the table; the number of rows read from a sheet is only affected by, `first_row`, `stop_in_empty_row` and `stop_in_row_function` (if specified).
`keep_empty_rows` is only checked once the first and last row of the table have been determined, to see whether to keep or drop empty rows between the first and the last row.

# Example

```julia
julia> using DataFrames, XLSX

julia> df = DataFrame(XLSX.readtable("myfile.xlsx", "mysheet"))
```

See also: [`XLSX.gettable`](@ref).
"""
function readtable(source::Union{AbstractString, IO}, sheet::Union{AbstractString, Int}; first_row::Union{Nothing, Int} = nothing, column_labels=nothing, header::Bool=true, infer_eltypes::Bool=false, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Nothing, Function}=nothing, enable_cache::Bool=false, keep_empty_rows::Bool=false)
    c = openxlsx(source, enable_cache=enable_cache) do xf
        gettable(getsheet(xf, sheet); first_row=first_row, column_labels=column_labels, header=header, infer_eltypes=infer_eltypes, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function, keep_empty_rows=keep_empty_rows)
    end
    return c
end

function readtable(source::Union{AbstractString, IO}, sheet::Union{AbstractString, Int}, columns::Union{ColumnRange, AbstractString}; first_row::Union{Nothing, Int} = nothing, column_labels=nothing, header::Bool=true, infer_eltypes::Bool=false, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Nothing, Function}=nothing, enable_cache::Bool=false, keep_empty_rows::Bool=false)
    c = openxlsx(source, enable_cache=enable_cache) do xf
        gettable(getsheet(xf, sheet), columns; first_row=first_row, column_labels=column_labels, header=header, infer_eltypes=infer_eltypes, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function, keep_empty_rows=keep_empty_rows)
    end
    return c
end
