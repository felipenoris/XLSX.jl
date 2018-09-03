var documenterSearchIndex = {"docs": [

{
    "location": "index.html#",
    "page": "Tutorial",
    "title": "Tutorial",
    "category": "page",
    "text": ""
},

{
    "location": "index.html#Tutorial-1",
    "page": "Tutorial",
    "title": "Tutorial",
    "category": "section",
    "text": ""
},

{
    "location": "index.html#Installation-1",
    "page": "Tutorial",
    "title": "Installation",
    "category": "section",
    "text": "julia> Pkg.add(\"XLSX\")"
},

{
    "location": "index.html#Getting-Started-1",
    "page": "Tutorial",
    "title": "Getting Started",
    "category": "section",
    "text": "The basic usage is to read an Excel file and read values.julia> import XLSX\n\njulia> xf = XLSX.readxlsx(\"myfile.xlsx\")\nXLSXFile(\"myfile.xlsx\") containing 3 Worksheets\n            sheetname size          range\n-------------------------------------------------\n              mysheet 4x2           A1:B4\n           othersheet 1x1           A1:A1\n                named 1x1           B4:B4\n\njulia> XLSX.sheetnames(xf)\n3-element Array{String,1}:\n \"mysheet\"\n \"othersheet\"\n \"named\"\n\njulia> sh = xf[\"mysheet\"] # get a reference to a Worksheet\n4×2 XLSX.Worksheet: [\"mysheet\"](A1:B4)\n\njulia> sh[\"B2\"] # From a sheet, you can access a cell value\n\"first\"\n\njulia> sh[\"A2:B4\"] # or a cell range\n3×2 Array{Any,2}:\n 1  \"first\"\n 2  \"second\"\n 3  \"third\"\n\njulia> XLSX.readdata(\"myfile.xlsx\", \"mysheet\", \"A2:B4\") # shorthand for all above\n3×2 Array{Any,2}:\n 1  \"first\"\n 2  \"second\"\n 3  \"third\"\n\njulia> sh[:] # all data inside worksheet\'s dimension\n4×2 Array{Any,2}:\n  \"HeaderA\"  \"HeaderB\"\n 1           \"first\"\n 2           \"second\"\n 3           \"third\"\n\njulia> xf[\"mysheet!A2:B4\"] # you can also query values from a file reference\n3×2 Array{Any,2}:\n 1  \"first\"\n 2  \"second\"\n 3  \"third\"\n\njulia> xf[\"NAMED_CELL\"] # you can even read named ranges\n\"B4 is a named cell from sheet \\\"named\\\"\"\n\njulia> xf[\"mysheet!A:B\"] # Column ranges are also supported\n4×2 Array{Any,2}:\n  \"HeaderA\"  \"HeaderB\"\n 1           \"first\"\n 2           \"second\"\n 3           \"third\"\nTo inspect the internal representation of each cell, use the getcell or getcellrange methods.The example above used xf = XLSX.readxlsx(filename) to open a file, so all file contents will be fetched at once from disk.You can also use XLSX.openxlsx to read file contents as needed (see section about streaming below)."
},

{
    "location": "index.html#Read-Tabular-Data-1",
    "page": "Tutorial",
    "title": "Read Tabular Data",
    "category": "section",
    "text": "The gettable method returns tabular data from a spreadsheet as a tuple (data, column_labels). You can use it to create a DataFrame from DataFrames.jl. Check the docstring for gettable method for more advanced options.There\'s also a helper method readtable to read from file directly, as shown in the following example.julia> using DataFrames, XLSX\n\njulia> df = DataFrame(XLSX.readtable(\"myfile.xlsx\", \"mysheet\")...)\n3×2 DataFrames.DataFrame\n│ Row │ HeaderA │ HeaderB  │\n├─────┼─────────┼──────────┤\n│ 1   │ 1       │ \"first\"  │\n│ 2   │ 2       │ \"second\" │\n│ 3   │ 3       │ \"third\"  │"
},

{
    "location": "index.html#Reading-Large-Excel-Files-and-Caching-1",
    "page": "Tutorial",
    "title": "Reading Large Excel Files and Caching",
    "category": "section",
    "text": "The method XLSX.openxlsx has a enable_cache option to control worksheet cells caching.Cache is enabled by default, so if you read a worksheet cell twice it will use the cached value instead of reading from disk in the second time.If enable_cache=false, worksheet cells will always be read from disk. This is useful when you want to read a spreadsheet that doesn\'t fit into memory.The following example shows how you would read worksheet cells, one row at a time, where myfile.xlsx is a spreadsheet that doesn\'t fit into memory.julia> XLSX.openxlsx(\"myfile.xlsx\", enable_cache=false) do f\n           sheet = f[\"mysheet\"]\n           for r in XLSX.eachrow(sheet)\n              # r is a `SheetRow`, values are read using column references\n              rn = XLSX.row_number(r) # `SheetRow` row number\n              v1 = r[1]    # will read value at column 1\n              v2 = r[\"B\"]  # will read value at column 2\n\n              println(\"v1=$v1, v2=$v2\")\n           end\n      end\nv1=HeaderA, v2=HeaderB\nv1=1, v2=first\nv1=2, v2=second\nv1=3, v2=thirdYou could also stream tabular data using XLSX.eachtablerow(sheet), which is the underlying iterator in gettable method. Check docstrings for XLSX.eachtablerow for more advanced options.julia> XLSX.openxlsx(\"myfile.xlsx\", enable_cache=false) do f\n           sheet = f[\"mysheet\"]\n           for r in XLSX.eachtablerow(sheet)\n               # r is a `TableRow`, values are read using column labels or numbers\n               rn = XLSX.row_number(r) # `TableRow` row number\n               v1 = r[1] # will read value at table column 1\n               v2 = r[:HeaderB] # will read value at column labeled `:HeaderB`\n\n               println(\"v1=$v1, v2=$v2\")\n            end\n       end\nv1=1, v2=first\nv1=2, v2=second\nv1=3, v2=third"
},

{
    "location": "index.html#Writing-Excel-Files-1",
    "page": "Tutorial",
    "title": "Writing Excel Files",
    "category": "section",
    "text": ""
},

{
    "location": "index.html#Create-New-Files-1",
    "page": "Tutorial",
    "title": "Create New Files",
    "category": "section",
    "text": "Opening a file in write mode with XLSX.openxlsx will open a new (blank) Excel file for editing.XLSX.openxlsx(\"my_new_file.xlsx\", mode=\"w\") do xf\n    sheet = xf[1]\n    XLSX.rename!(sheet, \"new_sheet\")\n    sheet[\"A1\"] = \"this\"\n    sheet[\"A2\"] = \"is a\"\n    sheet[\"A3\"] = \"new file\"\n    sheet[\"A4\"] = 100\nend"
},

{
    "location": "index.html#Edit-Existing-Files-1",
    "page": "Tutorial",
    "title": "Edit Existing Files",
    "category": "section",
    "text": "Opening a file in read-write mode with XLSX.openxlsx will open an existing Excel file for editing. This will preserve existing data in the original file.XLSX.openxlsx(\"my_new_file.xlsx\", mode=\"rw\") do xf\n    sheet = xf[1]\n    sheet[\"B1\"] = \"new data\"\nend"
},

{
    "location": "index.html#Export-Tabular-Data-1",
    "page": "Tutorial",
    "title": "Export Tabular Data",
    "category": "section",
    "text": "To export tabular data to Excel, use XLSX.writetable method.julia> import DataFrames, XLSX\n\njulia> df = DataFrames.DataFrame(integers=[1, 2, 3, 4], strings=[\"Hey\", \"You\", \"Out\", \"There\"], floats=[10.2, 20.3, 30.4, 40.5], dates=[Date(2018,2,20), Date(2018,2,21), Date(2018,2,22), Date(2018,2,23)], times=[Dates.Time(19,10), Dates.Time(19,20), Dates.Time(19,30), Dates.Time(19,40)], datetimes=[Dates.DateTime(2018,5,20,19,10), Dates.DateTime(2018,5,20,19,20), Dates.DateTime(2018,5,20,19,30), Dates.DateTime(2018,5,20,19,40)])\n4×6 DataFrames.DataFrame\n│ Row │ integers │ strings │ floats │ dates      │ times    │ datetimes           │\n├─────┼──────────┼─────────┼────────┼────────────┼──────────┼─────────────────────┤\n│ 1   │ 1        │ Hey     │ 10.2   │ 2018-02-20 │ 19:10:00 │ 2018-05-20T19:10:00 │\n│ 2   │ 2        │ You     │ 20.3   │ 2018-02-21 │ 19:20:00 │ 2018-05-20T19:20:00 │\n│ 3   │ 3        │ Out     │ 30.4   │ 2018-02-22 │ 19:30:00 │ 2018-05-20T19:30:00 │\n│ 4   │ 4        │ There   │ 40.5   │ 2018-02-23 │ 19:40:00 │ 2018-05-20T19:40:00 │\n\njulia> XLSX.writetable(\"df.xlsx\", DataFrames.columns(df), DataFrames.names(df))You can also export multiple tables to Excel, each table in a separate worksheet.julia> import DataFrames, XLSX\n\njulia> df1 = DataFrames.DataFrame(COL1=[10,20,30], COL2=[\"Fist\", \"Sec\", \"Third\"])\n3×2 DataFrames.DataFrame\n│ Row │ COL1 │ COL2  │\n├─────┼──────┼───────┤\n│ 1   │ 10   │ Fist  │\n│ 2   │ 20   │ Sec   │\n│ 3   │ 30   │ Third │\n\njulia> df2 = DataFrames.DataFrame(AA=[\"aa\", \"bb\"], AB=[10.1, 10.2])\n2×2 DataFrames.DataFrame\n│ Row │ AA │ AB   │\n├─────┼────┼──────┤\n│ 1   │ aa │ 10.1 │\n│ 2   │ bb │ 10.2 │\n\njulia> XLSX.writetable(\"report.xlsx\", REPORT_A=( DataFrames.columns(df1), DataFrames.names(df1) ), REPORT_B=( DataFrames.columns(df2), DataFrames.names(df2) ))"
},

{
    "location": "api.html#",
    "page": "API",
    "title": "API",
    "category": "page",
    "text": ""
},

{
    "location": "api.html#XLSX.CellDataFormat",
    "page": "API",
    "title": "XLSX.CellDataFormat",
    "category": "type",
    "text": "Keeps track of formatting information.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.CellRange",
    "page": "API",
    "title": "XLSX.CellRange",
    "category": "type",
    "text": "A CellRange represents a rectangular range of cells in a spreadsheet.\n\nCellRange(\"A1:C4\") denotes cells ranging from A1 (upper left corner) to C4 (bottom right corner).\n\nAs a convenience, @range_str macro is provided.\n\ncr = XLSX.range\"A1:C4\"\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.CellRef",
    "page": "API",
    "title": "XLSX.CellRef",
    "category": "type",
    "text": "A CellRef represents a cell location given by row and column identifiers.\n\nCellRef(\"A6\") indicates a cell located at column 1 and row 6.\n\nExample:\n\ncn = XLSX.CellRef(\"AB1\")\nprintln( XLSX.row_number(cn) ) # will print 1\nprintln( XLSX.column_number(cn) ) # will print 28\nprintln( string(cn) ) # will print out AB1\n\nAs a convenience, @ref_str macro is provided.\n\ncn = XLSX.ref\"AB1\"\nprintln( XLSX.row_number(cn) ) # will print 1\nprintln( XLSX.column_number(cn) ) # will print 28\nprintln( string(cn) ) # will print out AB1\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.CellValue",
    "page": "API",
    "title": "XLSX.CellValue",
    "category": "type",
    "text": "CellValue is a Julia type of a value read from a Spreadsheet.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.Relationship",
    "page": "API",
    "title": "XLSX.Relationship",
    "category": "type",
    "text": "Relationships are defined in ECMA-376-1 Section 9.2. This struct matches the Relationship tag attribute names.\n\nA Relashipship defines relations between the files inside a MSOffice package. Regarding Spreadsheets, there are two kinds of relationships:\n\n* package level: defined in `_rels/.rels`.\n* workbook level: defined in `xl/_rels/workbook.xml.rels`.\n\nThe function parse_relationships!(xf::XLSXFile) is used to parse package and workbook level relationships.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.SharedStrings",
    "page": "API",
    "title": "XLSX.SharedStrings",
    "category": "type",
    "text": "Shared String Table\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.SheetRowIterator",
    "page": "API",
    "title": "XLSX.SheetRowIterator",
    "category": "type",
    "text": "Iterates over Worksheet cells. See eachrow method docs. Each element is a SheetRow.\n\nImplementations: SheetRowStreamIterator, WorksheetCache.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.Workbook",
    "page": "API",
    "title": "XLSX.Workbook",
    "category": "type",
    "text": "Workbook is the result of parsing file xl/workbook.xml.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.XLSXFile",
    "page": "API",
    "title": "XLSX.XLSXFile",
    "category": "type",
    "text": "XLSXFile stores all XML data from an Excel file.\n\nfilepath is the filepath of the source file for this XLSXFile. data stored the raw XML data. It maps internal XLSX filenames to XMLDocuments. workbook is the result of parsing xl/workbook.xml.\n\n\n\n\n\n"
},

{
    "location": "api.html#Base.in-Tuple{XLSX.CellRef,XLSX.CellRange}",
    "page": "API",
    "title": "Base.in",
    "category": "method",
    "text": "Base.in(ref::CellRef, rng::CellRange) :: Bool\n\nChecks wether ref is a cell reference inside a range given by rng.\n\n\n\n\n\n"
},

{
    "location": "api.html#Base.issubset-Tuple{XLSX.CellRange,XLSX.CellRange}",
    "page": "API",
    "title": "Base.issubset",
    "category": "method",
    "text": "Base.issubset(subrng::CellRange, rng::CellRange)\n\nChecks wether subrng is a cell range contained in rng.\n\n\n\n\n\n"
},

{
    "location": "api.html#Base.iterate",
    "page": "API",
    "title": "Base.iterate",
    "category": "function",
    "text": "SheetRowStreamIterator(ws::Worksheet)\n\nCreates a reader for row elements in the Worksheet\'s XML. Will return a stream reader positioned in the first row element if it exists.\n\nIf there\'s no row element inside sheetData XML tag, it will close all streams and return nothing.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.add_relationship!-Tuple{XLSX.Workbook,String,String}",
    "page": "API",
    "title": "XLSX.add_relationship!",
    "category": "method",
    "text": "Adds new relationship. Returns new generated rId.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.add_shared_string!-Tuple{XLSX.Workbook,AbstractString,AbstractString}",
    "page": "API",
    "title": "XLSX.add_shared_string!",
    "category": "method",
    "text": "add_shared_string!(sheet, str_unformatted, [str_formatted]) :: Int\n\nAdd string to shared string table. Returns the 0-based index of the shared string in the shared string table.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.addsheet!",
    "page": "API",
    "title": "XLSX.addsheet!",
    "category": "function",
    "text": "addsheet!(workbook, [name]) :: Worksheet\n\nCreate a new worksheet with named name. If name is not provided, a unique name is created.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.column_bounds-Tuple{XLSX.SheetRow}",
    "page": "API",
    "title": "XLSX.column_bounds",
    "category": "method",
    "text": "column_bounds(sr::SheetRow)\n\nReturns a tuple with the first and last index of the columns for a SheetRow.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.column_number-Tuple{XLSX.CellRef}",
    "page": "API",
    "title": "XLSX.column_number",
    "category": "method",
    "text": "column_number(c::CellRef) :: Int\n\nReturns the column number of a given cell reference.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.decode_column_number-Tuple{AbstractString}",
    "page": "API",
    "title": "XLSX.decode_column_number",
    "category": "method",
    "text": "decode_column_number(column_name::AbstractString) :: Int\n\nConverts column name to a column number.\n\njulia> XLSX.decode_column_number(\"D\")\n4\n\nSee also: encode_column_number.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.default_cell_format-Tuple{XLSX.Worksheet,Union{Missing, Bool, Float64, Int64, Date, DateTime, Time, String}}",
    "page": "API",
    "title": "XLSX.default_cell_format",
    "category": "method",
    "text": "Returns the default CellDataFormat for a type\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.eachrow-Tuple{XLSX.Worksheet}",
    "page": "API",
    "title": "XLSX.eachrow",
    "category": "method",
    "text": "eachrow(sheet)\n\nCreates a row iterator for a worksheet.\n\nExample: Query all cells from columns 1 to 4.\n\nleft = 1  # 1st column\nright = 4 # 4th column\nfor sheetrow in XLSX.eachrow(sheet)\n    for column in left:right\n        cell = XLSX.getcell(sheetrow, column)\n\n        # do something with cell\n    end\nend\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.eachtablerow-Tuple{XLSX.Worksheet,Union{ColumnRange, AbstractString}}",
    "page": "API",
    "title": "XLSX.eachtablerow",
    "category": "method",
    "text": "eachtablerow(sheet, [columns]; [first_row], [column_labels], [header], [stop_in_empty_row], [stop_in_row_function])\n\nConstructs an iterator of table rows. Each element of the iterator is of type TableRow.\n\nheader is a boolean indicating wether the first row of the table is a table header.\n\nIf header == false and no names were supplied, column names will be generated following the column names found in the Excel file. Also, the column range will be inferred by the non-empty contiguous cells in the first row of the table.\n\nThe user can replace column names by assigning the optional names input variable with a Vector{Symbol}.\n\nstop_in_empty_row is a boolean indicating wether an empty row marks the end of the table. If stop_in_empty_row=false, the iterator will continue to fetch rows until there\'s no more rows in the Worksheet. The default behavior is stop_in_empty_row=true. Empty rows may be returned by the iterator when stop_in_empty_row=false.\n\nstop_in_row_function is a Function that receives a TableRow and returns a Bool indicating if the end of the table was reached.\n\nExample for stop_in_row_function:\n\nfunction stop_function(r)\n    v = r[:col_label]\n    return !ismissing(v) && v == \"unwanted value\"\nend\n\nExample code:\n\nfor r in XLSX.eachtablerow(sheet)\n    # r is a `TableRow`. Values are read using column labels or numbers.\n    rn = XLSX.row_number(r) # `TableRow` row number.\n    v1 = r[1] # will read value at table column 1.\n    v2 = r[:COL_LABEL2] # will read value at column labeled `:COL_LABEL2`.\nend\n\nSee also gettable.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.encode_column_number-Tuple{Int64}",
    "page": "API",
    "title": "XLSX.encode_column_number",
    "category": "method",
    "text": "encode_column_number(column_number::Int) :: String\n\nConverts column number to a column name.\n\nExample\n\njulia> XLSX.encode_column_number(4)\n\"D\"\n\nSee also: decode_column_number.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.excel_value_to_date-Tuple{Int64,Bool}",
    "page": "API",
    "title": "XLSX.excel_value_to_date",
    "category": "method",
    "text": "Converts Excel number to Date.\n\nSee also: isdate1904 function.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.excel_value_to_datetime-Tuple{Float64,Bool}",
    "page": "API",
    "title": "XLSX.excel_value_to_datetime",
    "category": "method",
    "text": "Converts Excel number to DateTime.\n\nThe decimal part represents the Time (see _time function). The integer part represents the Date.\n\nSee also: isdate1904 function.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.excel_value_to_time-Tuple{Float64}",
    "page": "API",
    "title": "XLSX.excel_value_to_time",
    "category": "method",
    "text": "Converts Excel number to Time. x must be between 0 and 1.\n\nTo represent Time, Excel uses the decimal part of a floating point number. 1 equals one day.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.filenames-Tuple{XLSX.XLSXFile}",
    "page": "API",
    "title": "XLSX.filenames",
    "category": "method",
    "text": "Lists internal files from the XLSX package.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.get_dimension-Tuple{XLSX.Worksheet}",
    "page": "API",
    "title": "XLSX.get_dimension",
    "category": "method",
    "text": "Retuns the dimension of this worksheet as a CellRange.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.get_shared_string_index-Tuple{XLSX.SharedStrings,AbstractString}",
    "page": "API",
    "title": "XLSX.get_shared_string_index",
    "category": "method",
    "text": "Checks if string is inside shared string table. Returns -1 if it\'s not in the shared string table. Returns the index of the string in the shared string table. The index is 0-based.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.getcell-Tuple{XLSX.Worksheet,XLSX.CellRef}",
    "page": "API",
    "title": "XLSX.getcell",
    "category": "method",
    "text": "getcell(sheet, ref)\n\nReturns an AbstractCell that represents a cell in the spreadsheet.\n\nExample:\n\njulia> xf = XLSX.readxlsx(\"myfile.xlsx\")\n\njulia> sheet = xf[\"mysheet\"]\n\njulia> cell = XLSX.getcell(sheet, \"A1\")\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.getcellrange-Tuple{XLSX.Worksheet,XLSX.CellRange}",
    "page": "API",
    "title": "XLSX.getcellrange",
    "category": "method",
    "text": "getcellrange(sheet, rng)\n\nReturns a matrix with cells as Array{AbstractCell, 2}. rng must be a valid cell range, as in \"A1:B2\".\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.getdata-Tuple{XLSX.Worksheet,XLSX.CellRef}",
    "page": "API",
    "title": "XLSX.getdata",
    "category": "method",
    "text": "getdata(sheet, ref)\n\nReturns a escalar or a matrix with values from a spreadsheet. ref can be a cell reference or a range.\n\nIndexing in a Worksheet will dispatch to getdata method.\n\nExample\n\njulia> f = XLSX.readxlsx(\"myfile.xlsx\")\n\njulia> sheet = f[\"mysheet\"]\n\njulia> v = sheet[\"A1:B4\"]\n\nSee also readdata.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.getdata-Tuple{XLSX.Worksheet,XLSX.Cell}",
    "page": "API",
    "title": "XLSX.getdata",
    "category": "method",
    "text": "getdata(ws::Worksheet, cell::Cell) :: CellValue\n\nReturns a Julia representation of a given cell value. The result data type is chosen based on the value of the cell as well as its style.\n\nFor example, date is stored as integers inside the spreadsheet, and the style is the information that is taken into account to chose Date as the result type.\n\nFor numbers, if the style implies that the number is visualized with decimals, the method will return a float, even if the underlying number is stored as an integer inside the spreadsheet XML.\n\nIf cell has empty value or empty String, this function will return missing.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.gettable-Tuple{XLSX.Worksheet,Union{ColumnRange, AbstractString}}",
    "page": "API",
    "title": "XLSX.gettable",
    "category": "method",
    "text": "gettable(sheet, [columns]; [first_row], [column_labels], [header], [infer_eltypes], [stop_in_empty_row], [stop_in_row_function]) -> data, column_labels\n\nReturns tabular data from a spreadsheet as a tuple (data, column_labels). data is a vector of columns. column_labels is a vector of symbols. Use this function to create a DataFrame from package DataFrames.jl.\n\nUse columns argument to specify which columns to get. For example, columns=\"B:D\" will select columns B, C and D. If columns is not given, the algorithm will find the first sequence of consecutive non-empty cells.\n\nUse first_row to indicate the first row from the table. first_row=5 will look for a table starting at sheet row 5. If first_row is not given, the algorithm will look for the first non-empty row in the spreadsheet.\n\nheader is a Bool indicating if the first row is a header. If header=true and column_labels is not specified, the column labels for the table will be read from the first row of the table. If header=false and column_labels is not specified, the algorithm will generate column labels. The default value is header=true.\n\nUse column_labels as a vector of symbols to specify names for the header of the table.\n\nUse infer_eltypes=true to get data as a Vector{Any} of typed vectors. The default value is infer_eltypes=false.\n\nstop_in_empty_row is a boolean indicating wether an empty row marks the end of the table. If stop_in_empty_row=false, the TableRowIterator will continue to fetch rows until there\'s no more rows in the Worksheet. The default behavior is stop_in_empty_row=true.\n\nstop_in_row_function is a Function that receives a TableRow and returns a Bool indicating if the end of the table was reached.\n\nExample for stop_in_row_function:\n\nfunction stop_function(r)\n    v = r[:col_label]\n    return !ismissing(v) && v == \"unwanted value\"\nend\n\nRows where all column values are equal to missing are dropped.\n\nExample code for gettable:\n\njulia> using DataFrames, XLSX\n\njulia> df = XLSX.openxlsx(\"myfile.xlsx\") do xf\n                DataFrame(XLSX.gettable(xf[\"mysheet\"])...)\n            end\n\nSee also: readtable.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.has_sst-Tuple{XLSX.Workbook}",
    "page": "API",
    "title": "XLSX.has_sst",
    "category": "method",
    "text": "has_sst(workbook::Workbook)\n\nChecks wether this workbook has a Shared String Table.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.internal_xml_file_isread-Tuple{XLSX.XLSXFile,String}",
    "page": "API",
    "title": "XLSX.internal_xml_file_isread",
    "category": "method",
    "text": "Returns true if the file data was read into xl.data.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.is_cache_enabled-Tuple{XLSX.Worksheet}",
    "page": "API",
    "title": "XLSX.is_cache_enabled",
    "category": "method",
    "text": "Indicates wether worksheet cache will be fed while reading worksheet cells.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.is_end_of_sheet_data-Tuple{EzXML.StreamReader}",
    "page": "API",
    "title": "XLSX.is_end_of_sheet_data",
    "category": "method",
    "text": "Detects a closing sheetData element\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.is_valid_cellname-Tuple{AbstractString}",
    "page": "API",
    "title": "XLSX.is_valid_cellname",
    "category": "method",
    "text": "is_valid_cellname(n::AbstractString) :: Bool\n\nChecks wether n is a valid name for a cell.\n\nCell names are bounded by A1 : XFD1048576.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.is_writable-Tuple{XLSX.XLSXFile}",
    "page": "API",
    "title": "XLSX.is_writable",
    "category": "method",
    "text": "is_writable(xl::XLSXFile)\n\nIndicates wether this XLSX file can be edited. This controls if assignment to worksheet cells is allowed. Writable XLSXFile instances are opened with XLSX.open_xlsx_template method.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.isdate1904-Tuple{XLSX.Workbook}",
    "page": "API",
    "title": "XLSX.isdate1904",
    "category": "method",
    "text": "isdate1904(wb) :: Bool\n\nReturns true if workbook follows date1904 convention.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.open_empty_template",
    "page": "API",
    "title": "XLSX.open_empty_template",
    "category": "function",
    "text": "open_empty_template(sheetname::AbstractString=\"\") :: XLSXFile\n\nReturns an empty, writable XLSXFile with 1 worksheet.\n\nsheetname is the name of the worksheet, defaults to Sheet1.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.open_internal_file_stream-Tuple{XLSX.XLSXFile,String}",
    "page": "API",
    "title": "XLSX.open_internal_file_stream",
    "category": "method",
    "text": "Open a file for streaming.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.open_xlsx_template-Tuple{AbstractString}",
    "page": "API",
    "title": "XLSX.open_xlsx_template",
    "category": "method",
    "text": "open_xlsx_template(filepath::AbstractString) :: XLSXFile\n\nOpen an Excel file as template for editing and saving to another file with XLSX.writexlsx.\n\nThe returned XLSXFile instance is in closed state.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.openxlsx-Tuple{Function,AbstractString}",
    "page": "API",
    "title": "XLSX.openxlsx",
    "category": "method",
    "text": "openxlsx(f::Function, filename::AbstractString; mode::AbstractString=\"r\", enable_cache::Bool=true)\n\nOpen XLSX file for reading and/or writing. It returns an opened XLSXFile that will be automatically closed after applying f to the file.\n\nDo syntax\n\nThis function should be used with do syntax, like in:\n\nXLSX.openxlsx(\"myfile.xlsx\") do xf\n    # read data from `xf`\nend\n\nFilemodes\n\nThe mode argument controls how the file is opened. The following modes are allowed:\n\nr : read mode. The existing data in filename will be accessible for reading. This is the default mode.\nw : write mode. Opens an empty file that will be written to filename.\nrw : edit mode. Opens filename for editing. The file will be saved to disk when the function ends.\n\nArguments\n\nfilename is the name of the file.\nmode is the file mode, as explained in the last section.\nenable_cache:\n\nIf enable_cache=true, all read worksheet cells will be cached. If you read a worksheet cell twice it will use the cached value instead of reading from disk in the second time.\n\nIf enable_cache=false, worksheet cells will always be read from disk. This is useful when you want to read a spreadsheet that doesn\'t fit into memory.\n\nThe default value is enable_cache=true.\n\nExamples\n\nRead from file\n\nThe following example shows how you would read worksheet cells, one row at a time, where myfile.xlsx is a spreadsheet that doesn\'t fit into memory.\n\njulia> XLSX.openxlsx(\"myfile.xlsx\", enable_cache=false) do xf\n          for r in XLSX.eachrow(xf[\"mysheet\"])\n              # read something from row `r`\n          end\n       end\n\nWrite a new file\n\nXLSX.openxlsx(\"new.xlsx\", mode=\"w\") do xf\n    sheet = xf[1]\n    sheet[1, :] = [1, Date(2018, 1, 1), \"test\"]\nend\n\nEdit an existing file\n\nXLSX.openxlsx(\"edit.xlsx\", mode=\"rw\") do xf\n    sheet = xf[1]\n    sheet[2, :] = [2, Date(2019, 1, 1), \"add new line\"]\nend\n\nSee also readxlsx method.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.parse_file_mode-Tuple{AbstractString}",
    "page": "API",
    "title": "XLSX.parse_file_mode",
    "category": "method",
    "text": "Parses filemode string to the tuple (read, write). See openxlsx.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.parse_relationships!-Tuple{XLSX.XLSXFile}",
    "page": "API",
    "title": "XLSX.parse_relationships!",
    "category": "method",
    "text": "Parses package level relationships defined in _rels/.rels. Prases workbook level relationships defined in xl/_rels/workbook.xml.rels.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.parse_workbook!-Tuple{XLSX.XLSXFile}",
    "page": "API",
    "title": "XLSX.parse_workbook!",
    "category": "method",
    "text": "parse_workbook!(xf::XLSXFile)\n\nUpdates xf.workbook from xf.data[\"xl/workbook.xml\"]\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.readtable-Tuple{AbstractString,Union{Int64, AbstractString}}",
    "page": "API",
    "title": "XLSX.readtable",
    "category": "method",
    "text": "readtable(filepath, sheet, [columns]; [first_row], [column_labels], [header], [infer_eltypes], [stop_in_empty_row], [stop_in_row_function]) -> data, column_labels\n\nReturns tabular data from a spreadsheet as a tuple (data, column_labels). data is a vector of columns. column_labels is a vector of symbols. Use this function to create a DataFrame from package DataFrames.jl.\n\nUse columns argument to specify which columns to get. For example, columns=\"B:D\" will select columns B, C and D. If columns is not given, the algorithm will find the first sequence of consecutive non-empty cells.\n\nUse first_row to indicate the first row from the table. first_row=5 will look for a table starting at sheet row 5. If first_row is not given, the algorithm will look for the first non-empty row in the spreadsheet.\n\nheader is a Bool indicating if the first row is a header. If header=true and column_labels is not specified, the column labels for the table will be read from the first row of the table. If header=false and column_labels is not specified, the algorithm will generate column labels. The default value is header=true.\n\nUse column_labels as a vector of symbols to specify names for the header of the table.\n\nUse infer_eltypes=true to get data as a Vector{Any} of typed vectors. The default value is infer_eltypes=false.\n\nstop_in_empty_row is a boolean indicating wether an empty row marks the end of the table. If stop_in_empty_row=false, the TableRowIterator will continue to fetch rows until there\'s no more rows in the Worksheet. The default behavior is stop_in_empty_row=true.\n\nstop_in_row_function is a Function that receives a TableRow and returns a Bool indicating if the end of the table was reached.\n\nExample for stop_in_row_function:\n\nfunction stop_function(r)\n    v = r[:col_label]\n    return !ismissing(v) && v == \"unwanted value\"\nend\n\nRows where all column values are equal to missing are dropped.\n\nExample code for readtable:\n\njulia> using DataFrames, XLSX\n\njulia> df = DataFrame(XLSX.readtable(\"myfile.xlsx\", \"mysheet\")...)\n\nSee also: `gettable`.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.readxlsx-Tuple{AbstractString}",
    "page": "API",
    "title": "XLSX.readxlsx",
    "category": "method",
    "text": "readxlsx(filepath) :: XLSXFile\n\nMain function for reading an Excel file. This function will read the whole Excel file into memory and return a closed XLSXFile.\n\nConsider using openxlsx for lazy loading of Excel file contents.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.relative_cell_position-Tuple{XLSX.CellRef,XLSX.CellRange}",
    "page": "API",
    "title": "XLSX.relative_cell_position",
    "category": "method",
    "text": "Returns (row, column) representing a ref position relative to rng.\n\nFor example, for a range \"B2:D4\", we have:\n\n\"C3\" relative position is (2, 2)\n\"B2\" relative position is (1, 1)\n\"C4\" relative position is (3, 2)\n\"D4\" relative position is (3, 3)\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.row_number-Tuple{XLSX.CellRef}",
    "page": "API",
    "title": "XLSX.row_number",
    "category": "method",
    "text": "row_number(c::CellRef) :: Int\n\nReturns the row number of a given cell reference.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.sheet_column_numbers-Tuple{XLSX.Index}",
    "page": "API",
    "title": "XLSX.sheet_column_numbers",
    "category": "method",
    "text": "Returns real sheet column numbers (based on cellref)\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.sheetcount-Tuple{XLSX.Workbook}",
    "page": "API",
    "title": "XLSX.sheetcount",
    "category": "method",
    "text": "Counts the number of sheets in the Workbook.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.sheetnames-Tuple{XLSX.Workbook}",
    "page": "API",
    "title": "XLSX.sheetnames",
    "category": "method",
    "text": "Lists Worksheet names for this Workbook.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.split_cellname-Tuple{AbstractString}",
    "page": "API",
    "title": "XLSX.split_cellname",
    "category": "method",
    "text": "split_cellname(n::AbstractString) -> column_name, row_number\n\nSplits a string representing a cell name to its column name and row number.\n\nExample\n\njulia> XLSX.split_cellname(\"AB:12\")\n(\"AB:\", 12)\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.split_cellrange-Tuple{AbstractString}",
    "page": "API",
    "title": "XLSX.split_cellrange",
    "category": "method",
    "text": "split_cellrange(n::AbstractString) -> start_name, stop_name\n\nSplits a string representing a cell range into its cell names.\n\nExample\n\njulia> XLSX.split_cellrange(\"AB12:CD24\")\n(\"AB12\", \"CD24\")\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.split_column_range-Tuple{AbstractString}",
    "page": "API",
    "title": "XLSX.split_column_range",
    "category": "method",
    "text": "Returns tuple (columnnamestart, columnnamestop).\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.sst_formatted_string-Tuple{XLSX.Workbook,Int64}",
    "page": "API",
    "title": "XLSX.sst_formatted_string",
    "category": "method",
    "text": "sst_formatted_string(wb, index) :: String\n\nLooks for a formatted string inside the Shared Strings Table (sst). index starts at 0.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.sst_unformatted_string-Tuple{XLSX.Workbook,Int64}",
    "page": "API",
    "title": "XLSX.sst_unformatted_string",
    "category": "method",
    "text": "sst_unformatted_string(wb, index) :: String\n\nLooks for a string inside the Shared Strings Table (sst). index starts at 0.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.styles_add_font-Tuple{XLSX.Workbook,Array{Union{Pair{String,Pair{String,String}}, AbstractString},1}}",
    "page": "API",
    "title": "XLSX.styles_add_font",
    "category": "method",
    "text": "Defines a custom font. Returns the index to be used as the fontId in a cellXf definition.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.styles_add_numFmt-Tuple{XLSX.Workbook,AbstractString}",
    "page": "API",
    "title": "XLSX.styles_add_numFmt",
    "category": "method",
    "text": "Defines a custom number format to render numbers, dates or text. Returns the index to be used as the numFmtId in a cellXf definition.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.styles_cell_xf-Tuple{XLSX.Workbook,Int64}",
    "page": "API",
    "title": "XLSX.styles_cell_xf",
    "category": "method",
    "text": "Returns the xf XML node element for style index. index is 0-based.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.styles_cell_xf_numFmtId-Tuple{XLSX.Workbook,Int64}",
    "page": "API",
    "title": "XLSX.styles_cell_xf_numFmtId",
    "category": "method",
    "text": "Queries numFmtId from cellXfs -> xf nodes.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.styles_get_cellXf_with_numFmtId-Tuple{XLSX.Workbook,Int64}",
    "page": "API",
    "title": "XLSX.styles_get_cellXf_with_numFmtId",
    "category": "method",
    "text": "Cell Xf element follows the XML format below. This function queries the 0-based index of the first xf element that has the provided numFmtId. Returns -1 if not found.\n\n<styleSheet ...\n    <cellXfs count=\"5\">\n            <xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\"/>\n            <xf applyNumberFormat=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"14\" xfId=\"0\"/>\n            <xf applyNumberFormat=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"20\" xfId=\"0\"/>\n            <xf applyNumberFormat=\"1\" borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"22\" xfId=\"0\"/>\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.styles_numFmt_formatCode-Tuple{XLSX.Workbook,AbstractString}",
    "page": "API",
    "title": "XLSX.styles_numFmt_formatCode",
    "category": "method",
    "text": "Queries numFmt formatCode field by numFmtId.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.table_column_numbers-Tuple{XLSX.Index}",
    "page": "API",
    "title": "XLSX.table_column_numbers",
    "category": "method",
    "text": "Returns an iterator for table column numbers.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.table_column_to_sheet_column_number-Tuple{XLSX.Index,Int64}",
    "page": "API",
    "title": "XLSX.table_column_to_sheet_column_number",
    "category": "method",
    "text": "Maps table column index (1-based) -> sheet column index (cellref based)\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.unformatted_text-Tuple{EzXML.Node}",
    "page": "API",
    "title": "XLSX.unformatted_text",
    "category": "method",
    "text": "unformatted_text(el::EzXML.Node) :: String\n\nHelper function to gather unformatted text from Excel data files. It looks at all childs of el for tag name t and returns a join of all the strings found.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.writetable-Tuple{AbstractString,Any,Any}",
    "page": "API",
    "title": "XLSX.writetable",
    "category": "method",
    "text": "writetable(filename, data, columnnames; [overwrite], [sheetname])\n\ndata is a vector of columns. columnames is a vector of column labels. overwrite is a Bool to control if filename should be overwritten if already exists. sheetname is the name for the worksheet.\n\nExample using DataFrames.jl:\n\nimport DataFrames, XLSX\ndf = DataFrames.DataFrame(integers=[1, 2, 3, 4], strings=[\"Hey\", \"You\", \"Out\", \"There\"], floats=[10.2, 20.3, 30.4, 40.5])\nXLSX.writetable(\"df.xlsx\", DataFrames.columns(df), DataFrames.names(df))\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.writetable-Tuple{AbstractString}",
    "page": "API",
    "title": "XLSX.writetable",
    "category": "method",
    "text": "writetable(filename::AbstractString; overwrite::Bool=false, kw...)\nwritetable(filename::AbstractString, tables::Vector{Tuple{String, Vector{Any}, Vector{String}}}; overwrite::Bool=false)\n\nWrite multiple tables.\n\nkw is a variable keyword argument list. Each element should be in this format: sheetname=( data, column_names ), where data is a vector of columns and column_names is a vector of column labels.\n\nExample:\n\nimport DataFrames, XLSX\n\ndf1 = DataFrames.DataFrame(COL1=[10,20,30], COL2=[\"Fist\", \"Sec\", \"Third\"])\ndf2 = DataFrames.DataFrame(AA=[\"aa\", \"bb\"], AB=[10.1, 10.2])\n\nXLSX.writetable(\"report.xlsx\", REPORT_A=( DataFrames.columns(df1), DataFrames.names(df1) ), REPORT_B=( DataFrames.columns(df2), DataFrames.names(df2) ))\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.writexlsx-Tuple{AbstractString,XLSX.XLSXFile}",
    "page": "API",
    "title": "XLSX.writexlsx",
    "category": "method",
    "text": "writexlsx(output_filepath, xlsx_file; [overwrite=false])\n\nWrites an Excel file given by xlsx_file::XLSXFile to file at path output_filepath.\n\nIf overwrite=true, output_filepath will be overwritten if it exists.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.xlsx_encode-Tuple{XLSX.Worksheet,AbstractString}",
    "page": "API",
    "title": "XLSX.xlsx_encode",
    "category": "method",
    "text": "Returns the datatype and value for val to be inserted into ws.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.xmldocument-Tuple{XLSX.XLSXFile,String}",
    "page": "API",
    "title": "XLSX.xmldocument",
    "category": "method",
    "text": "xmldocument(xl::XLSXFile, filename::String) :: EzXML.Document\n\nUtility method to find the XMLDocument associated with a given package filename. Returns xl.data[filename] if it exists. Throws an error if it doesn\'t.\n\n\n\n\n\n"
},

{
    "location": "api.html#XLSX.xmlroot-Tuple{XLSX.XLSXFile,String}",
    "page": "API",
    "title": "XLSX.xmlroot",
    "category": "method",
    "text": "xmlroot(xl::XLSXFile, filename::String) :: EzXML.Node\n\nUtility method to return the root element of a given XMLDocument from the package. Returns EzXML.root(xl.data[filename]) if it exists.\n\n\n\n\n\n"
},

{
    "location": "api.html#API-1",
    "page": "API",
    "title": "API",
    "category": "section",
    "text": "Modules = [XLSX]"
},

]}
