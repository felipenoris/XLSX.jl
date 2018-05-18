
import XLSX

#
# examples from docstrings
#

data_directory = joinpath(dirname(@__FILE__), "..", "data")

v = XLSX.getdata(joinpath(data_directory, "myfile.xlsx"), "mysheet", "A1:B4")

f = XLSX.openxlsx(joinpath(data_directory, "myfile.xlsx"))
sheet = f["mysheet"]
v = sheet["A1:B4"]
close(f)

xf = XLSX.openxlsx(joinpath(data_directory, "myfile.xlsx"))
sheet = xf["mysheet"]
cell = XLSX.getcell(sheet, "A1")
close(xf)

xf = XLSX.openxlsx(joinpath(data_directory, "myfile.xlsx"))
sheet = xf["mysheet"]
left = 1  # 1st column
right = 4 # 4th column
for sheetrow in XLSX.eachrow(sheet)
    for column in left:right
        cell = XLSX.getcell(sheetrow, column)

        # do something with cell
    end
end

using DataFrames, XLSX

df = DataFrame(XLSX.gettable(joinpath(data_directory, "myfile.xlsx"), "mysheet")...)

XLSX.decode_column_number("D")
XLSX.encode_column_number(4)

cn = XLSX.CellRef("AB1")
println( XLSX.row_number(cn) ) # will print 1
println( XLSX.column_number(cn) ) # will print 28
println( string(cn) ) # will print out AB1

cn = XLSX.ref"AB1"
println( XLSX.row_number(cn) ) # will print 1
println( XLSX.column_number(cn) ) # will print 28
println( string(cn) ) # will print out AB1
cr = XLSX.range"A1:C4"

XLSX.getcellrange(joinpath(data_directory, "myfile.xlsx"), "mysheet", "A1:B4")

#
# examples from README.md
#

xf = XLSX.openxlsx(joinpath(data_directory, "myfile.xlsx"))
XLSX.sheetnames(xf)
sh = xf["mysheet"]
sh["B2"] # access a cell value
sh["A2:B4"] # access a range
XLSX.getdata(joinpath(data_directory, "myfile.xlsx"), "mysheet", "A2:B4") # shorthand for all above
xf["mysheet!A2:B4"] # you can also query values from a file reference
xf["NAMED_CELL"] # you can even read named ranges
xf["mysheet!A:B"] # Column ranges are also supported
sh[:] # all data inside worksheet's dimension
XLSX.getdata(sh) # same as sh[:]
close(xf)

using DataFrames, XLSX

df = DataFrame(XLSX.gettable(joinpath(data_directory, "myfile.xlsx"), "mysheet")...)

f = XLSX.openxlsx(joinpath(data_directory, "myfile.xlsx"), enable_cache=false)
sheet = f["mysheet"]
for r in XLSX.eachrow(sheet)
    # r is a `SheetRow`, values are read using column references
    rn = XLSX.row_number(r) # `SheetRow` row number
    v1 = r[1]    # will read value at column 1
    v2 = r["B"]  # will read value at column 2
end

for r in XLSX.eachtablerow(sheet)
	# r is a `TableRow`, values are read using column labels or numbers
	rn = XLSX.row_number(r) # `TableRow` row number
	v1 = r[1] # will read value at table column 1
	v2 = r[:HeaderB] # will read value at column labeled `:HeaderB`
end
