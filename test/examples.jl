
import XLSX

#
# examples from docstrings
#

data_directory = joinpath(dirname(@__FILE__), "..", "data")

v = XLSX.getdata(joinpath(data_directory, "myfile.xlsx"), "mysheet", "A1:B4")

f = XLSX.read(joinpath(data_directory, "myfile.xlsx"))

sheet = f["mysheet"]

sheet = XLSX.getsheet(joinpath(data_directory, "myfile.xlsx"), "mysheet")

cell = XLSX.getcell(sheet, "A1")

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

xf = XLSX.read(joinpath(data_directory, "myfile.xlsx"))
XLSX.sheetnames(xf)
sh = xf["mysheet"]
sh["B2"] # access a cell value
sh["A2:B4"] # access a range
XLSX.getdata(joinpath(data_directory, "myfile.xlsx"), "mysheet", "A2:B4") # shorthand for all above
xf["mysheet!A2:B4"] # you can also query values from a file reference
xf["NAMED_CELL"] # you can even read named ranges
sh[:] # all data inside worksheet's dimension
XLSX.getdata(sh) # same as sh[:]

using DataFrames, XLSX

df = DataFrame(XLSX.gettable(joinpath(data_directory, "myfile.xlsx"), "mysheet")...)
