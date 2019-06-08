
# Tutorial

## Setup

First, make sure you have **XLSX.jl** package installed.

```julia
julia> using Pkg

julia> Pkg.add("XLSX")
```

## Getting Started

The basic usage is to read an Excel file and read values.

```julia
julia> import XLSX

julia> xf = XLSX.readxlsx("myfile.xlsx")
XLSXFile("myfile.xlsx") containing 3 Worksheets
            sheetname size          range
-------------------------------------------------
              mysheet 4x2           A1:B4
           othersheet 1x1           A1:A1
                named 1x1           B4:B4

julia> XLSX.sheetnames(xf)
3-element Array{String,1}:
 "mysheet"
 "othersheet"
 "named"

julia> sh = xf["mysheet"] # get a reference to a Worksheet
4×2 XLSX.Worksheet: ["mysheet"](A1:B4)

julia> sh[2, 2] # access element "B2" (2nd row, 2nd column)
"first"

julia> sh["B2"] # you can also use the cell name
"first"

julia> sh["A2:B4"] # or a cell range
3×2 Array{Any,2}:
 1  "first"
 2  "second"
 3  "third"

julia> XLSX.readdata("myfile.xlsx", "mysheet", "A2:B4") # shorthand for all above
3×2 Array{Any,2}:
 1  "first"
 2  "second"
 3  "third"

julia> sh[:] # all data inside worksheet's dimension
4×2 Array{Any,2}:
  "HeaderA"  "HeaderB"
 1           "first"
 2           "second"
 3           "third"

julia> xf["mysheet!A2:B4"] # you can also query values using a sheet reference
3×2 Array{Any,2}:
 1  "first"
 2  "second"
 3  "third"

julia> xf["NAMED_CELL"] # you can even read named ranges
"B4 is a named cell from sheet \"named\""

julia> xf["mysheet!A:B"] # Column ranges are also supported
4×2 Array{Any,2}:
  "HeaderA"  "HeaderB"
 1           "first"
 2           "second"
 3           "third"
```

To inspect the internal representation of each cell, use the `getcell` or `getcellrange` methods.

The example above used `xf = XLSX.readxlsx(filename)` to open a file, so all file contents are fetched at once from disk.

You can also use `XLSX.openxlsx` to read file contents as needed (see [Reading Large Excel Files and Caching](@ref)).

## Read Tabular Data

The `gettable` method returns tabular data from a spreadsheet as a tuple `(data, column_labels)`.
You can use it to create a `DataFrame` from [DataFrames.jl](https://github.com/JuliaData/DataFrames.jl).
Check the docstring for `gettable` method for more advanced options.

There's also a helper method `readtable` to read from file directly, as shown in the following example.
In this case, the [... operator](https://docs.julialang.org/en/v1/base/base/#...)
will splat the tuple `(data, column_labels)` into the constructor of `DataFrame`.

```julia
julia> using DataFrames, XLSX

julia> df = DataFrame(XLSX.readtable("myfile.xlsx", "mysheet")...)
3×2 DataFrames.DataFrame
│ Row │ HeaderA │ HeaderB  │
├─────┼─────────┼──────────┤
│ 1   │ 1       │ "first"  │
│ 2   │ 2       │ "second" │
│ 3   │ 3       │ "third"  │
```

## Reading Large Excel Files and Caching

The method `XLSX.openxlsx` has a `enable_cache` option to control worksheet cells caching.

Cache is enabled by default, so if you read a worksheet cell twice it will use the cached value instead of reading from disk
in the second time.

If `enable_cache=false`, worksheet cells will always be read from disk.
This is useful when you want to read a spreadsheet that doesn't fit into memory.

The following example shows how you would read worksheet cells, one row at a time,
where `myfile.xlsx` is a spreadsheet that doesn't fit into memory.

```julia
julia> XLSX.openxlsx("myfile.xlsx", enable_cache=false) do f
           sheet = f["mysheet"]
           for r in XLSX.eachrow(sheet)
              # r is a `SheetRow`, values are read using column references
              rn = XLSX.row_number(r) # `SheetRow` row number
              v1 = r[1]    # will read value at column 1
              v2 = r["B"]  # will read value at column 2

              println("v1=$v1, v2=$v2")
           end
      end
v1=HeaderA, v2=HeaderB
v1=1, v2=first
v1=2, v2=second
v1=3, v2=third
```

You could also stream tabular data using `XLSX.eachtablerow(sheet)`, which is the underlying iterator in `gettable` method.
Check docstrings for `XLSX.eachtablerow` for more advanced options.

```julia
julia> XLSX.openxlsx("myfile.xlsx", enable_cache=false) do f
           sheet = f["mysheet"]
           for r in XLSX.eachtablerow(sheet)
               # r is a `TableRow`, values are read using column labels or numbers
               rn = XLSX.row_number(r) # `TableRow` row number
               v1 = r[1] # will read value at table column 1
               v2 = r[:HeaderB] # will read value at column labeled `:HeaderB`

               println("v1=$v1, v2=$v2")
            end
       end
v1=1, v2=first
v1=2, v2=second
v1=3, v2=third
```

## Writing Excel Files

### Create New Files

Opening a file in `write` mode with `XLSX.openxlsx` will open a new (blank) Excel file for editing.

```julia
XLSX.openxlsx("my_new_file.xlsx", mode="w") do xf
    sheet = xf[1]
    XLSX.rename!(sheet, "new_sheet")
    sheet["A1"] = "this"
    sheet["A2"] = "is a"
    sheet["A3"] = "new file"
    sheet["A4"] = 100

    # will add a row from "A5" to "E5"
    sheet["A5"] = collect(1:5) # equivalent to `sheet["A5", dim=2] = collect(1:4)`

    # will add a column from "B1" to "B4"
    sheet["B1", dim=1] = collect(1:4)

    # will add a matrix from "A7" to "C9"
    sheet["A7:C9"] = [ 1 2 3 ; 4 5 6 ; 7 8 9 ]
end
```

### Edit Existing Files

Opening a file in `read-write` mode with `XLSX.openxlsx` will open an existing Excel file for editing.
This will preserve existing data in the original file.

```julia
XLSX.openxlsx("my_new_file.xlsx", mode="rw") do xf
    sheet = xf[1]
    sheet["B1"] = "new data"
end
```

### Export Tabular Data

Given a sheet reference, use the `XLSX.writetable!` method. Anchor cell defaults to cell `"A1"`.

```julia
using XLSX, Test

filename = "myfile.xlsx"

columns = Vector()
push!(columns, [1, 2, 3])
push!(columns, ["a", "b", "c"])

labels = [ "column_1", "column_2"]

XLSX.openxlsx(filename, mode="w") do xf
    sheet = xf[1]
    XLSX.writetable!(sheet, columns, labels, anchor_cell=XLSX.CellRef("B2"))
end

# read data back
XLSX.openxlsx(filename) do xf
    sheet = xf[1]
    @test sheet["B2"] == "column_1"
    @test sheet["C2"] == "column_2"
    @test sheet["B3"] == 1
    @test sheet["B4"] == 2
    @test sheet["B5"] == 3
    @test sheet["C3"] == "a"
    @test sheet["C4"] == "b"
    @test sheet["C5"] == "c"
end
```

You can also use `XLSX.writetable` to write directly to a new file (see next section).

### Export Tabular Data as a DataFrame

To export tabular data to Excel, use `XLSX.writetable` method.

```julia
julia> import DataFrames, XLSX

julia> df = DataFrames.DataFrame(integers=[1, 2, 3, 4], strings=["Hey", "You", "Out", "There"], floats=[10.2, 20.3, 30.4, 40.5], dates=[Date(2018,2,20), Date(2018,2,21), Date(2018,2,22), Date(2018,2,23)], times=[Dates.Time(19,10), Dates.Time(19,20), Dates.Time(19,30), Dates.Time(19,40)], datetimes=[Dates.DateTime(2018,5,20,19,10), Dates.DateTime(2018,5,20,19,20), Dates.DateTime(2018,5,20,19,30), Dates.DateTime(2018,5,20,19,40)])
4×6 DataFrames.DataFrame
│ Row │ integers │ strings │ floats │ dates      │ times    │ datetimes           │
├─────┼──────────┼─────────┼────────┼────────────┼──────────┼─────────────────────┤
│ 1   │ 1        │ Hey     │ 10.2   │ 2018-02-20 │ 19:10:00 │ 2018-05-20T19:10:00 │
│ 2   │ 2        │ You     │ 20.3   │ 2018-02-21 │ 19:20:00 │ 2018-05-20T19:20:00 │
│ 3   │ 3        │ Out     │ 30.4   │ 2018-02-22 │ 19:30:00 │ 2018-05-20T19:30:00 │
│ 4   │ 4        │ There   │ 40.5   │ 2018-02-23 │ 19:40:00 │ 2018-05-20T19:40:00 │

julia> XLSX.writetable("df.xlsx", collect(DataFrames.eachcol(df)), DataFrames.names(df))
```

You can also export multiple tables to Excel, each table in a separate worksheet.

```julia
julia> import DataFrames, XLSX

julia> df1 = DataFrames.DataFrame(COL1=[10,20,30], COL2=["Fist", "Sec", "Third"])
3×2 DataFrames.DataFrame
│ Row │ COL1 │ COL2  │
├─────┼──────┼───────┤
│ 1   │ 10   │ Fist  │
│ 2   │ 20   │ Sec   │
│ 3   │ 30   │ Third │

julia> df2 = DataFrames.DataFrame(AA=["aa", "bb"], AB=[10.1, 10.2])
2×2 DataFrames.DataFrame
│ Row │ AA │ AB   │
├─────┼────┼──────┤
│ 1   │ aa │ 10.1 │
│ 2   │ bb │ 10.2 │

julia> XLSX.writetable("report.xlsx", REPORT_A=( collect(DataFrames.eachcol(df1)), DataFrames.names(df1) ), REPORT_B=( collect(DataFrames.eachcol(df2)), DataFrames.names(df2) ))
```
