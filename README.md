
# XLSX.jl

[![License](http://img.shields.io/badge/license-MIT-brightgreen.svg?style=flat)](LICENSE)
[![Build Status](https://travis-ci.org/felipenoris/XLSX.jl.svg?branch=master)](https://travis-ci.org/felipenoris/XLSX.jl)
[![codecov.io](http://codecov.io/github/felipenoris/XLSX.jl/coverage.svg?branch=master)](http://codecov.io/github/felipenoris/XLSX.jl?branch=master)

Excel file parser written in pure Julia.

## Usage

```julia
julia> import XLSX

julia> xf = XLSX.read("Book1.xlsx")
XLSXFile("Book1.xlsx")

julia> XLSX.sheetnames(xf)
2-element Array{String,1}:
 "Sheet1"
 "Sheet2"

julia> sh = xf["Sheet1"]
XLSX.Worksheet: "Sheet1". Dimension: "B2:C8".

julia> sh["C3"] # access a cell value
21.2

julia> sh["B3:C4"] # access a range
2×2 Array{Any,2}:
 10.5          21.2
   2018-03-21    2018-03-22

julia> sh[:] # all data inside worksheet's dimension
7×2 Array{Any,2}:
     "B2"             "C2"
   10.5             21.2
     2018-03-21       2018-03-22
     2018-03-21       2018-03-22
 true            false
    1                2
     "palavra1"       "palavra2"

julia> XLSX.getdata(sh) # same as sh[:]
7×2 Array{Any,2}:
     "B2"             "C2"
   10.5             21.2
     2018-03-21       2018-03-22
     2018-03-21       2018-03-22
 true            false
    1                2
     "palavra1"       "palavra2"
```

## References

* [ECMA Open XML White Paper](https://www.ecma-international.org/news/TC45_current_work/OpenXML%20White%20Paper.pdf)

* [ECMA-376](https://www.ecma-international.org/publications/standards/Ecma-376.htm)

* [Excel file limits](https://support.office.com/en-gb/article/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)

## Alternative Packages

* [ExcelFiles.jl](https://github.com/davidanthoff/ExcelFiles.jl)

* [ExcelReaders.jl](https://github.com/davidanthoff/ExcelReaders.jl)

* [XLSXReader.jl](https://github.com/mpastell/XLSXReader.jl)

* [Taro.jl](https://github.com/aviks/Taro.jl)
