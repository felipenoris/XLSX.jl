
# XLSX.jl

## Introduction

**XLSX.jl** is a Julia package to read and write [Excel](https://products.office.com/excel) spreadsheet files.

Internally, an Excel XLSX file is just a [Zip](https://en.wikipedia.org/wiki/Zip_(file_format)) file with a set of [XML](https://en.wikipedia.org/wiki/XML) files inside.
The formats for these XML files are described in the [Standard ECMA-376](https://www.ecma-international.org/publications/standards/Ecma-376.htm).

This package follows the EMCA-376 to parse and generate XLSX files.

## Requirements

* Julia v1.0

* Linux, macOS or Windows.

## Installation

From a Julia session, run:

```julia
julia> using Pkg

julia> Pkg.add("XLSX")
```

## Source Code

The source code for this package is hosted at [https://github.com/felipenoris/XLSX.jl](https://github.com/felipenoris/XLSX.jl).

## License

The source code for the package **XLSX.jl** is licensed under the [MIT License](https://raw.githubusercontent.com/felipenoris/XLSX.jl/master/LICENSE).

## References

* [ECMA Open XML White Paper](https://www.ecma-international.org/news/TC45_current_work/OpenXML%20White%20Paper.pdf)

* [ECMA-376](https://www.ecma-international.org/publications/standards/Ecma-376.htm)

* [Excel file limits](https://support.office.com/en-gb/article/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)

## Alternative Packages

* [ExcelFiles.jl](https://github.com/davidanthoff/ExcelFiles.jl)

* [ExcelReaders.jl](https://github.com/davidanthoff/ExcelReaders.jl)

* [XLSXReader.jl](https://github.com/mpastell/XLSXReader.jl)

* [Taro.jl](https://github.com/aviks/Taro.jl)
