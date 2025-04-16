
# Formatting Guide

## Excel Formatting

Each cell in an Excel spreadsheet may refer to an Excel `style`. Multiple cells can 
refer to the same `style` and therefore have a uniform appearance. A `style` defines
the cell's `alignment` directly (as part of the `style` definition), but it may also 
refer to further formatting definitions for `font`, `fill`, `border`, `format`. 
Multiple `style`s may each refer to the same `fill` definition or the same `font` 
definition, etc and therefore share these formatting characteristics.
This hierarchy can be shown like this:

                `Cell`
                  │
               `Style` => `Alignment`
                  │
  ┌──────────┬────┴─────┬─────────┐
  │          │          │         │
`font`     `fill`    `border`  `format`

A family of setter functions is provided to set each of the format attributes Excel uses.
These are applied to cells, and the functions deal with the relationships between the 
individual attributes, the overarching `style` and the cell(s) themselves.

## Setting format attributes of a cell

Set the font attributes of a cell using [`XLSX.setFont`](@ref). For example, to set cells `A1` and 
`A5` in the `general` sheet of a workbook to specific values, use:

```julia

julia> using XLSX

julia> f=XLSX.opentemplate("general.xlsx")
XLSXFile("general.xlsx") containing 13 Worksheets
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

julia> s=f["general"]
10×6 XLSX.Worksheet: ["general"](A1:F10)

julia> XLSX.setFont(s, "A1"; name="Arial", size=24, color="blue", bold=true)
2

XLSX.setFont(s, "A5"; name="Arial", size=24, color="blue", bold=true)
2
```

The function returns the `fontId` that has been used to define this combination 
of attributes.

There are more `font` attributes that can be set. Setting attributes for a cell 
that already has some, merges the new attributes with the old. Thus:

```julia
julia> XLSX.setFont(s, "A5"; italic=true, under="double", bold=false)
3
```

will over-ride the `bold` setting and add a double underline and make the font 
italic. However, the color, font name and size will all remain unchanged. This 
combination of attributes is unique, so a new `fontId` has been created.

The other set attribute functions behave in similar ways. See [`XLSX.setBorder`](@ref), 
[`XLSX.setFill`](@ref), [`XLSX.setFormat`](@ref) and [`XLSX.setAlignment`](@ref).

## Indexing multiple cells at once

Each of the setter functions can be applied to multiple cells at once using cell-ranges, 
row- or column-ranges or non-contiguous ranges. Additionally, indexing can use integer
indices for rows and columns, vectors of index values, unit- or step-ranges. This makes 
it easy to apply formatting to many cells at once.

Thus, for example:

```julia

julia> using XLSX

julia> f=XLSX.newxlsx()
XLSXFile("C:\Users\Tim Gebbels\.julia\artifacts\c0b84c4a80d13f58b3409f4a77d4a11455b5609e\blank.xlsx") containing 1 Worksheet
            sheetname size          range        
-------------------------------------------------
               Sheet1 1x1           A1:A1

julia> s=f[1]
1×1 XLSX.Worksheet: ["Sheet1"](A1:A1)

julia> s[1:100, 1:100] = "" # Can't set format attributes on `EmptyCell`s. This simply sets them to `missing` instead.
""

julia> XLSX.setFont(s, "A1:CV100"; name="Arial", size=24, color="blue", bold=true)
-1                          # Returns -1 on a range because a single `fontId` is unlikely to be possible.

julia> XLSX.setBorder(s, "A1:CV100"; allsides = ["style" => "thin", "color" => "black"])
-1

julia> XLSX.setAlignment(s, [10, 50, 90], 1:100; wrapText=true) # wrap text in the specified rows.
-1

julia>  XLSX.setAlignment(s, 1:100, 2:2:100; rotation=90) # rotate text 90° every second column.
-1
```

It is even possible to use defined names to index these functions:

```julia

julia> XLSX.addDefinedName(s, "my_name", "A1,B20,C30") # Define a non-contiguous named range.
XLSX.DefinedNameValue(Sheet1!A1,Sheet1!B20,Sheet1!C30, Bool[1, 1, 1])

julia> XLSX.setFill(s, "my_name"; pattern="solid", fgColor="coral")
-1
```

When setting format attributes over a range of cells as decribed, the new attributes are merged 
with existing on a cell by cell basis. If you set the font name on a range of cells that previously 
all had different font colors, the color differences will persist even as the font name is applied 
to the range consistently.

## Setting uniform attributes

Sometime it is useful to be able to apply a fully consistent set of format attributes to a range of 
cells, over-riding any pre-existing differences. This is the purpose of the `setUniformAttribute` 
family of functions. These functions update the attributes of the first cell in the range and then 
apply the relevant attribute Id to the rest of the cells in the range. Thus:

```julia
julia> XLSX.setBorder(s, "A1:CV100"; allsides = ["style" => "thin", "color" => "black"]) # set every cell individually
-1

julia> XLSX.setUniformBorder(s, "A1:CV100"; allsides = ["color" => "green"], diagonal = ["direction"=>"both", "color"=>"red"])
2 # This is the `borderId` that has now been uniformly applied to every cell.
```

This updates the border color in cell A1 to be green and adds red diagonal lines across the cell. 
It then applies all the `font` attributes of cell A1 uniformly to all the other cells in the range, 
overriding their previous attributes.

All the format setter functions have `setUniformAttribute` versions, too. See [`XLSX.setUniformBorder`](@ref), 
[`XLSX.setUniformFill`](@ref), [`XLSX.setUniformFormat`](@ref) and [`XLSX.setUniformAlignment`](@ref).

It is possible to use each of these functions in turn to ensure every possible attribute is consistently 
applied to a range of cells. However, if perfect uniformity is required, then `setUniformStyle` is 
considerably more efficient. It will simply take the `styleID` of the first cell in the range and apply 
it uniformly to each cell in the range. This ensures that all of font, fill, border, format, and 
alignment are all completely consistent across the range:

```julia

julia> XLSX.setUniformStyle(s, "A1:CV100") # set all formatting attributes to be uniformly tha same as cell A1.
7    # this is the `styleId` that has now been applied to all cells in the range
```

## Copying formatting attributes

It is possible to use non-contiguous ranges to copy format attributes from any cell to any other cells, 
whether you are also updating the source cell's format or not.

```julia

julia> XLSX.setBorder(s, "BB50"; allsides = ["style" => "medium", "color" => "yellow"])
3 # Cell BB50 has the border format I want!

julia> XLSX.setUniformBorder(s, "BB50,A1:CV100") # Make cell BB50 the first (reference) cell in a non-contiguous range.
3

julia> XLSX.setUniformStyle(s, "BB50,A1:CV100") # Or if I want to apply all formatting attributes from BB50 to the range.
11
```

## Setting column width and row height

Two functions offer the ability to set the column width and row height within a worksheet. These can use 
all of the indexing options described above. For example:

```julia

julia> XLSX.setRowHeight(s, "A2:A5"; height=25) # Rows 1 to 5 (columns ignored)

julia> XLSX.setColumnWidth(s, 5:5:100; width=50) # Every 5th column.
```

Excel applies some padding to user specified widths and heights. The two functions described here attempt 
to do something similar but it is not an exact match to what Excel does. User specified row heights and 
column widths will therefore differ by a small amount from the values you would see setting the same 
widths in Excel itself.

## Working with Merged Cells

Worksheets may contain merged cells. XLSX.jl provides functions to identify the merged cells in a worksheet, 
to determine if a cell is part of a merged range and to dtermine the value of a merged cell range from any 
cell in that range.

```julia

julia> using XLSX

julia> f=XLSX.opentemplate("customXml.xlsx")
XLSXFile("customXml.xlsx") containing 2 Worksheets
            sheetname size          range        
-------------------------------------------------
              Mock-up 116x11        A1:K116
     Document History 17x3          A1:C17

julia> XLSX.getMergedCells(f[1])
25-element Vector{XLSX.CellRange}:
 D49:H49
 D72:J72
 F94:J94
 F96:J96
 F84:J84
 F86:J86
 D62:J63
 D51:J53
 D55:J60
 D92:J92
 D82:J82
 D74:J74
 D67:J68
 D47:H47
 D9:H9
 D11:G11
 D12:G12
 D14:E14
 D16:E16
 D32:F32
 D38:J38
 D34:J34
 D18:E18
 D20:E20
 D13:G13

julia> XLSX.isMergedCell(f[1], "D13")
true

julia> XLSX.isMergedCell(f[1], "H13")
false

julia> XLSX.getMergedBaseCell(f[1], "E18") # E18 is a merged cell. The base cell in the merged range is D18.
(baseCell = D18, baseValue = "Here") # The base cell in the merged range is D18 and it's value is "Here".
```

It is also possible to create new merged cells:

```julia

julia> XLSX.isMergedCell(f[1], "F5")
false

julia> XLSX.isMergedCell(f[1], "J8")
false

julia> XLSX.mergeCells(s, "F5:J8")

julia> s["F5"] = pi
π = 3.1415926535897...

julia> XLSX.isMergedCell(f[1], "J8")
true

julia> XLSX.isMergedCell(f[1], "F5")
true

julia> XLSX.getMergedBaseCell(f[1], "J8")
(baseCell = F5, baseValue = 3.141592653589793)
```

It is not allowed to create new merged cells that overlap at all with any existing merged cells.

!!! warning
    It is possible to write into a merged cell using `XLSX.jl`.

    ```julia

    julia> XLSX.isMergedCell(f[1], "J8")
    true

    julia> f[1]["J8"] = "This cell is merged"
    "This cell is merged"

    julia> XLSX.isMergedCell(f[1], "J8")
    true

    julia> XLSX.getMergedBaseCell(f[1], "J8")
    (baseCell = F5, baseValue = 3.141592653589793)

    julia> f[1]["J8"]
    "This cell is merged"                                                                           

    ```

    The cell remains merged, and this is how Excel will see it. The assigned cell value won't be 
    visible in Excel, but it can be referenced in a formula, etc. This is prevented in Excel 
    itself by the UI (unless some clever VBA indirection is used). There is currently no check 
    to prevent this in `XLSX.jl`. See [#241](https://github.com/felipenoris/XLSX.jl/issues/241)