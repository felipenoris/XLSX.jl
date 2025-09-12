
module XLSX

import Artifacts
import Dates
import Printf.@printf
import ZipArchives
import XML
import Tables
import Unicode
import Colors
import Base.convert
import UUIDs
import Mmap
import Base.Threads

import PrecompileTools as PCT    # this is a small dependency.

export
    # Files and worksheets
    XLSXFile, readxlsx, openxlsx, opentemplate, newxlsx, writexlsx, savexlsx,
    Worksheet, sheetnames, sheetcount, hassheet, rename!, addsheet!, copysheet!, deletesheet!, 
    # Cells & data
    CellRef, row_number, column_number, eachrow, eachtablerow,
    readdata, getdata, gettable, readtable, readto, writetable, writetable!,
    addDefinedName,
    # Formats
    setFormat, setFont, setBorder, setFill, setAlignment,
    setUniformFormat, setUniformFont, setUniformBorder, setUniformFill, setUniformAlignment, setUniformStyle,
    setConditionalFormat,
    setColumnWidth, setRowHeight,
    getMergedCells, isMergedCell, getMergedBaseCell, mergeCells
    
const SPREADSHEET_NAMESPACE_XPATH_ARG = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const EXCEL_MAX_COLS = 16_384     # total columns supported by Excel per sheet
const EXCEL_MAX_ROWS = 1_048_576  # total rows supported by Excel per sheet (including headers)

include("types.jl")
include("formula.jl")
include("cellref.jl")
include("sst.jl")
include("stream.jl")
include("table.jl")
include("tables_interface.jl")
include("relationship.jl")
include("read.jl")
include("workbook.jl")
include("worksheet.jl")
include("cell.jl")
include("styles.jl")
include("cellformat-helpers.jl")
include("cellformats.jl")
include("conditional-format-helpers.jl") # must load before conditional-formats.jl
include("conditional-formats.jl")
include("write.jl")
include("fileArray.jl")

PCT.@setup_workload begin
    # Putting some things in `@setup_workload` instead of `@compile_workload` can reduce the size of the
    # precompile file and potentially make loading faster.
    s=IOBuffer()
    t=IOBuffer()
    PCT.@compile_workload begin
        # all calls in this block will be precompiled, regardless of whether
        # they belong to your package or not (on Julia 1.8 and higher)
        f=openxlsx(joinpath(_relocatable_data_path(), "blank.xlsx"), mode="rw")
        f[1]["A1:Z26"] = "hello World"
        openxlsx(s, mode="w") do xf
            xf[1][1:26, 1:26] = pi
        end
        _ = XLSX.readtable(seekstart(s), 1, "A:Z")
        f= openxlsx(seekstart(s), mode="rw")
        f[1][1:26, 1:26] = pi
        setConditionalFormat(f[1], :, :cellIs)
        setConditionalFormat(f[1], "A1:Z26", :colorScale)
        setBorder(f[1], collect(1:26), 1:26, allsides=["style"=>"thin", "color"=>"black"])
        _ = getdata(f[1], "A1:A20")
        writexlsx(t, f)
    end
end

end # module XLSX
