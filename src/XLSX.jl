
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

end # module XLSX
