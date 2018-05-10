
__precompile__(true)
module XLSX

import ZipFile, EzXML, Missings

# https://github.com/fhs/ZipFile.jl/issues/39
if !method_exists(Base.nb_available, Tuple{ZipFile.ReadableFile})
	Base.nb_available(f::ZipFile.ReadableFile) = f.uncompressedsize - f._pos
end

include("structs.jl")
include("cellref.jl")
include("iterator.jl")
include("sst.jl")
include("relationship.jl")
include("read.jl")
include("workbook.jl")
include("worksheet.jl")
include("cell.jl")
include("styles.jl")
include("helper.jl")
include("deprecated.jl")

end # module XLSX
