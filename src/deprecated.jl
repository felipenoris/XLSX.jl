
function read(s::AbstractString)
    warn("`XLSX.read(filepath)` is deprecated. Use `XLSX.readxlsx(filepath)` instead. Hint: also consider using `XLSX.openxlsx(filepath)` for loading contents on demand.")
    readxlsx(s)
end

function getcell(filepath::AbstractString, sheet::Union{AbstractString, Int}, ref)
	warn("`XLSX.getcell(filepath, sheet, ref)` is deprecated and will be removed.")
    xf = openxlsx(filepath, enable_cache=false)
    c = getcell(getsheet(xf, sheet), ref )
    close(xf)
    return c
end

function getcell(filepath::AbstractString, sheetref::AbstractString)
	warn("`XLSX.getcell(filepath, ref)` is deprecated and will be removed.")
    xf = openxlsx(filepath, enable_cache=false)
    c = getcell(xf, sheetref)
    close(xf)
    return c
end

function getcellrange(filepath::AbstractString, sheet::Union{AbstractString, Int}, rng)
	warn("`XLSX.getcellrange(filepath, sheet, range)` is deprecated and will be removed.")
    xf = openxlsx(filepath, enable_cache=false)
    c = getcellrange(getsheet(xf, sheet), rng )
    close(xf)
    return c
end

function getcellrange(filepath::AbstractString, sheetref::AbstractString)
	warn("`XLSX.getcellrange(filepath, ref) is deprecated and will be removed.`")
    xf = openxlsx(filepath, enable_cache=false)
    c = getcellrange(xf, sheetref)
    close(xf)
    return c
end
