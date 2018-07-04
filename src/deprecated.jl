
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

function getdata(filepath::AbstractString, sheet::Union{AbstractString, Int}, ref)
    warn("`XLSX.getdata(filepath, sheet, ref)` is deprecated. Use `XLSX.readdata(filepath, sheet, ref)` instead.")
    return readdata(filepath, sheet, ref)
end

function getdata(filepath::AbstractString, sheetref::AbstractString)
    warn("`XLSX.getdata(filepath, sheetref)` is deprecated. Use `XLSX.readdata(filepath, sheetref)` instead.")
    return readdata(filepath, sheetref)
end

function gettable(filepath::AbstractString, sheet::Union{AbstractString, Int}; first_row::Int = 1, column_labels::Vector{Symbol}=Vector{Symbol}(), header::Bool=true, infer_eltypes::Bool=false, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Function, Void}=nothing, enable_cache::Bool=false)
    warn("`XLSX.gettable(filepath, sheet, ...)` is deprecated. Use `XLSX.readtable(filepath, sheet, ...)` instead.")
    return readtable(filepath, sheet, first_row=first_row, column_labels=column_labels, header=header, infer_eltypes=infer_eltypes, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function, enable_cache=enable_cache)
end

function dimension(ws::Worksheet)
    warn("`XLSX.dimension(sheet)` is deprecated. Use `XLSX.get_dimension(sheet)` instead.")
    return get_dimension(ws)
end
