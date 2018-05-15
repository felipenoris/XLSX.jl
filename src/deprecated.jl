
function read(s::AbstractString)
    warn("`XLSX.read(filepath)` is deprecated. Use `XLSX.readxlsx(filepath)` instead. Hint: also consider using `XLSX.openxlsx(filepath)` for loading contents on demand.")
    readxlsx(s)
end
