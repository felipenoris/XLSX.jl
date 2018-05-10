
function read(s::AbstractString)
	warn("`XLSX.read(filepath)` is deprecated. Use `XLSX.readxlsx(filepath)`.")
	readxlsx(s)
end
