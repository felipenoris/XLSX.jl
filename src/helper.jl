
#
# Helper Functions
#

function getcell(filepath::AbstractString, sheet::Union{AbstractString, Int}, ref)
	xf = openxlsx(filepath, enable_cache=false)
	c = getcell(getsheet(xf, sheet), ref )
	close(xf)
	return c
end

function getcell(filepath::AbstractString, sheetref::AbstractString)
	xf = openxlsx(filepath, enable_cache=false)
	c = getcell(xf, sheetref)
	close(xf)
	return c
end

function getcellrange(filepath::AbstractString, sheet::Union{AbstractString, Int}, rng)
	xf = openxlsx(filepath, enable_cache=false)
	c = getcellrange(getsheet(xf, sheet), rng )
	close(xf)
	return c
end

function getcellrange(filepath::AbstractString, sheetref::AbstractString)
	xf = openxlsx(filepath, enable_cache=false)
	c = getcellrange(xf, sheetref)
	close(xf)
	return c
end

function getdata(filepath::AbstractString, sheet::Union{AbstractString, Int}, ref)
	xf = openxlsx(filepath, enable_cache=false)
	c = getdata(getsheet(xf, sheet), ref )
	close(xf)
	return c
end

function getdata(filepath::AbstractString, sheetref::AbstractString)
	xf = openxlsx(filepath, enable_cache=false)
	c = getdata(xf, sheetref)
	close(xf)
	return c
end

function gettable(filepath::AbstractString, sheet::Union{AbstractString, Int}; first_row::Int = 1, column_labels::Vector{Symbol}=Vector{Symbol}(), header::Bool=true, infer_eltypes::Bool=false, stop_in_empty_row::Bool=true, stop_in_row_function::Union{Function, Void}=nothing, enable_cache::Bool=false)
	xf = openxlsx(filepath, enable_cache=enable_cache)
	c = gettable(getsheet(xf, sheet); first_row=first_row, column_labels=column_labels, header=header, infer_eltypes=infer_eltypes, stop_in_empty_row=stop_in_empty_row, stop_in_row_function=stop_in_row_function)
	close(xf)
	return c
end
