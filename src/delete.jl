#=
Delete a worksheet by using either its name or ID; this action will generate a new file 
containing the remaining worksheets without altering the original file. 
When a worksheet is removed, the data types of cells, cell formatting, 
and the relationships between worksheets will not be preserved in the new file.
=#

# https://github.com/felipenoris/XLSX.jl/issues/80
using XLSX
function delete_ws_by_name(file_path::AbstractString, sheet_name::AbstractString, new_file::AbstractString)
    wb = readxlsx(file_path)
    if hassheet(wb, sheet_name)
        deleted_sheet_id= getsheet(wb, sheet_name).sheetId
        sheet_count = sheetcount(wb)
        if sheet_count == 1
            openxlsx(new_file, mode="w") do xf
            end
            return "Because the file has just this sheet '$sheet_name', an empty file was created."
        end
        openxlsx(new_file, mode="w") do xf
        end
        names = sheetnames(wb)
        sheet_id_new_file = 1
        for id in 1:sheet_count
            if id == deleted_sheet_id
                continue
            end
            openxlsx(new_file, mode="rw") do xf
                if sheet_id_new_file == 1
                    if xf[sheet_id_new_file] != names[id]
                        rename!(xf[sheet_id_new_file], names[id])
                    end
                else
                    addsheet!(xf, names[id])
                end
                ws_dimension = get_dimension(wb[id])
                bottom = row_number(ws_dimension.stop)
                right = column_number(ws_dimension.stop)
                for r in 1:bottom
                    for c in 1:right
                        cell_data = getdata(wb[id], r, c)
                        setdata!(xf[sheet_id_new_file], r, c, cell_data)
                    end
                end
            end
            sheet_id_new_file += 1
        end
    else
        error("The Sheet '$sheet_name' is not exists in the file!")
    end    
end


function delete_ws_by_id(file_path::AbstractString, sheet_id::Int64, new_file::AbstractString)
    wb = readxlsx(file_path)
    sheet_count = sheetcount(wb)
    if sheet_id <= sheet_count
        if sheet_count == 1
            openxlsx(new_file, mode="w") do xf
            end
            return "Because the file has just this sheet '$sheet_id', an empty file was created."
        end
        openxlsx(new_file, mode="w") do xf
        end
        sheet_id_new_file = 1
        for id in 1:sheet_count
            if id == sheet_id
                continue
            end
            openxlsx(new_file, mode="rw") do xf
                if sheet_id_new_file == 1
                    if xf[sheet_id_new_file] != wb[id].name
                        rename!(xf[sheet_id_new_file], wb[id].name)
                    end
                else
                    addsheet!(xf, wb[id].name)
                end
                ws_dimension = get_dimension(wb[id])
                bottom = row_number(ws_dimension.stop)
                right = column_number(ws_dimension.stop)
                for r in 1:bottom
                    for c in 1:right
                        cell_data = getdata(wb[id], r, c)
                        setdata!(xf[sheet_id_new_file], r, c, cell_data)
                    end
                end
            end
            sheet_id_new_file += 1
        end
    else
        error("The Sheet '$sheet_id' is not exists in the file!")
    end
end