
import XLSX
using Base.Test, Missings

data_directory = joinpath(dirname(@__FILE__), "..", "data")

ef_blank_ptbr_1904 = XLSX.read(joinpath(data_directory, "blank_ptbr_1904.xlsx"))
ef_Book1 = XLSX.read(joinpath(data_directory, "Book1.xlsx"))
ef_Book_1904 = XLSX.read(joinpath(data_directory, "Book_1904.xlsx"))
ef_book_1904_ptbr = XLSX.read(joinpath(data_directory, "book_1904_ptbr.xlsx"))
ef_book_sparse = XLSX.read(joinpath(data_directory, "book_sparse.xlsx"))
ef_book_sparse_2 = XLSX.read(joinpath(data_directory, "book_sparse_2.xlsx"))

@test ef_Book1.filepath == joinpath(data_directory, "Book1.xlsx")
@test length(keys(ef_Book1.data)) > 0

@test ef_Book_1904.filepath == joinpath(data_directory, "Book_1904.xlsx")
@test length(keys(ef_Book_1904.data)) > 0

@test !XLSX.isdate1904(ef_Book1)
@test XLSX.isdate1904(ef_Book_1904)
@test XLSX.isdate1904(ef_blank_ptbr_1904)
@test XLSX.isdate1904(ef_book_1904_ptbr)

@test XLSX.sheetnames(ef_Book1) == [ "Sheet1", "Sheet2" ]
@test XLSX.sheetcount(ef_Book1) == 2
@test ef_Book1["Sheet1"].name == "Sheet1"
@test ef_Book1[1].name == "Sheet1"

@test XLSX.sst_unformatted_string(ef_Book1.workbook, 0) == "B2" # index is 0-based
@test XLSX.sst_unformatted_string(ef_Book1, 0) == "B2"
@test XLSX.sst_unformatted_string(ef_Book1, "0") == "B2"

@test_throws ErrorException XLSX.get_relationship_target_by_id(ef_Book1.workbook, "indalid_id")
@test_throws ErrorException XLSX.get_relationship_target_by_type(ef_Book1.workbook, "indalid_type")
@test !XLSX.has_relationship_by_type(ef_Book1.workbook, "invalid_type")

# Cell names
@test !XLSX.is_valid_cellname("A0")
@test XLSX.is_valid_cellname("A1")
@test XLSX.is_valid_cellname("XFD1048576")
@test !XLSX.is_valid_cellname("XFD1048577")
@test XLSX.is_valid_cellname("XFD1")
@test !XLSX.is_valid_cellname("ZFD1")
@test XLSX.is_valid_column_name("A")
@test XLSX.is_valid_column_name("AZ")
@test XLSX.is_valid_column_name("AAZ")
@test !XLSX.is_valid_column_name("AAAZ")
@test !XLSX.is_valid_column_name(":")
@test !XLSX.is_valid_column_name("É")

cn = XLSX.CellRef("A1")
@test string(cn) == "A1"
@test XLSX.column_name(cn) == "A"
@test XLSX.row_number(cn) == 1
@test XLSX.column_number(cn) == 1

cn = XLSX.CellRef("AB1")
@test string(cn) == "AB1"
@test XLSX.column_name(cn) == "AB"
@test XLSX.row_number(cn) == 1
@test XLSX.column_number(cn) == 28

cn = XLSX.CellRef("AMI1")
@test string(cn) == "AMI1"
@test XLSX.column_name(cn) == "AMI"
@test XLSX.row_number(cn) == 1
@test XLSX.column_number(cn) == 1023

cn = XLSX.CellRef("XFD1048576")
@test string(cn) == "XFD1048576"
@test XLSX.column_name(cn) == "XFD"
@test XLSX.row_number(cn) == 1048576
@test XLSX.column_number(cn) == 16384

@test XLSX.encode_column_number(1) == "A"
@test XLSX.encode_column_number(28) == "AB"
@test XLSX.encode_column_number(1023) == "AMI"
@test XLSX.encode_column_number(16384) == "XFD"

@test XLSX.CellRef(12, 2).name == "B12"

cr = XLSX.range"A1:C4"
@test string(cr) == "A1:C4"
@test XLSX.row_number(cr.start) == 1
@test XLSX.column_number(cr.start) == 1
@test XLSX.row_number(cr.stop) == 4
@test XLSX.column_number(cr.stop) == 3
@test size(cr) == (4, 3)

cr = XLSX.range"B2:C8"
@test XLSX.ref"B2" ∈ cr
@test XLSX.ref"B3" ∈ cr
@test XLSX.ref"C2" ∈ cr
@test XLSX.ref"C3" ∈ cr
@test XLSX.ref"C8" ∈ cr
@test XLSX.ref"A1" ∉ cr
@test XLSX.ref"C9" ∉ cr
@test XLSX.ref"D4" ∉ cr
@test size(cr) == (7, 2)

fullrng = XLSX.range"B2:E5"
@test fullrng ⊆ fullrng
@test XLSX.range"B3:D4" ⊆ fullrng
@test !issubset(XLSX.range"A1:E5", fullrng)

@test XLSX.is_valid_cellrange("B2:C8")
@test !XLSX.is_valid_cellrange("A:B")
@test_throws AssertionError XLSX.CellRange("Z10:A1")
@test_throws AssertionError XLSX.CellRange("Z1:A1")

# hashing and equality
@test XLSX.CellRef("AMI1") == XLSX.CellRef("AMI1")
@test hash(XLSX.CellRef("AMI1")) == hash(XLSX.CellRef("AMI1"))
@test XLSX.CellRange("A1:C4") == XLSX.CellRange("A1:C4")
@test hash(XLSX.CellRange("A1:C4")) == hash(XLSX.CellRange("A1:C4"))

# worksheet for book1
@test XLSX.dimension(ef_Book1["Sheet1"]) == XLSX.range"B2:C8"
@test XLSX.isdate1904(ef_Book1["Sheet1"]) == false

# relative cell position
rng = XLSX.range"B2:D4"
@test XLSX.relative_cell_position(XLSX.ref"C3", rng) == (2, 2)
@test XLSX.relative_cell_position(XLSX.ref"B2", rng) == (1, 1)
@test XLSX.relative_cell_position(XLSX.ref"C4", rng) == (3, 2)
@test XLSX.relative_cell_position(XLSX.ref"D4", rng) == (3, 3)

# getindex
f = XLSX.read(joinpath(data_directory, "Book1.xlsx"))
sheet1 = f["Sheet1"]
@test sheet1["B2"] == "B2"
@test isapprox(sheet1["C3"], 21.2)
@test sheet1["B5"] == Date(2018, 3, 21)
@test sheet1["B8"] == "palavra1"

XLSX.getcell(sheet1, "B2")
XLSX.getcellrange(sheet1, "B2:C3")

sheet2 = f[2]
sheet2_data = [ 1 2 3 ; 4 5 6 ; 7 8 9 ]
@test sheet2_data == sheet2["A1:C3"]
@test sheet2_data == sheet2[:]
@test sheet2[:] == XLSX.getdata(sheet2)

# Time and DateTime
@test XLSX._time(0.82291666666666663) == Dates.Time(Dates.Hour(19), Dates.Minute(45))
@test XLSX._datetime(43206.805447106482, false) == Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51))

# General number formats
f = XLSX.read(joinpath(data_directory, "general.xlsx"))
sheet = f["general"]
@test sheet["A1"] == "text"
@test sheet["B1"] == "regular text"
@test sheet["A2"] == "integer"
@test sheet["B2"] == 102
@test sheet["A3"] == "float"
@test isapprox(sheet["B3"], 102.2)
@test sheet["A4"] == "date"
@test sheet["B4"] == Date(1983, 4, 16)
@test sheet["A5"] == "hour"
@test sheet["B5"] == Dates.Time(Dates.Hour(19), Dates.Minute(45))
@test sheet["A6"] == "datetime"
@test sheet["B6"] == Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51))

# Book1.xlsx
f = XLSX.read(joinpath(data_directory, "Book1.xlsx"))
sheet = f["Sheet1"]
@test ismissing(sheet["A1"])
@test sheet["B2"] == "B2"
@test sheet["C2"] == "C2"
@test isapprox(sheet["B3"], 10.5)
@test isapprox(sheet["C3"], 21.2)
@test sheet["B4"] == Date(2018, 3, 21)
@test sheet["C4"] == Date(2018, 3, 22)
@test sheet["B5"] == Date(2018, 3, 21)
@test sheet["C5"] == Date(2018, 3, 22)
@test sheet["B6"] == true
@test sheet["C6"] == false
@test sheet["B7"] == 1
@test sheet["C7"] == 2
@test sheet["B8"] == "palavra1"
@test sheet["C8"] == "palavra2"

# book_1904_ptbr.xlsx
f = XLSX.read(joinpath(data_directory, "book_1904_ptbr.xlsx"))

@test f["Plan1"][:] == Any[ "Coluna A" "Coluna B" "Coluna C" "Coluna D";
                            10 10.5 Date(2018, 3, 22) "linha 2";
                            20 20.5 Date(2017, 12, 31) "linha 3";
                            30 30.5 Date(2018, 1, 1) "linha 4" ]

@test f["Plan2"]["A1"] == "Merge de A1:D1"
@test ismissing(f["Plan2"]["B1"])
@test f["Plan2"]["C2"] == "C2"
@test f["Plan2"]["D3"] == "D3"

# numbers.xlsx
f = XLSX.read(joinpath(data_directory, "numbers.xlsx"))
floats = f["float"][:]
for n in floats
    if !ismissing(n)
        @test isa(n, Float64)
    end
end

ints = f["int"][:]
for n in ints
    if !ismissing(n)
        @test isa(n, Int)
    end
end

error_sheet = f["error"]
@test error_sheet["A1"] == "errors"
@test !XLSX.iserror(XLSX.getcell(error_sheet, "A1"))
@test XLSX.iserror(XLSX.getcell(error_sheet, "A2"))
@test ismissing(error_sheet["A2"])
@test ismissing(error_sheet["A3"])
@test ismissing(error_sheet["A4"])
emptycell = XLSX.getcell(error_sheet, "B1")
@test !XLSX.iserror(emptycell)
@test ismissing(XLSX.getdata(error_sheet, emptycell))
@test XLSX.row_number(emptycell) == 1
@test XLSX.column_number(emptycell) == 2

# Column Range

cr = XLSX.ColumnRange("B:D")
@test cr.start == 2
@test cr.stop == 4
@test length(cr) == 3
@test_throws AssertionError XLSX.ColumnRange("B1:D3")
@test_throws AssertionError XLSX.ColumnRange("D:A")
@test collect(cr) == [ "B", "C", "D" ]

# CellRange iterator
rng = XLSX.CellRange("A2:C4")
@test collect(rng) == [ XLSX.CellRef("A2"), XLSX.CellRef("B2"), XLSX.CellRef("C2"), XLSX.CellRef("A3"), XLSX.CellRef("B3"), XLSX.CellRef("C3"), XLSX.CellRef("A4"), XLSX.CellRef("B4"), XLSX.CellRef("C4") ]

# Table

f = XLSX.read(joinpath(data_directory, "book_sparse.xlsx"))
s = f["Sheet1"]

report = Vector{String}()
for r in XLSX.eachrow(s)
    if !isempty(r)
        push!(report, string(XLSX.row_number(r), " - ", XLSX.column_bounds(r)))

        if XLSX.row_number(r) == 2
            @test XLSX.last_column_index(r, 2) == 2
        elseif XLSX.row_number(r) == 3
            @test XLSX.last_column_index(r, 3) == 4
        elseif XLSX.row_number(r) == 6
            @test XLSX.last_column_index(r, 1) == 4
            @test XLSX.last_column_index(r, 2) == 4
            @test XLSX.last_column_index(r, 3) == 4
            @test XLSX.last_column_index(r, 4) == 4
            @test_throws AssertionError XLSX.last_column_index(r, 5)
        elseif XLSX.row_number(r) == 9
            @test XLSX.last_column_index(r, 2) == 3
            @test XLSX.last_column_index(r, 3) == 3
            @test XLSX.last_column_index(r, 5) == 5
        end
    end
end
@test report == [ "2 - (2, 2)", "3 - (3, 4)", "6 - (1, 4)", "9 - (2, 5)"]

f = XLSX.read(joinpath(data_directory, "general.xlsx"))
s = f["table"]
data, col_names = XLSX.gettable(s)
@test col_names == [ Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G")]

test_data = Vector{Any}(6)
test_data[1] = collect(1:8)
test_data[2] = [ "Str1", missing, "Str1", "Str1", "Str2", "Str2", "Str2", "Str2" ]
test_data[3] = [ Date(2018, 4, 21) + Dates.Day(i) for i in 0:7 ]
test_data[4] = [ missing, missing, missing, missing, missing, "a", "b", missing ]
test_data[5] = [ 0.2001132319, 0.2793987377, 0.0950591677, 0.0744023067, 0.8242278091, 0.6205883578, 0.9174151018, 0.6749604883 ]
test_data[6] = [ missing for i in 1:8 ]

function check_test_data(data::Vector{Any}, test_data::Vector{Any})

    @test length(data) == length(test_data)

    function size_of_data(d::Vector{Any})
        isempty(d) && return (0, 0)
        return length(d[1]), length(d)
    end

    rows, cols = size_of_data(test_data)

    for col in 1:cols
        @test length(data[col]) == length(test_data[col])
    end

    for row in 1:rows, col in 1:cols
        test_value = test_data[col][row]
        value = data[col][row]
        if ismissing(test_value)
            @test ismissing(value)
        else
            if isa(test_value, Float64)
                @test isapprox(value, test_value)
            else
                @test value == test_value
            end
        end
    end

    nothing
end

check_test_data(data, test_data)

@test XLSX.infer_eltype(data[1]) == Int
@test XLSX.infer_eltype(data[2]) == Union{Missing, String}
@test XLSX.infer_eltype(data[3]) == Date
@test XLSX.infer_eltype(data[4]) == Union{Missing, String}
@test XLSX.infer_eltype(data[5]) == Float64
@test XLSX.infer_eltype(data[6]) == Any
@test XLSX.infer_eltype([1, "1", 10.2]) == Any
@test XLSX.infer_eltype(Vector{Int}()) == Int

data_inferred, col_names = XLSX.gettable(s, infer_eltypes=true)
@test eltype(data_inferred[1]) == Int
@test eltype(data_inferred[2]) == Union{Missing, String}
@test eltype(data_inferred[3]) == Date
@test eltype(data_inferred[4]) == Union{Missing, String}
@test eltype(data_inferred[5]) == Float64
@test eltype(data_inferred[6]) == Any

s = f["table2"]
test_data = Vector{Any}(3)
test_data[1] = [ "A1", "A2", "A3", missing ]
test_data[2] = [ "B1", "B2", missing, "B4"]
test_data[3] = [ missing, missing, missing, missing ]

data, col_names = XLSX.gettable(s)

@test col_names == [:HA, :HB, :HC]
check_test_data(data, test_data)

for (ri, rowdata) in enumerate(XLSX.TableRowIterator(s))
    if ismissing(test_data[1][ri])
        @test ismissing(rowdata[:HA])
    else
        @test rowdata[:HA] == test_data[1][ri]
    end
end

override_col_names = [:ColumnA, :ColumnB, :ColumnC]
data, col_names = XLSX.gettable(s, column_labels=override_col_names)

@test col_names == override_col_names
check_test_data(data, test_data)

data, col_names = XLSX.gettable(s, "A:B", first_row=1)
test_data_AB_cols = Vector{Any}(2)
test_data_AB_cols[1] = test_data[1]
test_data_AB_cols[2] = test_data[2]
@test col_names == [:HA, :HB]
check_test_data(data, test_data_AB_cols)

data, col_names = XLSX.gettable(s, "A:B")
test_data_AB_cols = Vector{Any}(2)
test_data_AB_cols[1] = test_data[1]
test_data_AB_cols[2] = test_data[2]
@test col_names == [:HA, :HB]
check_test_data(data, test_data_AB_cols)

data, col_names = XLSX.gettable(s, "B:B", first_row=2)
@test col_names == [:B1]
@test length(data) == 1
@test length(data[1]) == 1
@test data[1][1] == "B2"

data, col_names = XLSX.gettable(s, "B:C")
@test col_names == [ :HB, :HC ]
test_data_BC_cols = Vector{Any}(2)
test_data_BC_cols[1] = ["B1", "B2"]
test_data_BC_cols[2] = [ missing, missing]
check_test_data(data, test_data_BC_cols)

data, col_names = XLSX.gettable(s, "B:C", first_row=2, header=false)
@test col_names == [ :B, :C ]
check_test_data(data, test_data_BC_cols)

s = f["table3"]
test_data = Vector{Any}(3)
test_data[1] = [ missing, missing, "B5" ]
test_data[2] = [ "C3", missing, missing ]
test_data[3] = [ missing, "D4", missing ]
data, col_names = XLSX.gettable(s)
@test col_names == [:H1, :H2, :H3]
check_test_data(data, test_data)
@test_throws ErrorException XLSX.find_row(XLSX.SheetRowIterator(s), 20)

s = f["table4"]
data, col_names = XLSX.gettable(s)
@test col_names == [:H1, :H2, :H3]
check_test_data(data, test_data)

empty_sheet = XLSX.getsheet(joinpath(data_directory, "general.xlsx"), "empty")
@test_throws ErrorException XLSX.gettable(empty_sheet)
itr = XLSX.SheetRowIterator(empty_sheet)
@test_throws ErrorException XLSX.find_row(itr, 1)

f = XLSX.read(joinpath(data_directory,"general.xlsx"))
tb5 = f["table5"]
test_data = Vector{Any}(1)
test_data[1] = [1, 2, 3, 4, 5]
data, col_names = XLSX.gettable(tb5)
@test col_names == [ :HEADER ]
check_test_data(data, test_data)
tb6 = f["table6"]
data, col_names = XLSX.gettable(tb6, first_row=3)
@test col_names == [ :HEADER ]
check_test_data(data, test_data)
tb7 = f["table7"]
data, col_names = XLSX.gettable(tb7, first_row=3)
@test col_names == [ :HEADER ]
check_test_data(data, test_data)

sheet_lookup = f["lookup"]
test_data = Vector{Any}(3)
test_data[1] = [ 10, 20, 30]
test_data[2] = [ "name1", "name2", "name3" ]
test_data[3] = [ 100, 200, 300 ]
data, col_names = XLSX.gettable(sheet_lookup)
@test col_names == [:ID, :NAME, :VALUE]
check_test_data(data, test_data)

header_error_sheet = f["header_error"]
@test_throws AssertionError XLSX.gettable(header_error_sheet)

#
# Helper functions
#

test_data = Vector{Any}(3)
test_data[1] = [ missing, missing, "B5" ]
test_data[2] = [ "C3", missing, missing ]
test_data[3] = [ missing, "D4", missing ]

data, col_names = XLSX.gettable(joinpath(data_directory, "general.xlsx"), "table4")
@test col_names == [:H1, :H2, :H3]
check_test_data(data, test_data)

@test XLSX.getdata(joinpath(data_directory, "general.xlsx"), "table4", "E12") == "H1"
test_data = Array{Any, 2}(2, 1)
test_data[1, 1] = "H2"
test_data[2, 1] = "C3"
@test XLSX.getdata(joinpath(data_directory, "general.xlsx"), "table4", "F12:F13") == test_data
