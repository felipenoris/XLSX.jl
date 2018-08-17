
import XLSX
using Base.Test, Missings

data_directory = joinpath(Pkg.dir("XLSX"), "data")

ef_blank_ptbr_1904 = XLSX.readxlsx(joinpath(data_directory, "blank_ptbr_1904.xlsx"))
ef_Book1 = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
ef_Book_1904 = XLSX.readxlsx(joinpath(data_directory, "Book_1904.xlsx"))
ef_book_1904_ptbr = XLSX.readxlsx(joinpath(data_directory, "book_1904_ptbr.xlsx"))
ef_book_sparse = XLSX.readxlsx(joinpath(data_directory, "book_sparse.xlsx"))
ef_book_sparse_2 = XLSX.readxlsx(joinpath(data_directory, "book_sparse_2.xlsx"))

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
@test !XLSX.is_valid_cellname("A")
@test !XLSX.is_valid_cellname("1")
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

@test XLSX.is_valid_sheet_cellname("Sheet1!A2")
@test !XLSX.is_valid_sheet_cellname("Sheet1!A2:B3")
@test !XLSX.is_valid_sheet_cellname("Sheet1!A0")
@test !XLSX.is_valid_sheet_cellname("A1")
@test !XLSX.is_valid_sheet_cellname("Sheet1!")
@test !XLSX.is_valid_sheet_cellname("Sheet1")
@test !XLSX.is_valid_sheet_cellname("Sheet1!A")
@test !XLSX.is_valid_sheet_cellname("Sheet1!1")

@test XLSX.is_valid_sheet_cellrange("Sheet1!A1:B4")
@test !XLSX.is_valid_sheet_cellrange("Sheet1!:B4")
@test !XLSX.is_valid_sheet_cellrange("Sheet1!A1:")
@test !XLSX.is_valid_sheet_cellrange("Sheet1!:")
@test !XLSX.is_valid_sheet_cellrange("A1:B4")
@test !XLSX.is_valid_sheet_cellrange("Sheet1!")
@test !XLSX.is_valid_sheet_cellrange("Sheet1")
@test !XLSX.is_valid_sheet_cellrange("mysheet!A1")
@test XLSX.is_valid_sheet_cellrange("mysheet!A1:A4")

@test XLSX.is_valid_sheet_column_range("Sheet1!A:B")
@test XLSX.is_valid_sheet_column_range("Sheet1!AB:BC")
@test !XLSX.is_valid_sheet_column_range("A:B")
@test !XLSX.is_valid_sheet_column_range("Sheet1!")
@test !XLSX.is_valid_sheet_column_range("Sheet1")

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

v_column_numbers = [1
,    15
,    22
,    23
,    24
,    25
,    26
,    27
,    28
,    29
,    30
,    38
,    39
,    40
,    41
,    42
,    43
,    44
,    45
,    46
,    47
,    48
,    49
,    50
,    51
,    52
,    53
,    54
,    55
,    56
,    57
,    58
,    59
,    60
,    61
,    74
,    75
,    76
,    77
,    78
,    79
,    80
,    81
,    82
,    83
,    84
,    85
,    86
,   284
,   285
,   286
,   287
,   288
,   289
,   296
,   297
,   299
,   300
,   301
,   700
,   701
,   702
,   703
,   704
,   705
,   706
,   727
,   728
,   729
,   730
,   731
,  1008
,  1013
,  1014
,  1015
,  1016
,  1017
,  1018
,  1023
,  1024
,  1376
,  1377
,  1378
,  1379
,  1380
,  1381
,  3379
,  3380
,  3381
,  3382
,  3383
,  3403
,  3404
,  3405
,  3406
,  3407
, 16250
, 16251
, 16354
, 16355
, 16384]

v_column_names = [  "A"
, "O"
, "V"
, "W"
, "X"
, "Y"
, "Z"
, "AA"
, "AB"
, "AC"
, "AD"
, "AL"
, "AM"
, "AN"
, "AO"
, "AP"
, "AQ"
, "AR"
, "AS"
, "AT"
, "AU"
, "AV"
, "AW"
, "AX"
, "AY"
, "AZ"
, "BA"
, "BB"
, "BC"
, "BD"
, "BE"
, "BF"
, "BG"
, "BH"
, "BI"
, "BV"
, "BW"
, "BX"
, "BY"
, "BZ"
, "CA"
, "CB"
, "CC"
, "CD"
, "CE"
, "CF"
, "CG"
, "CH"
, "JX"
, "JY"
, "JZ"
, "KA"
, "KB"
, "KC"
, "KJ"
, "KK"
, "KM"
, "KN"
, "KO"
, "ZX"
, "ZY"
, "ZZ"
, "AAA"
, "AAB"
, "AAC"
, "AAD"
, "AAY"
, "AAZ"
, "ABA"
, "ABB"
, "ABC"
, "ALT"
, "ALY"
, "ALZ"
, "AMA"
, "AMB"
, "AMC"
, "AMD"
, "AMI"
, "AMJ"
, "AZX"
, "AZY"
, "AZZ"
, "BAA"
, "BAB"
, "BAC"
, "DYY"
, "DYZ"
, "DZA"
, "DZB"
, "DZC"
, "DZW"
, "DZX"
, "DZY"
, "DZZ"
, "EAA"
, "WZZ"
, "XAA"
, "XDZ"
, "XEA"
, "XFD"]

@assert length(v_column_names) == length(v_column_numbers) "Test script is wrong."

for i in 1:length(v_column_names)
    @test XLSX.encode_column_number(v_column_numbers[i]) == v_column_names[i]
    @test XLSX.decode_column_number(v_column_names[i]) == v_column_numbers[i]
end

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
@test XLSX.get_dimension(ef_Book1["Sheet1"]) == XLSX.range"B2:C8"
@test XLSX.isdate1904(ef_Book1["Sheet1"]) == false

# relative cell position
rng = XLSX.range"B2:D4"
@test XLSX.relative_cell_position(XLSX.ref"C3", rng) == (2, 2)
@test XLSX.relative_cell_position(XLSX.ref"B2", rng) == (1, 1)
@test XLSX.relative_cell_position(XLSX.ref"C4", rng) == (3, 2)
@test XLSX.relative_cell_position(XLSX.ref"D4", rng) == (3, 3)
@test XLSX.relative_cell_position(XLSX.EmptyCell(XLSX.ref"D4"), rng) == (3, 3)

# SheetCellRef, SheetCellRange, SheetColumnRange
ref = XLSX.SheetCellRef("Sheet1!A2")
@test string(ref) == "Sheet1!A2"
@test ref.sheet == "Sheet1"
@test ref.cellref == XLSX.CellRef("A2")
@test XLSX.SheetCellRef("Sheet1!A2") == XLSX.SheetCellRef("Sheet1!A2")
@test hash(XLSX.SheetCellRef("Sheet1!A2")) == hash(XLSX.SheetCellRef("Sheet1!A2"))

ref = XLSX.SheetCellRange("Sheet1!A1:B4")
@test ref.sheet == "Sheet1"
@test ref.rng == XLSX.CellRange("A1:B4")
@test_throws AssertionError XLSX.SheetCellRange("Sheet1!B4:A1")
@test XLSX.SheetCellRange("Sheet1!A1:B4") == XLSX.SheetCellRange("Sheet1!A1:B4")
@test hash(XLSX.SheetCellRange("Sheet1!A1:B4")) == hash(XLSX.SheetCellRange("Sheet1!A1:B4"))

ref = XLSX.SheetColumnRange("Sheet1!A:B")
@test string(ref) == "Sheet1!A:B"
@test ref.sheet == "Sheet1"
@test ref.colrng == XLSX.ColumnRange("A:B")
@test XLSX.SheetColumnRange("Sheet1!A:B") == XLSX.SheetColumnRange("Sheet1!A:B")
@test hash(XLSX.SheetColumnRange("Sheet1!A:B")) == hash(XLSX.SheetColumnRange("Sheet1!A:B"))

# getindex
f = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
sheet1 = f["Sheet1"]
@test sheet1["B2"] == "B2"
@test isapprox(sheet1["C3"], 21.2)
@test sheet1["B5"] == Date(2018, 3, 21)
@test sheet1["B8"] == "palavra1"

@test XLSX.getcell(sheet1, "B2") == XLSX.Cell(XLSX.CellRef("B2"), "s", "", "0", "")
XLSX.getcellrange(sheet1, "B2:C3")
XLSX.getcellrange(f, "Sheet1!B2:C3")
@test_throws ErrorException XLSX.getcellrange(f, "B2:C3")

# a cell can be put in a dict
c = XLSX.getcell(sheet1, "B2")
dct = Dict("a" => c)
@test dct["a"] == XLSX.Cell(XLSX.CellRef("B2"), "s", "", "0", "")

# equality and hash
@test XLSX.Cell(XLSX.CellRef("B2"), "s", "", "0", "") == XLSX.Cell(XLSX.CellRef("B2"), "s", "", "0", "")
@test hash(dct["a"]) == hash(XLSX.Cell(XLSX.CellRef("B2"), "s", "", "0", ""))

sheet2 = f[2]
sheet2_data = [ 1 2 3 ; 4 5 6 ; 7 8 9 ]
@test sheet2_data == sheet2["A1:C3"]
@test sheet2_data == sheet2[:]
@test sheet2[:] == XLSX.getdata(sheet2)

# Time and DateTime
@test XLSX.excel_value_to_time(0.82291666666666663) == Dates.Time(Dates.Hour(19), Dates.Minute(45))
@test XLSX.time_to_excel_value( XLSX.excel_value_to_time(0.2) ) == 0.2
@test XLSX.excel_value_to_datetime(43206.805447106482, false) == Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51))
@test XLSX.excel_value_to_datetime(XLSX.datetime_to_excel_value(Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51)),false ), false ) == Date(2018, 4, 16) + Dates.Time(Dates.Hour(19), Dates.Minute(19), Dates.Second(51))


dt = Date(2018,4,1)
@test XLSX.excel_value_to_date(XLSX.date_to_excel_value( dt, false), false) == dt
@test XLSX.excel_value_to_date(XLSX.date_to_excel_value( dt, true), true) == dt

# General number formats
f = XLSX.openxlsx(joinpath(data_directory, "general.xlsx"))
show(IOBuffer(), f)
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
@test f["general!B7"] == -220.0
@test f["general!B8"] == -2000
@test f["general!B9"] == 100000000000000
@test f["general!B10"] == -100000000000000

# Defined Names
@test XLSX.is_defined_name_value_a_reference(XLSX.SheetCellRef("Sheet1!A1"))
@test XLSX.is_defined_name_value_a_reference(XLSX.SheetCellRange("Sheet1!A1:B2"))
@test !XLSX.is_defined_name_value_a_reference(1)
@test !XLSX.is_defined_name_value_a_reference(1.2)
@test !XLSX.is_defined_name_value_a_reference("Hey")
@test !XLSX.is_defined_name_value_a_reference(Missings.missing)

@test f["SINGLE_CELL"] == "single cell A2"
@test f["RANGE_B4C5"] == Any["range B4:C5" "range B4:C5"; "range B4:C5" "range B4:C5"]
@test f["CONST_DATE"] == 43383
@test isapprox(f["CONST_FLOAT"], 10.2)
@test f["CONST_INT"] == 100
@test f["LOCAL_INT"] == 2000
@test f["named_ranges_2"]["LOCAL_INT"] == 2000
@test f["named_ranges"]["LOCAL_INT"] == 1000
@test f["named_ranges"]["LOCAL_NAME"] == "Hey You"
@test f["named_ranges_2"]["LOCAL_NAME"] == "out there in the cold"
@test f["named_ranges"]["SINGLE_CELL"] == "single cell A2"

@test_throws ErrorException f["header_error"]["LOCAL_REF"]
@test f["named_ranges"]["LOCAL_REF"][1] == 10
@test f["named_ranges"]["LOCAL_REF"][2] == 20
@test f["named_ranges_2"]["LOCAL_REF"][1] == "local"
@test f["named_ranges_2"]["LOCAL_REF"][2] == "reference"

close(f)

@test XLSX.readdata(joinpath(data_directory, "general.xlsx"), "SINGLE_CELL") == "single cell A2"
@test XLSX.readdata(joinpath(data_directory, "general.xlsx"), "RANGE_B4C5") == Any["range B4:C5" "range B4:C5"; "range B4:C5" "range B4:C5"]

# Book1.xlsx
f = XLSX.readxlsx(joinpath(data_directory, "Book1.xlsx"))
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

@test XLSX.getdata(f, XLSX.SheetCellRef("Sheet1!B2")) == "B2"
@test XLSX.getdata(f, XLSX.SheetCellRange("Sheet1!B2:B3"))[1] == "B2"
@test XLSX.getdata(f, XLSX.SheetCellRange("Sheet1!B2:B3"))[2] == 10.5
@test f["Sheet1!B2"] == "B2"
@test f["Sheet1!B2:B3"][1] == "B2"
@test f["Sheet1!B2:B3"][2] == 10.5
@test string(XLSX.SheetCellRange("Sheet1!B2:B3")) == "Sheet1!B2:B3"

# book_1904_ptbr.xlsx
f = XLSX.readxlsx(joinpath(data_directory, "book_1904_ptbr.xlsx"))

@test f["Plan1"][:] == Any[ "Coluna A" "Coluna B" "Coluna C" "Coluna D";
                            10 10.5 Date(2018, 3, 22) "linha 2";
                            20 20.5 Date(2017, 12, 31) "linha 3";
                            30 30.5 Date(2018, 1, 1) "linha 4" ]

@test f["Plan2"]["A1"] == "Merge de A1:D1"
@test ismissing(f["Plan2"]["B1"])
@test f["Plan2"]["C2"] == "C2"
@test f["Plan2"]["D3"] == "D3"

# numbers.xlsx
f = XLSX.readxlsx(joinpath(data_directory, "numbers.xlsx"))
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
@test XLSX.iserror(XLSX.getcell(f, "error!A2"))
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
@test string(cr) == "B:D"
@test cr.start == 2
@test cr.stop == 4
@test length(cr) == 3
@test_throws AssertionError XLSX.ColumnRange("B1:D3")
@test_throws AssertionError XLSX.ColumnRange("D:A")
@test collect(cr) == [ "B", "C", "D" ]
@test XLSX.ColumnRange("B:D") == XLSX.ColumnRange("B:D")
@test hash(XLSX.ColumnRange("B:D")) == hash(XLSX.ColumnRange("B:D"))

# CellRange iterator
rng = XLSX.CellRange("A2:C4")
@test collect(rng) == [ XLSX.CellRef("A2"), XLSX.CellRef("B2"), XLSX.CellRef("C2"), XLSX.CellRef("A3"), XLSX.CellRef("B3"), XLSX.CellRef("C3"), XLSX.CellRef("A4"), XLSX.CellRef("B4"), XLSX.CellRef("C4") ]

# Table

f = XLSX.readxlsx(joinpath(data_directory, "book_sparse.xlsx"))
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

f = XLSX.openxlsx(joinpath(data_directory, "general.xlsx"))
f["general"][:];
close(f)

f = XLSX.readxlsx(joinpath(data_directory, "general.xlsx"))
s = f["table"]
s[:];
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
        if ismissing(test_value) || ( isa(test_value, AbstractString) && isempty(test_value) )
            @test ismissing(value) || ( isa(value, AbstractString) && isempty(value) )
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

function stop_function(r::XLSX.TableRow)
    v = r[Symbol("Column C")]
    return !Missings.ismissing(v) && v == "Str2"
end

function never_reaches_stop(r::XLSX.TableRow)
    v = r[Symbol("Column C")]
    return !Missings.ismissing(v) && v == "never was found"
end

data, col_names = XLSX.gettable(s, stop_in_row_function=never_reaches_stop)
check_test_data(data, test_data)

data, col_names = XLSX.gettable(s, stop_in_row_function=stop_function)
@test col_names == [ Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G")]

test_data = Vector{Any}(6)
test_data[1] = collect(1:4)
test_data[2] = [ "Str1", missing, "Str1", "Str1" ]
test_data[3] = [ Date(2018, 4, 21) + Dates.Day(i) for i in 0:3 ]
test_data[4] = [ missing, missing, missing, missing ]
test_data[5] = [ 0.2001132319, 0.2793987377, 0.0950591677, 0.0744023067 ]
test_data[6] = [ missing for i in 1:4 ]

check_test_data(data, test_data)

data, col_names = XLSX.gettable(s, stop_in_empty_row=false)
@test col_names == [ Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G")]

test_data = Vector{Any}(6)
test_data[1] = [1, 2, 3, 4, 5, 6, 7, 8, "trash" ]
test_data[2] = [ "Str1", missing, "Str1", "Str1", "Str2", "Str2", "Str2", "Str2", missing ]
test_data[3] = Any[ Date(2018, 4, 21) + Dates.Day(i) for i in 0:7 ]
push!(test_data[3], "trash")
test_data[4] = [ missing, missing, missing, missing, missing, "a", "b", missing, missing ]
test_data[5] = [ 0.2001132319, 0.2793987377, 0.0950591677, 0.0744023067, 0.8242278091, 0.6205883578, 0.9174151018, 0.6749604883, "trash" ]
test_data[6] = Any[ missing for i in 1:8 ]
push!(test_data[6], "trash")

check_test_data(data, test_data)

# queries based on ColumnRange
x = XLSX.getcellrange(s, XLSX.ColumnRange("B:D"))
@test size(x) == (11, 3)
y = XLSX.getcellrange(s, "B:D")
@test size(y) == (11, 3)
@test x == y
@test_throws AssertionError XLSX.getcellrange(s, "D:B")
@test_throws ErrorException XLSX.getcellrange(s, "A:C1")

d = XLSX.getdata(s, "B:D")
@test size(d) == (11, 3)
@test_throws ErrorException XLSX.getdata(s, "A:C1")
@test d[1, 1] == "Column B"
@test d[1, 2] == "Column C"
@test d[1, 3] == "Column D"
@test d[9, 1] == 8
@test d[9, 2] == "Str2"
@test d[9, 3] == Date(2018, 4, 28)
@test d[10, 1] == "trash"
@test ismissing(d[10, 2])
@test d[10, 3] == "trash"
@test ismissing(d[11, 1])
@test ismissing(d[11, 2])
@test ismissing(d[11, 3])

d2 = f["table!B:D"]
@test size(d) == size(d2)
@test all(d .=== d2)

@test_throws ErrorException f["table!B1:D"]
@test_throws AssertionError f["table!D:B"]

s = f["table2"]
test_data = Vector{Any}(3)
test_data[1] = [ "A1", "A2", "A3", missing ]
test_data[2] = [ "B1", "B2", missing, "B4"]
test_data[3] = [ missing, missing, missing, missing ]

data, col_names = XLSX.gettable(s)

@test col_names == [:HA, :HB, :HC]
check_test_data(data, test_data)

for (ri, rowdata) in enumerate(XLSX.eachtablerow(s))
    if ismissing(test_data[1][ri])
        @test ismissing(rowdata[:HA])
    else
        @test rowdata[:HA] == test_data[1][ri]
    end

    @test XLSX.table_columns_count(rowdata) == 3
    @test XLSX.row_number(rowdata) == ri
    @test XLSX.get_column_labels(rowdata) == col_names
    @test XLSX.get_column_label(rowdata, 1) == :HA
    @test XLSX.get_column_label(rowdata, 2) == :HB
    @test XLSX.get_column_label(rowdata, 3) == :HC

    @test_throws ErrorException XLSX.getdata(rowdata, :INVALID_COLUMN)
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
@test_throws ErrorException XLSX.find_row(XLSX.eachrow(s), 20)

for r in XLSX.eachrow(s)
    @test isempty(XLSX.getcell(r, "A"))
    @test XLSX.getdata(s, XLSX.getcell(r, "B")) == "H1"
    @test r[2] == "H1"
    @test r["B"] == "H1"
    break
end

@test XLSX._find_first_row_with_data(s, 5) == 5
@test_throws ErrorException XLSX._find_first_row_with_data(s, 7)

s = f["table4"]
data, col_names = XLSX.gettable(s)
@test col_names == [:H1, :H2, :H3]
check_test_data(data, test_data)

xf = XLSX.openxlsx(joinpath(data_directory, "general.xlsx"))
empty_sheet = XLSX.getsheet(xf, "empty")
@test_throws ErrorException XLSX.gettable(empty_sheet)
itr = XLSX.eachrow(empty_sheet)
@test_throws ErrorException XLSX.find_row(itr, 1)
@test_throws ErrorException XLSX.getsheet(xf, "invalid_sheet")
close(xf)

f = XLSX.readxlsx(joinpath(data_directory,"general.xlsx"))
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

@test XLSX.is_valid_fixed_sheet_cellname("named_ranges!\$A\$2")
@test XLSX.is_valid_fixed_sheet_cellrange("named_ranges!\$B\$4:\$C\$5")
@test !XLSX.is_valid_fixed_sheet_cellname("named_ranges!A2")
@test !XLSX.is_valid_fixed_sheet_cellrange("named_ranges!B4:C5")
@test XLSX.SheetCellRef("named_ranges!\$A\$2") == XLSX.SheetCellRef("named_ranges!A2")
@test XLSX.SheetCellRange("named_ranges!\$B\$4:\$C\$5") == XLSX.SheetCellRange("named_ranges!B4:C5")

# Consecutive passes should yield the same results
f = XLSX.openxlsx(joinpath(data_directory, "general.xlsx"))
sl = f["lookup"]
data, col_names = XLSX.gettable(sl)
@test col_names == [:ID, :NAME, :VALUE]
check_test_data(data, test_data)

data, col_names = XLSX.gettable(sl)
@test col_names == [:ID, :NAME, :VALUE]
check_test_data(data, test_data)

#
# Helper functions
#

test_data = Vector{Any}(3)
test_data[1] = [ missing, missing, "B5" ]
test_data[2] = [ "C3", missing, missing ]
test_data[3] = [ missing, "D4", missing ]

data, col_names = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4")
@test col_names == [:H1, :H2, :H3]
check_test_data(data, test_data)

@test XLSX.readdata(joinpath(data_directory, "general.xlsx"), "table4", "E12") == "H1"
test_data = Array{Any, 2}(2, 1)
test_data[1, 1] = "H2"
test_data[2, 1] = "C3"

@test XLSX.readdata(joinpath(data_directory, "general.xlsx"), "table4", "F12:F13") == test_data

#
# Write
#

f = XLSX.openxlsxtemplate(joinpath(data_directory, "general.xlsx"))
filename_copy = "general_copy.xlsx"
XLSX.writexlsx(filename_copy, f)
@test isfile(filename_copy)

f_copy = XLSX.readxlsx(filename_copy)

s = f_copy["table"]
s[:];
data, col_names = XLSX.gettable(s)
@test col_names == [ Symbol("Column B"), Symbol("Column C"), Symbol("Column D"), Symbol("Column E"), Symbol("Column F"), Symbol("Column G")]

test_data = Vector{Any}(6)
test_data[1] = collect(1:8)
test_data[2] = [ "Str1", missing, "Str1", "Str1", "Str2", "Str2", "Str2", "Str2" ]
test_data[3] = [ Date(2018, 4, 21) + Dates.Day(i) for i in 0:7 ]
test_data[4] = [ missing, missing, missing, missing, missing, "a", "b", missing ]
test_data[5] = [ 0.2001132319, 0.2793987377, 0.0950591677, 0.0744023067, 0.8242278091, 0.6205883578, 0.9174151018, 0.6749604883 ]
test_data[6] = [ missing for i in 1:8 ]
check_test_data(data, test_data)
rm(filename_copy)

# Edit cells
f = XLSX.openxlsxtemplate(joinpath(data_directory, "general.xlsx"))
s = f["general"]
@test_throws ErrorException s["A1"] = :sym
XLSX.rename!(s, "renamed_sheet")
s["A1"] = "Hey You!"
s["B1"] = "Out there in the cold..."
s["A2"] = "Getting lonely getting old..."
s["B2"] = "Can you feel me?"
s["A3"] = 1000
s["B3"] = 99.99

# create a new sheet
s = XLSX.addsheet!(f, "my_new_sheet_1")
s = XLSX.addsheet!(f, "my_new_sheet_2")
s["B1"] = "This is a new sheet"
s["B2"] = "This is a new sheet"
s = XLSX.addsheet!(f)
s["B1"] = "unnamed sheet"

XLSX.writexlsx("general_copy_2.xlsx", f, rewrite=true)
@test isfile("general_copy_2.xlsx")

f = XLSX.openxlsx("general_copy_2.xlsx")
s = f["renamed_sheet"]
@test s["A1"] == "Hey You!"
@test s["B1"] == "Out there in the cold..."
@test s["A2"] == "Getting lonely getting old..."
@test s["B2"] == "Can you feel me?"
@test s["A3"] == 1000
@test s["B3"] == 99.99
f["my_new_sheet_1"];
@test f["my_new_sheet_2"]["B1"] == "This is a new sheet"
@test f["my_new_sheet_2"]["B2"] == "This is a new sheet"
@test f["Sheet1"]["B1"] == "unnamed sheet"
close(f)
rm("general_copy_2.xlsx")

# export table from template
col_names = ["Integers", "Strings", "Floats", "Booleans", "Dates", "Times", "DateTimes"]
data = Vector{Any}(7)
data[1] = [1, 2, Missings.missing, 4]
data[2] = ["Hey", "You", "Out", "There"]
data[3] = [101.5, 102.5, Missings.missing, 104.5]
data[4] = [ true, false, Missings.missing, true]
data[5] = [ Date(2018, 2, 1), Date(2018, 3, 1), Date(2018,5,20), Date(2018, 6, 2)]
data[6] = [ Dates.Time(19, 10), Dates.Time(19, 20), Dates.Time(19, 30), Dates.Time(19, 40) ]
data[7] = [ Dates.DateTime(2018, 5, 20, 19, 10), Dates.DateTime(2018, 5, 20, 19, 20), Dates.DateTime(2018, 5, 20, 19, 30), Dates.DateTime(2018, 5, 20, 19, 40)]

XLSX.writetable("output_table.xlsx", data, col_names, rewrite=true, sheetname="report", anchor_cell="B2")
@test isfile("output_table.xlsx")

read_data, read_column_names = XLSX.readtable("output_table.xlsx", "report")
@test length(read_column_names) == length(col_names)
for c in 1:length(col_names)
    @test Symbol(col_names[c]) == read_column_names[c]
end
check_test_data(read_data, data)
rm("output_table.xlsx")

report_1_column_names = ["HEADER_A", "HEADER_B"]
report_1_data = Vector{Any}(2)
report_1_data[1] = [1, 2, 3]
report_1_data[2] = ["A", "B", ""]

report_2_column_names = ["COLUMN_A", "COLUMN_B"]
report_2_data = Vector{Any}(2)
report_2_data[1] = [Date(2017,2,1), Date(2018,2,1)]
report_2_data[2] = [10.2, 10.3]

XLSX.writetable("output_tables.xlsx", rewrite=true, REPORT_A=(report_1_data, report_1_column_names), REPORT_B=(report_2_data, report_2_column_names))
@test isfile("output_tables.xlsx")

data, labels = XLSX.readtable("output_tables.xlsx", "REPORT_A")
@test labels[1] == :HEADER_A
@test labels[2] == :HEADER_B
check_test_data(data, report_1_data)

data, labels = XLSX.readtable("output_tables.xlsx", "REPORT_B")
@test labels[1] == :COLUMN_A
@test labels[2] == :COLUMN_B
check_test_data(data, report_2_data)

XLSX.writetable("output_tables.xlsx", [ ("REPORT_A", report_1_data, report_1_column_names), ("REPORT_B", report_2_data, report_2_column_names) ], rewrite=true)

data, labels = XLSX.readtable("output_tables.xlsx", "REPORT_A")
@test labels[1] == :HEADER_A
@test labels[2] == :HEADER_B
check_test_data(data, report_1_data)

data, labels = XLSX.readtable("output_tables.xlsx", "REPORT_B")
@test labels[1] == :COLUMN_A
@test labels[2] == :COLUMN_B
check_test_data(data, report_2_data)

report_1_column_names = [:HEADER_A, :HEADER_B]
report_2_column_names = [:COLUMN_A, :COLUMN_B]

XLSX.writetable("output_tables.xlsx", [ ("REPORT_A", report_1_data, report_1_column_names), ("REPORT_B", report_2_data, report_2_column_names) ], rewrite=true)

data, labels = XLSX.readtable("output_tables.xlsx", "REPORT_A")
@test labels[1] == :HEADER_A
@test labels[2] == :HEADER_B
check_test_data(data, report_1_data)

data, labels = XLSX.readtable("output_tables.xlsx", "REPORT_B")
@test labels[1] == :COLUMN_A
@test labels[2] == :COLUMN_B
check_test_data(data, report_2_data)

rm("output_tables.xlsx")

using XLSX: CellValue, id, getcell, setdata!, CellRef
xfile = XLSX.open_default_template()
wb = XLSX.get_workbook(xfile)
sheet = xfile["Sheet1"]

datefmt = XLSX.styles_add_numFmt(wb, "yyyymmdd")
numfmt = XLSX.styles_add_numFmt(wb, "\$* \#,\#\#0.00;\$* (\#,\#\#0.00);\$* \"-\"??;[Magenta]@")

#Check format id numbers dont intersect with predefined formats or each other
@test datefmt == 164
@test numfmt == 165

font = XLSX.styles_add_font(wb, XLSX.FontAttribute["b", "sz"=>("val"=>"24")])
xroot = XLSX.styles_xmlroot(wb)
fontnodes = find(xroot, "/xpath:styleSheet/xpath:fonts/xpath:font", XLSX.SPREADSHEET_NAMESPACE_XPATH_ARG)
fontnode = fontnodes[font+1] # XML is zero indexed so we need to add 1 to get the right node

# Check the font was written correctly
@test string(fontnode) == "<font><b/><sz val=\"24\"/></font>"

textstyle = XLSX.styles_add_cell_xf(wb, Dict("applyFont"=>"true", "fontId"=>"$font"))
datestyle = XLSX.styles_add_cell_xf(wb, Dict("applyNumberFormat"=>"1", "numFmtId"=>"$datefmt"))
numstyle = XLSX.styles_add_cell_xf(wb, Dict("applyFont"=>"1", "applyNumberFormat"=>"1", "fontId"=>"$font", "numFmtId"=>"$numfmt"))

xf = XLSX.styles_get_cellXf_with_numFmtId(wb, 1000)
@test xf == XLSX.EmptyCellDataFormat()
@test isempty(xf)
@test id(xf) == ""

@test textstyle isa XLSX.CellDataFormat
@test !isempty(textstyle)
@test id(textstyle) == "1"

@test XLSX.styles_get_cellXf_with_numFmtId(wb, datefmt) == datestyle
@test XLSX.styles_numFmt_formatCode(wb, string(datefmt)) == "yyyymmdd"
@test datestyle isa XLSX.CellDataFormat
@test !isempty(datestyle)
@test id(datestyle) == "2"

@test XLSX.styles_get_cellXf_with_numFmtId(wb, numfmt) == numstyle
@test XLSX.styles_numFmt_formatCode(wb, string(numfmt)) == "\$* \#,\#\#0.00;\$* (\#,\#\#0.00);\$* &quot;-&quot;??;[Magenta]@"
@test numstyle isa XLSX.CellDataFormat
@test !isempty(numstyle)
@test id(numstyle) == "3"

setdata!(sheet, CellRef("A1"), CellValue(Date(2011, 10, 13), datestyle))
setdata!(sheet, CellRef("A2"), CellValue(1000, numstyle))
setdata!(sheet, CellRef("A3"), CellValue(1000.10, numstyle))
setdata!(sheet, CellRef("A4"), CellValue(-1000.10, numstyle))
setdata!(sheet, CellRef("A5"), CellValue(0, numstyle))
setdata!(sheet, CellRef("A6"), CellValue("hello", numstyle))
setdata!(sheet, CellRef("B1"), CellValue("hello world", textstyle))

@test sheet["A1"] == Date(2011, 10, 13)
cell = getcell(sheet, "A1")
@test cell.style == id(datestyle)
formatid = XLSX.styles_cell_xf_numFmtId(wb, parse(Int, cell.style))
@test formatid == datefmt

cellstyle = getcell(sheet, "A2").style
@test cellstyle == id(numstyle)
formatid = XLSX.styles_cell_xf_numFmtId(wb, parse(Int, cellstyle))
@test formatid == numfmt

@test sheet["A2"] == 1000
@test sheet["A3"] == 1000.10
@test XLSX.getcell(sheet, "A3").style == cellstyle
@test sheet["A4"] == -1000.10
@test XLSX.getcell(sheet, "A4").style == cellstyle
@test sheet["A5"] == 0
@test XLSX.getcell(sheet, "A5").style == cellstyle
@test sheet["A6"] == "hello"
@test XLSX.getcell(sheet, "A6").style == cellstyle

@test sheet["B1"] == "hello world"
@test XLSX.getcell(sheet, "B1").style == id(textstyle)

sheet["B2"] = CellValue("hello world", textstyle)
@test sheet["B2"] == "hello world"
@test XLSX.getcell(sheet, "B2").style == id(textstyle)

# Check CellDataFormat only works with CellValues
@test_throws MethodError XLSX.CellValue([1,2,3,4], textstyle)

close(xfile)

using XLSX: styles_is_datetime, styles_add_numFmt, styles_add_cell_xf
# Capitalized and single character numfmts
xfile = XLSX.open_default_template()
wb = XLSX.get_workbook(xfile)
sheet = xfile["Sheet1"]

fmt = styles_add_numFmt(wb, "yyyy m d")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test styles_is_datetime(wb, style)

fmt = styles_add_numFmt(wb, "h:m:s")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test styles_is_datetime(wb, style)

fmt = styles_add_numFmt(wb, "0.00")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test !styles_is_datetime(wb, style)
@test XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, "[red]yyyy m d")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test styles_is_datetime(wb, style)
fmt = styles_add_numFmt(wb, "[red] h:m:s")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test styles_is_datetime(wb, style)
fmt = styles_add_numFmt(wb, "[red] 0.00; [magenta] 0.00")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test !styles_is_datetime(wb, style)
@test XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, "YYYY M D")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test styles_is_datetime(wb, style)
fmt = styles_add_numFmt(wb, "H:M:S")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test styles_is_datetime(wb, style)

fmt = styles_add_numFmt(wb, "m")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test styles_is_datetime(wb, style)

fmt = styles_add_numFmt(wb, "M")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test styles_is_datetime(wb, style)

fmt = styles_add_numFmt(wb, "y")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test !styles_is_datetime(wb, style)

fmt = styles_add_numFmt(wb, "[s]")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test styles_is_datetime(wb, style)

fmt = styles_add_numFmt(wb, "am/pm")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test styles_is_datetime(wb, style)

fmt = styles_add_numFmt(wb, "a/p")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test styles_is_datetime(wb, style)

fmt = styles_add_numFmt(wb, "\"Monday\"")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test !styles_is_datetime(wb, style)
@test !XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, "0.00*m")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test !styles_is_datetime(wb, style)
@test XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, "0.00_m")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test !styles_is_datetime(wb, style)
@test XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, "0.00\\d")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test !styles_is_datetime(wb, style)
@test XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, "[red][>1.5]000")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test !XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, "0.#")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, "\"hello.\" ###")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test !XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, ".??")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, "#e+00")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, "0e00")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, "# ??/??")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, "*.00")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test !XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, "\\.00")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test !XLSX.styles_is_float(wb, style)

fmt = styles_add_numFmt(wb, "00_.")
style = styles_add_cell_xf(wb, Dict("numFmtId"=>"$fmt"))
@test !XLSX.styles_is_float(wb, style)

close(xfile)

#
# Write one row at a time
#

sheetname = "New Sheet"
filename = "test_file.xlsx"
if isfile(filename)
    rm(filename)
end

data = [
    1 "a" Date(2018, 1, 1);
    2 missing Date(2018, 1, 2);
    missing "c" Date(2018, 1, 3);
]

XLSX.openxlsx(filename, write=true) do xf
    sheet = xf[1]
    XLSX.rename!(sheet, sheetname)

    sheet["A1"] = data[1, :]
    sheet[2, :] = data[2, :]
    sheet[2, 1] = "test overwrite"
    sheet[3, 2:3] = data[3, 2:3]
end

@test isfile(filename)
XLSX.openxlsx(filename) do xf
    sheet = xf[sheetname]
    read_data = sheet[:]

    @test isequal(read_data[1, :], data[1, :])
    @test isequal(read_data[2, :], vcat(["test overwrite"], data[2, 2:end]))
    @test isequal(read_data[3, :], data[3, :])
end

# test overwrite file
@test isfile(filename)
new_data = [1 2 3;]
XLSX.openxlsx(filename, write=true, read=false) do xf
    sheet = xf[1]
    sheet[1, :] = new_data[1,:]
end

XLSX.openxlsx(filename) do xf
    sheet = xf[1]
    read_data = sheet[:]

    @test isequal(read_data, new_data)
end

# test edit file
XLSX.openxlsx(filename, write=true, read=true) do xf
    sheet=xf[1]
    sheet[1, 2] = "hello"
    sheet["B6"] = 5
end

XLSX.openxlsx(filename) do xf
    sheet = xf[1]
    read_data = sheet[:]

    @test read_data[1, 1] == new_data[1, 1]
    @test read_data[1, 2] == "hello"
    @test read_data[1, 3] == new_data[1, 3]
    @test read_data[6, 2] == 5
end

# test writing throws error if flag not set
XLSX.openxlsx(filename) do xf
    sheet = xf[1]
    @test_throws AssertionError sheet[1, 1] = "failure"
end

xf = XLSX.emptyfile("page")
@test XLSX.sheetnames(xf) == ["page"]
close(xf)

rm(filename)
