
import XLSX
using Base.Test, Missings

ef_blank_ptbr_1904 = XLSX.read("blank_ptbr_1904.xlsx")
ef_Book1 = XLSX.read("Book1.xlsx")
ef_Book_1904 = XLSX.read("Book_1904.xlsx")
ef_book_1904_ptbr = XLSX.read("book_1904_ptbr.xlsx")
ef_book_sparse = XLSX.read("book_sparse.xlsx")
ef_book_sparse_2 = XLSX.read("book_sparse_2.xlsx")

@test ef_Book1.filepath == "Book1.xlsx"
@test length(keys(ef_Book1.data)) > 0

@test ef_Book_1904.filepath == "Book_1904.xlsx"
@test length(keys(ef_Book_1904.data)) > 0

@test !XLSX.isdate1904(ef_Book1)
@test XLSX.isdate1904(ef_Book_1904)
@test XLSX.isdate1904(ef_blank_ptbr_1904)
@test XLSX.isdate1904(ef_book_1904_ptbr)

@test XLSX.sheetnames(ef_Book1) == [ "Sheet1", "Sheet2" ]
@test XLSX.sheetcount(ef_Book1) == 2
@test ef_Book1["Sheet1"].name == "Sheet1"
@test ef_Book1[1].name == "Sheet1"

@test XLSX.unformatted_text(ef_Book1.workbook.sst[1]) == "B2"
@test XLSX.sst_unformatted_string(ef_Book1.workbook, 0) == "B2" # index is 0-based
@test XLSX.sst_unformatted_string(ef_Book1, 0) == "B2"
@test XLSX.sst_unformatted_string(ef_Book1, "0") == "B2"

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
@test cn.column_name == "A"
@test cn.row_number == 1
@test XLSX.row_number(cn) == 1
@test XLSX.column_number(cn) == 1

cn = XLSX.CellRef("AB1")
@test string(cn) == "AB1"
@test cn.column_name == "AB"
@test cn.row_number == 1
@test XLSX.row_number(cn) == 1
@test XLSX.column_number(cn) == 28

cn = XLSX.CellRef("AMI1")
@test string(cn) == "AMI1"
@test cn.column_name == "AMI"
@test cn.row_number == 1
@test XLSX.row_number(cn) == 1
@test XLSX.column_number(cn) == 1023

cn = XLSX.CellRef("XFD1048576")
@test string(cn) == "XFD1048576"
@test cn.column_name == "XFD"
@test cn.row_number == 1048576
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
@test !XLSX.is_valid_cellrange("Z10:A1") # start cell should be at the top left corner of the range
@test !XLSX.is_valid_cellrange("A10:A1")
@test !XLSX.is_valid_cellrange("Z1:A1")

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
f = XLSX.read("Book1.xlsx")
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
f = XLSX.read("general.xlsx")
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
f = XLSX.read("Book1.xlsx")
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
f = XLSX.XLSXFile("book_1904_ptbr.xlsx")

@test f["Plan1"][:] == Any[ "Coluna A" "Coluna B" "Coluna C" "Coluna D";
                            10 10.5 Date(2018, 3, 22) "linha 2";
                            20 20.5 Date(2017, 12, 31) "linha 3";
                            30 30.5 Date(2018, 1, 1) "linha 4" ]

@test f["Plan2"]["A1"] == "Merge de A1:D1"
@test ismissing(f["Plan2"]["B1"])
@test f["Plan2"]["C2"] == "C2"
@test f["Plan2"]["D3"] == "D3"
