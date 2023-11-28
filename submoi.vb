Private Sub btnXuatNhanQLTS()
    Dim sheetnameBangTheoDoiTS As String
    Dim WorksheetTS As Worksheet

    sheetnameBangTheoDoiTS = "BANG THEO DOI TS"
    Set WorksheetTS = ThisWorkbook.Worksheets(sheetnameBangTheoDoiTS)

    Dim lrow As Long
    lrow = WorksheetTS.Range("J9:" & "J" & Rows.Count).End(xlDown).Row

    Dim DataBangTheoDoiTS As Range
    Set DataBangTheoDoiTS = WorksheetTS.Range("A9:AA" & lrow)

    Dim DataArrayBangTheoDoiTS As Variant
    DataArrayBangTheoDoiTS = DataBangTheoDoiTS.Value2

    Dim DataInNhanQLTS() As String
    ReDim DataInNhanQLTS(1 To UBound(DataArrayBangTheoDoiTS, 1), 1 To 2) ' Tao ra mang gom n dong theo tich chon x, va 2 cot

    Dim i As Long, cntX As Long
    cntX = 0
    For i = LBound(DataArrayBangTheoDoiTS, 1) To UBound(DataArrayBangTheoDoiTS, 1)
        If DataArrayBangTheoDoiTS(i, 26) = "x" Then ' Chon In Nhan QL TS
                cntX = cntX + 1
                DataInNhanQLTS(cntX, 1) = DataArrayBangTheoDoiTS(i, 8)   ' Ten/Loai thiet bi
                DataInNhanQLTS(cntX, 2) = DataArrayBangTheoDoiTS(i, 22)  ' Ma TS Quy uoc
        End If
    Next i

    Dim NhanQLTSAddress() As String
    ReDim NhanQLTSAddress(1 To 50, 1 To 2)

    NhanQLTSAddress(1, 1) = "$D$5"
    NhanQLTSAddress(1, 2) = "$D$6"
    NhanQLTSAddress(2, 1) = "$H$5"
    NhanQLTSAddress(2, 2) = "$H$6"
    NhanQLTSAddress(3, 1) = "$L$5"
    NhanQLTSAddress(3, 2) = "$L$6"
    NhanQLTSAddress(4, 1) = "$P$5"
    NhanQLTSAddress(4, 2) = "$P$6"
    NhanQLTSAddress(5, 1) = "$T$5"
    NhanQLTSAddress(5, 2) = "$T$6"

    NhanQLTSAddress(6, 1) = "$D$12"
    NhanQLTSAddress(6, 2) = "$D$13"
    NhanQLTSAddress(7, 1) = "$H$12"
    NhanQLTSAddress(7, 2) = "$H$13"
    NhanQLTSAddress(8, 1) = "$L$12"
    NhanQLTSAddress(8, 2) = "$L$13"
    NhanQLTSAddress(9, 1) = "$P$12"
    NhanQLTSAddress(9, 2) = "$P$13"
    NhanQLTSAddress(10, 1) = "$T$12"
    NhanQLTSAddress(10, 2) = "$T$13"

    NhanQLTSAddress(11, 1) = "$D$19"
    NhanQLTSAddress(11, 2) = "$D$20"
    NhanQLTSAddress(12, 1) = "$H$19"
    NhanQLTSAddress(12, 2) = "$H$20"
    NhanQLTSAddress(13, 1) = "$L$19"
    NhanQLTSAddress(13, 2) = "$L$20"
    NhanQLTSAddress(14, 1) = "$P$19"
    NhanQLTSAddress(14, 2) = "$P$20"
    NhanQLTSAddress(15, 1) = "$T$19"
    NhanQLTSAddress(15, 2) = "$T$20"

    NhanQLTSAddress(16, 1) = "$D$26"
    NhanQLTSAddress(16, 2) = "$D$27"
    NhanQLTSAddress(17, 1) = "$H$26"
    NhanQLTSAddress(17, 2) = "$H$27"
    NhanQLTSAddress(18, 1) = "$L$26"
    NhanQLTSAddress(18, 2) = "$L$27"
    NhanQLTSAddress(19, 1) = "$P$26"
    NhanQLTSAddress(19, 2) = "$P$27"
    NhanQLTSAddress(20, 1) = "$T$26"
    NhanQLTSAddress(20, 2) = "$T$27"

    NhanQLTSAddress(21, 1) = "$D$33"
    NhanQLTSAddress(21, 2) = "$D$34"
    NhanQLTSAddress(22, 1) = "$H$33"
    NhanQLTSAddress(22, 2) = "$H$34"
    NhanQLTSAddress(23, 1) = "$L$33"
    NhanQLTSAddress(23, 2) = "$L$34"
    NhanQLTSAddress(24, 1) = "$P$33"
    NhanQLTSAddress(24, 2) = "$P$34"
    NhanQLTSAddress(25, 1) = "$T$33"
    NhanQLTSAddress(25, 2) = "$T$34"

    NhanQLTSAddress(26, 1) = "$D$40"
    NhanQLTSAddress(26, 2) = "$D$41"
    NhanQLTSAddress(27, 1) = "$H$40"
    NhanQLTSAddress(27, 2) = "$H$41"
    NhanQLTSAddress(28, 1) = "$L$40"
    NhanQLTSAddress(28, 2) = "$L$41"
    NhanQLTSAddress(29, 1) = "$P$40"
    NhanQLTSAddress(29, 2) = "$P$41"
    NhanQLTSAddress(30, 1) = "$T$40"
    NhanQLTSAddress(30, 2) = "$T$41"

    NhanQLTSAddress(31, 1) = "$D$47"
    NhanQLTSAddress(31, 2) = "$D$48"
    NhanQLTSAddress(32, 1) = "$H$47"
    NhanQLTSAddress(32, 2) = "$H$48"
    NhanQLTSAddress(33, 1) = "$L$47"
    NhanQLTSAddress(33, 2) = "$L$48"
    NhanQLTSAddress(34, 1) = "$P$47"
    NhanQLTSAddress(34, 2) = "$P$48"
    NhanQLTSAddress(35, 1) = "$T$47"
    NhanQLTSAddress(35, 2) = "$T$48"

    NhanQLTSAddress(36, 1) = "$D$54"
    NhanQLTSAddress(36, 2) = "$D$55"
    NhanQLTSAddress(37, 1) = "$H$54"
    NhanQLTSAddress(37, 2) = "$H$55"
    NhanQLTSAddress(38, 1) = "$L$54"
    NhanQLTSAddress(38, 2) = "$L$55"
    NhanQLTSAddress(39, 1) = "$P$54"
    NhanQLTSAddress(39, 2) = "$P$55"
    NhanQLTSAddress(40, 1) = "$T$54"
    NhanQLTSAddress(40, 2) = "$T$55"

    NhanQLTSAddress(41, 1) = "$D$61"
    NhanQLTSAddress(41, 2) = "$D$62"
    NhanQLTSAddress(42, 1) = "$H$61"
    NhanQLTSAddress(42, 2) = "$H$62"
    NhanQLTSAddress(43, 1) = "$L$61"
    NhanQLTSAddress(43, 2) = "$L$62"
    NhanQLTSAddress(44, 1) = "$P$61"
    NhanQLTSAddress(44, 2) = "$P$62"
    NhanQLTSAddress(45, 1) = "$T$61"
    NhanQLTSAddress(45, 2) = "$T$62"

    NhanQLTSAddress(46, 1) = "$D$68"
    NhanQLTSAddress(46, 2) = "$D$69"
    NhanQLTSAddress(47, 1) = "$H$68"
    NhanQLTSAddress(47, 2) = "$H$69"
    NhanQLTSAddress(48, 1) = "$L$68"
    NhanQLTSAddress(48, 2) = "$L$69"
    NhanQLTSAddress(49, 1) = "$P$68"
    NhanQLTSAddress(49, 2) = "$P$69"
    NhanQLTSAddress(50, 1) = "$T$68"
    NhanQLTSAddress(50, 2) = "$T$69"


    Dim SoFilePDFExport As Long
    SoFilePDFExport = Application.WorksheetFunction.RoundUp(cntX / 50, 0)

    Dim FileName As Variant
    Dim FileNameExport As Variant
    Dim ReportSheet As Worksheet
    Dim Title As String
    Dim InitialFileName As String
    
    Set ReportSheet = ThisWorkbook.Worksheets(sheetnameBangTheoDoiTS)
    Title = "L" & ChrW(432) & "u nh" & ChrW(227) & "n Qu" & ChrW(7843) & "n l" & ChrW(253) & " T" & ChrW(224) & "i s" & ChrW(7843) & "n"
    ' Title = Luu nhan Quan ly Tai san
    InitialFileName = "Nhan Quan ly Tai san - " & Format(Now, "yyyymmdd")
    
    FileName = Application.GetSaveAsFilename(InitialFileName:=InitialFileName, _
                FileFilter:="PDF Files (*.pdf), *.pdf", _
                Title:=Title)

    FileNameExport = Replace(FileName, ".pdf", "")

    Dim NhanQLTSTemplateSheet As Worksheet
    Set NhanQLTSTemplateSheet = ThisWorkbook.Sheets("Nhan QL TS")


    If FileName = False Then
        Exit Sub
    Else
        Dim j As Long, n As Long, m As Long
        For j = 1 To SoFilePDFExport
            For n = 1 To 50
                NhanQLTSTemplateSheet.Range(NhanQLTSAddress(n, 1)).Value = DataInNhanQLTS((j - 1) * 50 + n, 1)
                NhanQLTSTemplateSheet.Range(NhanQLTSAddress(n, 2)).Value = DataInNhanQLTS((j - 1) * 50 + n, 2)
            Next n

            FileName = FileNameExport & " - " & j & " of " & SoFilePDFExport
            NhanQLTSTemplateSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=FileName, _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            For m = 1 To 50
                NhanQLTSTemplateSheet.Range(NhanQLTSAddress(m, 1)).ClearContents
                NhanQLTSTemplateSheet.Range(NhanQLTSAddress(m, 2)).ClearContents
            Next m

        Next j

    End If
End Sub
Private Sub btnXuatNhanKKTS()
    Dim sheetnameBangTheoDoiTS As String
    Dim WorksheetTS As Worksheet

    sheetnameBangTheoDoiTS = "BANG THEO DOI TS"
    Set WorksheetTS = ThisWorkbook.Worksheets(sheetnameBangTheoDoiTS)

    Dim lrow As Long
    lrow = WorksheetTS.Range("J9:" & "J" & Rows.Count).End(xlDown).Row

    Dim DataBangTheoDoiTS As Range
    Set DataBangTheoDoiTS = WorksheetTS.Range("A9:AA" & lrow)

    Dim DataArrayBangTheoDoiTS As Variant
    DataArrayBangTheoDoiTS = DataBangTheoDoiTS.Value2

    Dim DataInNhanKKTS() As String
    ReDim DataInNhanKKTS(1 To UBound(DataArrayBangTheoDoiTS, 1), 1 To 3) ' Tao ra mang gom n dong theo tich chon x, va 3 cot

    Dim i As Long, cntX As Long
    cntX = 0
    For i = LBound(DataArrayBangTheoDoiTS, 1) To UBound(DataArrayBangTheoDoiTS, 1)
        If DataArrayBangTheoDoiTS(i, 27) = "x" Then ' Chon In Nhan KK TS
                cntX = cntX + 1
                DataInNhanKKTS(cntX, 1) = DataArrayBangTheoDoiTS(i, 8)   ' Ten/Loai thiet bi
                DataInNhanKKTS(cntX, 2) = DataArrayBangTheoDoiTS(i, 22)  ' Ma TS Quy uoc
                DataInNhanKKTS(cntX, 3) = DataArrayBangTheoDoiTS(i, 4)   ' Phong/Ban su dung
        End If
    Next i

    Dim NhanKKAddress() As String
    ReDim NhanKKAddress(1 To 50, 1 To 3)

    NhanKKAddress(1, 1) = "$D$5"
    NhanKKAddress(1, 2) = "$D$6"
    NhanKKAddress(1, 3) = "$D$7"
    NhanKKAddress(2, 1) = "$H$5"
    NhanKKAddress(2, 2) = "$H$6"
    NhanKKAddress(2, 3) = "$H$7"
    NhanKKAddress(3, 1) = "$L$5"
    NhanKKAddress(3, 2) = "$L$6"
    NhanKKAddress(3, 3) = "$L$7"
    NhanKKAddress(4, 1) = "$P$5"
    NhanKKAddress(4, 2) = "$P$6"
    NhanKKAddress(4, 3) = "$P$7"
    NhanKKAddress(5, 1) = "$T$5"
    NhanKKAddress(5, 2) = "$T$6"
    NhanKKAddress(5, 3) = "$T$7"

    NhanKKAddress(6, 1) = "$D$13"
    NhanKKAddress(6, 2) = "$D$14"
    NhanKKAddress(6, 3) = "$D$15"
    NhanKKAddress(7, 1) = "$H$13"
    NhanKKAddress(7, 2) = "$H$14"
    NhanKKAddress(7, 3) = "$H$15"
    NhanKKAddress(8, 1) = "$L$13"
    NhanKKAddress(8, 2) = "$L$14"
    NhanKKAddress(8, 3) = "$L$15"
    NhanKKAddress(9, 1) = "$P$13"
    NhanKKAddress(9, 2) = "$P$14"
    NhanKKAddress(9, 3) = "$P$15"
    NhanKKAddress(10, 1) = "$T$13"
    NhanKKAddress(10, 2) = "$T$14"
    NhanKKAddress(10, 3) = "$T$15"

    NhanKKAddress(11, 1) = "$D$21"
    NhanKKAddress(11, 2) = "$D$22"
    NhanKKAddress(11, 3) = "$D$23"
    NhanKKAddress(12, 1) = "$H$21"
    NhanKKAddress(12, 2) = "$H$22"
    NhanKKAddress(12, 3) = "$H$23"
    NhanKKAddress(13, 1) = "$L$21"
    NhanKKAddress(13, 2) = "$L$22"
    NhanKKAddress(13, 3) = "$L$23"
    NhanKKAddress(14, 1) = "$P$21"
    NhanKKAddress(14, 2) = "$P$22"
    NhanKKAddress(14, 3) = "$P$23"
    NhanKKAddress(15, 1) = "$T$21"
    NhanKKAddress(15, 2) = "$T$22"
    NhanKKAddress(15, 3) = "$T$23"

    NhanKKAddress(16, 1) = "$D$29"
    NhanKKAddress(16, 2) = "$D$30"
    NhanKKAddress(16, 3) = "$D$31"
    NhanKKAddress(17, 1) = "$H$29"
    NhanKKAddress(17, 2) = "$H$30"
    NhanKKAddress(17, 3) = "$H$31"
    NhanKKAddress(18, 1) = "$L$29"
    NhanKKAddress(18, 2) = "$L$30"
    NhanKKAddress(18, 3) = "$L$31"
    NhanKKAddress(19, 1) = "$P$29"
    NhanKKAddress(19, 2) = "$P$30"
    NhanKKAddress(19, 3) = "$P$31"
    NhanKKAddress(20, 1) = "$T$29"
    NhanKKAddress(20, 2) = "$T$30"
    NhanKKAddress(20, 3) = "$T$31"

    NhanKKAddress(21, 1) = "$D$37"
    NhanKKAddress(21, 2) = "$D$38"
    NhanKKAddress(21, 3) = "$D$39"
    NhanKKAddress(22, 1) = "$H$37"
    NhanKKAddress(22, 2) = "$H$38"
    NhanKKAddress(22, 3) = "$H$39"
    NhanKKAddress(23, 1) = "$L$37"
    NhanKKAddress(23, 2) = "$L$38"
    NhanKKAddress(23, 3) = "$L$39"
    NhanKKAddress(24, 1) = "$P$37"
    NhanKKAddress(24, 2) = "$P$38"
    NhanKKAddress(24, 3) = "$P$39"
    NhanKKAddress(25, 1) = "$T$37"
    NhanKKAddress(25, 2) = "$T$38"
    NhanKKAddress(25, 3) = "$T$39"

    NhanKKAddress(26, 1) = "$D$45"
    NhanKKAddress(26, 2) = "$D$46"
    NhanKKAddress(26, 3) = "$D$47"
    NhanKKAddress(27, 1) = "$H$45"
    NhanKKAddress(27, 2) = "$H$46"
    NhanKKAddress(27, 3) = "$H$47"
    NhanKKAddress(28, 1) = "$L$45"
    NhanKKAddress(28, 2) = "$L$46"
    NhanKKAddress(28, 3) = "$L$47"
    NhanKKAddress(29, 1) = "$P$45"
    NhanKKAddress(29, 2) = "$P$46"
    NhanKKAddress(29, 3) = "$P$47"
    NhanKKAddress(30, 1) = "$T$45"
    NhanKKAddress(30, 2) = "$T$46"
    NhanKKAddress(30, 3) = "$T$47"

    NhanKKAddress(31, 1) = "$D$53"
    NhanKKAddress(31, 2) = "$D$54"
    NhanKKAddress(31, 3) = "$D$55"
    NhanKKAddress(32, 1) = "$H$53"
    NhanKKAddress(32, 2) = "$H$54"
    NhanKKAddress(32, 3) = "$H$55"
    NhanKKAddress(33, 1) = "$L$53"
    NhanKKAddress(33, 2) = "$L$54"
    NhanKKAddress(33, 3) = "$L$55"
    NhanKKAddress(34, 1) = "$P$53"
    NhanKKAddress(34, 2) = "$P$54"
    NhanKKAddress(34, 3) = "$P$55"
    NhanKKAddress(35, 1) = "$T$53"
    NhanKKAddress(35, 2) = "$T$54"
    NhanKKAddress(35, 3) = "$T$55"

    NhanKKAddress(36, 1) = "$D$61"
    NhanKKAddress(36, 2) = "$D$62"
    NhanKKAddress(36, 3) = "$D$63"
    NhanKKAddress(37, 1) = "$H$61"
    NhanKKAddress(37, 2) = "$H$62"
    NhanKKAddress(37, 3) = "$H$63"
    NhanKKAddress(38, 1) = "$L$61"
    NhanKKAddress(38, 2) = "$L$62"
    NhanKKAddress(38, 3) = "$L$63"
    NhanKKAddress(39, 1) = "$P$61"
    NhanKKAddress(39, 2) = "$P$62"
    NhanKKAddress(39, 3) = "$P$63"
    NhanKKAddress(40, 1) = "$T$61"
    NhanKKAddress(40, 2) = "$T$62"
    NhanKKAddress(40, 3) = "$T$63"

    NhanKKAddress(41, 1) = "$D$69"
    NhanKKAddress(41, 2) = "$D$70"
    NhanKKAddress(41, 3) = "$D$71"
    NhanKKAddress(42, 1) = "$H$69"
    NhanKKAddress(42, 2) = "$H$70"
    NhanKKAddress(42, 3) = "$H$71"
    NhanKKAddress(43, 1) = "$L$69"
    NhanKKAddress(43, 2) = "$L$70"
    NhanKKAddress(43, 3) = "$L$71"
    NhanKKAddress(44, 1) = "$P$69"
    NhanKKAddress(44, 2) = "$P$70"
    NhanKKAddress(44, 3) = "$P$71"
    NhanKKAddress(45, 1) = "$T$69"
    NhanKKAddress(45, 2) = "$T$70"
    NhanKKAddress(45, 3) = "$T$71"

    NhanKKAddress(46, 1) = "$D$77"
    NhanKKAddress(46, 2) = "$D$78"
    NhanKKAddress(46, 3) = "$D$79"
    NhanKKAddress(47, 1) = "$H$77"
    NhanKKAddress(47, 2) = "$H$78"
    NhanKKAddress(47, 3) = "$H$79"
    NhanKKAddress(48, 1) = "$L$77"
    NhanKKAddress(48, 2) = "$L$78"
    NhanKKAddress(48, 3) = "$L$79"
    NhanKKAddress(49, 1) = "$P$77"
    NhanKKAddress(49, 2) = "$P$78"
    NhanKKAddress(49, 3) = "$P$79"
    NhanKKAddress(50, 1) = "$T$77"
    NhanKKAddress(50, 2) = "$T$78"
    NhanKKAddress(50, 3) = "$T$79 "


    Dim SoFilePDFExport As Long
    SoFilePDFExport = Application.WorksheetFunction.RoundUp(cntX / 50, 0)
    Dim FileName As Variant
    Dim FileNameExport As Variant
    Dim ReportSheet As Worksheet
    Dim Title As String
    Dim InitialFileName As String
    
    Set ReportSheet = ThisWorkbook.Worksheets(sheetnameBangTheoDoiTS)
    Title = "L" & ChrW(432) & "u nh" & ChrW(227) & "n Qu" & ChrW(7843) & "n l" & ChrW(253) & " T" & ChrW(224) & "i s" & ChrW(7843) & "n"
    ' Title = Luu nhan Quan ly Tai san
    InitialFileName = "Nhan Kiem Ke Tai san - " & Format(Now, "yyyymmdd")
    
    FileName = Application.GetSaveAsFilename(InitialFileName:=InitialFileName, _
                FileFilter:="PDF Files (*.pdf), *.pdf", _
                Title:=Title)

    FileNameExport = Replace(FileName, ".pdf", "")

    Dim NhanKKTSTemplateSheet As Worksheet
    Set NhanKKTSTemplateSheet = ThisWorkbook.Sheets("Nhan KK TS")


    If FileName = False Then
        Exit Sub
    Else
        Dim j As Long, n As Long, m As Long
        For j = 1 To SoFilePDFExport
            For n = 1 To 50
                NhanKKTSTemplateSheet.Range(NhanKKAddress(n, 1)).Value = DataInNhanKKTS((j - 1) * 50 + n, 1)
                NhanKKTSTemplateSheet.Range(NhanKKAddress(n, 2)).Value = DataInNhanKKTS((j - 1) * 50 + n, 2)
                NhanKKTSTemplateSheet.Range(NhanKKAddress(n, 3)).Value = DataInNhanKKTS((j - 1) * 50 + n, 3)
            Next n

            FileName = FileNameExport & " - " & j & " of " & SoFilePDFExport
            NhanKKTSTemplateSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=FileName, _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            For m = 1 To 50
                NhanKKTSTemplateSheet.Range(NhanKKAddress(m, 1)).ClearContents
                NhanKKTSTemplateSheet.Range(NhanKKAddress(m, 2)).ClearContents
                NhanKKTSTemplateSheet.Range(NhanKKAddress(m, 3)).ClearContents
            Next m

        Next j

    End If
End Sub