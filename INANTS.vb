Option Explicit
Option Base 1

' **********************************************
' I. Ribbon ************************************
' **********************************************

' ****** 1. Ribbon ******
Public ActivateStatus As Boolean

Private Sub tggbtnHighlightRowsAndCols(control As IRibbonControl, pressed As Boolean)
    ActivateStatus = pressed
End Sub

Private Sub Callback(control As IRibbonControl)
    Dim ControlString As String
    Select Case control.ID
        Case "btnDaiDung"
            ControlString = control.ID
        Case "btnVersion"
            ControlString = control.ID
        Case "btnTaoDanhSachTaiSan"
            ControlString = control.ID
        Case "btnXuatNhanQLTS"
            ControlString = control.ID
        Case "btnXuatNhanKKTS"
            ControlString = control.ID
    End Select
    
    Application.Run ControlString
End Sub

' ****** 2. Procedure ******
Private Sub btnDaiDung()
    ActiveWorkbook.FollowHyperlink Address:="https://daidung.com/", NewWindow:=True
End Sub


Private Sub btnVersion()
    MsgBox "Add-in version", vbOKOnly, "Theo doi Tai san - DDC Add-in"
End Sub

Private Sub btnTaoDanhSachTaiSan()
    Dim sheetName As String
    Dim lastRow As Long
    Dim sttCotSoLuong As Long
    Dim tenCotSoLuong As String
    Dim WorksheetTS As Worksheet
    
    
    sheetName = "BANG THEO DOI TS"
    sttCotSoLuong = 10
    tenCotSoLuong = "J"
    Set WorksheetTS = ThisWorkbook.Worksheets(sheetName)
    
    WorksheetTS.Activate
    
    lastRow = WorksheetTS.Range(tenCotSoLuong & Rows.Count).End(xlUp).Row
    
    Dim soLuong As Long
    Dim numOfRowFillDown As Long
    
    soLuong = Range(tenCotSoLuong & lastRow).Value
    
    numOfRowFillDown = lastRow + soLuong - 1
    
    WorksheetTS.Range("B" & lastRow & ":T" & numOfRowFillDown).Select
    Selection.FillDown
    WorksheetTS.Range("J" & lastRow & ":J" & numOfRowFillDown).Value = 1

    Call ThongBaoDaChayXong
End Sub


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
    ReDim DataInNhanQLTS(1 To UBound(DataArrayBangTheoDoiTS, 1) * 24, 1 To 3) ' Tao ra mang gom n dong theo tich chon x, va 3 cot

    Dim i As Long, cntX As Long
    cntX = 0
    For i = LBound(DataArrayBangTheoDoiTS, 1) To UBound(DataArrayBangTheoDoiTS, 1)
        If DataArrayBangTheoDoiTS(i, 26) = "x" Then ' Chon In Nhan QL TS
                cntX = cntX + 1
                DataInNhanQLTS(cntX, 1) = DataArrayBangTheoDoiTS(i, 8)   ' Ten/Loai thiet bi
                DataInNhanQLTS(cntX, 2) = DataArrayBangTheoDoiTS(i, 22)  ' Ma TS Quy uoc
                DataInNhanQLTS(cntX, 3) = DataArrayBangTheoDoiTS(i, 17)  ' Tgian cap phat
        End If
    Next i

    Dim NhanQLTSAddress() As String
    ReDim NhanQLTSAddress(1 To 24, 1 To 3)

     NhanQLTSAddress(1, 1) = "$E$5"
    NhanQLTSAddress(1, 2) = "$E$6"
    NhanQLTSAddress(1, 3) = "$E$7"
    NhanQLTSAddress(2, 1) = "$K$5"
    NhanQLTSAddress(2, 2) = "$K$6"
    NhanQLTSAddress(2, 3) = "$K$7"
    NhanQLTSAddress(3, 1) = "$Q$5"
    NhanQLTSAddress(3, 2) = "$Q$6"
    NhanQLTSAddress(3, 3) = "$Q$7"
    
    NhanQLTSAddress(4, 1) = "$E$13"
    NhanQLTSAddress(4, 2) = "$E$14"
    NhanQLTSAddress(4, 3) = "$E$15"
    NhanQLTSAddress(5, 1) = "$K$13"
    NhanQLTSAddress(5, 2) = "$K$14"
    NhanQLTSAddress(5, 3) = "$K$15"
    NhanQLTSAddress(6, 1) = "$Q$13"
    NhanQLTSAddress(6, 2) = "$Q$14"
    NhanQLTSAddress(6, 3) = "$Q$15"
    
    NhanQLTSAddress(7, 1) = "$E$21"
    NhanQLTSAddress(7, 2) = "$E$22"
    NhanQLTSAddress(7, 3) = "$E$23"
    NhanQLTSAddress(8, 1) = "$K$21"
    NhanQLTSAddress(8, 2) = "$K$22"
    NhanQLTSAddress(8, 3) = "$K$23"
    NhanQLTSAddress(9, 1) = "$Q$21"
    NhanQLTSAddress(9, 2) = "$Q$22"
    NhanQLTSAddress(9, 3) = "$Q$23"
    

    NhanQLTSAddress(10, 1) = "$E$29"
    NhanQLTSAddress(10, 2) = "$E$30"
    NhanQLTSAddress(10, 3) = "$E$31"
    NhanQLTSAddress(11, 1) = "$K$29"
    NhanQLTSAddress(11, 2) = "$K$30"
    NhanQLTSAddress(11, 3) = "$K$31"
    NhanQLTSAddress(12, 1) = "$Q$29"
    NhanQLTSAddress(12, 2) = "$Q$30"
    NhanQLTSAddress(12, 3) = "$Q$31"


    NhanQLTSAddress(13, 1) = "$E$37"
    NhanQLTSAddress(13, 2) = "$E$38"
    NhanQLTSAddress(13, 3) = "$E$39"
    NhanQLTSAddress(14, 1) = "$K$37"
    NhanQLTSAddress(14, 2) = "$K$38"
    NhanQLTSAddress(14, 3) = "$K$39"
    NhanQLTSAddress(15, 1) = "$Q$37"
    NhanQLTSAddress(15, 2) = "$Q$38"
    NhanQLTSAddress(15, 3) = "$Q$39"

    NhanQLTSAddress(16, 1) = "$E$45"
    NhanQLTSAddress(16, 2) = "$E$46"
    NhanQLTSAddress(16, 3) = "$E$47"
    NhanQLTSAddress(17, 1) = "$K$45"
    NhanQLTSAddress(17, 2) = "$K$46"
    NhanQLTSAddress(17, 3) = "$K$47"
    NhanQLTSAddress(18, 1) = "$Q$45"
    NhanQLTSAddress(18, 2) = "$Q$46"
    NhanQLTSAddress(18, 3) = "$Q$47"
    

    NhanQLTSAddress(19, 1) = "$E$53"
    NhanQLTSAddress(19, 2) = "$E$54"
    NhanQLTSAddress(19, 3) = "$E$55"
    NhanQLTSAddress(20, 1) = "$K$53"
    NhanQLTSAddress(20, 2) = "$K$54"
    NhanQLTSAddress(20, 3) = "$K$55"
    NhanQLTSAddress(21, 1) = "$Q$53"
    NhanQLTSAddress(21, 2) = "$Q$54"
    NhanQLTSAddress(21, 3) = "$Q$55"


    NhanQLTSAddress(22, 1) = "$E$61"
    NhanQLTSAddress(22, 2) = "$E$62"
    NhanQLTSAddress(22, 3) = "$E$63"
    NhanQLTSAddress(23, 1) = "$K$61"
    NhanQLTSAddress(23, 2) = "$K$62"
    NhanQLTSAddress(23, 3) = "$K$63"
    NhanQLTSAddress(24, 1) = "$Q$61"
    NhanQLTSAddress(24, 2) = "$Q$62"
    NhanQLTSAddress(24, 3) = "$Q$63"



    'NhanQLTSAddress(25, 1) = "$D$69"
    'NhanQLTSAddress(25, 2) = "$D$70"
    'NhanQLTSAddress(25, 3) = "$E$71"
    'NhanQLTSAddress(26, 1) = "$J$69"
    'NhanQLTSAddress(26, 2) = "$J$70"
    'NhanQLTSAddress(26, 3) = "$K$71"
    'NhanQLTSAddress(27, 1) = "$P$69"
    'NhanQLTSAddress(27, 2) = "$P$70"
    'NhanQLTSAddress(27, 3) = "$Q$71"


    'NhanQLTSAddress(28, 1) = "$D$77"
    'NhanQLTSAddress(28, 2) = "$D$78"
    'NhanQLTSAddress(28, 3) = "$E$79"
    'NhanQLTSAddress(29, 1) = "$J$77"
    'NhanQLTSAddress(29, 2) = "$J$78"
    'NhanQLTSAddress(29, 3) = "$k$79"
    'NhanQLTSAddress(30, 1) = "$P$77"
    'NhanQLTSAddress(30, 2) = "$P$78"
   'NhanQLTSAddress(30, 3) = "$Q$79"


    'NhanQLTSAddress(31, 1) = "$D$85"
    'NhanQLTSAddress(31, 2) = "$D$86"
    'NhanQLTSAddress(31, 3) = "$E$87"
    'NhanQLTSAddress(32, 1) = "$J$85"
    'NhanQLTSAddress(32, 2) = "$J$86"
    'NhanQLTSAddress(32, 3) = "$K$87"
    'NhanQLTSAddress(33, 1) = "$P$85"
    'NhanQLTSAddress(33, 2) = "$P$86"
    'NhanQLTSAddress(33, 3) = "$Q$87"



    Dim SoFilePDFExport As Long
    SoFilePDFExport = Application.WorksheetFunction.RoundUp(cntX / 24, 0)

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
            For n = 1 To 24
                NhanQLTSTemplateSheet.Range(NhanQLTSAddress(n, 1)).Value = DataInNhanQLTS((j - 1) * 24 + n, 1)
                NhanQLTSTemplateSheet.Range(NhanQLTSAddress(n, 2)).Value = DataInNhanQLTS((j - 1) * 24 + n, 2)
                NhanQLTSTemplateSheet.Range(NhanQLTSAddress(n, 3)).Value = DataInNhanQLTS((j - 1) * 24 + n, 3)
            Next n

            FileName = FileNameExport & " - " & j & " of " & SoFilePDFExport
            NhanQLTSTemplateSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=FileName, _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            For m = 1 To 24
                NhanQLTSTemplateSheet.Range(NhanQLTSAddress(m, 1)).ClearContents
                NhanQLTSTemplateSheet.Range(NhanQLTSAddress(m, 2)).ClearContents
                NhanQLTSTemplateSheet.Range(NhanQLTSAddress(m, 3)).ClearContents
            Next m

        Next j

    End If

    Call ThongBaoDaChayXong
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
    ReDim DataInNhanKKTS(1 To UBound(DataArrayBangTheoDoiTS, 1) * 27, 1 To 4) ' Tao ra mang gom n dong theo tich chon x, va 4 cot

    Dim i As Long, cntX As Long
    cntX = 0
    For i = LBound(DataArrayBangTheoDoiTS, 1) To UBound(DataArrayBangTheoDoiTS, 1)
        If DataArrayBangTheoDoiTS(i, 27) = "x" Then ' Chon In Nhan KK TS
                cntX = cntX + 1
                DataInNhanKKTS(cntX, 1) = DataArrayBangTheoDoiTS(i, 8)   ' Ten/Loai thiet bi
                DataInNhanKKTS(cntX, 2) = DataArrayBangTheoDoiTS(i, 24)  ' Ma TS Kiem ke
                DataInNhanKKTS(cntX, 3) = DataArrayBangTheoDoiTS(i, 4)   ' Phong/Ban su dung
                DataInNhanKKTS(cntX, 4) = DataArrayBangTheoDoiTS(i, 25)   ' Ngay Kiem Ke
        End If
    Next i

    Dim NhanKKAddress() As String
    ReDim NhanKKAddress(1 To 27, 1 To 4)

    NhanKKAddress(1, 1) = "$D$5"
    NhanKKAddress(1, 2) = "$D$6"
    NhanKKAddress(1, 3) = "$E$7"
    NhanKKAddress(1, 4) = "$E$8"
    NhanKKAddress(2, 1) = "$J$5"
    NhanKKAddress(2, 2) = "$J$6"
    NhanKKAddress(2, 3) = "$K$7"
    NhanKKAddress(2, 4) = "$K$8"
    NhanKKAddress(3, 1) = "$P$5"
    NhanKKAddress(3, 2) = "$P$6"
    NhanKKAddress(3, 3) = "$Q$7"
    NhanKKAddress(3, 4) = "$Q$8"
    
    NhanKKAddress(4, 1) = "$D$14"
    NhanKKAddress(4, 2) = "$D$15"
    NhanKKAddress(4, 3) = "$E$16"
    NhanKKAddress(4, 4) = "$E$17"
    NhanKKAddress(5, 1) = "$J$14"
    NhanKKAddress(5, 2) = "$J$15"
    NhanKKAddress(5, 3) = "$K$16"
    NhanKKAddress(5, 4) = "$K$17"
    NhanKKAddress(6, 1) = "$P$14"
    NhanKKAddress(6, 2) = "$P$15"
    NhanKKAddress(6, 3) = "$Q$16"
    NhanKKAddress(6, 4) = "$Q$17"
    
    NhanKKAddress(7, 1) = "$D$23"
    NhanKKAddress(7, 2) = "$D$24"
    NhanKKAddress(7, 3) = "$E$25"
    NhanKKAddress(7, 4) = "$E$26"
    NhanKKAddress(8, 1) = "$J$23"
    NhanKKAddress(8, 2) = "$J$24"
    NhanKKAddress(8, 3) = "$K$25"
    NhanKKAddress(8, 4) = "$K$26"
    NhanKKAddress(9, 1) = "$P$23"
    NhanKKAddress(9, 2) = "$P$24"
    NhanKKAddress(9, 3) = "$Q$25"
    NhanKKAddress(9, 4) = "$Q$26"
    
    NhanKKAddress(10, 1) = "$D$32"
    NhanKKAddress(10, 2) = "$D$33"
    NhanKKAddress(10, 3) = "$E$34"
    NhanKKAddress(10, 4) = "$E$35"
    NhanKKAddress(11, 1) = "$J$32"
    NhanKKAddress(11, 2) = "$J$33"
    NhanKKAddress(11, 3) = "$K$34"
    NhanKKAddress(11, 4) = "$K$35"
    NhanKKAddress(12, 1) = "$P$32"
    NhanKKAddress(12, 2) = "$P$33"
    NhanKKAddress(12, 3) = "$Q$34"
    NhanKKAddress(12, 4) = "$Q$35"
    
    NhanKKAddress(13, 1) = "$D$41"
    NhanKKAddress(13, 2) = "$D$42"
    NhanKKAddress(13, 3) = "$E$43"
    NhanKKAddress(13, 4) = "$E$44"
    NhanKKAddress(14, 1) = "$J$41"
    NhanKKAddress(14, 2) = "$J$42"
    NhanKKAddress(14, 3) = "$K$43"
    NhanKKAddress(14, 4) = "$K$44"
    NhanKKAddress(15, 1) = "$P$41"
    NhanKKAddress(15, 2) = "$P$42"
    NhanKKAddress(15, 3) = "$Q$43"
    NhanKKAddress(15, 4) = "$Q$44"

    NhanKKAddress(16, 1) = "$D$50"
    NhanKKAddress(16, 2) = "$D$51"
    NhanKKAddress(16, 3) = "$E$52"
    NhanKKAddress(16, 4) = "$E$53"
    NhanKKAddress(17, 1) = "$J$50"
    NhanKKAddress(17, 2) = "$J$51"
    NhanKKAddress(17, 3) = "$K$52"
    NhanKKAddress(17, 4) = "$K$53"
    NhanKKAddress(18, 1) = "$P$50"
    NhanKKAddress(18, 2) = "$P$51"
    NhanKKAddress(18, 3) = "$Q$52"
    NhanKKAddress(18, 4) = "$Q$53"
    
    NhanKKAddress(19, 1) = "$D$59"
    NhanKKAddress(19, 2) = "$D$60"
    NhanKKAddress(19, 3) = "$E$61"
    NhanKKAddress(19, 4) = "$E$62"
    NhanKKAddress(20, 1) = "$J$59"
    NhanKKAddress(20, 2) = "$J$60"
    NhanKKAddress(20, 3) = "$K$61"
    NhanKKAddress(20, 4) = "$K$62"
    NhanKKAddress(21, 1) = "$P$59"
    NhanKKAddress(21, 2) = "$P$60"
    NhanKKAddress(21, 3) = "$Q$61"
    NhanKKAddress(21, 4) = "$Q$62"
    
    NhanKKAddress(22, 1) = "$D$68"
    NhanKKAddress(22, 2) = "$D$69"
    NhanKKAddress(22, 3) = "$E$70"
    NhanKKAddress(22, 4) = "$E$71"
    NhanKKAddress(23, 1) = "$J$68"
    NhanKKAddress(23, 2) = "$J$69"
    NhanKKAddress(23, 3) = "$K$70"
    NhanKKAddress(23, 4) = "$K$71"
    NhanKKAddress(24, 1) = "$P$68"
    NhanKKAddress(24, 2) = "$P$69"
    NhanKKAddress(24, 3) = "$Q$70"
    NhanKKAddress(24, 4) = "$Q$71"
    
    NhanKKAddress(25, 1) = "$D$77"
    NhanKKAddress(25, 2) = "$D$78"
    NhanKKAddress(25, 3) = "$E$79"
    NhanKKAddress(25, 4) = "$E$80"
    NhanKKAddress(26, 1) = "$J$77"
    NhanKKAddress(26, 2) = "$J$78"
    NhanKKAddress(26, 3) = "$K$79"
    NhanKKAddress(26, 4) = "$K$80"
    NhanKKAddress(27, 1) = "$P$77"
    NhanKKAddress(27, 2) = "$P$78"
    NhanKKAddress(27, 3) = "$Q$79"
    NhanKKAddress(27, 4) = "$Q$80"
    
    'NhanKKAddress(28, 1) = "$D$86"
    'NhanKKAddress(28, 2) = "$D$87"
    'NhanKKAddress(28, 3) = "$E$88"
    'NhanKKAddress(28, 4) = "$E$89"
    'NhanKKAddress(29, 1) = "$J$86"
   'NhanKKAddress(29, 2) = "$J$87"
    'NhanKKAddress(29, 3) = "$K$88"
   'NhanKKAddress(29, 4) = "$K$89"
   'NhanKKAddress(30, 1) = "$P$86"
    'NhanKKAddress(30, 2) = "$P$87"
    'NhanKKAddress(30, 3) = "$Q$88"
    'NhanKKAddress(30, 4) = "$Q$89"

    'NhanKKAddress(31, 1) = "$D$95"
    'NhanKKAddress(31, 2) = "$D$96"
    'NhanKKAddress(31, 3) = "$E$97"
    'NhanKKAddress(31, 4) = "$E$98"
    'NhanKKAddress(32, 1) = "$J$95"
    'NhanKKAddress(32, 2) = "$J$96"
    'NhanKKAddress(32, 3) = "$K$97"
    'NhanKKAddress(32, 4) = "$K$98"
    'NhanKKAddress(33, 1) = "$P$95"
    'NhanKKAddress(33, 2) = "$P$96"
    'NhanKKAddress(33, 3) = "$Q$97"
    'NhanKKAddress(33, 4) = "$Q$98"
    
   


    Dim SoFilePDFExport As Long
    SoFilePDFExport = Application.WorksheetFunction.RoundUp(cntX / 27, 0)
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
            For n = 1 To 27
                NhanKKTSTemplateSheet.Range(NhanKKAddress(n, 1)).Value = DataInNhanKKTS((j - 1) * 27 + n, 1)
                NhanKKTSTemplateSheet.Range(NhanKKAddress(n, 2)).Value = DataInNhanKKTS((j - 1) * 27 + n, 2)
                NhanKKTSTemplateSheet.Range(NhanKKAddress(n, 3)).Value = DataInNhanKKTS((j - 1) * 27 + n, 3)
                NhanKKTSTemplateSheet.Range(NhanKKAddress(n, 4)).Value = DataInNhanKKTS((j - 1) * 27 + n, 4)
            Next n

            FileName = FileNameExport & " - " & j & " of " & SoFilePDFExport
            NhanKKTSTemplateSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=FileName, _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            For m = 1 To 27
                NhanKKTSTemplateSheet.Range(NhanKKAddress(m, 1)).ClearContents
                NhanKKTSTemplateSheet.Range(NhanKKAddress(m, 2)).ClearContents
                NhanKKTSTemplateSheet.Range(NhanKKAddress(m, 3)).ClearContents
                NhanKKTSTemplateSheet.Range(NhanKKAddress(m, 4)).ClearContents
            Next m

        Next j
'Commit test
    End If
    Call ThongBaoDaChayXong
    
End Sub

' **********************************************
' II. Helper Procedures ************************
' **********************************************
Private Sub ThongBaoDaChayXong()
    ' Thong bao [Da chay xong]
    CreateObject("WScript.Shell").Popup ChrW(272) & ChrW(227) & " ch" & ChrW(7841) & "y xong.", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", 0 + 64
End Sub




' **********************************************
' III. Trigger (callback) **********************
' **********************************************


' ************************* End Module Ribbon *************************

