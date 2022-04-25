Attribute VB_Name = "Module1"
Option Explicit
Public hoba, vestas, yenile As Boolean
Public id As Integer
Public idrow, idcol, sonkol As Integer
Public dbPath
Public ws As Worksheet
Public tarih1 As Date
Public targetval
Public kanatno, opno
Public SQL, sql1, proje As String
Public modal, maxset, k As Integer
Public fark As Integer
Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export "C:\Users\ftek\OneDrive - TPI Composites Inc\Masaüstü\Vba\" & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    MsgBox "Export is ready"
End Sub

Sub goster()
Dim wstracker As Worksheet
Set wstracker = ThisWorkbook.Worksheets("NCR TRACKER")
On Error GoTo err
        Dim sonuc As Integer
        On Error Resume Next
        sonuc = InputBox("Enter the password", "Database path change")
        If sonuc = "1474" Then
            wstracker.Unprotect Password:="4135911"
            Dim fname As Variant
            fname = Application.GetOpenFilename(filefilter:="Access Files,*.acc*")
            wstracker.Cells(1, 9) = fname
            With wstracker
                .Protect Password:="4135911", AllowFiltering:=True
                .EnableSelection = xlNoRestrictions
            End With
        Else
            MsgBox "Incorrect Password"
            Exit Sub
        End If
        On Error GoTo 0
Exit Sub
err:
    MsgBox "Ýnternete baðlý olduðunu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran görüntüsü at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "Sýkýntý var :("
    Exit Sub
End Sub

Sub refresh()
'On Error GoTo err
Application.DisplayStatusBar = True
Application.Calculation = xlManual
Application.ScreenUpdating = False
Dim ws, ws4, ws5 As Worksheet
Set ws = ThisWorkbook.Worksheets("Finish Status")
Set ws4 = ThisWorkbook.Worksheets("Gecikmeler")
Set ws5 = ThisWorkbook.Worksheets("Ýstasyon CT")
Dim i, j, k, t, tamamlanma As Integer
ws.Unprotect Password:="1709"
ws4.Unprotect Password:="1709"

Dim cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rscoc As ADODB.Recordset
Dim rsvers As ADODB.Recordset
Dim rsgecikme, rsset As ADODB.Recordset
Dim rsct As ADODB.Recordset
Dim SQL, sql2, sql3, sqlset, sqlvers, versiyon, sqlgecikme, sqlvak, sqlct As String
Dim lastrow, lastrow2, lastcolumn As Long
Dim durus, protime, ektime, slotsayisi, slot, leadtime, gecikme, istsayi, plantime, x, hedeflead As Double
Dim bladeno, atla, opno, sonsetler As Integer
Dim bitmis As Boolean
Dim bas, bit As Date
dbPath = ws.Range("B1").Value
Set cnn = New ADODB.Connection
cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=Sarper1709;"
    sqlvers = "SELECT VERS FROM VERS"
    Set rsvers = New ADODB.Recordset
    rsvers.Open sqlvers, cnn
    If rsvers.EOF And rsvers.BOF Then
            rsvers.Close
            cnn.Close
            Set rsvers = Nothing
            Set cnn = Nothing
            MsgBox "Herhangi bir kayýt bulunamadý server baðlantýnýzý kontrol edin", vbCritical, "No Records"
            Exit Sub
    End If
        versiyon = rsvers.Fields(0).Value
        If ws.Cells(1, 3) <> versiyon Then
        MsgBox "Eski bir sürüm kullanýyorsunuz baðlantý yapýlmadý!"
        On Error Resume Next
        cnn.Close
        Exit Sub
        End If
'On Error Resume Next
lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
lastrow = ws.Range("C" & Rows.Count).End(xlUp).Row
slotsayisi = 12
ws.Range("A7:AL" & lastrow).EntireRow.Delete

            sqlset = "SELECT TOP 1 COC.SetSayi, Count(COC.SetSayi) AS SaySetSayi FROM COC GROUP BY COC.SetSayi ORDER BY COC.Setsayi Desc"
            Set rsset = New ADODB.Recordset
            rsset.Open sqlset, cnn
            sonsetler = rsset.Fields(0).Value
            sonsetler = sonsetler - 20
    SQL = "SELECT Hafta, Setsayi, Kanat FROM COC WHERE Setsayi > " & sonsetler & " AND Kanat > 435 Order By [Setsayi] Asc, [Kanat] Asc"
    Set rs = New ADODB.Recordset
    rs.Open SQL, cnn
    ws.Cells(7, 1).CopyFromRecordset rs
sonkol = 42
sql3 = "SELECT * From Prod"
Set rs3 = New ADODB.Recordset
rs3.Open sql3, cnn
hedeflead = 155
For i = 7 To lastrow
    On Error GoTo hop
    'ws.Range(Cells(i, 1), Cells(i, 3)).Borders.LineStyle = xlContinuous ' ColorIndex:=1, Weight:=xlThin
    'ws.Range(Cells(i, sonkol + 1), Cells(i, sonkol + 6)).Borders.LineStyle = xlContinuous ' ColorIndex:=1, Weight:=xlThin
    bladeno = ws.Cells(i, 3)
    SQL = "SELECT Prod.Kanat, Sum(Prod.Hedef) AS ToplaHedef, Sum(Prod.Süre) AS ToplaSüre FROM Prod GROUP BY Prod.Kanat"
    Set rs = New ADODB.Recordset
    rs.Open SQL, cnn
    rs.Filter = "Kanat = '" & bladeno & "'"
    plantime = rs.Fields("ToplaHedef").Value
    opno = 0
    SQL = "SELECT Kanat, [Baþlangýç], [Operasyon No] FROM Prod WHERE (Kanat = " & bladeno & " AND [Operasyon No] = " & 0 & ") "
    Set rs2 = New ADODB.Recordset
    rs2.Open SQL, cnn
    bas = rs2.Fields("Baþlangýç").Value
    SQL = "SELECT TOP 1 * FROM PROD WHERE Kanat = " & bladeno & " ORDER BY [Bitiþ] DESC"
    Set rs2 = New ADODB.Recordset
    rs2.Open SQL, cnn
    bit = rs2.Fields("Bitiþ").Value
    SQL = "SELECT COC FROM COC WHERE Kanat = " & bladeno & ""
    Set rs2 = New ADODB.Recordset
    rs2.Open SQL, cnn
    If rs2.Fields("COC").Value = "Yes" Then
    ws.Cells(i, 2).Interior.ColorIndex = 43
    End If
    SQL = "SELECT [Not] FROM COC WHERE Kanat = " & bladeno & ""
    Set rs2 = New ADODB.Recordset
    rs2.Open SQL, cnn
    ws.Cells(i, sonkol + 6) = rs2.Fields("Not").Value
    ws.Cells(i, sonkol + 6).HorizontalAlignment = xlLeft
    ws.Cells(i, sonkol + 6).Font.Size = 9
    ws.Columns(sonkol + 6).AutoFit
    For j = 4 To sonkol + 5
        slot = ws.Cells(6, j)
        If j <= sonkol Then
           rs3.Filter = "Kanat = '" & bladeno & "' and [Operasyon No] = '" & ws.Cells(3, j).Value & "' "
           Do Until rs3.EOF
              If rs3.Fields("Vardiya2").Value = "X" Then 'iþ bitmemiþ
                        ws.Cells(i, j) = DateDiff("n", rs3.Fields("Baþlangýç"), Now) / 60
                        If ws.Cells(i, j) > ws.Cells(5, j) Then
                            ws.Cells(i, j).Font.Color = vbRed
                            ws.Cells(i, j).Interior.ColorIndex = 27
                            ws.Cells(i, j).BorderAround ColorIndex:=1, Weight:=xlThin
                            gecikme = gecikme + (ws.Cells(i, j) - ws.Cells(5, j))
                        ElseIf ws.Cells(i, j) < ws.Cells(5, j) Then
                            ws.Cells(i, j).Font.Color = vbBlack
                            ws.Cells(i, j).Interior.ColorIndex = 27
                            ws.Cells(i, j).BorderAround ColorIndex:=1, Weight:=xlThin
                        End If
              ElseIf rs3.Fields("Süre") <> "" Then
                        ws.Cells(i, j).Value = rs3.Fields("Süre")
                        If ws.Cells(i, j) > ws.Cells(5, j) Then
                            ws.Cells(i, j).Font.Color = vbRed
                            ws.Cells(i, j).Interior.ColorIndex = 43
                            ws.Cells(i, j).BorderAround ColorIndex:=1, Weight:=xlThin
                            gecikme = gecikme + (ws.Cells(i, j) - ws.Cells(5, j))
                        ElseIf ws.Cells(i, j) <= ws.Cells(5, j) Then
                            ws.Cells(i, j).Font.Color = vbBlack
                            ws.Cells(i, j).Interior.ColorIndex = 43
                            ws.Cells(i, j).BorderAround ColorIndex:=1, Weight:=xlThin
                        End If
                        istsayi = istsayi + 1
              Else:
              ws.Cells(i, j).Font.Color = vbBlack
              ws.Cells(i, j).Interior.Color = vbWhite
              End If
              protime = ws.Cells(i, j) + protime
              If ws.Cells(i, j).Value = "" Then ektime = ws.Cells(5, j) + ektime
              ws.Cells(i, j).NumberFormat = "0.0"
              rs3.MoveNext
           Loop
        ws.Cells(i, j).BorderAround ColorIndex:=1, Weight:=xlThin
        ElseIf j = sonkol + 1 Then
        On Error Resume Next
        'MsgBox bladeno & " protime " & protime & " plan " & plantime & " kalan " & plantime - protime
                'If i Mod 3 = 1 Then atla = 0 Else If i Mod 3 = 2 Then atla = 1 Else If i Mod 3 = 0 Then atla = 2
                If i <= lastrow + 2 And ektime > 0 And ws.Cells(i, sonkol + 2) <> 1 And ws.Cells(i, 1) <> 0 Then
                x = DateAdd("ww", Right(ws.Cells(i, 1).Value, 2), DateSerial(Year(Date), 1, 1)) + 3
                ws.Cells(i, j) = (DateDiff("n", Now, x) / 60) - ektime
                    
                    'Haftalýk plan haftasonu nerde kaldý
                    If ws.Cells(i, j) < 0 Then
                    ws.Cells(i, j).Font.Color = vbRed
                            x = 0
                            plantime = ws.Cells(i, j)
                            Do Until plantime > 0 Or sonkol - x > 4
                                plantime = ws.Cells(5, sonkol - x).Value + plantime
                                x = x + 1
                            Loop
                            ws.Cells(i, sonkol - x + 1).BorderAround ColorIndex:=3, Weight:=xlThick
                    Else
                    ws.Cells(i, j).Font.Color = vbBlack
                    End If
               Else
               ws.Cells(i, j) = ""
               ws.Cells(i, j).Font.Color = vbBlack
               End If
        ElseIf j = sonkol + 2 Then
           ' MsgBox bladeno & istsayi
             ws.Cells(i, j) = istsayi / (sonkol - 3)
             ws.Cells(i, j).NumberFormat = "%0"
             If ws.Cells(i, j) = 1 Then
             ws.Cells(i, j).Interior.ColorIndex = 43
             ws.Cells(i, j).EntireRow.Hidden = True
             Else: ws.Cells(i, j).Interior.ColorIndex = 27
             End If
        ElseIf j = sonkol + 3 Then
             ws.Cells(i, j) = gecikme
        ElseIf j = sonkol + 4 Then
             If ws.Cells(i, sonkol + 2) <> 1 Then
             ws.Cells(i, j) = (DateDiff("n", bas, Now) / 60) + ektime
             Else
             ws.Cells(i, j) = (DateDiff("n", bit, bas) / 60)
             End If
        ElseIf j = sonkol + 5 Then
             ws.Cells(i, j) = 168 / (ws.Cells(i, j - 1) / slotsayisi) / 3
    End If
    Next j
    gecikme = 0
    protime = 0
    ektime = 0
    istsayi = 0
hop:
Next i

ws.Range(Cells(7, 1), Cells(i - 1, sonkol + 6)).Borders.LineStyle = xlContinuous ' ColorIndex:=1, Weight:=xlThin

    sqlvak = "SELECT Vak FROM Mod"
    Set rs2 = New ADODB.Recordset
    rs2.Open sqlvak, cnn
    ThisWorkbook.Worksheets("Vardiya Aktarýmý").Cells(2, 1).Value = rs2.Fields("Vak").Value


sqlgecikme = "SELECT Kanat, Prod.[Operasyon Tanýmý], Prod.[Ekip], Prod.[Fark], Prod.[Açýklama]" & _
      "FROM Prod WHERE Prod.[Açýklama] IS NOT NULL " & _
      "ORDER BY Prod.Kanat"

Set rsgecikme = New ADODB.Recordset
rsgecikme.Open sqlgecikme, cnn
If rsgecikme.EOF And rsgecikme.BOF Then
    rsgecikme.Close
    Set rsgecikme = Nothing
    MsgBox "Herhangi bir kayýt bulunamadý", vbCritical, "No Records"
    With ws4.ListObjects("Tablo1")
    .AutoFilter.ShowAllData
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Rows.Delete
    End If
End With
End If
'On Error Resume Next
With ws4.ListObjects("Tablo1")
    .AutoFilter.ShowAllData
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Rows.Delete
    End If
    Call .Range(2, 1).CopyFromRecordset(rsgecikme)
End With


sqlct = "SELECT Prod.Kanat, Prod.Ekip, Min(Prod.Baþlangýç) AS EnAzBaþlangýç, Max(Prod.Bitiþ) AS EnÇokBitiþ FROM Prod GROUP BY Prod.Kanat, Prod.Ekip" ' HAVING (((Prod.Ekip)<>""))"
Set rsct = New ADODB.Recordset
rsct.Open sqlct, cnn
If rsct.EOF And rsct.BOF Then
    rsct.Close
    Set rsct = Nothing
    MsgBox "Herhangi bir kayýt bulunamadý", vbCritical, "No Records"
    With ws5.ListObjects("Tablo2")
    .AutoFilter.ShowAllData
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Rows.Delete
    End If
End With
End If
'On Error Resume Next
With ws5.ListObjects("Tablo2")
    .AutoFilter.ShowAllData
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Rows.Delete
    End If
    Call .Range(2, 1).CopyFromRecordset(rsct)
End With

ws.Cells(1, 16).Value = "Vestas Finish Tracker"
ws.Cells(1, 26).Value = "Güncelleme : " & Now()
'Call aUTorefresh
On Error Resume Next
rs.Close
rs2.Close
rs3.Close
rsvers.Close
cnn.Close
Set rs = Nothing
Set rs2 = Nothing
Set rs3 = Nothing
Set rsvers = Nothing
Set cnn = Nothing
With ws.Range(Cells(7, 1), Cells(lastrow, sonkol + 5))
.HorizontalAlignment = xlCenter
.Font.Size = 9
End With
ws.Columns("D:AL").AutoFit
ws4.Columns("A:E").AutoFit
Application.AutoCorrect.AutoExpandListRange = True
Application.Calculation = xlAutomatic
ws.Range("A7:B" & lastrow).Locked = False
ws.Range("AV7:AV" & lastrow).Locked = False
ws4.Protect Password:="1709"
ws.Protect Password:="1709", UserInterfaceOnly:=True

Exit Sub
err:
    MsgBox "Ýnternete baðlý olduðunu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran görüntüsü at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "Sýkýntý var :("
    On Error Resume Next
    Set rs = Nothing
    Exit Sub
End Sub

Public Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
Dim element As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function
Sub aUTorefresh()
    yenile = False
    ' If yenile = False Then Exit Sub Else Application.OnTime Now + TimeValue("00:05:00"), "refresh"
End Sub
