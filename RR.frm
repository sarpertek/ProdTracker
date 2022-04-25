VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RR 
   Caption         =   "Kanat Demold"
   ClientHeight    =   2796
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3816
   OleObjectBlob   =   "RR.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub adamtext_Change()
If TypeName(Me.ActiveControl) = "TextBox" Then
        With Me.ActiveControl
            If Not IsNumeric(.Value) And .Value <> vbNullString Then
                MsgBox "Sadece sayý girilmeli"
                .Value = vbNullString
            End If
        End With
    End If
End Sub

Private Sub CommandButton1_Click()
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rsvers As ADODB.Recordset
Dim versiyon, sqlvers As String
proje = "X"
Set ws = ThisWorkbook.Worksheets("Finish Status")
ws.Unprotect Password:="1709"
dbPath = ws.Range("B1").Value
    If Date1.Value = "" Or saattext.Value = "" Then
        If RR.Caption <> "Kanat Demold" Then MsgBox "Tüm alanlarý doldurun", vbCritical, "Gerekli Alanlar"
    Exit Sub
    End If
    If RR.Caption = "Operasyon Bitir" And adamtext.Value = "" Then
    MsgBox "Tüm alanlarý doldurun", vbCritical, "Gerekli Alanlar"
    Exit Sub
    End If
    tarih1 = Date1.Value & " " & saattext.Value
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
        
    If RR.Caption = "Kanat Demold" Then
            hoba = False
            Application.ScreenUpdating = False
            Set ws = ThisWorkbook.Worksheets("Finish Status")
            Dim rsset As ADODB.Recordset
            Dim rst As ADODB.Recordset
            Dim varMyArray As Variant
            Set rst = New ADODB.Recordset
            Set rsset = New ADODB.Recordset
            dbPath = ws.Range("B1").Value
            Set cnn = New ADODB.Connection
            lastrow = ws.Range("C" & Rows.Count).End(xlUp).Row
            If lastrow < 7 Then lastrow = 6
            Set cnn = New ADODB.Connection
            cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=Sarper1709;"
            sqlset = "SELECT TOP 1 COC.SetSayi, Count(COC.SetSayi) AS SaySetSayi FROM COC GROUP BY COC.SetSayi ORDER BY COC.Setsayi Desc"
            rsset.Open sqlset, cnn
            rst.Open Source:="COC", ActiveConnection:=cnn, _
            CursorType:=adOpenDynamic, LockType:=adLockOptimistic, _
            Options:=adCmdTable
            With rst
            .AddNew
            .Fields("Kanat").Value = adamtext.Value
            .Fields("COC").Value = "NO"
            If rsset.Fields(1).Value >= 3 Then .Fields("Setsayi") = rsset.Fields(0).Value + 1 Else .Fields("Setsayi") = rsset.Fields(0).Value
            .Fields("Sýra").Value = lastrow + 1
            .Update
            End With
            ws.Cells(lastrow + 1, 3).Value = cell
            tarih2 = tarih1 + (1 / 48)
            sonkanat = adamtext.Value
            sql1 = "UPDATE Ops SET [Kanat] = " & sonkanat
            sql2 = "INSERT INTO Prod ([Kanat], [Operasyon No], [Operasyon Tanýmý], [Hedef], [Ekip]) SELECT [Kanat], [Operasyon No], [Operasyon Tanýmý], [Hedef], [Ekip] from Ops ORDER BY [Operasyon No]"
            sql3 = "UPDATE Prod SET [Baþlangýç] = '" & tarih1 & "', [Bitiþ] = '" & tarih2 & "' WHERE Kanat = " & sonkanat & " AND [Operasyon No] = " & 0
            proje = "X"
            'bassayi = "TE Insert Trim"
            'sql4 = "UPDATE Prod SET [Baþlangýç] = '" & tarih1 & "', [Vardiya2] = '" & proje & "', [Bitiþ] = '" & tarih2 & "' WHERE Kanat = " & sonkanat & " AND [Operasyon Tanýmý] = '" & bassayi & "' "
            '''bassayi = "TE Insert"
            '''sql5 = "UPDATE Prod SET [Baþlangýç] = '" & tarih1 & "', [Vardiya2] = '" & proje & "', [Bitiþ] = '" & tarih2 & "'  WHERE Kanat = " & sonkanat & " AND [Operasyon Tanýmý] = '" & bassayi & "' "
            'sql5 = "UPDATE Prod SET [Baþlangýç] = '" & tarih1 & "', [Bitiþ] = '" & tarih2 & "' WHERE Kanat = " & sonkanat & " AND [Operasyon No] < " & bassayi & "" '([Operasyon No] = '0' OR [Operasyon No] = '1' OR [Operasyon No] = '2') "
            cnn.Execute sql1
            cnn.Execute sql2
            cnn.Execute sql3
            'cnn.Execute sql4
            '''cnn.Execute sql5
            cnn.Close
         'End If
         lastrow = ws.Range("C" & Rows.Count).End(xlUp).Row
     
     ElseIf RR.Caption = "Operasyon Baþlat" Then
            kanatno = ws.Cells(idrow, 3)
            opno = ws.Cells(3, idcol)
            sql1 = "UPDATE Prod SET [Baþlangýç] = '" & tarih1 & "', [Vardiya1] = '" & ComboBox1.Value & "', [Vardiya2] = '" & proje & "' WHERE Kanat = " & kanatno & " AND [Operasyon No] = " & opno & " "
            cnn.Execute sql1
     ElseIf RR.Caption = "Operasyon Bitir" Then
            kanatno = ws.Cells(idrow, 3)
            opno = ws.Cells(3, idcol)
            
            SQL = "SELECT Baþlangýç, Hedef FROM Prod WHERE Kanat = " & kanatno & " AND [Operasyon No] = " & opno & " "
            Set rs = New ADODB.Recordset
            rs.Open SQL, cnn
            If (tarih1 - rs.Fields(0).Value) * 24 > rs.Fields(1).Value Then
                    'Reason.Show
                        reas = ((tarih1 - rs.Fields(0).Value) * 24) - rs.Fields(1).Value
                        reas = reas & " saatlik gecikme nedeni"
                        sonuc = InputBox("Gecikme nedenini yazýn", reas)
                        If sonuc <> "" Then
                        sql1 = "UPDATE Prod SET [Bitiþ] = '" & tarih1 & "', [Açýklama] = '" & sonuc & "' WHERE Kanat = " & kanatno & " AND [Operasyon No] = " & opno & " "
                        cnn.Execute sql1
                        Else
                            MsgBox "Gecikme nedeni belirt"
                            Exit Sub
                        End If
            End If
            sql2 = "UPDATE Prod SET [Bitiþ] = '" & tarih1 & "', [Vardiya2] = '" & ComboBox1.Value & "', [Adam] = " & adamtext.Value & " WHERE Kanat = " & kanatno & " AND [Operasyon No] = " & opno & " "
            cnn.Execute sql2
     End If
     
On Error Resume Next
    rs.Close
    rs2.Close
    rs3.Close
    rsvers.Close
    cnn.Close
    Set rs = Nothing
    Set cnn = Nothing
   
Call refresh
Unload Me
End Sub

Private Sub Date1_Change()
RemoveChar = 1
If Len(Date1.Value) = "11" Then
            Date1.Value = Left(Date1.Value, Len(Date1.Value) - RemoveChar)
End If
If Len(Date1.Value) = "1" Or Len(Date1.Value) = "4" Or Len(Date1.Value) = "7" Or Len(Date1.Value) = "8" Or Len(Date1.Value) = "9" Or Len(Date1.Value) = "10" Then
    If Not IsNumeric(Right(Date1.Value, 1)) Then
            If Len(Date1.Value) < 1 Then RemoveChar = 0
            Date1.Value = Left(Date1.Value, Len(Date1.Value) - RemoveChar)
    End If
End If
If Len(Date1.Value) = "2" Then
    If Right(Date1.Value, 2) < 32 Then
    Else
    MsgBox "Gün 31'den fazla olamaz!"
    Date1.Value = ""
    End If
End If
If Len(Date1.Value) = "5" Then
    If Right(Date1.Value, 2) < 13 Then
    Else
    MsgBox "Ay 12'den fazla olamaz!"
    Date1.Value = ""
    End If
End If
If Len(Date1.Value) = "3" Or Len(Date1.Value) = "6" Then
        If Not Right(Date1.Value, 1) = "." Then
            If Len(Date1.Value) < 1 Then RemoveChar = 0
            Date1.Value = Left(Date1.Value, Len(Date1.Value) - RemoveChar)
        End If
End If
End Sub




Private Sub saattext_Change()
RemoveChar = 1
If Len(saattext.Value) = "6" Then
            saattext.Value = Left(saattext.Value, Len(saattext.Value) - RemoveChar)
End If
If Len(saattext.Value) <> "3" Then
    If Not IsNumeric(Right(saattext.Value, 1)) Then
            If Len(saattext.Value) < 1 Then RemoveChar = 0
            saattext.Value = Left(saattext.Value, Len(saattext.Value) - RemoveChar)
    End If
End If
If Len(saattext.Value) = "2" Then
    If Right(saattext.Value, 2) < 25 Then
    Else
    MsgBox "Saat 24'den fazla olamaz!"
    saattext.Value = ""
    End If
End If
If Len(saattext.Value) = "5" Then
    If Right(saattext.Value, 2) < 60 Then
    Else
    MsgBox "Dakika 60'dan fazla olamaz!"
    saattext.Value = ""
    End If
End If
If Len(saattext.Value) = "3" Then
        If Not Right(saattext.Value, 1) = ":" Then
            If Len(saattext.Value) < 1 Then RemoveChar = 0
            saattext.Value = Left(saattext.Value, Len(saattext.Value) - RemoveChar)
        End If
End If
End Sub
Private Sub UserForm_Initialize()
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Finish Status")
ComboBox1.List = Array("A", "B", "C", "D")
    If hoba = False Then
    RR.Caption = "Kanat Demold"
    Label5.Caption = "Demold Tarihi"
    Label3.Caption = "Demold Saati"
    Label4.Caption = ""
    CommandButton1.Caption = "Demold"
    adamlabel.Caption = "Kanat No"
    adamlabel.Visible = True
    adamtext.Visible = True
    Label9.Visible = False
    ComboBox1.Visible = False
    ElseIf targetval = "" Then
    RR.Caption = "Operasyon Baþlat"
    Label5.Caption = ws.Cells(6, idcol) & " Baþlangýç Tarihi"
    Label3.Caption = ws.Cells(6, idcol) & " Baþlangýç Saati"
    Label4.Caption = "Kanat " & ws.Cells(idrow, 3)
    CommandButton1.Caption = "Baþlat"
    ElseIf targetval <> "" Then
    RR.Caption = "Operasyon Bitir"
    adamlabel.Visible = True
    adamtext.Visible = True
    Label5.Caption = ws.Cells(6, idcol) & " Bitiþ Tarihi"
    Label3.Caption = ws.Cells(6, idcol) & " Bitiþ Saati"
    Label4.Caption = "Kanat " & ws.Cells(idrow, 3)
    CommandButton1.Caption = "Bitir"
    End If
    
End Sub

 
