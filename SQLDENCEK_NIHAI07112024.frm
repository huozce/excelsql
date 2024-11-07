VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SQLDENCEK 
   Caption         =   "UserForm1"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19545
   OleObjectBlob   =   "SQLDENCEK_NIHAI07112024.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SQLDENCEK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CheckBox1_Click()

End Sub



Private Sub btnFetchData_Click()
    ' Deðiþkenlerin tanýmý
    Dim conn As Object, rs As Object
    Dim strConn As String, sql As String, filterSql As String
    Dim ws As Worksheet
    Dim maxRows As Long
    Dim i As Long
    Dim filterValues As String
    Dim valueArray() As String
    Dim j As Long
    Dim formattedValues As String
    
    ' Çalýþma sayfasýný tanýmlama
    Set ws = ActiveSheet
    
    If ServerGiris.txtUsername.Text = "" And ServerGiris.txtPassword.Text = "" Then
        ' Use Windows Authentication
        strConn = "Provider=SQLOLEDB;Data Source=" & ServerGiris.txtServer.Text & ";Initial Catalog=" & txtDatabaseName.Text & ";Integrated Security=SSPI;"
    Else
    strConn = "Provider=SQLOLEDB;Data Source=" & ServerGiris.txtServer.Text & ";Initial Catalog=" & databaseName & ";" & _
              "User ID=" & ServerGiris.txtUsername.Text & ";Password=" & ServerGiris.txtPassword.Text & ";"

    End If
    Dim tableName As String: tableName = txtTableName.Value
    Dim selectedColumns As String: selectedColumns = IIf(UCase(txtSelectedColumns.Value) = "S", "*", txtSelectedColumns.Value)
    
    ' Toplam sýra inputunu ayarla
    If UCase(txtMaxRows.Value) = "S" Then
        maxRows = 0 ' Tüm sýralarý çek
    Else
        maxRows = CLng(txtMaxRows.Value)
    End If

    ' Filtre þartlarý
    If chkApplyFilter.Value = True Then
        ' Process the filter value for multiple values separated by &
        filterValues = Trim(txtFilterValue.Value)
        valueArray = Split(filterValues, "&") ' Split the input into an array
        
        ' Format the values for SQL IN clause
        formattedValues = ""
        For j = LBound(valueArray) To UBound(valueArray)
            formattedValues = formattedValues & "'" & Trim(valueArray(j)) & "'" ' Add quotes around each value
            If j < UBound(valueArray) Then
                formattedValues = formattedValues & "," ' Add a comma if it's not the last value
            End If
        Next j
        
        ' Construct the SQL filter condition
        filterSql = " WHERE " & txtFilterColumn.Value & " IN (" & formattedValues & ")"
    Else
        filterSql = ""
    End If
    
    ' SQL sorgusunu seçilmiþ seçeneklere göre ayarla
    If maxRows > 0 Then
        sql = "SELECT TOP " & maxRows & " " & selectedColumns & " FROM " & tableName & filterSql
    Else
        sql = "SELECT " & selectedColumns & " FROM " & tableName & filterSql
    End If
    
    ' ADODB oluþturup baðlantýyý aç
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    On Error GoTo ConnectionError
    
    conn.Open strConn
    On Error GoTo 0
    
    ' Sorguyu iþle veriyi çek
    rs.Open sql, conn
    ws.Cells.Clear
    
    ' Sayfaya veriyle doldu
    For i = 0 To rs.Fields.Count - 1
        ws.Cells(1, i + 1).Value = rs.Fields(i).Name
    Next i
    ws.Range("A2").CopyFromRecordset rs

    ' Sil
    rs.Close
    conn.Close
    Set rs = Nothing
    MsgBox "Tablo '" & tableName & "' verileri baþarýyla çalýþma sayfasýna yüklendi!"

    Exit Sub

ConnectionError:
    MsgBox "Veritabanýna baðlanýrken hata oluþtu: " & Err.Description
    If Not conn Is Nothing Then conn.Close
    Set conn = Nothing
    
End Sub





Private Sub chkApplyFilter_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub txtDatabaseName_Change()

End Sub

Private Sub txtFilterColumn_Change()

End Sub

Private Sub txtFilterValue_Change()

End Sub

Private Sub txtMaxRows_Change()

End Sub

Private Sub txtServerName_Change()

End Sub

Private Sub txtTableName_Change()

End Sub

Private Sub UserForm_Click()

End Sub
Private Sub UserForm_Initialize()
    ' Check if "Settings" sheet exists, if not create and hide it
    Dim wsSettings As Worksheet
    On Error Resume Next
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    On Error GoTo 0

    If wsSettings Is Nothing Then
        Set wsSettings = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsSettings.Name = "Settings"
        wsSettings.Visible = xlSheetVeryHidden
    End If

    ' Load saved values from Settings sheet
    With wsSettings
        ' ServerGiris.txtServer.Value = .Cells(1, 1).Value
        ' txtDatabase.Value = .Cells(1, 2).Value
        txtTableName.Value = .Cells(1, 3).Value
        txtSelectedColumns.Value = .Cells(1, 4).Value
        txtMaxRows.Value = .Cells(1, 5).Value
        txtFilterColumn.Value = .Cells(1, 6).Value
        txtFilterValue.Value = .Cells(1, 7).Value
        chkApplyFilter.Value = .Cells(1, 8).Value
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Save current values to Settings sheet before closing
    Dim wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    With wsSettings
        ' .Cells(1, 1).Value = txtServer.Value
         ' .Cells(1, 2).Value = ServerGiris.txtDatabase.Value
        .Cells(1, 3).Value = txtTableName.Value
        .Cells(1, 4).Value = txtSelectedColumns.Value
        .Cells(1, 5).Value = txtMaxRows.Value
        .Cells(1, 6).Value = txtFilterColumn.Value
        .Cells(1, 7).Value = txtFilterValue.Value
        .Cells(1, 8).Value = chkApplyFilter.Value
    End With
End Sub

