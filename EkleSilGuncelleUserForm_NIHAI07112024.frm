VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EkleSilGuncelleUserForm 
   Caption         =   "UserForm1"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14985
   OleObjectBlob   =   "EkleSilGuncelleUserForm_NIHAI07112024.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EkleSilGuncelleUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnAdd_Click()
     ' Kullanýcýya dikkatli olmasý gerektiðini bildiren uyarý
    Dim warningResponse As VbMsgBoxResult
    warningResponse = MsgBox("Bu iþlem seçili satýrlardaki verileri SQL veritabanýna ekleyecektir. Devam etmek istiyor musunuz?", vbExclamation + vbYesNo, "Uyarý")

    ' Kullanýcý "Hayýr" seçerse makroyu sonlandýr
    If warningResponse = vbNo Then Exit Sub ' END IFIYOK

    ' Deðiþkenleri tanýmla
    Dim conn As Object
    Dim strConn As String
    Dim ws As Worksheet
    Dim serverName As String
    Dim databaseName As String
    Dim tableName As String
    Dim primaryKey As String
    Dim columnNames As String
    Dim columnValues As String
    Dim sqlInsert As String
    Dim cell As Range
    Dim i As Long, j As Long
    Dim lastColumn As Long
    Dim hasUpdatedDateTime As Boolean
    Dim currentDateTime As String

    ' Çalýþma sayfasýný ayarla
    Set ws = ActiveSheet

   ' Get connection info from ServerGiris form
    serverName = ServerGiris.txtServer.Text
    databaseName = txtDatabaseName.Text
    tableName = txtTableName.Text

    ' Connection string with ServerGiris values
   If ServerGiris.txtUsername.Text = "" And ServerGiris.txtPassword.Text = "" Then
        ' Use Windows Authentication
        strConn = "Provider=SQLOLEDB;Data Source=" & serverName & ";Initial Catalog=" & databaseName & ";Integrated Security=SSPI;"
    Else
    strConn = "Provider=SQLOLEDB;Data Source=" & ServerGiris.txtServer.Text & ";Initial Catalog=" & databaseName & ";" & _
              "User ID=" & ServerGiris.txtUsername.Text & ";Password=" & ServerGiris.txtPassword.Text & ";"

    End If
    ' ADODB baðlantýsýný baþlat
    Set conn = CreateObject("ADODB.Connection")
    On Error GoTo ConnectionError
     conn.ConnectionTimeout = 5
    conn.Open strConn
    On Error GoTo 0

    ' Sütun isimlerini elde etmek için baþlýk satýrýný kontrol et
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    hasUpdatedDateTime = False
    columnNames = ""

    For j = 1 To lastColumn
        Dim columnName As String
        columnName = ws.Cells(1, j).Value
        
        ' UpdatedDateTime sütunu olup olmadýðýný kontrol et
        If LCase(columnName) = "updateddatetime" Then
            hasUpdatedDateTime = True
        Else
            columnNames = columnNames & "[" & columnName & "], "
        End If
    Next j

    ' Eðer UpdatedDateTime varsa, sütun adlarýna ekle
    If hasUpdatedDateTime Then
        columnNames = columnNames & "[UpdatedDateTime], "
    End If
    columnNames = Left(columnNames, Len(columnNames) - 2) ' Son virgülü kaldýr

    ' Þu anki tarih ve saat
    currentDateTime = Format(Now, "yyyy-mm-dd hh:nn:ss")

    ' Her seçili hücre için yeni satýr ekleyin
    For Each cell In Selection
        ' Sadece satýrýn ilk hücresi seçilmiþse iþlem yap
        If cell.Column = Selection(1).Column Then
            columnValues = ""
            ' Seçili satýrdaki her sütun hücresini oku
            For j = 1 To lastColumn
                If LCase(ws.Cells(1, j).Value) <> "updateddatetime" Then
                    columnValues = columnValues & "'" & ws.Cells(cell.row, j).Text & "', "
                End If
            Next j

            ' Eðer UpdatedDateTime varsa currentDateTime ekle
            If hasUpdatedDateTime Then
                columnValues = columnValues & "'" & currentDateTime & "', "
            End If

            columnValues = Left(columnValues, Len(columnValues) - 2) ' Son virgülü kaldýr

            ' Ekleme SQL ifadesini oluþtur
            sqlInsert = "INSERT INTO " & tableName & " (" & columnNames & ") VALUES (" & columnValues & ")"
            conn.Execute sqlInsert
        End If
    Next cell

    ' Excel'deki mevcut verileri temizle (ilk satýr baþlýk olarak kalýr)
    ws.Rows("2:" & ws.Rows.Count).ClearContents

    ' SQL'den yeni verileri al
    Dim sqlSelect As String
    Dim rs As Object
    sqlSelect = "SELECT * FROM " & tableName
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlSelect, conn

    ' Yeni verileri Excel'e yaz
    If Not rs.EOF Then
        ws.Range("A2").CopyFromRecordset rs
    End If

    ' Kayýtlarý kapat ve baðlantýyý sonlandýr
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    MsgBox "Seçili satýrlardaki veriler SQL'e eklendi ve yeni tablo güncellendi!"

    Exit Sub

ConnectionError:
    MsgBox "Veritabanýna baðlanýrken bir hata oluþtu: " & Err.Description
    If Not conn Is Nothing Then
    conn.Close
    Set conn = Nothing
    End If
End Sub
Private Sub btnDelete_Click()
    ' Kullanýcýya dikkatli olmasý gerektiðini bildiren uyarý
    Dim warningResponse As VbMsgBoxResult
    warningResponse = MsgBox("Bu iþlem seçili satýrlardaki kayýtlarý SQL veritabanýndan silecektir. Devam etmek istiyor musunuz?", vbExclamation + vbYesNo, "Uyarý")

    ' Kullanýcý "Hayýr" seçerse makroyu sonlandýr
    If warningResponse = vbNo Then Exit Sub ' END IFIYOK

    ' Deðiþkenleri tanýmla
    Dim conn As Object
    Dim strConn As String
    Dim ws As Worksheet
    Dim primaryKey As String
    Dim primaryKeyValue As String
    Dim cell As Range
    Dim sqlDelete As String
    Dim sqlSelect As String
    Dim rs As Object

    ' Çalýþma sayfasýný ayarla
    Set ws = ActiveSheet

    ' Sunucu, veritabaný adý, tablo adý ve birincil anahtar bilgilerini al
    Dim serverName As String, databaseName As String, tableName As String
    ' Get connection info from ServerGiris form
    serverName = ServerGiris.txtServer.Text
    databaseName = txtDatabaseName.Text
    tableName = txtTableName.Text
    primaryKey = txtPrimaryKey.Text

    ' Connection string with ServerGiris values
  If ServerGiris.txtUsername.Text = "" And ServerGiris.txtPassword.Text = "" Then
        ' Use Windows Authentication
        strConn = "Provider=SQLOLEDB;Data Source=" & serverName & ";Initial Catalog=" & databaseName & ";Integrated Security=SSPI;"
  Else
    strConn = "Provider=SQLOLEDB;Data Source=" & ServerGiris.txtServer.Text & ";Initial Catalog=" & databaseName & ";" & _
              "User ID=" & ServerGiris.txtUsername.Text & ";Password=" & ServerGiris.txtPassword.Text & ";"
  End If

    ' ADODB baðlantýsýný baþlat
    Set conn = CreateObject("ADODB.Connection")
     On Error GoTo ConnectionError
    conn.Open strConn
    On Error GoTo 0

    ' SQL toplu silme iþlemi için sorgu baþlat
    sqlDelete = "DELETE FROM " & tableName & " WHERE [" & primaryKey & "] IN ("

    ' Seçili hücre aralýðýný kontrol et
    For Each cell In Selection
        ' Sadece satýrýn ilk hücresi seçilmiþse iþlem yap
        If cell.Column = Selection(1).Column Then
            primaryKeyValue = ws.Cells(cell.row, cell.Column).Value
            If Len(primaryKeyValue) > 0 Then
                sqlDelete = sqlDelete & "'" & primaryKeyValue & "', "
            End If
        End If
    Next cell

    ' Sorgunun sonundaki virgül ve boþluðu kaldýr, ardýndan kapat
    sqlDelete = Left(sqlDelete, Len(sqlDelete) - 2) & ")"

    ' Kayýtlarý SQL'de sil
    conn.Execute sqlDelete

    ' Excel'deki mevcut verileri temizle (ilk satýr baþlýk olarak kalýr)
    ws.Rows("2:" & ws.Rows.Count).ClearContents

    ' SQL'den yeni verileri al
    sqlSelect = "SELECT * FROM " & tableName
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlSelect, conn

    ' Yeni verileri Excel'e yaz
    If Not rs.EOF Then
        ws.Range("A2").CopyFromRecordset rs
    End If

    ' Kayýtlarý kapat ve baðlantýyý sonlandýr
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    MsgBox "Seçili satýrlara karþýlýk gelen SQL kayýtlarý silindi ve yeni veriler güncellendi!"

    Exit Sub

ConnectionError:
    MsgBox "Veritabanýna baðlanýrken bir hata oluþtu: " & Err.Description
    If Not conn Is Nothing Then
        conn.Close
    Set conn = Nothing
    End If
End Sub







Private Sub TextBox1_Change()

End Sub


Private Sub btnBackup_Click()
    ' Retrieve server and database details from UserForm
    Dim serverName As String
    Dim databaseName As String
    Dim tableName As String
    Dim backupTableName As String
    Dim strConn As String
    Dim conn As Object
    Dim formattedDate As String

   serverName = ServerGiris.txtServer.Text
    databaseName = txtDatabaseName.Text
    tableName = txtTableName.Text

    ' Connection string with ServerGiris values
 If ServerGiris.txtUsername.Text = "" And ServerGiris.txtPassword.Text = "" Then
        ' Use Windows Authentication
        strConn = "Provider=SQLOLEDB;Data Source=" & serverName & ";Initial Catalog=" & databaseName & ";Integrated Security=SSPI;"
    Else
    strConn = "Provider=SQLOLEDB;Data Source=" & ServerGiris.txtServer.Text & ";Initial Catalog=" & databaseName & ";" & _
              "User ID=" & ServerGiris.txtUsername.Text & ";Password=" & ServerGiris.txtPassword.Text & ";"
End If

    ' Initialize ADODB connection
    Set conn = CreateObject("ADODB.Connection")
    
    On Error GoTo ConnectionError
    conn.ConnectionTimeout = 5
    
    conn.Open strConn
    On Error GoTo 0

    ' Get the current date and time to create a unique backup table name
    formattedDate = Format(Now, "yyyymmdd_HHmmss")
    backupTableName = tableName & "_Backup_" & formattedDate

    ' Create a backup of the table
    Dim sqlBackup As String
    sqlBackup = "SELECT * INTO " & backupTableName & " FROM " & tableName

    On Error GoTo BackupError
    conn.Execute sqlBackup

    ' Close the connection
    conn.Close
    Set conn = Nothing

    MsgBox "Backup created successfully as: " & backupTableName
    Exit Sub

ConnectionError:
    MsgBox "Error connecting to the database: " & Err.Description
    If Not conn Is Nothing Then
    conn.Close ' END IFIYOK
    Set conn = Nothing
    Exit Sub
    End If
BackupError:
    MsgBox "Error creating backup: " & Err.Description
    If Not conn Is Nothing Then
    conn.Close
    Set conn = Nothing
    End If
End Sub


Private Sub topluGuncelle_Click()
    ' Retrieve server and database details from UserForm
    Dim serverName As String
    Dim databaseName As String
    Dim primaryKey As String
    Dim tableName As String
    
   serverName = ServerGiris.txtServer.Text
    databaseName = txtDatabaseName.Text
    tableName = txtTableName.Text
    primaryKey = txtPrimaryKey.Text

  
    Dim conn As Object
    Dim strConn As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sqlInsert As String
    Dim sqlUpdate As String
    Dim primaryKeyValue As String
    Dim i As Long, j As Long
    Dim columnNames As String
    Dim columnValues As String
    Dim updateSet As String
    Dim recordExists As Boolean
    Dim rs As Object
    Dim formattedDate As String
    Dim hasDateTimeColumn As Boolean
    Dim userResponse As VbMsgBoxResult
    Dim backupTableName As String
    Dim confirmUpdate As VbMsgBoxResult
    Dim viewUpdatedData As VbMsgBoxResult
    
    ' Set worksheet (assuming ActiveSheet)
    Set ws = ActiveSheet
  ' Connection string with ServerGiris values
    If ServerGiris.txtUsername.Text = "" And ServerGiris.txtPassword.Text = "" Then
        ' Use Windows Authentication
        strConn = "Provider=SQLOLEDB;Data Source=" & serverName & ";Initial Catalog=" & databaseName & ";Integrated Security=SSPI;"
    Else
    strConn = "Provider=SQLOLEDB;Data Source=" & ServerGiris.txtServer.Text & ";Initial Catalog=" & databaseName & ";" & _
              "User ID=" & ServerGiris.txtUsername.Text & ";Password=" & ServerGiris.txtPassword.Text & ";"
    End If
    ' Initialize ADODB connection
    Set conn = CreateObject("ADODB.Connection")
    
     On Error GoTo ConnectionError
     conn.ConnectionTimeout = 5
    conn.Open strConn
    On Error GoTo 0

    ' Get the current date and time in 24-hour format
    formattedDate = Format(Now, "yyyy-mm-dd HH:mm:ss")

    ' Check if "UpdatedDateTime" column exists in SQL table
    hasDateTimeColumn = False
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & tableName & "' AND COLUMN_NAME = 'UpdatedDateTime'", conn
    If Not rs.EOF Then hasDateTimeColumn = True
    rs.Close

    ' Add UpdatedDateTime column if it doesn't exist
    If Not hasDateTimeColumn Then conn.Execute "ALTER TABLE " & tableName & " ADD UpdatedDateTime DATETIME" ' END IFIYOK

    ' Ask the user if they want to create a backup
    userResponse = MsgBox("Veri güncellenmeden önce bir yedekleme almak ister misiniz?", vbYesNo + vbQuestion, "Yedekleme Al")
    
    If userResponse = vbYes Then
        ' Create backup table
        backupTableName = tableName & "_Backup_" & Format(Now, "yyyymmdd_HHmmss")
        conn.Execute "SELECT * INTO " & backupTableName & " FROM " & tableName
    End If

    ' Find the last row in the worksheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' Confirmation before updating the database
    confirmUpdate = MsgBox("Veritabanýný güncellemek istediðinize emin misiniz?", vbYesNo + vbExclamation, "Güncelleme Onayý")
    
    If confirmUpdate = vbNo Then
        MsgBox "Güncelleme iþlemi iptal edildi.", vbInformation
        conn.Close
        Set conn = Nothing
        Exit Sub
    End If

    ' Loop through rows and insert/update records in SQL
    For i = 2 To lastRow ' Assume headers are in row 1
        columnNames = ""
        columnValues = ""
        updateSet = ""
        
        ' Loop through columns to build SQL statements
        For j = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If ws.Cells(1, j).Value = primaryKey Then
                primaryKeyValue = ws.Cells(i, j).Value
            Else
                columnNames = columnNames & "[" & ws.Cells(1, j).Value & "], "
                columnValues = columnValues & "'" & ws.Cells(i, j).Text & "', "
                updateSet = updateSet & "[" & ws.Cells(1, j).Value & "] = '" & ws.Cells(i, j).Text & "', "
            End If
        Next j
        
        ' Remove trailing commas
        If Len(columnNames) > 0 Then
            columnNames = Left(columnNames, Len(columnNames) - 2)
        End If
        If Len(columnValues) > 0 Then
            columnValues = Left(columnValues, Len(columnValues) - 2)
        End If
        If Len(updateSet) > 0 Then
            updateSet = Left(updateSet, Len(updateSet) - 2)
        End If

        ' Add UpdatedDateTime to the update statement
        updateSet = updateSet & ", [UpdatedDateTime] = '" & formattedDate & "'"

        ' Check if record with the primary key exists
        recordExists = False
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open "SELECT COUNT(*) AS RecordCount FROM " & tableName & " WHERE [" & primaryKey & "] = '" & primaryKeyValue & "'", conn
        If Not rs.EOF Then
            If rs("RecordCount") > 0 Then recordExists = True
        End If
        rs.Close
        
        ' Insert or update the record
        If recordExists Then
            sqlUpdate = "UPDATE " & tableName & " SET " & updateSet & " WHERE [" & primaryKey & "] = '" & primaryKeyValue & "'"
            conn.Execute sqlUpdate
        Else
            sqlInsert = "INSERT INTO " & tableName & " (" & primaryKey & ", " & columnNames & ", [UpdatedDateTime]) VALUES ('" & primaryKeyValue & "', " & columnValues & ", '" & formattedDate & "')"
            conn.Execute sqlInsert
        End If
    Next i

    ' Close the connection
    conn.Close
    Set conn = Nothing

    ' Ask the user if they want to see the updated table
    viewUpdatedData = MsgBox("Veriler baþarýyla güncellendi! Güncellenmiþ tabloyu görmek ister misiniz?", vbYesNo + vbInformation, "Güncelleme Baþarýlý")

    If viewUpdatedData = vbYes Then
        ' Retrieve and display the updated data
        Dim updatedConn As Object
        Set updatedConn = CreateObject("ADODB.Connection")
        updatedConn.Open strConn
        RetrieveAndDisplayData updatedConn, tableName
        updatedConn.Close
        Set updatedConn = Nothing
    End If

    MsgBox "Veriler baþarýyla SQL'e eklendi veya güncellendi!" & IIf(userResponse = vbYes, " Yedekleme alýndý: " & backupTableName, ""), vbInformation
    Exit Sub

ConnectionError:
    MsgBox "Veritabanýna baðlanýrken bir hata oluþtu: " & Err.Description
    If Not conn Is Nothing Then
    conn.Close
    Set conn = Nothing
    End If
End Sub



Private Sub RetrieveAndDisplayData(conn As Object, tableName As String)
    Dim sqlQuery As String
    Dim rs As Object
    Dim ws As Worksheet
    Dim row As Long, col As Long

    ' Set worksheet (use ActiveSheet or specify another sheet)
    Set ws = ActiveSheet
    ws.Cells.Clear ' Clear existing content if you want to refresh the display

    ' SQL query to retrieve all data from the table
    sqlQuery = "SELECT * FROM " & tableName

    ' Open recordset
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlQuery, conn

    ' Check if the recordset is empty
    If rs.EOF Then
        MsgBox "Veri bulunamadý!", vbInformation
        rs.Close
        Exit Sub
    End If

    ' Populate headers in the first row
    For col = 0 To rs.Fields.Count - 1
        ws.Cells(1, col + 1).Value = rs.Fields(col).Name
    Next col

    ' Populate data starting from the second row
    row = 2
    Do Until rs.EOF
        For col = 0 To rs.Fields.Count - 1
            ws.Cells(row, col + 1).Value = rs.Fields(col).Value
        Next col
        rs.MoveNext
        row = row + 1
    Loop

    ' Clean up
    rs.Close
    Set rs = Nothing
End Sub











Private Sub txtServerName_Change()

End Sub

Private Sub UserForm_Initialize()
    ' Load saved values into text boxes
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    If ws Is Nothing Then
        ' Create a hidden worksheet if it doesn't exist
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Settings"
        ws.Visible = xlSheetVeryHidden
    End If
    On Error GoTo 0

    ' Load values from the hidden worksheet into text boxes
    txtTableName.Text = ws.Range("A5").Value
    txtDatabaseName.Text = ws.Range("A9").Value
    ' ServerGiris.txtServer.Text = ws.Range("A3").Value
    txtPrimaryKey.Text = ws.Range("A4").Value
     
End Sub

Private Sub UserForm_Terminate()
    ' Save text box values when the form is closed
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Settings")

    ws.Range("A5").Value = txtTableName.Text
    ws.Range("A9").Value = txtDatabaseName.Text
     ' ws.Range("A3").Value = ServerGiris.txtServer.Text
    ws.Range("A4").Value = txtPrimaryKey.Text
    
End Sub


