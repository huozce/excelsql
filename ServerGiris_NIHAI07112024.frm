VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ServerGiris 
   Caption         =   "UserForm2"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9105.001
   OleObjectBlob   =   "ServerGiris_NIHAI07112024.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ServerGiris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub CommandButton1_Click()
    Dim strConn As String
    
    ' Check if username and password are provided
    If txtUsername.Text = "" And txtPassword.Text = "" Then
        ' Use Windows Authentication if no username and password are provided
        strConn = "Provider=SQLOLEDB;Data Source=" & txtServer.Text & ";Integrated Security=SSPI;"
    Else
        ' Use SQL Server Authentication if username and password are provided
        strConn = "Provider=SQLOLEDB;Data Source=" & txtServer.Text & ";User ID=" & txtUsername.Text & ";Password=" & txtPassword.Text & ";"
    End If
    
    ' Attempt connection
    On Error Resume Next
    Set conn = New ADODB.Connection
    conn.Open strConn
    On Error GoTo 0
    
    If conn.State = adStateOpen Then
        MsgBox "Connected successfully!", vbInformation
        Set conn = Nothing
        ' Navigate to EkleSilUserForm if connection is successful
        UserForm1.Show
        Unload Me
    Else
        MsgBox "Connection failed. Please check your details.", vbExclamation
        ' Still navigate to the next form if connection fails
        UserForm1.Show
        Unload Me
    End If
End Sub

Public username As String
Public password As String

Private Sub btnConnect_Click()
    Dim strConn As String
    
    ' Check if username and password are provided
    If txtUsername.Text = "" And txtPassword.Text = "" Then
        ' Use Windows Authentication if no username and password are provided
        strConn = "Provider=SQLOLEDB;Data Source=" & txtServer.Text & ";Integrated Security=SSPI;"
    Else
    
       
        ' Use SQL Server Authentication if username and password are provided
        strConn = "Provider=SQLOLEDB;Data Source=" & txtServer.Text & ";User ID=" & txtUsername.Text & ";Password=" & txtPassword.Text & ";"
    End If
    
    ' Attempt connection
    On Error Resume Next
    Set conn = New ADODB.Connection
    conn.Open strConn
    On Error GoTo 0
    
    If conn.State = adStateOpen Then
        MsgBox "Connected successfully!", vbInformation
        Set conn = Nothing
        ' Navigate to EkleSilUserForm if connection is successful
        EkleSilGuncelleUserForm.Show
        Unload Me
    Else
        MsgBox "Connection failed. Please check your details.", vbExclamation
        ' Still navigate to the next form if connection fails
        EkleSilGuncelleUserForm.Show
        Unload Me
    End If
End Sub



Private Sub btnConnectTopla_Click()
    Dim strConn As String
    
    ' Check if username and password are provided
    If txtUsername.Text = "" And txtPassword.Text = "" Then
        ' Use Windows Authentication if no username and password are provided
        strConn = "Provider=SQLOLEDB;Data Source=" & txtServer.Text & ";Integrated Security=SSPI;"
    Else
        ' Use SQL Server Authentication if username and password are provided
        strConn = "Provider=SQLOLEDB;Data Source=" & txtServer.Text & ";User ID=" & txtUsername.Text & ";Password=" & txtPassword.Text & ";"
    End If
    
    ' Attempt connection
    On Error Resume Next
    Set conn = New ADODB.Connection
    Set conn.ConnectionTimeout = 10
    conn.Open strConn
    On Error GoTo 0
    
    If conn.State = adStateOpen Then
        MsgBox "Connected successfully!", vbInformation
        Set conn = Nothing
        ' Navigate to EkleSilUserForm if connection is successful
        SQLDENCEK.Show
        Unload Me
    Else
        MsgBox "Connection failed. Please check your details.", vbExclamation
        ' Still navigate to the next form if connection fails
        SQLDENCEK.Show
        Unload Me
    End If
End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub txtServer_Change()

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
         ServerGiris.txtServer.Value = .Cells(1, 8).Value
        ' txtDatabase.Value = .Cells(1, 2).Value
        ServerGiris.txtUsername.Value = .Cells(1, 10).Value
        ServerGiris.txtPassword.Value = .Cells(1, 11).Value
        ServerGiris.CheckBox1.Value = .Cells(1, 12).Value
       
       
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Save current values to Settings sheet before closing
    Dim wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    With wsSettings
       
       If ServerGiris.CheckBox1.Value = True Then
            .Cells(1, 10).Value = ServerGiris.txtUsername.Value
            .Cells(1, 11).Value = ServerGiris.txtPassword.Value
            .Cells(1, 8).Value = ServerGiris.txtServer.Value
         Else
            ' Clear saved Username and Password if checkbox is not selected
            .Cells(1, 8).ClearContents
            .Cells(1, 10).ClearContents
            .Cells(1, 11).ClearContents
        End If
        
    End With
End Sub

