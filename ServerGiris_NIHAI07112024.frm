VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ServerGiris 
   Caption         =   "UserForm2"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465.001
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
    Set conn.ConnectionTimeout = 5
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
    Set conn.ConnectionTimeout = 1
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

Private Sub Label1_Click()

End Sub

Private Sub txtPassword_Change()

End Sub

Private Sub txtServer_Change()

End Sub

Private Sub txtUsername_Change()

End Sub


Private Sub UserForm_Click()

End Sub
