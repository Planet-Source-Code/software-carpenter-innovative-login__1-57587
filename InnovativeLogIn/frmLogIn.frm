VERSION 5.00
Begin VB.Form frmLogIn 
   BackColor       =   &H00E0E0E0&
   Caption         =   " "
   ClientHeight    =   1155
   ClientLeft      =   2355
   ClientTop       =   1950
   ClientWidth     =   2685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   100
   Icon            =   "frmLogIn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   797.201
   ScaleMode       =   0  'User
   ScaleWidth      =   2521.354
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   -360
      Top             =   480
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   240
      Left            =   75
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblKeyin2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please key-in your password..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   -2235
      TabIndex        =   2
      Top             =   120
      Width           =   2235
   End
   Begin VB.Label lblKeyin1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please key-in your password..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   120
      Width           =   2235
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Rodelio Martinez Rodriguez
'E-mail: rodeliorodriguez@yahoo.com

Option Explicit
Private HiddenPassword As String
Private Pressed(100) As String
Private ChrValues(100) As String
Private b As Integer, PasswordLenght As Integer
Private EnteredValues As String
Private Const PasswordValue = "My Unique Password Mask"

Private Sub DetermineAccess()
    Dim MatchFound As Boolean, MaximumTry As Byte, Success As Integer
    MaximumTry = 2
    Static Tries As Byte
    For b = 0 To 100
        If Pressed(b) <> vbNullString Then
            HiddenPassword = HiddenPassword & Pressed(b)
        End If
    Next
    MatchFound = UserIsExisting(HiddenPassword)
    If MatchFound = False Then
        If Tries = MaximumTry Then
            MsgBox ("Access denied... system will now abort.   "), vbInformation, "Access Failure"
            Unload Me
        End If
        frmDenied.Show 1
        Tries = Tries + 1
        ClearValues
        txtPassword = vbNullString
    Else
        MsgBox "Access granted!" & vbCrLf & vbCrLf & "I'm hoping to earn your vote... Have a nice day! ", vbExclamation, "Access Status"
        Unload Me
    End If
    Exit Sub
HandleError:
    Screen.MousePointer = vbDefault
    MsgBox Error$ & " - " & Str$(Err), vbExclamation, "Application Load Error"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    ClearValues
End Sub

Private Sub Timer1_Timer()
    DoEvents
    Static lbl2 As Boolean
    Static lbl1 As Boolean
    If lbl1 = False Then
        lblKeyin1.Move lblKeyin1.Left + 20
    End If
    If lbl2 = True Then
        lblKeyin2.Move lblKeyin2.Left + 20
    End If
    If lblKeyin1.Left + lblKeyin1.Width >= Me.Width Then
        If lbl2 = False Then lbl2 = True
    End If
    If lblKeyin2.Left + lblKeyin2.Width >= Me.Width Then
        If lbl1 = True Then lbl1 = False
    End If
    If lblKeyin1.Left >= Me.Width Then
        lblKeyin1.Left = 0 - lblKeyin1.Width
        lbl1 = True
    End If
    If lblKeyin2.Left >= Me.Width Then
        lblKeyin2.Left = 0 - lblKeyin2.Width
        lbl2 = False
    End If
End Sub

Private Sub txtPassword_Change()
    If txtPassword = vbNullString Then
        ClearValues
    End If
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        Pressed(txtPassword.SelStart) = ""
        PasswordLenght = PasswordLenght - 1
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Static CursorPosition As Integer
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPassword <> vbNullString Then DetermineAccess
    Else
        If KeyAscii <> 8 Then
            CheckKeypresses KeyAscii, txtPassword.SelStart
            CursorPosition = txtPassword.SelStart + 1
            txtPassword = Mid$(PasswordValue, 1, PasswordLenght)
            txtPassword.SelStart = CursorPosition
            KeyAscii = 0
        Else
            If txtPassword.SelStart > 1 Then
                Pressed(txtPassword.SelStart - 1) = ""
            End If
            PasswordLenght = PasswordLenght - 1
        End If
    End If
End Sub

Private Function UserIsExisting(ByVal UserPassword As String) As Boolean
    If StrComp(HiddenPassword, "rodelio", 0) = 0 Then
        UserIsExisting = True
    Else
        UserIsExisting = False
    End If
End Function

Public Sub CheckKeypresses(ByVal KeyAscii As Integer, Position As Integer)
    Dim a As Integer, Start As Integer
    If Pressed(Position) <> vbNullString Then
        Start = Position
        For a = Start To 99
            ChrValues(a) = Pressed(a)
        Next
        For a = Start To 99
            Pressed(a + 1) = ChrValues(a)
        Next
    End If
    Pressed(Position) = Chr$(KeyAscii)
    EnteredValues = ""
    For a = 0 To 100
        EnteredValues = EnteredValues & Pressed(a)
    Next
    PasswordLenght = Len(EnteredValues)
End Sub

Public Sub ClearValues()
    HiddenPassword = ""
    EnteredValues = ""
    PasswordLenght = 0
    For b = 0 To 100
        Pressed(b) = ""
        ChrValues(b) = ""
    Next
End Sub
