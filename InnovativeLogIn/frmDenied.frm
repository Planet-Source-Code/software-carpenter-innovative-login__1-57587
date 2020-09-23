VERSION 5.00
Begin VB.Form frmDenied 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1365
   ClientLeft      =   4740
   ClientTop       =   2910
   ClientWidth     =   4485
   ControlBox      =   0   'False
   FillColor       =   &H80000012&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   1365
   ScaleWidth      =   4485
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   720
   End
   Begin VB.Label lblPress 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press any key to continue..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   870
      TabIndex        =   1
      Top             =   1080
      Width           =   2745
   End
   Begin VB.Label lblPress2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press any key to continue..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   -2745
      TabIndex        =   2
      Top             =   1080
      Width           =   2745
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmDenied.frx":0000
      Top             =   247
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Access Denied!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00987758&
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmDenied"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Rodelio Martinez Rodriguez
'E-mail: rodeliorodriguez@yahoo.com

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Unload Me
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Unload Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Unload Me
End Sub

Private Sub Timer1_Timer()
    DoEvents
    Static lbl2 As Boolean
    Static lbl1 As Boolean
    If lbl1 = False Then
        lblPress.Move lblPress.Left + 20
    End If
    If lbl2 = True Then
        lblPress2.Move lblPress2.Left + 20
    End If
    If lblPress.Left + lblPress.Width >= Me.Width Then
        If lbl2 = False Then lbl2 = True
    End If
    If lblPress2.Left + lblPress2.Width >= Me.Width Then
        If lbl1 = True Then lbl1 = False
    End If
    If lblPress.Left >= Me.Width Then
        lblPress.Left = 0 - lblPress.Width
        lbl1 = True
    End If
    If lblPress2.Left >= Me.Width Then
        lblPress2.Left = 0 - lblPress2.Width
        lbl2 = False
    End If
End Sub
