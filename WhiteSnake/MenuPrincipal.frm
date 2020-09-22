VERSION 5.00
Begin VB.Form MenuPrincipal 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C$@R Soft WHITESnake 1.0"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   405
      Top             =   3735
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1555
      Left            =   810
      Picture         =   "MenuPrincipal.frx":0000
      ScaleHeight     =   1560
      ScaleWidth      =   5250
      TabIndex        =   2
      Top             =   1530
      Width           =   5250
   End
   Begin VB.Timer Timer1 
      Left            =   405
      Top             =   3240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   2265
      Left            =   1665
      TabIndex        =   0
      Top             =   3150
      Width           =   3480
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   420
         Index           =   3
         Left            =   630
         MouseIcon       =   "MenuPrincipal.frx":25CA
         TabIndex        =   7
         Top             =   1755
         Width           =   2130
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Acerca de..."
         Height          =   420
         Index           =   2
         Left            =   630
         MouseIcon       =   "MenuPrincipal.frx":28D4
         TabIndex        =   6
         Top             =   1170
         Width           =   2130
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Records"
         Enabled         =   0   'False
         Height          =   420
         Index           =   1
         Left            =   630
         MouseIcon       =   "MenuPrincipal.frx":2BDE
         TabIndex        =   5
         Top             =   630
         Width           =   2130
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Jugar"
         Default         =   -1  'True
         Height          =   420
         Index           =   0
         Left            =   630
         MouseIcon       =   "MenuPrincipal.frx":2EE8
         TabIndex        =   4
         Top             =   180
         Width           =   2130
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1500
      Index           =   1
      Left            =   45
      TabIndex        =   3
      Top             =   90
      Width           =   9735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   45
      Width           =   9690
   End
End
Attribute VB_Name = "MenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a       As Byte
Dim Ancho   As Integer
Dim Alto    As Integer

Private Sub Command1_Click(Index As Integer)
MousePointer = 11
Select Case Index
    Case 0
        Principal.Show 1
        MousePointer = 1
    Case 1
        MousePointer = 1
    Case 2
        Acercade.Show 1
        MousePointer = 1
    Case 3
        Unload Me
End Select
End Sub

Private Sub Form_Load()
Picture1.Visible = False
Frame1.Visible = False
a = 1
Ancho = Picture1.Width
Alto = Picture1.Height
Picture1.Width = 1
Picture1.Height = 1
Picture1.Top = Me.Height / 2
Picture1.Left = Me.Width / 2

Timer1.Enabled = True
Timer1.Interval = 100


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim IResult As Integer
IResult = MsgBox("Esta seguro que desea salir de C$@R Soft WHITESnake 1.0.?", vbYesNo + vbQuestion, "WHITESnake 1.0")
If IResult = vbNo Then
    Cancel = True
    MousePointer = 1
End If
End Sub

Private Sub Timer1_Timer()
'Largo 28
Label1(0).Caption = Label1(0).Caption & Mid(Bienvenido, a, 1)
Label1(1).Caption = Label1(1).Caption & Mid(Bienvenido, a, 1)
If a = 5 Then
    Label1(0).FontSize = 60
    Label1(1).FontSize = 60
End If
If a = 8 Then
    Label1(0).FontSize = 48
    Label1(1).FontSize = 48
End If

If a = 13 Then
    Label1(0).FontSize = 32
    Label1(1).FontSize = 32
End If

If a = 17 Then
    Label1(0).FontSize = 28
    Label1(1).FontSize = 28

End If
If a = 20 Then
    Label1(0).FontSize = 24
    Label1(1).FontSize = 24
End If
a = a + 1
If a = 29 Then
    Timer1.Enabled = False
    Timer2.Enabled = True
    Picture1.Visible = True
    Timer2.Interval = 1
    a = 1
End If
End Sub

Private Sub Timer2_Timer()
If Picture1.Width < Ancho Then
    Picture1.Width = Picture1.Width + 100
    Picture1.Top = (Me.Height - Picture1.Height) / 2
Else
    Picture1.Width = Ancho
    If Picture1.Top > 1000 Then
        Picture1.Top = Picture1.Top - 60
    Else
        Frame1.Visible = True
        Timer2.Enabled = False
    End If
End If

If Picture1.Height < Alto Then
    Picture1.Height = Picture1.Height + 20
Else
    Picture1.Height = Alto - 13
End If

Picture1.Left = (Me.Width - Picture1.Width) / 2
End Sub
