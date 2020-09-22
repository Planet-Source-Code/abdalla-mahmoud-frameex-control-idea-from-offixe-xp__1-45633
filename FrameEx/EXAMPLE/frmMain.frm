VERSION 5.00
Object = "*\A..\SOURCE\Frame.vbp"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "An Example Of The FrameEx Control"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin Frame.FrameEx FrameEx1 
      Height          =   1935
      Left            =   120
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3413
      BodyColor       =   -2147483632
      BackColor       =   -2147483633
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "If ypu clicked here ther first control will have focus"
      HighlightColor  =   16711680
      Enabled         =   -1  'True
      MousePointer    =   0
      MouseIcon       =   "frmMain.frx":0000
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   6
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   5
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country :"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone :"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   510
      End
   End
   Begin Frame.FrameEx FrameEx2 
      Height          =   1335
      Left            =   120
      Top             =   2280
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2355
      BodyColor       =   -2147483632
      BackColor       =   -2147483633
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "You will see the control name on the form's caption"
      HighlightColor  =   255
      Enabled         =   -1  'True
      MousePointer    =   0
      MouseIcon       =   "frmMain.frx":001C
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   10
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   9
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class :"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   8
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Language :"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   810
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":0038
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   6015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Call MsgBox("Please vote for me.", vbInformation, "Vote")
End Sub

Private Sub FrameEx1_ControlFocus(FocusControl As Object)
    Caption = FocusControl.Name & "(" & FocusControl.Index & ")"
End Sub

Private Sub FrameEx2_ControlFocus(FocusControl As Object)
    Caption = FocusControl.Name & "(" & FocusControl.Index & ")"
End Sub

