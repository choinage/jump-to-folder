VERSION 5.00
Begin VB.Form tool 
   Caption         =   "¹¤¾ßÏä"
   ClientHeight    =   1755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5700
   LinkTopic       =   "tool"
   MaxButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5700
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command1 
      Caption         =   "´ò¿ªÄ¿Â¼"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Left            =   5160
      Top             =   360
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   15
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   15
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "tool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Shell "explorer.exe " & Chr(34) & App.Path & Chr(34), vbNormalFocus
End Sub

Private Sub Form_Load()
    Text1.Text = Now

    Timer1.Enabled = True

    Timer1.Interval = 1000
End Sub

Private Sub Timer1_Timer()
    Text1.Text = Now
End Sub
