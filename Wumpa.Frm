VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wompus for Windows. (c) Yadav"
   ClientHeight    =   5970
   ClientLeft      =   105
   ClientTop       =   2025
   ClientWidth     =   5895
   FillStyle       =   0  'Solid
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Wumpa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   398
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4830
      Left            =   375
      ScaleHeight     =   15.395
      ScaleMode       =   0  'User
      ScaleWidth      =   16.83
      TabIndex        =   0
      Top             =   405
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4965
      Picture         =   "Wumpa.frx":030A
      Top             =   5370
      Width           =   660
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   5115
      Left            =   285
      Top             =   270
      Width           =   5340
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MAKE THE GREEN BALL REACH THE HOME CELL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   270
      TabIndex        =   1
      Top             =   5565
      Width           =   4485
   End
   Begin VB.Menu MenuGame 
      Caption         =   "&Game"
      Begin VB.Menu MenuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu MenuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MenuCustomise 
      Caption         =   "&Customise"
      Begin VB.Menu MenuColour 
         Caption         =   "&Colour"
         Begin VB.Menu MenuColourScheme 
            Caption         =   "De&fault"
            Index           =   0
         End
         Begin VB.Menu MenuColourScheme 
            Caption         =   "Scheme &A"
            Index           =   1
         End
         Begin VB.Menu MenuColourScheme 
            Caption         =   "Scheme &B"
            Index           =   2
         End
         Begin VB.Menu MenuColourScheme 
            Caption         =   "Scheme &C"
            Index           =   3
         End
         Begin VB.Menu MenuColourScheme 
            Caption         =   "Scheme &D"
            Index           =   4
         End
         Begin VB.Menu MenuColourScheme 
            Caption         =   "Scheme &E"
            Index           =   5
         End
      End
      Begin VB.Menu MenuBlockers 
         Caption         =   "&Blockers"
         Begin VB.Menu MenuBlockMode 
            Caption         =   "&Box"
            Index           =   0
         End
         Begin VB.Menu MenuBlockMode 
            Caption         =   "&Line"
            Index           =   1
         End
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MenuAbout 
         Caption         =   "&About Wompus"
      End
   End
   Begin VB.Menu MenuCopy 
      Caption         =   "&Copy"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case 37: Direction = "W"
        Case 38: Direction = "N"
        Case 39: Direction = "E"
        Case 40: Direction = "S"
        End Select

End Sub

Private Sub Form_Load()

Form1.Top = 0.5 * (Screen.Height - Form1.Height)
Form1.Left = 0.5 * (Screen.Width - Form1.Width)

End Sub

Private Sub Form_Paint()
    
    'Showblocks

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Private Sub MenuAbout_Click()
    
    ShowAbout

End Sub

Private Sub MenuBlockMode_Click(Index As Integer)

    BlockerMode = Index
    Form1.Picture1.Cls
    Showblocks


End Sub

Private Sub MenuColourScheme_Click(Index As Integer)

    SetColourScheme Index
    Showblocks

End Sub
Private Sub MenuCopy_Click()
'
'    This procedure is used for game development only.
'
'    For Y = 1 To MaxY: For X = 1 To MaxX
'    a$ = a$ + LTrim(Str$(Tile(X, Y)))
'    Next: Next
'
'    Clipboard.Clear: Clipboard.SetText a$
'
End Sub
Private Sub MenuExit_Click()
    
    Unload Me   ': End

End Sub

Private Sub MenuNew_Click()
    
    StartNewGame
    Form1.Picture1.Cls
    Showblocks

End Sub

Private Sub Picture1_Paint()
 
    Showblocks
                       
End Sub
