VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3585
   ClientLeft      =   4995
   ClientTop       =   2715
   ClientWidth     =   6855
   ControlBox      =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3585
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2925
      TabIndex        =   1
      Top             =   5655
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Click Mouse or Press Enter Key to Begin Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   6705
   End
   Begin VB.Image Image1 
      Height          =   2550
      Left            =   180
      Picture         =   "About.frx":030A
      Top             =   420
      Width           =   6345
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   6
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3495
      Left            =   45
      Top             =   60
      Width           =   6765
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'   Center the form on screen
    Form2.Top = 0.5 * (Screen.Height - Form2.Height)
    Form2.Left = 0.5 * (Screen.Width - Form2.Width)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Unload Me
       
End Sub

Private Sub Image1_Click()

    Unload Me
    
End Sub

Private Sub Label1_Click()

    Unload Me
    
End Sub

Private Sub Command1_Click()

    Unload Me
    
End Sub

