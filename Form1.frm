VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   6570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox skin 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2760
      List            =   "Form1.frx":001F
      TabIndex        =   3
      Text            =   "Skinz"
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label button2cap 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Button2"
      Height          =   195
      Left            =   8160
      TabIndex        =   2
      Top             =   2280
      Width           =   555
   End
   Begin VB.Image button2 
      Height          =   285
      Left            =   8040
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label button1cap 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Button1"
      Height          =   195
      Left            =   8160
      TabIndex        =   1
      Top             =   1920
      Width           =   555
   End
   Begin VB.Image button1 
      Height          =   285
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label Heading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Heading"
      Height          =   195
      Left            =   8160
      TabIndex        =   0
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image UM 
      Height          =   135
      Left            =   840
      Top             =   4080
      Width           =   135
   End
   Begin VB.Image CLOSE 
      Height          =   345
      Left            =   2400
      Top             =   1440
      Width           =   345
   End
   Begin VB.Image MIN 
      Height          =   345
      Left            =   1320
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image max 
      Height          =   345
      Left            =   840
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image RS 
      Height          =   90
      Left            =   0
      Top             =   2400
      Width           =   210
   End
   Begin VB.Image UR 
      Height          =   210
      Left            =   600
      Top             =   720
      Width           =   90
   End
   Begin VB.Image BR 
      Height          =   135
      Left            =   240
      Top             =   1320
      Width           =   210
   End
   Begin VB.Image BM 
      Height          =   120
      Left            =   120
      Top             =   2520
      Width           =   105
   End
   Begin VB.Image BL 
      Height          =   90
      Left            =   240
      Top             =   3960
      Width           =   105
   End
   Begin VB.Image LS 
      Height          =   225
      Left            =   120
      Top             =   1680
      Width           =   90
   End
   Begin VB.Image UL 
      Height          =   120
      Left            =   120
      Top             =   840
      Width           =   195
   End
   Begin VB.Image bak 
      Height          =   135
      Left            =   240
      Top             =   3600
      Width           =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub button1_Click()
LoadControls
End Sub

Private Sub button1cap_Click()
LoadControls
End Sub

Private Sub button2_Click()
SaveControls
End Sub

Private Sub button2cap_Click()
SaveControls
End Sub

Private Sub CLOSE_Click()
End
End Sub

Private Sub Command1_Click()
UnStretchIt
LoadGFX (Dir1.Path)
StretchIt
GFX
End Sub

Private Sub Command2_Click()
End
End Sub





Private Sub Command3_Click()
SaveControls
End Sub

Private Sub Command4_Click()
LoadControls
End Sub

Private Sub Form_Load()
'Dir1.Path = App.Path & "\skinz\Default 2"

'UnStretchIt
'LoadGFX (App.Path & "\skinz\sinter")
'StretchIt
'Refreshit

'LoadControls
GFX

skin.Left = (Me.Width / 2) - (skin.Width / 2)
button1.Left = skin.Left
button2.Left = skin.Left + skin.Width - button2.Width

button1.Top = skin.Top + skin.Height + 200
button2.Top = button1.Top
Refreshit
End Sub

Sub LoadGFX(FileDir As String)

bak.Picture = LoadPicture(FileDir & "\Back_Pic.bmp")

UL.Picture = LoadPicture(FileDir & "\UpLeft_pic.bmp")
UR.Picture = LoadPicture(FileDir & "\UpRight_pic.bmp")
UM.Picture = LoadPicture(FileDir & "\UpMiddle_pic.bmp")

LS.Picture = LoadPicture(FileDir & "\LeftSide_pic.bmp")
RS.Picture = LoadPicture(FileDir & "\RightSide_pic.bmp")

BL.Picture = LoadPicture(FileDir & "\BottomLeft_pic.bmp")
BM.Picture = LoadPicture(FileDir & "\BottomMiddle_pic.bmp")
BR.Picture = LoadPicture(FileDir & "\BottomRight_pic.bmp")

button1.Picture = LoadPicture(FileDir & "\Button_pic.bmp")
button2.Picture = LoadPicture(FileDir & "\Button_pic.bmp")

End Sub


Sub UnStretchIt()
bak.Stretch = False

UL.Stretch = False
UR.Stretch = False
UM.Stretch = False

LS.Stretch = False
RS.Stretch = False

BL.Stretch = False
BM.Stretch = False
BR.Stretch = False
End Sub

Sub Refreshit()
bak.Refresh

UL.Refresh
UR.Refresh
UM.Refresh

LS.Refresh
RS.Refresh

BL.Refresh
BM.Refresh
BR.Refresh
End Sub

Sub StretchIt()
bak.Stretch = True

UL.Stretch = True
UR.Stretch = True
UM.Stretch = True

LS.Stretch = True
RS.Stretch = True

BL.Stretch = True
BM.Stretch = True
BR.Stretch = True
End Sub

Sub GFX()
If Me.WindowState = vbMinimized Then Exit Sub

Heading.Top = UM.Height / 8
Heading.Left = Me.Width / 2

bak.Left = 0
bak.Top = 0
bak.Width = Me.Width
bak.Height = Me.Height

UL.Top = 0
UL.Left = 0

BL.Top = Me.Height - BL.Height
BL.Left = 0

LS.Top = UL.Height
LS.Height = Me.Height - UL.Height - BL.Height
LS.Left = 0

UM.Width = Me.Width - UL.Width - UR.Width
UM.Left = UL.Width
UM.Top = 0

UR.Top = 0
UR.Left = Me.Width - UR.Width

BR.Left = Me.Width - BR.Width
BR.Top = Me.Height - BR.Height

RS.Left = Me.Width - RS.Width
RS.Height = Me.Height - UR.Height - BR.Height
RS.Top = UR.Height

BM.Top = Me.Height - BM.Height
BM.Left = BL.Width
BM.Width = Me.Width - BL.Width - BR.Width

button1cap.Left = (button1.Left + (button1.Width / 4))
button1cap.Top = (button1.Top + (button1.Height / 6))

button2cap.Left = (button2.Left + (button2.Width / 4))
button2cap.Top = (button2.Top + (button2.Height / 6))
End Sub

Private Sub Form_Resize()
GFX
End Sub

Private Sub max_Click()
If Me.WindowState = vbMaximized Then Me.WindowState = vbNormal Else Me.WindowState = vbMaximized
End Sub

Private Sub MIN_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub skin_Click()

If skin.Text = skin.List(1) Then
UnStretchIt
LoadGFX (App.Path & "\skinz\O' Juice")
StretchIt
End If

If skin.Text = skin.List(2) Then
UnStretchIt
LoadGFX (App.Path & "\skinz\logika")
StretchIt
End If

If skin.Text = skin.List(3) Then
UnStretchIt
LoadGFX (App.Path & "\skinz\GT")
StretchIt
End If

If skin.Text = skin.List(4) Then
UnStretchIt
LoadGFX (App.Path & "\skinz\Simple")
StretchIt
End If

If skin.Text = skin.List(5) Then
UnStretchIt
LoadGFX (App.Path & "\skinz\sinter")
StretchIt
End If

If skin.Text = skin.List(6) Then
UnStretchIt
LoadGFX (App.Path & "\skinz\UrbanTech2")
StretchIt
End If

If skin.Text = skin.List(7) Then
UnStretchIt
LoadGFX (App.Path & "\skinz\saph")
StretchIt
End If

If skin.Text = skin.List(8) Then
UnStretchIt
LoadGFX (App.Path & "\skinz\Moon")
StretchIt
End If

GFX
End Sub
