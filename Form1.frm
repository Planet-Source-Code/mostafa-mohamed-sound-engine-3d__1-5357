VERSION 5.00
Object = "{5E379E14-C4AF-11D3-94B2-DCF27482A336}#3.0#0"; "Sound Engine.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "3D sound engine sample"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sound2"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sound1"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "Sound1"
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "Sound2"
         Height          =   195
         Left            =   2280
         TabIndex        =   3
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   "Listener"
         Height          =   195
         Left            =   1440
         TabIndex        =   2
         Top             =   1320
         Width           =   555
      End
   End
   Begin Sound.SoundEngine SoundEngine1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Move the red labels by the mouse to set its position from the listiner,the click its button"
      Height          =   1695
      Left            =   3960
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "For more samples and links and controls visit my site at http://go.to/mostafa"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s1 As Integer
Dim s2 As Integer
Dim MouseStartx As Single
Dim MouseStarty As Single

Private Sub Command1_Click()
SoundEngine1.Setlistenerpos Label1.Left, 0, Label1.Top 'set the listener position to label1
SoundEngine1.Playwave s1, Label3.Left, 0, Label3.Top 'play  the wave and set its  position
End Sub

Private Sub Command2_Click()
SoundEngine1.Setlistenerpos Label1.Left, 0, Label1.Top 'set the listener position to label1
SoundEngine1.Playwave s2, Label2.Left, 0, Label2.Top 'play  the wave and set its  position
End Sub

Private Sub Command3_Click()
SoundEngine1.EndEngine 'end the engine
End

End Sub

Private Sub Form_Load()
SoundEngine1.Setform (Me.hWnd) 'intalize the engine
s1 = SoundEngine1.Addwave(App.Path & "\bump1.wav") 'load the first wave
s2 = SoundEngine1.Addwave(App.Path & "\death1.wav") 'load the second
's1 and s2 are the pointer to the wave channel
End Sub

Private Sub Form_Unload(Cancel As Integer)
SoundEngine1.EndEngine 'end the engine
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MouseStartx = Picture1.ScaleX(x, 1, 3)
MouseStarty = Picture1.ScaleY(y, 1, 3)
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
  If Button = vbLeftButton Then
          x = Picture1.ScaleX(x, 1, 3)
    y = Picture1.ScaleY(y, 1, 3)
         Label3.Left = Label3.Left - (MouseStartx - x)
    
 Label3.Top = Label3.Top - (MouseStarty - y)

    End If
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MouseStartx = Picture1.ScaleX(x, 1, 3)
MouseStarty = Picture1.ScaleY(y, 1, 3)
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
  If Button = vbLeftButton Then
     x = Picture1.ScaleX(x, 1, 3)
    y = Picture1.ScaleY(y, 1, 3)
         Label2.Left = Label2.Left - (MouseStartx - x)
    
 Label2.Top = Label2.Top - (MouseStarty - y)

    End If
End Sub
