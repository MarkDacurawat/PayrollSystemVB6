VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SplashScreen 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Splash Form"
   ClientHeight    =   2040
   ClientLeft      =   5160
   ClientTop       =   3525
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   9270
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll System"
      BeginProperty Font 
         Name            =   "Poppins"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "SplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
        Timer1.Enabled = True
        Timer1_Timer
End Sub
Private Sub Timer1_Timer()
    ProgressBar1.Value = ProgressBar1.Value + 10
    
    If ProgressBar1.Value = 80 Then
        ProgressBar1.Value = ProgressBar1 + 20
    If ProgressBar1.Value >= ProgressBar1.Max Then
        Timer1.Enabled = False
    End If
        MsgBox "Loading Complete!", vbInformation, "Info"
        Unload Me
        adminLogin.Show
    End If
    
End Sub
