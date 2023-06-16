VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form adminSignup 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "adminSignup"
   ClientHeight    =   5625
   ClientLeft      =   7320
   ClientTop       =   2265
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc registerAdodc 
      Height          =   330
      Left            =   480
      Top             =   1080
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"adminSignup.frx":0000
      OLEDBString     =   $"adminSignup.frx":008D
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "adminAccounts"
      Caption         =   "registerAdodc"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton registerBtn 
      BackColor       =   &H0080FF80&
      Cancel          =   -1  'True
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Poppins SemiBold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   3255
   End
   Begin VB.TextBox password 
      DataField       =   "Password"
      DataSource      =   "registerAdodc"
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3720
      Width           =   3255
   End
   Begin VB.TextBox username 
      DataField       =   "Username"
      DataSource      =   "registerAdodc"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox fullName 
      DataField       =   "FullName"
      DataSource      =   "registerAdodc"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label createAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I already have an Account"
      BeginProperty Font 
         Name            =   "Poppins"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   840
      MousePointer    =   10  'Up Arrow
      TabIndex        =   8
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Poppins SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Poppins SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FullName: "
      BeginProperty Font 
         Name            =   "Poppins SemiBold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Register"
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
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "adminSignup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub createAccount_Click()
    Unload Me
    adminLogin.Show
End Sub

Private Sub Form_Load()
    registerAdodc.Recordset.AddNew
End Sub

Private Sub registerBtn_Click()
    
    If Len(fullName.Text) <= 0 Or Len(username.Text) <= 0 Or Len(password.Text) <= 0 Then
        MsgBox "Please fill out all the fields!"
    Else
        registerAdodc.Recordset.Update
        If registerAdodc.Recordset.EOF Then
            MsgBox "Can't Add a new User..."
        Else
            MsgBox "Successfully Added!"
        End If
    End If
    
End Sub

