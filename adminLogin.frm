VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form adminLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "adminLogin"
   ClientHeight    =   4995
   ClientLeft      =   7320
   ClientTop       =   2430
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   4470
   Begin MSAdodcLib.Adodc loginAdodc 
      Height          =   375
      Left            =   600
      Top             =   1080
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   $"adminLogin.frx":0000
      OLEDBString     =   $"adminLogin.frx":008D
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM adminAccounts"
      Caption         =   "loginAdodc"
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
   Begin VB.CommandButton loginBtn 
      BackColor       =   &H0080FF80&
      Cancel          =   -1  'True
      Caption         =   "Login"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   3255
   End
   Begin VB.CommandButton showHide 
      Caption         =   "SHOW"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox password 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox username 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label createAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I don't have an Account"
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
      TabIndex        =   7
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Label Label2 
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
      Left            =   480
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label7 
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
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Login"
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
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "adminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backBtn_Click()
    Unload Me
    TypeOfAccount.Show
End Sub

Private Sub createAccount_Click()
    Unload Me
    adminSignup.Show
End Sub

Private Sub loginBtn_Click()
If Len(username.Text) <= 0 Or Len(password.Text) <= 0 Then
    MsgBox "Please fill out all fields!"
Else
    loginAdodc.RecordSource = " SELECT * FROM adminAccounts WHERE Username = '" + username.Text + "' AND Password ='" + password.Text + "' "
    loginAdodc.Refresh
    
    If loginAdodc.Recordset.EOF Then
        MsgBox "Invalid User"
    Else
        User = loginAdodc.Recordset.Fields("ID").Value
        MsgBox "Login Successfully!"
        Unload Me
        payrollDashboard.Show
    End If
End If

End Sub

Private Sub showHide_Click()
    If showHide.Caption = "SHOW" Then
        password.PasswordChar = ""
        showHide.Caption = "HIDE"
    Else
        password.PasswordChar = "*"
        showHide.Caption = "SHOW"
    End If
End Sub

