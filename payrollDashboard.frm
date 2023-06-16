VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form payrollDashboard 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "payrollDashboard"
   ClientHeight    =   5280
   ClientLeft      =   5460
   ClientTop       =   2580
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox tardiness 
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton grossPayCompute 
      BackColor       =   &H0080FF80&
      Cancel          =   -1  'True
      Caption         =   "Compute"
      BeginProperty Font 
         Name            =   "Poppins SemiBold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox department 
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox employeeName 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox numberOfDaysWorked 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox ratePerHour 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc getdataAdodc 
      Height          =   375
      Left            =   7800
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   $"payrollDashboard.frx":0000
      OLEDBString     =   $"payrollDashboard.frx":008D
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM adminAccounts"
      Caption         =   "getdataAdodc"
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tardiness:"
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
      Left            =   3480
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label totalTxt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   7080
      TabIndex        =   11
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Department:"
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
      Left            =   3480
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name:"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NO. of Hours worked:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rate per Hour:"
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
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label userNameOutput 
      BackStyle       =   0  'Transparent
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
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll Dashboard"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label createAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome!"
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
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "payrollDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    getdataAdodc.RecordSource = "SELECT * FROM adminAccounts WHERE ID = " & User
    getdataAdodc.Refresh
    
    If getdataAdodc.Recordset.EOF Then
        MsgBox "Invalid User"
    Else
        userNameOutput.Caption = getdataAdodc.Recordset.Fields("FullName").Value
    End If
End Sub

Private Sub grossPayCompute_Click()
    If Len(ratePerHour.Text) <= 0 Or Len(numberOfDaysWorked.Text) <= 0 Then
        MsgBox "Please fill out all the fields!"
    Else
        Dim grossPay As String
        Dim finalTotal As String
        
        grossPay = Format((CDbl(ratePerHour.Text) * CDbl(numberOfDaysWorked.Text)), "#,##0")
        totalTxt.Caption = (grossPay - CDbl(tardiness.Text))
    End If
End Sub
