VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form payrollDashboard 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "payrollDashboard"
   ClientHeight    =   6420
   ClientLeft      =   7635
   ClientTop       =   2115
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "payrollDashboard.frx":0000
   ScaleHeight     =   6420
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton calculatorBtn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Calculate Rate Per Hour"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   2535
   End
   Begin VB.TextBox totalTxt 
      DataField       =   "Department"
      DataSource      =   "employee"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox tardiness 
      DataField       =   "Tardiness"
      DataSource      =   "employee"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Text            =   "0"
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton grossPayCompute 
      BackColor       =   &H0080FF80&
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox numberOfDaysWorked 
      DataField       =   "NoOfHoursWorked"
      DataSource      =   "employee"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Text            =   "0"
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox ratePerHour 
      DataField       =   "RatePerHpur"
      DataSource      =   "employee"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Text            =   "0"
      Top             =   1800
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc getdataAdodc 
      Height          =   375
      Left            =   2400
      Top             =   120
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
      Connect         =   $"payrollDashboard.frx":1E67B
      OLEDBString     =   $"payrollDashboard.frx":1E708
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
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
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tardiness Minutes:"
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
      TabIndex        =   8
      Top             =   3360
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
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   2295
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
      Left            =   480
      TabIndex        =   3
      Top             =   1440
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
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll System"
      BeginProperty Font 
         Name            =   "Poppins"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3135
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
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "payrollDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calculatorBtn_Click()
    Unload Me
    calculator.Show
End Sub

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
    If Len(ratePerHour.Text) <= 0 Then
        MsgBox "Please fill out Rate Per Hour field!"
    ElseIf Not IsNumeric(ratePerHour.Text) Or Not IsNumeric(numberOfDaysWorked.Text) Then
        MsgBox "Please input Number only!"
    ElseIf Len(numberOfDaysWorked.Text) <= 0 Then
        MsgBox "Please fill out Number of Hours Worked field!"
    Else
        Dim grossPay As Double
        Dim finalTotal As Double
        
        grossPay = Format((CDbl(ratePerHour.Text) * CDbl(numberOfDaysWorked.Text)), "#,##0")
        totalTxt.Text = (grossPay - CDbl(tardiness.Text))
    End If
End Sub
