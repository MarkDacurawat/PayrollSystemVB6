VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form payrollDashboard 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "payrollDashboard"
   ClientHeight    =   5280
   ClientLeft      =   6705
   ClientTop       =   1950
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Cancel          =   -1  'True
      Caption         =   "View"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Save"
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4200
      Width           =   855
   End
   Begin MSAdodcLib.Adodc employee 
      Height          =   375
      Left            =   1440
      Top             =   4920
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
      RecordSource    =   "SELECT * FROM employee"
      Caption         =   "employee"
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
   Begin VB.TextBox tardiness 
      DataField       =   "Tardiness"
      DataSource      =   "employee"
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Text            =   "0"
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton grossPayCompute 
      BackColor       =   &H00FFFF80&
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox department 
      DataField       =   "Department"
      DataSource      =   "employee"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox employeeName 
      DataField       =   "EmployeeName"
      DataSource      =   "employee"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox numberOfDaysWorked 
      DataField       =   "NoOfHoursWorked"
      DataSource      =   "employee"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Text            =   "0"
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox ratePerHour 
      DataField       =   "RatePerHpur"
      DataSource      =   "employee"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Text            =   "0"
      Top             =   2640
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc getdataAdodc 
      Height          =   375
      Left            =   0
      Top             =   4920
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
      Connect         =   $"payrollDashboard.frx":011A
      OLEDBString     =   $"payrollDashboard.frx":01A7
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
      Left            =   4080
      TabIndex        =   15
      Top             =   4440
      Width           =   975
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
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   4440
      Width           =   1335
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
      Top             =   240
      Width           =   2415
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
      Left            =   720
      TabIndex        =   1
      Top             =   480
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
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "payrollDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    If CInt(totalTxt.Caption) = 0 Then
        MsgBox "Complete all fields first!"
    Else
        employee.Recordset.Update
        If employee.Recordset.EOF Then
            MsgBox "Can't save employee information!"
        Else
            If MsgBox("Do you want to make another computation?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
                employeeName.Text = ""
                ratePerHour.Text = ""
                numberOfDaysWorked.Text = ""
                tardiness.Text = ""
                department.Text = ""
                totalTxt.Caption = 0
            Else
                MsgBox "Saved!"
                Unload Me
                employeeList.Show
            End If
        End If
    End If
    
End Sub

Private Sub Command2_Click()
    Unload Me
    employeeList.Show
End Sub

Private Sub Form_Load()
    employee.Recordset.AddNew
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
