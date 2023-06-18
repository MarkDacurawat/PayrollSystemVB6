VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form employeeList 
   Caption         =   "EmployeeListHistory"
   ClientHeight    =   5205
   ClientLeft      =   5070
   ClientTop       =   2190
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   9345
   Begin VB.CommandButton searchBtn 
      BackColor       =   &H0080FF80&
      Cancel          =   -1  'True
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Poppins SemiBold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox searchBar 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc employee 
      Height          =   330
      Left            =   2400
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      Connect         =   $"employeeList.frx":0000
      OLEDBString     =   $"employeeList.frx":008D
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
   Begin VB.CommandButton loginBtn 
      BackColor       =   &H008080FF&
      Caption         =   "Delete"
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
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid employeeGrid 
      Bindings        =   "employeeList.frx":011A
      Height          =   2415
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4260
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Employee List"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "employeeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub loginBtn_Click()
    employee.Recordset.Delete
    If employee.Recordset.EOF Then
        MsgBox "Can't delete!"
    Else
        MsgBox "Successfully Deleted!"
    End If
End Sub

Private Sub searchBtn_Click()
    employee.RecordSource = "SELECT * FROM employee WHERE EmployeeName = '%" + CStr(searchBar.Text) + "%'"
    employee.Refresh
        If employee.Recordset.EOF Then
            MsgBox "No data found"
        End If
End Sub
