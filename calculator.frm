VERSION 5.00
Begin VB.Form calculator 
   Caption         =   "Calculator"
   ClientHeight    =   4785
   ClientLeft      =   7095
   ClientTop       =   3120
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   Picture         =   "calculator.frx":0000
   ScaleHeight     =   4785
   ScaleWidth      =   4815
   Begin VB.TextBox ratePerHourTotal 
      DataField       =   "Department"
      DataSource      =   "employee"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton convertBtn 
      BackColor       =   &H0080FF80&
      Caption         =   "Convert"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   3255
   End
   Begin VB.TextBox noOfHoursPerDay 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox ratePerDay 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label createAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<- Back To Main Page"
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
      Left            =   -240
      MousePointer    =   10  'Up Arrow
      TabIndex        =   8
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Per Hour:"
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
      Left            =   600
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No. Of Hours Per Day:"
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
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rate per Day:"
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
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Per Hour Checker"
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
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub convertBtn_Click()
    If Len(ratePerDay.Text) <= 0 Then
        MsgBox "Please fill out Rate Per Day field!"
    ElseIf Not IsNumeric(ratePerDay.Text) Or Not IsNumeric(noOfHoursPerDay.Text) Then
        MsgBox "Please input Number only!"
    ElseIf Len(noOfHoursPerDay.Text) <= 0 Then
        MsgBox "Please fill out Number of Hours Per Day field!"
    Else
        Dim convertedText As Double
        
        convertedText = Format((ratePerDay.Text / noOfHoursPerDay), "0.00")
        ratePerHourTotal.Text = convertedText
        
    End If
End Sub

Private Sub createAccount_Click()
    Unload Me
    payrollDashboard.Show
End Sub
