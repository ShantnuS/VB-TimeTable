VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Random Subject Settings"
   ClientHeight    =   3030
   ClientLeft      =   8865
   ClientTop       =   4950
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin VB.CommandButton cmdSettingSave 
      Caption         =   "Save"
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox cmbEndTime 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Text            =   "End Time"
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cmbStartTime 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Text            =   "Start Time"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblAWESOME 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Between"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lbltextblock 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose time period of random subject:"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   120
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartTime As Integer
Dim EndTime As Integer
Dim StartDay As Integer
Dim EndDay As Integer

Private Sub cmdSettingSave_Click()
'Check if end time is larger than start time
If (cmbEndTime.ListIndex - cmbStartTime.ListIndex) < 0 Then
    MsgBox "End Time must be after Start Time!"
Else
    'Save the settings
    StartTime = cmbStartTime.ListIndex + 1
    EndTime = cmbEndTime.ListIndex + 1
    SaveSetting "shantnu", "timetable", "StartTime", StartTime
    SaveSetting "shantnu", "timetable", "EndTime", EndTime
    
    'Display Confirmation Message
    MsgBox "Settings Saved"
    
    'Close the form
    frmSettings.Hide
End If

End Sub

Private Sub Form_Load()
'Add values to start Time combobox
cmbStartTime.AddItem "0701-0800"
cmbStartTime.AddItem "0801-0900"
cmbStartTime.AddItem "0901-1000"
cmbStartTime.AddItem "1001-1100"
cmbStartTime.AddItem "1101-1200"
cmbStartTime.AddItem "1201-1300"
cmbStartTime.AddItem "1301-1400"
cmbStartTime.AddItem "1401-1500"
cmbStartTime.AddItem "1501-1600"
cmbStartTime.AddItem "1601-1700"
cmbStartTime.AddItem "1701-1800"
cmbStartTime.AddItem "1801-1900"
cmbStartTime.AddItem "1901-2000"
cmbStartTime.AddItem "2001-2100"
cmbStartTime.AddItem "2101-2200"
cmbStartTime.AddItem "2201-2300"
cmbStartTime.AddItem "2301-0000"

'Add values to end time combobox
cmbEndTime.AddItem "0701-0800"
cmbEndTime.AddItem "0801-0900"
cmbEndTime.AddItem "0901-1000"
cmbEndTime.AddItem "1001-1100"
cmbEndTime.AddItem "1101-1200"
cmbEndTime.AddItem "1201-1300"
cmbEndTime.AddItem "1301-1400"
cmbEndTime.AddItem "1401-1500"
cmbEndTime.AddItem "1501-1600"
cmbEndTime.AddItem "1601-1700"
cmbEndTime.AddItem "1701-1800"
cmbEndTime.AddItem "1801-1900"
cmbEndTime.AddItem "1901-2000"
cmbEndTime.AddItem "2001-2100"
cmbEndTime.AddItem "2101-2200"
cmbEndTime.AddItem "2201-2300"
cmbEndTime.AddItem "2301-0000"
End Sub

