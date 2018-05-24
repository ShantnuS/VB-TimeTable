VERSION 5.00
Begin VB.Form FrmAddAppo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Appointment"
   ClientHeight    =   4500
   ClientLeft      =   7185
   ClientTop       =   3795
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7920
   Begin VB.ListBox lstDesc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   13
      Top             =   1080
      Width           =   975
   End
   Begin VB.ListBox lstSubj 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   5040
      TabIndex        =   12
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton cmdBACK 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   615
   End
   Begin VB.ComboBox cmbTime 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   10
      Text            =   "Time"
      Top             =   3720
      Width           =   1935
   End
   Begin VB.ComboBox cmbDay 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "FrmAddAppo.frx":0000
      Left            =   360
      List            =   "FrmAddAppo.frx":0002
      TabIndex        =   9
      Text            =   "Day"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton cmdTesting 
      Appearance      =   0  'Flat
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Appointment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   5
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtAppoTitle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   285
      Left            =   360
      MaxLength       =   22
      TabIndex        =   0
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox txtAppoDescription 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   975
      Left            =   360
      MaxLength       =   104
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label lblCharacters 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "characters remaining :104"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label lblsaysubjectlist 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject List:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblTitleAppo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Appointment"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   7695
   End
   Begin VB.Label lblTesting 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "(empty)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblSayDESCRIPTION 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblSayTITLE 
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3495
      Index           =   0
      Left            =   4800
      Top             =   720
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3495
      Index           =   1
      Left            =   120
      Top             =   720
      Width           =   4575
   End
End
Attribute VB_Name = "FrmAddAppo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Appointment adding form!

'Variables
Dim Tempstring As String
Dim recnum As Integer
Dim DayID As Integer
Dim TimeID As Integer
Dim WeekID As Integer


Private Sub cmbSubject_Change()

End Sub

Private Sub cmdAdd_Click()

'Finds the DayID based on the listindex
DayID = cmbDay.ListIndex + 1

'Finds the timeID based on listindex
TimeID = cmbTime.ListIndex + 1

'Checks if day and time have been selected
If cmbTime.ListIndex = -1 Or cmbDay.ListIndex = -1 Or txtAppoTitle = "" Then
    cmbTime.ForeColor = &HFF&
    cmbDay.ForeColor = &HFF&
    MsgBox "Check the Date, Time and Title!"
Else
    'Day and Time IDs are Day and Time IDs
    DayID = DayID
    TimeID = TimeID

    'Clicking add button adds appointment to appointment data file
    Appo.AppoTitle = txtAppoTitle
    Appo.AppoDescription = txtAppoDescription
    Appo.AppoID = LOF(1) / Len(Appo) + 1
    Appo.AppoDayID = DayID
    Appo.AppoTimeID = TimeID
    Appo.AppoWeekID = WeekID
    'Appo.AppoWeekID = 2 '(This was for testing only - changes WeekID to 2 manually).
    
    Put #1, Appo.AppoID, Appo
    
    'Message box and clear text boxes
    MsgBox "Appointment Added!"
    txtAppoTitle.Text = ""
    txtAppoDescription = ""

End If

'"refills" the table on frmtable
FrmTable.PopulatesGrid

End Sub

Private Sub cmdBACK_Click()

'Hides the current form and shows the main form
FrmTable.Show
FrmAddAppo.Hide

End Sub

Private Sub cmdTesting_Click()

'if lbltesting doesn't have code yet, this displays
lblTesting = "The WeekID is:  " & WeekID

End Sub

Private Sub Form_Load()

'Get the value of the WeekID
WeekID = GetSetting("shantnu", "timetable", "WeekID", 1)

'The Characters remaining when form loads
lblCharacters = "Characters Remaining: " & 104 - Len(txtAppoDescription)

'Adds the name of the days into the combobox of days
cmbDay.AddItem "Monday"
cmbDay.AddItem "Tuesday"
cmbDay.AddItem "Wednesday"
cmbDay.AddItem "Thursday"
cmbDay.AddItem "Friday"
cmbDay.AddItem "Saturday"
cmbDay.AddItem "Sunday"

'Adds the name of time on the combobox of time
cmbTime.AddItem "0701-0800"
cmbTime.AddItem "0801-0900"
cmbTime.AddItem "0901-1000"
cmbTime.AddItem "1001-1100"
cmbTime.AddItem "1101-1200"
cmbTime.AddItem "1201-1300"
cmbTime.AddItem "1301-1400"
cmbTime.AddItem "1401-1500"
cmbTime.AddItem "1501-1600"
cmbTime.AddItem "1601-1700"
cmbTime.AddItem "1701-1800"
cmbTime.AddItem "1801-1900"
cmbTime.AddItem "1901-2000"
cmbTime.AddItem "2001-2100"
cmbTime.AddItem "2101-2200"
cmbTime.AddItem "2201-2300"
cmbTime.AddItem "2301-0000"


'Gets the title of the subjects and puts them into the listbox
For a = 1 To LOF(2) / Len(Subj)
    Get #2, a, Subj
        lstSubj.AddItem Subj.SubjTitle
        lstDesc.AddItem Subj.SubjDescription
Next

'Makes the lstdesc invisible
lstDesc.Visible = False

'Makes developer testing tools invisible
cmdTesting.Visible = False
lblTesting.Visible = False

End Sub


Private Sub lblTitleAppo_Click()
If cmdTesting.Visible = False Then
    cmdTesting.Visible = True
    lblTesting.Visible = True
Else
    cmdTesting.Visible = False
    lblTesting.Visible = False
End If
End Sub

Private Sub lstSubj_Click()
'adds the text from the listbox to the title
    txtAppoTitle.Text = lstSubj

'If lstSubj.ListIndex = lstDesc.ListIndex Then
    lstDesc.ListIndex = lstSubj.ListIndex
    txtAppoDescription.Text = lstDesc
'Else: End If

End Sub

Private Sub txtAppoDescription_Change()

txtAppoDescription.ForeColor = &H80000008

'Characters Remaining and Limit Reached
If (104 - Len(txtAppoDescription)) > 0 Then
    lblCharacters = "Characters Remaining: " & 104 - Len(txtAppoDescription)
Else
    lblCharacters = "LIMIT REACHED!"
End If

End Sub

Private Sub txtAppoTitle_Change()

txtAppoTitle.ForeColor = &H80000008

End Sub

