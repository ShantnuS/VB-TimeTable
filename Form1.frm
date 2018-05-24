VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form FrmTable 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TimeTable"
   ClientHeight    =   8235
   ClientLeft      =   5040
   ClientTop       =   1635
   ClientWidth     =   12555
   ForeColor       =   &H00FFFFC0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   549
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   837
   Begin VB.CommandButton cmdRndSetting 
      Caption         =   "Random Subject Settings"
      Height          =   615
      Left            =   11520
      TabIndex        =   13
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Random Subject"
      Height          =   615
      Left            =   11280
      TabIndex        =   12
      Top             =   5760
      Width           =   1095
   End
   Begin VB.ComboBox cmbWeekID 
      Height          =   315
      Left            =   10080
      TabIndex        =   11
      Text            =   "(Pick Week)"
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "GO!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      Picture         =   "Form1.frx":0000
      TabIndex        =   10
      Top             =   240
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   7095
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   12515
      _Version        =   393216
      BackColor       =   16777152
      Rows            =   18
      BackColorFixed  =   14737632
      FocusRect       =   0
      Appearance      =   0
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   615
      Left            =   10080
      Picture         =   "Form1.frx":0B8C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Timer ColourTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5760
      Top             =   0
   End
   Begin VB.CommandButton cmdAddSubj 
      Caption         =   "Add Subject"
      Height          =   735
      Left            =   11280
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddAppo 
      Caption         =   "Add Appointment"
      Height          =   735
      Left            =   10080
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Timer ClockTimer 
      Interval        =   99
      Left            =   6240
      Top             =   0
   End
   Begin VB.CommandButton cmdTestButton 
      Height          =   1215
      Left            =   9960
      Picture         =   "Form1.frx":117F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label lblTest 
      BackColor       =   &H00FF80FF&
      Caption         =   "(click test button)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   7560
      Width           =   11295
   End
   Begin VB.Label lblAWESOME 
      BackStyle       =   0  'Transparent
      Caption         =   "TimeTable Program!"
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
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Height          =   3735
      Left            =   10080
      TabIndex        =   7
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FrmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TimeTable Program for A2 Level computing
'Last updated 17/11/2014
'Shantnu Singh
'**************************************************************************************************

'Variable declarations
Dim Lastweek As Date 'lastweek needs to be date
Public WeekID As Integer 'Arithmetics needs to be done on WeekID, it is integer
Dim Appointment As Appointment 'The appointment is an appointment
Dim RandomSubj As Integer 'Random SubjectID for random subject
Dim DoNum As Integer
Dim RandomDay As Integer 'Random DayID for random subject
Dim RandomTime As Integer 'Random TimeID for random subject
Dim StartTime As Integer 'Minimum value for random TimeID
Dim EndTime As Integer 'Maximum value for random TimeID
Dim WeekPointer As Integer 'Allows specific appointments with certain WeekID to be opened


'Random RGB Colour (REALLY IMPORTANT)
Public Function RandomRGBColour() As Long
    RandomRGBColour = RGB( _
        Int(Rnd() * 256), _
        Int(Rnd() * 256), _
        Int(Rnd() * 256))
End Function

'**************************************************************************************************

Private Sub ClockTimer_Timer() 'Timer for time and date on top

'This is to display the time at the top of the screen
lblTime = Time

'Shows current date at top
lblDate = Date

End Sub

Private Sub cmdAddAppo_Click() 'Add Appointment button

'Shows form to add appointment
FrmAddAppo.Show

'Saves the WeekID to be used on the other forms
SaveSetting "shantnu", "timetable", "lastWeek", Lastweek

End Sub

Private Sub cmdAddSubj_Click() 'Add subject button

'Shows the add subject form
FrmAddSubj.Show

End Sub

Private Sub cmdPrint_Click() 'Print button

'Opens form of extreme printing!
FrmPrint.Show

End Sub

Private Sub cmdRandom_Click()
'calls random generator sub when clicked
RandomGen
End Sub

Private Sub cmdRefresh_Click()
'calls populate grid sub when clicked
PopulatesGrid
End Sub

Private Sub cmdRndSetting_Click()
'Shows random generator settings form
frmSettings.Show
End Sub

Private Sub cmdTestButton_Click() 'A button for testing

'This is only for testing. It shows if the variables are changing correctly
lblTest = "WeekID is :" & WeekID & "   Lastweek is: " & Lastweek

End Sub


Private Sub ColourTimer_Timer() 'A Timer to change the colour of the title
'Changes colour of title
lblAWESOME.ForeColor = RandomRGBColour
End Sub

Private Sub Daylbl_Click(Index As Integer)
'Refreshes the flex grid
PopulatesGrid
End Sub

Private Sub Form_Load() 'When Form opens
lblTest.Visible = False
cmdTestButton.Visible = False


'This allows only certain appointments to load up (only one with certain weekID)
WeekPointer = WeekID

'Opening Files when the form opens
Open App.Path & "/AppoData.txt" For Random As #1 Len = Len(Appo)
Open App.Path & "/SubjData.txt" For Random As #2 Len = Len(Subj)
'Open App.Path & "/Appointments.txt" For Random As #3 Len = Len(Apment)

'This is the WeekID from the registry, sets default value of 1 if registry value is empty
WeekID = GetSetting("shantnu", "timetable", "WeekID", 1)

'This is the LastWeek from the registry, sets default value of today if registry value is empty
Lastweek = GetSetting("shantnu", "timetable", "lastWeek", Date)

'This is the algorithm that adds one to the WeekID every 7 days and also makes lastweek today
If Date >= Lastweek + 7 Then
    WeekID = WeekID + 1
    Lastweek = Date
Else: End If

'FLEX GRID
    'Grid has more columns
    MSHFlexGrid1.Cols = 8
    MSHFlexGrid1.Rows = 18
    'Allow Resizing
    MSHFlexGrid1.AllowUserResizing = flexResizeBoth
    MSHFlexGrid1.RowSizingMode = flexRowSizeAll

'Fills the table upon form load
For e = 1 To LOF(1) / Len(Appo)
Get #1, e, Appo
    If Appo.AppoWeekID = WeekPointer + 1 Then
        MSHFlexGrid1.TextMatrix(Appo.AppoTimeID, Appo.AppoDayID) = Appo.AppoTitle
    Else: End If
Next

'Fills table with Titles
GridTitleTime
GridTitleDay

'Adds weekid to combo box at top
For U = 1 To WeekID Step 1
    cmbWeekID.AddItem (U)
Next

'This part is for testing only, manually changes max-WeekID to 2.
'For U = 1 To 2 Step 1
'    cmbWeekID.AddItem (U)
'Next


'donum value
DoNum = 0


End Sub

Private Sub GridTitleTime()
'Adds the time titles onto the flexgrid
MSHFlexGrid1.TextMatrix(1, 0) = "0701-0800"
MSHFlexGrid1.TextMatrix(2, 0) = "0801-0900"
MSHFlexGrid1.TextMatrix(3, 0) = "0901-1000"
MSHFlexGrid1.TextMatrix(4, 0) = "1001-1100"
MSHFlexGrid1.TextMatrix(5, 0) = "1101-1200"
MSHFlexGrid1.TextMatrix(6, 0) = "1201-1300"
MSHFlexGrid1.TextMatrix(7, 0) = "1301-1400"
MSHFlexGrid1.TextMatrix(8, 0) = "1401-1500"
MSHFlexGrid1.TextMatrix(9, 0) = "1501-1600"
MSHFlexGrid1.TextMatrix(10, 0) = "1601-1700"
MSHFlexGrid1.TextMatrix(11, 0) = "1701-1800"
MSHFlexGrid1.TextMatrix(12, 0) = "1801-1900"
MSHFlexGrid1.TextMatrix(13, 0) = "1901-2000"
MSHFlexGrid1.TextMatrix(14, 0) = "2001-2100"
MSHFlexGrid1.TextMatrix(15, 0) = "2101-2200"
MSHFlexGrid1.TextMatrix(16, 0) = "2201-2300"
MSHFlexGrid1.TextMatrix(17, 0) = "2301-0000"
End Sub

Private Sub GridTitleDay()
'Adds the day titles onto the flexgrid
MSHFlexGrid1.TextMatrix(0, 1) = "Monday"
MSHFlexGrid1.TextMatrix(0, 2) = "Tuesday"
MSHFlexGrid1.TextMatrix(0, 3) = "Wednesday"
MSHFlexGrid1.TextMatrix(0, 4) = "Thursday"
MSHFlexGrid1.TextMatrix(0, 5) = "Friday"
MSHFlexGrid1.TextMatrix(0, 6) = "Saturday"
MSHFlexGrid1.TextMatrix(0, 7) = "Sunday"

End Sub


Public Sub RandomGen()
'Get values of start and end time
StartTime = GetSetting("shantnu", "timetable", "StartTime", 1)
EndTime = GetSetting("shantnu", "timetable", "EndTime", 17)

'Generate random x and y values for grid
RandomTime = Int((EndTime - StartTime + 1) * Rnd + StartTime)
RandomDay = Int((Rnd * 7) + 1)

'finds a random number based on length of subject file
RandomSubj = Int((Rnd * (LOF(2) / Len(Subj))) + 1)

'Gets a subject with the same subject id as the random number
Get #2, RandomSubj, Subj

    'Checks if random grid is empty or not
    If MSHFlexGrid1.TextMatrix(RandomTime, RandomDay) = "" Then
            'Fill the random empty box with random subject
            MSHFlexGrid1.TextMatrix(RandomTime, RandomDay) = Subj.SubjTitle
            
            'Change the backcolour of the random subject cell
            MSHFlexGrid1.Row = RandomTime
            MSHFlexGrid1.Col = RandomDay
            MSHFlexGrid1.CellBackColor = &HFFFF00
            
            'Add the random subject to the appointment file
            Appo.AppoID = LOF(1) / Len(Appo) + 1
            Appo.AppoDayID = RandomDay
            Appo.AppoTimeID = RandomTime
            Appo.AppoDescription = Subj.SubjDescription
            Appo.AppoTitle = Subj.SubjTitle
            Appo.AppoWeekID = WeekID
            
            Put #1, Appo.AppoID, Appo
    Else: End If
    
End Sub

Public Sub PopulatesGrid()

'Appointment

'Clear existing grid
MSHFlexGrid1.Clear

'Refill titles
GridTitleDay
GridTitleTime

'Get data from file one
WeekPointer = cmbWeekID.ListIndex

If WeekPointer = -1 Then
    WeekPointer = 0
Else: End If

For e = 1 To LOF(1) / Len(Appo)
Get #1, e, Appo
    If Appo.AppoWeekID = WeekPointer + 1 Then
        MSHFlexGrid1.TextMatrix(Appo.AppoTimeID, Appo.AppoDayID) = Appo.AppoTitle
    Else: End If
Next

End Sub
Private Sub Form_Unload(Cancel As Integer) 'When Form closes

'This saves the value of lastweek to the reqistry when form is unloaded
SaveSetting "shantnu", "timetable", "lastWeek", Lastweek

'This saves the value of WeekID to the registry when form is unloaded
SaveSetting "shantnu", "timetable", "WeekID", WeekID

'Unload the whole program


End Sub

Private Sub lblAWESOME_Click() 'Toggles colour changing title

'Disables/enables the flashing colour title
If ColourTimer.Enabled = False Then
    ColourTimer.Enabled = True
Else
    ColourTimer.Enabled = False
    lblAWESOME.ForeColor = &H80000012
End If

End Sub

Private Sub lblTime_Click()
If lblTest.Visible = False Then
    cmdTestButton.Visible = True
    lblTest.Visible = True
Else
    lblTest.Visible = False
    cmdTestButton.Visible = False
End If

End Sub

Private Sub MSHFlexGrid1_Click()
'Makes the appointment description show up in the description label

For n = 1 To LOF(1) / Len(Appo)
Get #1, n, Appo
    If Appo.AppoTitle = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.Col) Then
        lblDesc = Appo.AppoDescription
    Else: End If
Next n


'Adds appointment from previous cell to next empty cell when clicked
If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.Col) = "" Then
    If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row - 1, MSHFlexGrid1.Col) <> "" Then
        lblDesc = ""
        If MSHFlexGrid1.Row - 1 <> 0 Then
        
            'Title is value of previous row
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.Col) = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row - 1, MSHFlexGrid1.Col)
            
            'Adds this to the appointment file
            Appo.AppoTitle = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, MSHFlexGrid1.Col)
            Appo.AppoID = LOF(1) / Len(Appo) + 1
            Appo.AppoWeekID = WeekID
            Appo.AppoDayID = MSHFlexGrid1.Col
            Appo.AppoTimeID = MSHFlexGrid1.Row
            'Adds description
            Put #1, Appo.AppoID, Appo
    
        Else: End If
    Else: End If
Else: End If


End Sub
