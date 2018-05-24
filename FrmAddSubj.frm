VERSION 5.00
Begin VB.Form FrmAddSubj 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Subject"
   ClientHeight    =   4710
   ClientLeft      =   8970
   ClientTop       =   3975
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4710
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Subject"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtSubjDescription 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   720
      MaxLength       =   104
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox txtSubjTitle 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   720
      MaxLength       =   22
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label lblCharLeft 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblSayDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblSAYTITLE 
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblTitleSubj 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Subject"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3855
      Left            =   600
      Top             =   720
      Width           =   3495
   End
End
Attribute VB_Name = "FrmAddSubj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

'Clicking the add subject button adds the subject to the file
Subj.SubjTitle = txtSubjTitle
Subj.SubjDescription = txtSubjDescription
Subj.SubjID = LOF(2) / Len(Subj) + 1

Put #2, Subj.SubjID, Subj

'Messge box and clear text
MsgBox "Subject Added!"
txtSubjTitle.Text = ""
txtSubjDescription.Text = ""

End Sub

Private Sub cmdBACK_Click()

'Hides current form and goes to main form
FrmAddSubj.Hide
FrmTable.Show

End Sub

Private Sub Form_Load()

'The characters remaining part
lblCharLeft = "Characters Remaining: " & (104 - Len(txtSubjDescription))

End Sub

Private Sub txtSubjDescription_Change()

'Works out how many characters are remaining and if limit has been reached
If (104 - Len(txtSubjDescription)) > 0 Then
    lblCharLeft = "Characters Remaining: " & (104 - Len(txtSubjDescription))
Else
    lblCharLeft = "LIMIT REACHED"
End If

End Sub
