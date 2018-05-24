VERSION 5.00
Begin VB.Form FrmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print"
   ClientHeight    =   2310
   ClientLeft      =   8970
   ClientTop       =   4695
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   3645
   Begin VB.CommandButton cmdBACK 
      Caption         =   "Back"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.ComboBox cmbColour 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FrmPrint.frx":0000
      Left            =   240
      List            =   "FrmPrint.frx":000A
      TabIndex        =   3
      Text            =   "(Select)"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtCopies 
      Appearance      =   0  'Flat
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "(Type Here)"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblAWESOME 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print"
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
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Colour Mode:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Copies:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1575
      Left            =   120
      Top             =   480
      Width           =   3375
   End
End
Attribute VB_Name = "FrmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBACK_Click() 'Clickity click goes back
'Goes back to main form
FrmPrint.Hide
FrmTable.Show
End Sub

Private Sub cmdPrint_Click() 'None shall print!

If txtCopies = "" Then
    MsgBox "Enter number of copies!"
Else
    If IsNumeric(txtCopies) = False Then
        MsgBox "That is not a number!"
    Else
        If txtCopies < 0 Then
            MsgBox "That is not a valid number of copies!"
        Else
            If cmbColour.ListIndex = -1 Then
                MsgBox "Choose colour!"
            Else
                If txtCopies = 0 Then
                    MsgBox "No copies will be printed"
                Else
                    'Allows the form to be printed in landscape
                        Printer.Orientation = vbPRORLandscape
                    'The number of copies
                        Printer.Copies = Val(txtCopies)
                    'Prints the Timetable form
                        FrmTable.PrintForm
                    'Print in black and white or colour
                        Printer.ColorMode = cmbColour.ListIndex + 1
                End If
            End If
        End If
    End If
End If

End Sub


Private Sub txtCopies_Click()
'Change colour of text box and delete default text
txtCopies.ForeColor = &H80000008
txtCopies.Text = ""
End Sub
