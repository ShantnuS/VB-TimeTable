Attribute VB_Name = "Module"
'This is a Module :D (Treat it with respect)

'Defining the appointment record with correct lengths
Type Appointment
    AppoTitle As String * 22
    AppoDescription As String * 104
    AppoID As Integer
    AppoDayID As Integer
    AppoTimeID As Integer
    AppoWeekID As Integer
End Type

Public Appo As Appointment

'Defining the subject record with correct lengths
Type Subject
    SubjTitle As String * 22
    SubjDescription As String * 104
    SubjID As Integer
End Type

Public Subj As Subject




