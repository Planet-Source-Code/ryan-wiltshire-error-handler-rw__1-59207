VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Error Handler"
   ClientHeight    =   855
   ClientLeft      =   7035
   ClientTop       =   5895
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   3375
   Begin VB.CommandButton cmdError 
      Caption         =   "Generate Error"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox cmbError 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oErr As clsErrorHandler
Private Sub cmdError_Click()
On Error GoTo Handler

    Err.Raise Me.cmbError.Text              'Generate Error
    
Exit Sub
Handler:
    'Actions to take based on generic setup
    Select Case oErr.ErrorHandler(Me.Caption & ": Form Load")
        Case ExitForm:      Unload Me
        Case ExitProcedure: Exit Sub
        Case ExitProgram:   End
        Case ResumeAction:  Resume
        Case ResumeNext:    Resume Next
    End Select
End Sub
Private Sub PopulateCombo()
Dim X As Integer
    For X = 3 To 485                'Populate the combo with some error #'s
        Me.cmbError.AddItem X       'add the number to the combo
    Next X
    Me.cmbError.Text = "3"          'set the text equal to the first item
End Sub
Private Sub Form_Load()
    Set oErr = New clsErrorHandler  'Set the instance of the class
    PopulateCombo                   'Fill the combo with error #'s
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set oErr = Nothing
End Sub
