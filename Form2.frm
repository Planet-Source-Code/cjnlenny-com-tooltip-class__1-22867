VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Tooltip Demo"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   LinkTopic       =   "Form2"
   ScaleHeight     =   1530
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   612
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   1092
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   612
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   612
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1092
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Dim str(1 To 3) As String
    Dim ToolTipFactory As New clsToolTipFactory
    
    str(1) = "This is the first button's tooltip."

    str(2) = "And now we" & vbNewLine & " have a multi-line tooltip" & _
        vbNewLine & "that is very long and very neat!"
    
    str(3) = "And of course, each project's tooltips can " & vbNewLine & _
        "each be a different color entirely or all the same color"
        
    With ToolTipFactory
        .AssignToolTip Command1, str(1)
    
        .ForeColor = vbWhite
        .AssignToolTip Command2, str(2)

        .ForeColor = vbWhite
        .BackColor = vbRed
        .AssignToolTip Command3, str(3)
    End With
    Set ToolTipFactory = Nothing
End Sub



