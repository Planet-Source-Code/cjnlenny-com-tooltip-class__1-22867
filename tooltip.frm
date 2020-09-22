VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   Caption         =   "Cooper's Revolutionary Personalities"
   ClientHeight    =   3636
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5712
   Icon            =   "tooltip.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3636
   ScaleWidth      =   5712
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "Marquis de Lafayette"
      Height          =   2412
      Left            =   3720
      Picture         =   "tooltip.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1692
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "Thomas Gage"
      Height          =   2412
      Left            =   1980
      Picture         =   "tooltip.frx":1B55
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "Arthur St. Claire"
      Height          =   2412
      Left            =   240
      Picture         =   "tooltip.frx":2FD4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1692
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " Hover your mouse over the portrait to learn more."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   4212
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   3252
      Left            =   120
      Top             =   240
      Width           =   5412
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
    Dim str(1 To 3) As String
    Dim ToolTipFactory As New clsToolTipFactory
    
    str(1) = "Arthur St. Clair served as a brigadier general with Gen. George " & vbNewLine & _
        "Washington in the 1776-77 American Revolution battles " & vbNewLine & _
        "of Trenton and Princeton."

    str(2) = "British general and colonial governor in America, " & vbNewLine & _
        "whose aggressive actions against the colonists contributed " & vbNewLine & _
        "to the American Revolution."
    
    str(3) = "French military leader and statesman, who fought on " & vbNewLine & _
        "the side of the colonists during the American Revolution " & vbNewLine & _
        "and later took a prominent part in the French Revolution."
        
    With ToolTipFactory
        .ForeColor = vbYellow
        .BackColor = vbRed
        .AssignToolTip Command1, str(1)
    
        .ForeColor = vbWhite
        .BackColor = vbBlack
        .AssignToolTip Command2, str(2)

        .ForeColor = vbWhite
        .BackColor = vbRed
        .AssignToolTip Command3, str(3)
    End With
    Set ToolTipFactory = Nothing
End Sub

