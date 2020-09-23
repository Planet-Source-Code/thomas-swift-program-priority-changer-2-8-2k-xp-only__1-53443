VERSION 5.00
Begin VB.Form frmKill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kill"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   2490
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   75
      Top             =   1410
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   945
      Width           =   2100
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   250
      Left            =   1193
      TabIndex        =   3
      Top             =   570
      Width           =   915
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   510
      TabIndex        =   2
      Top             =   1485
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kill"
      Height          =   250
      Left            =   383
      TabIndex        =   1
      Top             =   570
      Width           =   690
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   150
      Width           =   2130
   End
End
Attribute VB_Name = "frmKill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Sub Populate()
List1.Clear
For myloop = 1 To frmPriority.lstvwProcesses.ListItems.Count '- 1
List1.AddItem frmPriority.lstvwProcesses.ListItems.item(myloop)
Next myloop
xListKillDupes List1
Combo1.Clear
Combo1.Text = List1.List(0)
For myloop2 = 0 To List1.ListCount - 1
Combo1.AddItem List1.List(myloop2)
Next myloop2
End Sub
Private Sub Command1_Click()
Dim mystring As String
mystring = Combo1
Killapp (Combo1)
Unload frmKill
End Sub
Private Sub Command2_Click()
Unload frmKill
End Sub


Private Sub Form_Load()
btnFlat Command1
btnFlat Command2
Call FormOnTop(Me.hWnd, True)
Populate
End Sub
Private Sub Timer1_Timer()
Text1.Text = Combo1
End Sub
Private Function btnFlat(Button As CommandButton)
SetWindowLong Command1.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
Command1.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
SetWindowLong Command2.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
Command2.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
    
End Function
