VERSION 5.00
Begin VB.Form frmRemove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remove"
   ClientHeight    =   450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1050
      Width           =   2745
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   250
      Left            =   1972
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove"
      Height          =   250
      Left            =   547
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Width           =   1200
   End
End
Attribute VB_Name = "frmRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Sub Command1_Click()
frmPriority.lstvwChangeList.ListItems.Remove (frmPriority.progindex)
frmPriority.ReDefineChangeList                                                              'update ChangeList
frmPriority.lstbxSystemDialog.AddItem Time & " : Removed process " & frmPriority.progname & " from Change List..."
frmPriority.SaveChangeList 'save ChangeList
frmPriority.lstvwChangeList.Visible = True
Unload frmRemove
End Sub
Private Sub Command2_Click()
frmPriority.lstvwChangeList.Visible = True
Unload frmRemove
End Sub
Private Sub Form_Load()
Call FormOnTop(Me.hWnd, True)
btnFlat Command1
btnFlat Command2
Me.Caption = "Remove item: " & frmPriority.progname
'Label1.Caption = "Remove item: " & frmPriority.progname
End Sub
Private Function btnFlat(Button As CommandButton)
SetWindowLong Command1.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
Command1.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
SetWindowLong Command2.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
Command2.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function
