VERSION 5.00
Begin VB.Form frmChangeList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Task"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2655
      Top             =   1845
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   128
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   360
      Width           =   1470
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   435
      TabIndex        =   8
      Top             =   1530
      Width           =   1830
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   250
      Left            =   2468
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   250
      Left            =   1388
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox combxTo 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmChangeList.frx":0000
      Left            =   3248
      List            =   "frmChangeList.frx":0002
      TabIndex        =   5
      Text            =   "High"
      Top             =   360
      Width           =   1455
   End
   Begin VB.ComboBox combxFrom 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmChangeList.frx":0004
      Left            =   1688
      List            =   "frmChangeList.frx":0006
      TabIndex        =   4
      Text            =   "Normal"
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Process 
      Height          =   285
      Left            =   2490
      TabIndex        =   1
      Top             =   2490
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "To:"
      Height          =   255
      Left            =   3248
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "From:"
      Height          =   255
      Left            =   1688
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Process Name:"
      Height          =   255
      Left            =   128
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmChangeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Sub cmdAdd_Click()
Dim listx As ListItem
On Error GoTo MaxExceeded
Call FormOnTop(Me.hWnd, False)
If Process.Text = "" Then
MsgBox "You must enter a process name !"
Combo1.SetFocus
Call Populate
Exit Sub
End If
Call FormOnTop(Me.hWnd, True)

If combxFrom = "Idle" Then
ElseIf combxFrom = "Above Normal" Then
ElseIf combxFrom = "Normal" Then
ElseIf combxFrom = "Below Normal" Then
ElseIf combxFrom = "High" Then
Else
Call FormOnTop(Me.hWnd, False)
MsgBox "You must select a priority !"
Call FormOnTop(Me.hWnd, True)
combxFrom.SetFocus
combxFrom.Text = "Normal"
Exit Sub
End If

If combxTo = "Idle" Then
ElseIf combxTo = "Above Normal" Then
ElseIf combxTo = "Normal" Then
ElseIf combxTo = "Below Normal" Then
ElseIf combxTo = "High" Then
Else
Call FormOnTop(Me.hWnd, False)
MsgBox "You must select a priority !"
Call FormOnTop(Me.hWnd, True)
combxTo.SetFocus
combxTo.Text = "High"
Exit Sub
End If

If combxFrom = "Normal" Then combxFrom.ListIndex = "2"
If combxTo = "High" Then combxTo.ListIndex = "4"
ChangeList(frmPriority.lstvwChangeList.ListItems.Count + 1).Process = Trim(Process.Text)

Set listx = frmPriority.lstvwChangeList.ListItems.Add(frmPriority.lstvwChangeList.ListItems.Count + 1, , ChangeList(frmPriority.lstvwChangeList.ListItems.Count + 1).Process)

Select Case combxFrom.ListIndex
    Case 0:
        listx.SubItems(1) = "Idle"
    Case 1:
        listx.SubItems(1) = "Above Normal"
    Case 2:
        listx.SubItems(1) = "Normal"
    Case 3:
        listx.SubItems(1) = "Below Normal"
    Case 4:
        listx.SubItems(1) = "High"
    Case Else

        Exit Sub
End Select

Select Case combxTo.ListIndex
    Case 0:
        listx.SubItems(2) = "Idle"
    Case 1:
        listx.SubItems(2) = "Above Normal"
    Case 2:
        listx.SubItems(2) = "Normal"
    Case 3:
        listx.SubItems(2) = "Below Normal"
    Case 4:
        listx.SubItems(2) = "High"
    Case Else

        Exit Sub
End Select

frmPriority.ReDefineChangeList
frmPriority.lstbxSystemDialog.AddItem Time & " : Added processes " & ChangeList(frmPriority.lstvwChangeList.ListItems.Count).Process & " to Change List..."
frmPriority.SaveChangeList
Call FormOnTop(frmPriority.hWnd, True)
DoEvents
Unload frmChangeList
Exit Sub
MaxExceeded:
Call FormOnTop(Me.hWnd, False)
MsgBox "Sorry, this program is set to control only 200 entrys !"
Call FormOnTop(Me.hWnd, True)
Call FormOnTop(frmPriority.hWnd, True)
DoEvents
Unload frmChangeList
End Sub

Private Sub cmdCancel_Click()

Unload frmChangeList 'Exit without doing anything

'Me.Hide
End Sub

Private Sub Combo1_Change()
Process.Text = Combo1
End Sub

Private Sub Form_Load()
btnFlat cmdCancel
btnFlat cmdAdd
Call FormOnTop(Me.hWnd, True)
DoEvents
combxFrom.Clear                                                                 'Populate Combo Boxes
combxFrom.Text = "Normal"
combxFrom.AddItem "High"
combxFrom.AddItem "Above Normal"
combxFrom.AddItem "Normal"
combxFrom.AddItem "Below Normal"
combxFrom.AddItem "Idle"
'combxFrom.AddItem "Highest"
combxTo.Clear
combxTo.Text = "High"
combxTo.AddItem "High"
combxTo.AddItem "Above Normal"
combxTo.AddItem "Normal"
combxTo.AddItem "Below Normal"
combxTo.AddItem "Idle"
'combxTo.AddItem "Highest"
Call Populate

End Sub
Private Sub Timer1_Timer()
Process.Text = Combo1
End Sub
Private Sub Populate()
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
Private Function btnFlat(Button As CommandButton)
SetWindowLong cmdCancel.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
cmdCancel.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
SetWindowLong cmdAdd.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
cmdAdd.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function

