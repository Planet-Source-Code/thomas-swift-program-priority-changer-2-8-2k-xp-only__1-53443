VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPriority 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Priority Changer (v2.8) rev. 4-26-2004 T.A.S. Independent Programming"
   ClientHeight    =   4185
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   9420
   Icon            =   "frmPiority.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   9420
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Kill A Process"
      Height          =   250
      Left            =   7088
      TabIndex        =   10
      Top             =   3870
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   330
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   4800
      Width           =   1005
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7455
      Top             =   4245
   End
   Begin VB.Frame fmeUpdateFrequency 
      Caption         =   "Update Frequency"
      Height          =   1575
      Left            =   4515
      TabIndex        =   3
      Top             =   240
      Width           =   1815
      Begin VB.OptionButton optCustom 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton opt1Second 
         Caption         =   "Every Second"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton opt2Seconds 
         Caption         =   "Every 2 Seconds"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton opt5Seconds 
         Caption         =   "Every 5 Seconds"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblCustom 
         Caption         =   "[ custom ]"
         Height          =   240
         Left            =   360
         TabIndex        =   8
         Top             =   1087
         Width           =   1335
      End
   End
   Begin VB.Frame fmeOptions 
      Caption         =   "Program Process Change Settings"
      Height          =   1980
      Left            =   30
      TabIndex        =   2
      Top             =   45
      Width           =   6375
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   250
         Left            =   795
         TabIndex        =   14
         Top             =   1650
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   250
         Left            =   1875
         TabIndex        =   13
         Top             =   1650
         Width           =   1770
      End
      Begin MSComctlLib.ListView lstvwChangeList 
         Height          =   1365
         Left            =   60
         TabIndex        =   15
         Top             =   225
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   2408
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Timer Timer1 
      Left            =   6990
      Top             =   4260
   End
   Begin VB.ListBox lstbxSystemDialog 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   443
      TabIndex        =   1
      Top             =   2250
      Width           =   5535
   End
   Begin MSComctlLib.ListView lstvwProcesses 
      Height          =   3540
      Left            =   6428
      TabIndex        =   11
      Top             =   285
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6244
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Current Running Processes"
      Height          =   255
      Left            =   6855
      TabIndex        =   12
      Top             =   45
      Width           =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "System Dialog"
      Height          =   255
      Left            =   2685
      TabIndex        =   0
      Top             =   2010
      Width           =   1065
   End
   Begin VB.Menu Command1 
      Caption         =   "Hide"
   End
   Begin VB.Menu mnuCHUpdate 
      Caption         =   "Check for update"
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuSysTrayShow 
         Caption         =   "Show"
      End
      Begin VB.Menu goober1 
         Caption         =   "---------------------"
      End
      Begin VB.Menu mnuSysTrayExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmPriority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*Priority: Process Priority Changer program for win98, winNT kernels           *
'*    The program operates on a few main principles. The handles and processes  *
'*    arrays declared at the beginning are used by the program to store that    *
'*    information whenever it updates.  The core of the program are the subs    *
'*    that change priority and what is found in the timer1.timer routine. Most  *
'*    of the rest of the code is just standard Windows event handling.          *
'********************************************************************************
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Dim Handles(1 To 200) As Long                                                   'stores the current process names
Dim Processes(1 To 200) As String 'stores the corresponding process handle
Dim CheckString As String
Public progname As String
Public progindex As Long
Private Sub cmdAdd_Click()
Text1.SetFocus
frmChangeList.Show                                                           'shows the other form for task addition
End Sub
Private Sub cmdRemove_Click()
progname = lstvwChangeList.SelectedItem.Text
progindex = lstvwChangeList.SelectedItem.Index
frmRemove.Show
End Sub
Private Sub Command1_Click()
Text1.SetFocus
Me.Visible = False
AddToTray Me, "T.A.S. Program Priority Changer", Me.Icon
End Sub
Private Sub Command2_Click()
Text1.SetFocus
frmKill.Show
End Sub

Private Sub Form_Load()
If App.PrevInstance Then End
btnFlat Command2
btnFlat cmdAdd
btnFlat cmdRemove
Me.Visible = False
AddToTray Me, "T.A.S. Program Priority Changer", Me.Icon

CheckString = "1"

lstvwProcesses.ListItems.Clear                                                  'setup List Controls
lstvwProcesses.ColumnHeaders.Clear
lstvwChangeList.ColumnHeaders.Add 1, "Name", "Name", 1640
lstvwChangeList.ColumnHeaders.Add 2, "From", "From", 1200
lstvwChangeList.ColumnHeaders.Add 3, "To", "To", 1200
lstvwProcesses.ColumnHeaders.Add 1, "Name", "Name", 1640
lstvwProcesses.ColumnHeaders.Add 2, "Priority", "Priority", 1200
lstvwProcesses.SortKey = 1
lstbxSystemDialog.AddItem Time & " : Starting program..."

Dim Itmx As ListItem
Dim j As Long
Dim cb As Long
Dim reply, reply2 As Long
Dim temp, temp2, temp3 As String
Dim cbNeeded As Long
Dim NumElements As Long
Dim ProcessIDs() As Long
Dim cbNeeded2 As Long
Dim NumElements2 As Long
Dim Modules(1 To 200) As Long
Dim lRet As Long
Dim ModuleName As String
Dim nSize As Long
Dim hProcess As Long
Dim I As Long
Dim Priority As Long
Dim fNum As Integer

On Error GoTo Errorhandler
fNum = FreeFile
Open (App.Path & "\setup.txt") For Input As fNum
  
Do Until EOF(fNum)
     Line Input #fNum, temp
     reply = InStr(1, temp, Chr(42))
     ChangeList(frmPriority.lstvwChangeList.ListItems.Count + 1).Process = Trim(Left(temp, reply - 1))
     Set Itmx = lstvwChangeList.ListItems.Add(frmPriority.lstvwChangeList.ListItems.Count + 1, , ChangeList(frmPriority.lstvwChangeList.ListItems.Count + 1).Process)
     reply2 = InStr(reply + 1, temp, Chr(42))
     temp3 = Trim(Right(temp, Len(temp) - reply2))
     temp2 = Trim(Mid(temp, reply + 1, (Len(temp) - (Len(ChangeList(frmPriority.lstvwChangeList.ListItems.Count).Process) + Len(temp3)) - 2)))

     Select Case temp2
     Case "High":
         Itmx.SubItems(1) = "High"
     Case "Idle":
         Itmx.SubItems(1) = "Idle"
     Case "Above Normal":
         Itmx.SubItems(1) = "Above Normal"
     Case "Normal":
         Itmx.SubItems(1) = "Normal"
     Case "Below Normal":
         Itmx.SubItems(1) = "Below Normal"
     Case "Highest":
         Itmx.SubItems(1) = "Highest"
     End Select
     
     Select Case temp3
     Case "High":
         Itmx.SubItems(2) = "High"
     Case "Idle":
         Itmx.SubItems(2) = "Idle"
     Case "Above Normal":
         Itmx.SubItems(2) = "Above Normal"
     Case "Normal":
         Itmx.SubItems(2) = "Normal"
     Case "Below Normal":
         Itmx.SubItems(2) = "Below Normal"
     Case "Highest":
         Itmx.SubItems(2) = "Highest"
     End Select
Loop

ReDefineChangeList
Close fNum

Rest:
                                                                                'Get the array containing the process id's for each process object
cb = 8
cbNeeded = 96
Do While cb <= cbNeeded
    cb = cb * 2
    ReDim ProcessIDs(cb / 4) As Long
    lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
Loop
         
NumElements = cbNeeded / 4
j = 1                                                                           'j keeps track of index for lstvwProcesses
For I = 1 To NumElements
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, ProcessIDs(I))                'Get a handle to the Process
    If hProcess <> 0 Then                                                       'Got a Process handle
        lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)         'Get an array of the module handles for the specified process
        If lRet <> 0 Then                                                       'If the Module Array is retrieved, Get the ModuleFileName
            ModuleName = Space(MAX_PATH)                                        'Prepare variables...
            nSize = 500
            lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize) 'Get process name
            Handles(I) = hProcess                                               'Assign handle to array
            Processes(I) = TrimNull(Trim(Right(ModuleName, _
                (Len(ModuleName) - InStrRev(ModuleName, "\")))))                 'Assign name to array
            Set Itmx = lstvwProcesses.ListItems.Add(j, , Processes(I))          'Add process to Process List
            Priority = GetPriorityClass(Handles(I))                             'Retrieve process priority
            AddProcessListSubItems Priority, Itmx                               'Add sub item info to Process List
            ChangePriority I, Priority                                          'Call Change Priority Sub Routine
            j = j + 1
        End If
        lRet = CloseHandle(hProcess)                                            'Close the handle to the process
    End If
Next
lstbxSystemDialog.ListIndex = lstbxSystemDialog.ListCount - 1                   'select the most recent entry
Timer1.Interval = 1000




Exit Sub

Errorhandler:                                                                   'Deals with no setup.txt
    Call FormOnTop(Me.hWnd, False)
    MsgBox ("There is no setup file. Please setup the Change List.")
    Call FormOnTop(Me.hWnd, True)
    GoTo Rest
End Sub
Private Sub Form_Unload(Cancel As Integer)
RemoveFromTray
End Sub

Private Sub MnuClose_Click()
Dim frm As Form
For Each frm In Forms
Unload frm
Next
End Sub

Private Sub mnuCHUpdate_Click()
Shell ("C:\Program Files\Internet Explorer\IEXPLORE.EXE " & "http://www.tas-independent-programming.com/My_Other_Programs.htm"), vbMaximizedFocus
Me.Visible = False
AddToTray Me, "T.A.S. Program Priority Changer", Me.Icon
End Sub

Sub opt1Second_Click()
If opt1Second.Value = True Then Timer1.Interval = 1000
lstbxSystemDialog.AddItem Time & " : Changed Update Frequency to 1 Second."
CheckString = "1"
End Sub

Sub opt2Seconds_Click()
If opt2Seconds.Value = True Then Timer1.Interval = 2000
lstbxSystemDialog.AddItem Time & " : Changed Update Frequency to 2 Seconds."
CheckString = "2"
End Sub

Sub opt5Seconds_Click()
If opt5Seconds.Value = True Then Timer1.Interval = 5000
lstbxSystemDialog.AddItem Time & " : Changed Update Frequency to 5 Seconds."
CheckString = "3"
End Sub

Sub optCustom_Click()
Dim reply As String
Dim temp As Double
Comehere:
If optCustom.Value = True Then
Call FormOnTop(Me.hWnd, False)
reply = InputBox("Enter desired update frequency in seconds.")
Call FormOnTop(Me.hWnd, True)
If IsNumeric(reply) Then
    If CLng(reply) > 65 Then
        Call FormOnTop(Me.hWnd, False)
        MsgBox "You have entered a value larger than 65 seconds. This is not supported. Using 65 seconds as max."
        Call FormOnTop(Me.hWnd, True)
        reply = CStr(65)
    End If
    temp = CLng(reply) * 1000
    Timer1.Interval = temp
    lstbxSystemDialog.AddItem Time & " : Changed update frequency to every " & reply & " seconds."
    lblCustom.Caption = "Every " & reply & " Seconds"
Else
 optCustom.Value = False
 If CheckString = "1" Then
 opt1Second.Value = True
 opt1Second_Click
 ElseIf CheckString = "2" Then
 opt2Seconds.Value = True
 opt2Seconds_Click
 Else
 opt5Seconds.Value = True
 opt5Seconds_Click
 End If
End If
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
lstvwProcesses.ListItems.Clear

Dim Itmx As ListItem
Dim j, k As Long
Dim cb As Long
Dim reply As Long
Dim temp As String * 200
Dim cbNeeded As Long
Dim NumElements As Long
Dim ProcessIDs() As Long
Dim cbNeeded2 As Long
Dim NumElements2 As Long
Dim Modules(1 To 200) As Long
Dim lRet As Long
Dim ModuleName As String
Dim nSize As Long
Dim hProcess As Long
Dim I As Long
Dim Priority As Long

                                                                               'Get the array containing the process id's for each process object
cb = 8
cbNeeded = 96
Do While cb <= cbNeeded
    cb = cb * 2
    ReDim ProcessIDs(cb / 4) As Long
    lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
Loop
         
NumElements = cbNeeded / 4
j = 1                                                                           'j keeps track of index for lstvwProcesses
0 For I = 1 To NumElements
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, ProcessIDs(I))                'Get a handle to the Process
    If hProcess <> 0 Then                                                       'Got a Process handle
        lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)         'Get an array of the module handles for the specified process
        If lRet <> 0 Then                                                       'If the Module Array is retrieved, Get the ModuleFileName
            ModuleName = Space(MAX_PATH)                                        'Prepare variables...
            nSize = 500
            lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize) 'Get process name
            Handles(I) = hProcess                                               'Assign handle to array
            Processes(I) = Trim(Right(ModuleName, _
                (Len(ModuleName) - InStrRev(ModuleName, "\"))))                 'Assign name to array
            Set Itmx = lstvwProcesses.ListItems.Add(j, , Processes(I))          'Add process to Process List
            Priority = GetPriorityClass(Handles(I))                             'Retrieve process priority
            AddProcessListSubItems Priority, Itmx                               'Add sub item info to Process List
            ChangePriority I, Priority                                          'Call Change Priority Sub Routine
            j = j + 1
        End If
        lRet = CloseHandle(hProcess)                                            'Close the handle to the process
    End If
Next
lstbxSystemDialog.ListIndex = lstbxSystemDialog.ListCount - 1                   'select the most recent entry
End Sub

Sub ReDefineChangeList()                                                        'This sub updates the Changelist type
Dim Itmx As ListItem                                                            'so that it corresponds to what is seen
Dim temp As String                                                              'in the lstvwChangeList

For I = 1 To lstvwChangeList.ListItems.Count
    Set Itmx = lstvwChangeList.ListItems.item(I)
    ChangeList(I).Process = Itmx.Text
    
    temp = Itmx.SubItems(1)
    Select Case temp
        Case "Idle":
            ChangeList(I).From = IDLE_PRIORITY_CLASS
        Case "Above Normal":
            ChangeList(I).From = ABOVE_NORMAL_PRIORITY_CLASS
        Case "Normal":
            ChangeList(I).From = NORMAL_PRIORITY_CLASS
        Case "Below Normal":
            ChangeList(I).From = BELOW_NORMAL_PRIORITY_CLASS
        Case "High":
            ChangeList(I).From = HIGH_PRIORITY_CLASS
        Case "Highest":
            ChangeList(I).From = REALTIME_PRIORITY_CLASS
    End Select
    
    temp = Itmx.SubItems(2)
    Select Case temp
        Case "Idle":
            ChangeList(I).To = IDLE_PRIORITY_CLASS
        Case "Above Normal":
            ChangeList(I).To = ABOVE_NORMAL_PRIORITY_CLASS
        Case "Normal":
            ChangeList(I).To = NORMAL_PRIORITY_CLASS
        Case "Below Normal":
            ChangeList(I).To = BELOW_NORMAL_PRIORITY_CLASS
        Case "High":
            ChangeList(I).To = HIGH_PRIORITY_CLASS
        Case "Highest":
            ChangeList(I).To = REALTIME_PRIORITY_CLASS
    End Select
Next I

End Sub

Sub AddProcessListSubItems(Priority As Long, Itmx As ListItem)                      'This routine adds to the lstvwProcesses priority header
'Public Const BELOW_NORMAL_PRIORITY_CLASS = &H4000
'Public Const ABOVE_NORMAL_PRIORITY_CLASS = 32768
Select Case Priority                                                                'Add Item case
    Case ABOVE_NORMAL_PRIORITY_CLASS:
         Itmx.SubItems(1) = "Above Normal"
    Case NORMAL_PRIORITY_CLASS:
         Itmx.SubItems(1) = "Normal"
    Case BELOW_NORMAL_PRIORITY_CLASS:
         Itmx.SubItems(1) = "Below Normal"
    Case IDLE_PRIORITY_CLASS:
         Itmx.SubItems(1) = "Idle"
    Case HIGH_PRIORITY_CLASS:
         Itmx.SubItems(1) = "High"
    Case REALTIME_PRIORITY_CLASS:
         Itmx.SubItems(1) = "Realtime"
End Select
End Sub

Sub ChangePriority(I As Long, Priority As Long)                                     'This routine checks the priority with the Changelist
Dim k As Integer                                                                    'and changes the priority if necessary

For k = 1 To lstvwChangeList.ListItems.Count                                        'Begin loop that cycles through Change List
    If InStr(1, UCase(Processes(I)), UCase(ChangeList(k).Process)) Then             'Is Process in Change List?
        If Priority = ChangeList(k).From Then                                       'Does Priority need to be changed?
            reply = SetPriorityClass(Handles(I), ChangeList(k).To)                  'change priority
            If reply = 0 Then                                                       'If error, then explain..
                reply = GetLastError
                If reply = 5 Then
                    lstbxSystemDialog.AddItem Time & " : Error #: " & reply & " occured for process: " & Processes(I)
                End If
            Else
                lstbxSystemDialog.AddItem Time & " : Changed priority of process: " & _
                    Processes(I)                                                    'Successfully changed
            End If
        End If
    End If
Next k
End Sub

Sub SaveChangeList()                                                                'This routine saves the Change List based on
Dim fNum As Integer                                                                 'the information in the lstvwChangeList treeview
Dim Itmx As ListItem
Dim temp As String
Dim I As Integer

fNum = FreeFile                                                                     'Get available File number

Open (App.Path & "\setup.txt") For Output As fNum                                   'Open File

For I = 1 To lstvwChangeList.ListItems.Count                                        'Cylce through list items and save
    Set Itmx = lstvwChangeList.ListItems.item(I)
    temp = Itmx.Text & Chr(42) & Itmx.SubItems(1) & Chr(42) & Itmx.SubItems(2)
    Print #fNum, temp                                                               'Print to file
Next I

lstbxSystemDialog.AddItem Time & " : Saved new Change List..."                      'Update console
Close fNum                                                                          'Close file

End Sub
Private Sub mnuSysTrayExit_Click()
Dim frm As Form
For Each frm In Forms
Unload frm
Next
End Sub
Private Sub mnuSysTrayShow_Click()
Timer2.Enabled = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Message As Long
   On Error Resume Next
    Message = X / Screen.TwipsPerPixelX
    Select Case Message
        Case WM_RBUTTONUP
            'Something useful I just found out:
            ' You need to verify the height, otherwise
            ' it'll pop up the menu mid-form, if the
            ' form is big enough
            'temp = GetY
            'If temp > (Screen.Height / Screen.TwipsPerPixelY) - 30 Then
                Me.Visible = True
                Call FormOnTop(Me.hWnd, True)
                'Call FormOnTop(Me.hWnd, False)
                RemoveFromTray
                'PopupMenu mnuSysTray
            'End If
        Case WM_LBUTTONUP
            'If temp > (Screen.Height / Screen.TwipsPerPixelY) - 30 Then
                Me.Visible = True
                Call FormOnTop(Me.hWnd, True)
                'Call FormOnTop(Me.hWnd, False)
                RemoveFromTray
            'End If
    End Select
End Sub
Private Sub Timer2_Timer()
Timer2.Enabled = False
Me.Visible = True
Call FormOnTop(Me.hWnd, True)
End Sub
Private Function btnFlat(Button As CommandButton)
SetWindowLong Command2.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
Command2.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
SetWindowLong cmdAdd.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
cmdAdd.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
SetWindowLong cmdRemove.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
cmdRemove.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function
Public Function TrimNull(item As String) As String
    Dim pos As Integer
        pos = InStr(item, Chr$(0))
        If pos Then item = Left$(item, pos - 1)
        TrimNull = item
End Function
