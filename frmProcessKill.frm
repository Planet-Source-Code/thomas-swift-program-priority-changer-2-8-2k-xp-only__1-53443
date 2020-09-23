VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcessKill 
   Caption         =   "Terminate Windows Processes"
   ClientHeight    =   3510
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComctlLib.ListView lvwProcess 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4683
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Timer tmrProcess 
      Interval        =   50000
      Left            =   240
      Top             =   3000
   End
End
Attribute VB_Name = "frmProcessKill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #####################################################################################
' #####################################################################################
' Filename:         frmProcessKill.frm
'
' Description:      Form to display the filtered processes and to allow the user to
'                       terminate selected or all of the processes.
'
' #####################################################################################
' #####################################################################################
Option Explicit

Dim m_strFilter As String
Private Sub Form_Load()
  'Set the default filter
    m_strFilter = "C:\"


End Sub
Private Sub tmrProcess_Timer()
    SetUpList
End Sub


'***************************************************************************************
'   Private Routines
'***************************************************************************************
'   Sub Name:  LoadNTProcess
'
'   Description:    Loads the NTProcesses and populates the listview.
'
'   Inputs:         NONE
'   Returns:        NONE
'
'***************************************************************************************
Public Sub LoadNTProcess()
  Dim cb As Long
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
         
    'Get the array containing the process id's for each process object
    cb = 8
    cbNeeded = 96
    Do While cb <= cbNeeded
        cb = cb * 2
        ReDim ProcessIDs(cb / 4) As Long
        lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
    Loop
         
    NumElements = cbNeeded / 4
    For I = 1 To NumElements
      'Get a handle to the Process
         hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(I))
      'Got a Process handle
         If hProcess <> 0 Then
           'Get an array of the module handles for the specified process
             lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
             
           'If the Module Array is retrieved, Get the ModuleFileName
            If lRet <> 0 Then
                ModuleName = Space(MAX_PATH)
                nSize = 500
                lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
                                     
                                     
                If CBool(InStr(1, (Left(ModuleName, lRet)), m_strFilter, vbTextCompare)) Then
                    AddListItem Left(ModuleName, lRet), ProcessIDs(I)
                End If
            End If
        End If
               
        'Close the handle to the process
        lRet = CloseHandle(hProcess)
    Next
End Sub


'***************************************************************************************
'   Sub Name:  AddListItem
'
'   Description:    Given the name of the process and the processid, add it to the list.
'
'   Inputs:         p_ProcessId     -->  The processid of the process.
'                   p_strProcess    --> The name of the process.
'   Returns:        NONE
'
'***************************************************************************************
Public Sub AddListItem(p_strProcess As String, p_ProcessID As Long)
  Dim Item As Object
        
    Set Item = lvwProcess.ListItems.Add(, , p_strProcess)
    With Item
        .SubItems(1) = p_ProcessID
    End With
End Sub


'***************************************************************************************
'   Sub Name:  SetUpList
'
'   Description:    Initialize the listview.  Called several times to update the list of
'                     of processes
'
'   Inputs:         NONE
'   Returns:        NONE
'
'***************************************************************************************
Public Sub SetUpList()
    With lvwProcess
        .ListItems.Clear
        With .ColumnHeaders
            .Clear
            .Add , , "Process", (lvwProcess.Width * (0.8))
            .Add , , "PID", (lvwProcess.Width * (0.175))
        End With
    
        .View = lvwReport
        .HideColumnHeaders = False
    End With
        

End Sub


'***************************************************************************************
'   Sub Name:       KillAllProcess
'
'   Description:    Steps through each entry in the listview and makes a call to
'                     KillProcessByID passing in the PID and terminating that process.
'
'   Inputs:         NONE
'   Returns:        NONE
'
'***************************************************************************************
Public Sub KillAllProcess()
  Dim I As Integer
  Dim Counter As Integer
  Dim lngSuccess As Long
  Dim dblPID As Double
    Counter = lvwProcess.ListItems.Count
    For I = 1 To Counter
        With lvwProcess.ListItems.Item(I)
            KillProcessById (.SubItems(1))
        End With
    Next I
End Sub


'***************************************************************************************
'   Sub Name:       KillSelectedProcess
'
'   Description:    Steps through each entry in the listview, checks to see if it is
'                     selected and makes a call to KillProcessByID passing in the PID
'                     and terminating that process.
'
'   Inputs:         NONE
'   Returns:        NONE
'
'***************************************************************************************
Public Sub KillSelectedProcess()
  Dim I As Integer
  Dim Counter As Integer
  Dim lngSuccess As Long
  Dim dblPID As Double
    
    Counter = lvwProcess.ListItems.Count
    For I = 1 To Counter
        With lvwProcess.ListItems.Item(I)
            If .Selected = True Then
                KillProcessById (.SubItems(1))
            End If
        End With
    Next I
End Sub
Public Function KillMyApp()
  Dim I As Integer
  Dim Counter As Integer
  Dim lngSuccess As Long
  Dim dblPID As Double
  Dim file As String
    
    Counter = lvwProcess.ListItems.Count
    For I = 1 To Counter
        With lvwProcess.ListItems.Item(I)
        file = lvwProcess.ListItems.Item(I)
        GetFileName file
        If LCase(file) = "iexplore.exe" Then
         KillProcessById (.SubItems(1))
        End If
        End With
        'With lvwProcess.ListItems.Item(i)
            'If .Selected = True Then
               ' KillProcessById (.SubItems(1))
           ' End If
       ' End With
    Next I
 SetUpList
End Function
Public Function GetFileName(file As String) As String
        Dim m
    Dim GetChr0 As String
    Dim GetChr1 As String
    For m = 1 To Len(file)
        GetChr0 = Right(file, m)
        GetChr1 = Left(GetChr0, 1)
        If GetChr1 = "\" Or GetChr1 = "/" Then
        file = Right(GetChr0, m - 1): Exit Function
    End If
    Next m
End Function
