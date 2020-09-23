Attribute VB_Name = "Killapps"
'**************************************
'Windows API/Global Declarations for :Ki
'     llApp
'**************************************
Const MAX_PATH& = 260
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
    End Type

'**************************************
' Name: KillApp
' Description:Kill any application or pr
'     ocess running if you know the .exe name.
'     (Only Windows 95/98)
' By: Fernando Vicente
'
'
' Inputs:myName: is the name of the app
'     that wou want to kill (ex. "app.exe")
'
' Returns:True if the application was ki
'     lled correctly
'
'Assumes:None
'
'Side Effects:None
'This code is copyrighted and has limite
'     d warranties.
'Please see http://www.Planet-Source-Code.com/xq/ASP/txtCodeId.2160/lngWId.1/qx/vb/scripts/ShowCode.htm
'
'
'for details.
'**************************************
Public Function Killapp(myName As String) As Boolean
    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim I As Integer
    'On Local Error GoTo Finish
    appCount = 0
    Const TH32CS_SNAPPROCESS As Long = 2&
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    Do While rProcessFound
        I = InStr(1, uProcess.szExeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szExeFile, I - 1))
        If Right$(szExename, Len(myName)) = LCase$(myName) Then
            Killapp = True
            appCount = appCount + 1
            myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
            AppKill = TerminateProcess(myProcess, 0&) 'exitCode
            Call CloseHandle(myProcess)
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    Call CloseHandle(hSnapshot)
Finish:
End Function
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
