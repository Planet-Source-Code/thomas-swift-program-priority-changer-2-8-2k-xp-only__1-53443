Attribute VB_Name = "KillProgram"
Option Explicit

'***************************************************************************************
'   API Declares
'***************************************************************************************
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, _
                                                        ByVal bInheritHandle As Long, _
                                                        ByVal dwProcId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, _
                                                        ByVal uExitCode As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

'***************************************************************************************
'   Types Used to Retrieve Information From Windows
'***************************************************************************************
Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long           ' This process
    th32DefaultHeapID As Long
    th32ModuleID As Long            ' Associated exe
    cntThreads As Long
    th32ParentProcessID As Long     ' This process's parent process
    pcPriClassBase As Long          ' Base priority of process threads
    dwFlags As Long
    szExeFile As String * 260       ' MAX_PATH
End Type
Public Function Killapp(myName As String) As Boolean
    'Private Const PROCESS_ALL_ACCESS = 0
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
            myProcess = OpenProcess(1&, -1&, uProcess.th32ProcessID)
            'myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
            AppKill = TerminateProcess(myProcess, 0&)
            'AppKill = TerminateProcess(myProcess, exitCode)
            Call CloseHandle(myProcess)
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    Call CloseHandle(hSnapshot)
Finish:
End Function
