Attribute VB_Name = "ShellWait"
Option Explicit

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds _
 As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal _
 lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal _
 lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
 ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, _
 ByVal lpCurrentDirectory As Long, lpStartupInfo As _
 STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Public TaskID As Long
Public Function ExecuteTask(cmdline As String) As Long
Dim proc As PROCESS_INFORMATION
Dim startup As STARTUPINFO
Dim ret As Long
startup.cb = Len(startup)
ret = CreateProcessA(0&, cmdline, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, startup, proc)
ExecuteTask = proc.hProcess
End Function
Public Function TaskRunning() As Boolean
Dim ret As Long
If TaskID = 0 Then Exit Function
ret = WaitForSingleObject(TaskID, 0)
TaskRunning = (ret <> 0)
If Not TaskRunning Then
ret = CloseHandle(TaskID)
TaskID = 0
End If
End Function

