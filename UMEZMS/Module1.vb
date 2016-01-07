Imports System.Runtime.InteropServices
Module Module1
    Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Public Declare Function OpenProcessAPI Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Int32, ByVal bInheritHandle As Int32, ByVal dwProcessId As Int32) As Int32
    Public Declare Function CloseHandleAPI Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
    Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As IntPtr, <Out()> ByRef lpdwProcessId As IntPtr) As Integer
    Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
    Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Integer, ByVal lppe As PROCESSENTRY32) As Integer
    Public Declare Function Process32Next Lib "kernel32" (ByVal hSapshot As Integer, ByVal lppe As PROCESSENTRY32) As Integer
    Public Const PROCESS_ALL_ACCESS = &H1F0FFF
    Public hWnd, Handlex, pid, hprocess As Integer
    Public Inited As Boolean

    Public Structure PROCESSENTRY32
        Dim dwSize As Integer
        Dim cntUseage As Integer
        Dim th32ProcessID As Integer
        Dim th32DefaultHeapID As Integer
        Dim th32ModuleID As Integer
        Dim cntThreads As Integer
        Dim th32ParentProcessID As Integer
        Dim pcPriClassBase As Integer
        Dim swFlags As Integer
        Dim szExeFile As String
    End Structure

    Public Function OpenProcess(Optional ByVal lpPID As Long = -1) As Long
        If lpPID = 0 And pid = 0 Then Exit Function
        If lpPID > 0 And pid = 0 Then pid = lpPID
        Handlex = OpenProcessAPI(PROCESS_ALL_ACCESS, False, pid)
        OpenProcess = Handlex
        If Handlex > 0 Then Inited = True
    End Function
    Public Function OpenProcessByProcessName(ByVal lpName As String) As Long
        Dim PE32 As PROCESSENTRY32
        Dim hSnapshot As Long
        pid = 0
        hSnapshot = CreateToolhelp32Snapshot(2, 0&) 'TH32CS_SNAPPROCESS = 2
        PE32.dwSize = Len(PE32)
        Process32First(hSnapshot, PE32)
        While pid = 0 And CBool(Process32Next(hSnapshot, PE32))
            If Right$(LCase$(Left$(PE32.szExeFile, InStr(1, PE32.szExeFile, Chr(0)) - 1)), Len(lpName)) = LCase$(lpName) Then
                pid = PE32.th32ProcessID
            End If
        End While
        CloseHandleAPI(hSnapshot)
        OpenProcessByProcessName = OpenProcess()
    End Function
    Public Function OpenProcessByWindow(ByVal lpWindowName As String, Optional ByVal lpClassName As String = vbNullString) As Long
        hWnd = FindWindow(lpClassName, lpWindowName)
        GetWindowThreadProcessId(hWnd, pid)
        OpenProcessByWindow = OpenProcess(pid)
        hprocess = OpenProcessAPI(PROCESS_ALL_ACCESS, False, pid)
    End Function


End Module
