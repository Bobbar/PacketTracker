Attribute VB_Name = "Module4"
Option Explicit

Public Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Public Declare Function GetLogicalDrives Lib "kernel32" () As Long

Const ICC_USEREX_CLASSES = &H200
Const SEM_NOGPFAULTERRORBOX As Long = 2

Public Function XPVB() As Boolean
    On Error Resume Next
    Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx iccex
    XPVB = (Err.Number = 0)
End Function

Public Function FindPath(Parent As String, Optional Child As String, Optional Divider As String = "\") As String
    On Error Resume Next
    If Right$(Parent, 1) = Divider Then Parent = Left$(Parent, Len(Parent) - 1)
    If Left$(Child, 1) = Divider Then Child = Mid$(Child, 2)
    FindPath = Parent & Divider & Child
    If Left$(FindPath, 1) = Divider Then FindPath = Right$(FindPath, Len(FindPath) - 1)
End Function



Public Function IsIDE() As Boolean
    IsIDE = (App.LogMode = 0)
End Function

'Public Sub UnloadApp()
'   If Not IsIDE() Then SetErrorMode SEM_NOGPFAULTERRORBOX
'End Sub
Public Sub UnloadApp()
    On Error Resume Next
    If Not IsIDE() Then
        SetErrorMode SEM_NOGPFAULTERRORBOX
        LoadLibrary "comctl32.dll"
    End If
End Sub

Public Function MyVer() As String
    On Error Resume Next
    Dim AppRevision As Integer, AppMinor As Integer, AppMajor As Integer
    AppRevision = App.Revision
    AppMinor = App.Minor
    AppMajor = App.Major
    If AppRevision >= 10 Then
        AppMinor = AppMinor + 1
        AppRevision = AppRevision - 10
    End If
    If AppRevision > 0 Then AppMinor = AppMinor + 1
    If AppMinor >= 10 Then
        AppMajor = AppMajor + 1
        AppMinor = AppMinor - 10
    End If
    MyVer = "V." & AppMajor & "." & AppMinor & IIf(AppRevision > 0, " Beta", "")
End Function

Public Function IsDrivePresent(Optional driveLetter As String = "C") As Boolean
    'partial credits http://vb.ncis.com.tw/
    On Error Resume Next
    Dim driveNUM As Long
    Dim LngDrive As Long
    
    driveNUM = Asc(driveLetter) - 65
    
    LngDrive = GetLogicalDrives
    If LngDrive And 2 ^ driveNUM Then IsDrivePresent = True
End Function

