Attribute VB_Name = "FlatButtons"
Option Explicit

Public Declare Function IsThemeActive Lib "uxtheme.dll" () As Boolean
Public Declare Function IsAppThemed Lib "uxtheme.dll" () As Boolean

Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)

Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Sub BTFlat(BT As CommandButton)
    On Error Resume Next
    If GetWindowLong&(BT.hwnd, -16) And &H8000& Then Exit Sub
    SetWindowLong BT.hwnd, -16, GetWindowLong&(BT.hwnd, -16) Or &H8000&
    BT.Refresh
End Sub

Public Sub BTFlatAll(ByRef OnForm As Form)
    'doesnt always work
    On Error Resume Next
    Dim a As CommandButton
    For Each a In OnForm
        BTFlat a
    Next
End Sub

Public Sub FlatAll(ByRef ParentHwnd As Long)
    On Error Resume Next
    EnumChildWindows ParentHwnd, AddressOf ButtonProc, 1
End Sub

Public Function ButtonProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim R As RECT
    If GetWindowLong&(hwnd, -16) And &H8000& Then Exit Function
    GetWindowRect hwnd, R
    SetWindowLong hwnd, -16, GetWindowLong&(hwnd, -16) Or &H8000&
    RedrawWindow hwnd, R, 1, 1
    ButtonProc = 1
End Function

