Attribute VB_Name = "Module1"
Option Explicit


 Public strUserTo, strSelectUserTo, strUserFrom, strCurUser As String
Public strTicketAction, strTicketStatus As String
Public bolHasTicket As Boolean
Public strServerAddress, strUsername, strPassword, strSearchUser, strPlant As String
Public strUserIndex() As String
Public bolOpenForm As Boolean
Public intFormHMax, intFormHMin As Integer
Public strReportType As String

Public dtStartDate As Date
Public dtEndDate As Date
Public sAddlMsg As String
Public bolCanPrint As Boolean
Public strTicketComment As String
Public strLatestComment As String
Public strSortMode As String
Public bolButtonOn As Boolean
Public bolPrinting As Boolean
Public FlexINLastSel(2) As Integer
Public FlexOUTLastSel(2) As Integer
Public FlexHISTLastSel(2) As Integer
Public FlexHISTSelDate As String


Public FlexHistLastTopRow As Integer
Public Const intRowH As Integer = 310
Public strLocalUser As String



Public intMovement As Integer
Public bolOptionClicked As Boolean

Public HistoryIcons() As StdPicture

Public WhichGrid As MSHFlexGrid

Public TicketHours(99) As Single
Public TicketAction(99) As String
Public TicketActionText(99) As String
Public TicketDate(99) As String


Public TotalTime As Single
Public LStep As Single
Public Entry As Integer





Public Declare Function GetTickCount Lib "KERNEL32" () As Long
Public StartTime, EndTime As Long




Private Declare Function CallWindowProc2 Lib "user32.dll" Alias "CallWindowProcA" ( _
ByVal lpPrevWndFunc As Long, _
ByVal hWnd As Long, _
ByVal Msg As Long, _
ByVal wParam As Long, _
ByVal lParam As Long) As Long

Private Declare Function SetWindowLong2 Lib "user32.dll" Alias "SetWindowLongA" ( _
ByVal hWnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long


Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A

Dim LocalHwnd As Long
Dim LocalPrevWndProc As Long
Dim MyForm As Form


Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long _
) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
) As Long
'Private Const GWL_WNDPROC = -4
Private lpPrevWndProc As Long
Private lpWndProcTmp As Long
Private Const WM_POWERBROADCAST As Long = &H218
Private Const PBT_APMRESUMEAUTOMATIC As Long = &H12
Private Const PBT_APMSUSPEND As Long = &H4

Public Sub Hook(ByVal gHW As Long, HKflg As Boolean)
Static IsHooked As Boolean
If HKflg Xor IsHooked Then
    If HKflg Then
        lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
    Else
        SetWindowLong gHW, GWL_WNDPROC, lpPrevWndProc
    End If
    IsHooked = HKflg
End If
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_POWERBROADCAST Then
        If wParam = PBT_APMRESUMEAUTOMATIC Then
        
            
            Form1.tmrRefresher.Enabled = True
            Form1.tmrDateTime.Enabled = True
            
            
        ElseIf wParam = PBT_APMSUSPEND Then
            
            Form1.tmrRefresher.Enabled = False
            Form1.tmrDateTime.Enabled = False
            
        End If
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function


Private Function WindowProc2(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim MouseKeys As Long
Dim Rotation As Long
Dim Xpos As Long
Dim Ypos As Long

If Lmsg = WM_MOUSEWHEEL Then
MouseKeys = wParam And 65535
Rotation = wParam / 65536
Xpos = lParam And 65535
Ypos = lParam / 65536
MyForm.MouseWheel MouseKeys, Rotation, Xpos, Ypos
End If
WindowProc2 = CallWindowProc2(LocalPrevWndProc, Lwnd, Lmsg, wParam, lParam)
End Function

Public Sub WheelHook(PassedForm As Form)

On Error Resume Next

Set MyForm = PassedForm
LocalHwnd = PassedForm.hWnd
LocalPrevWndProc = SetWindowLong2(LocalHwnd, GWL_WNDPROC, AddressOf WindowProc2)
End Sub

Public Sub WheelUnHook()

    Dim WorkFlag As Long

    On Error Resume Next

    WorkFlag = SetWindowLong2(LocalHwnd, GWL_WNDPROC, LocalPrevWndProc)
    Set MyForm = Nothing
End Sub



