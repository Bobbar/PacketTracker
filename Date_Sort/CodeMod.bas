Attribute VB_Name = "CodeMOd"
Public Sub B2S_BSort_Date(DList() As String, LBox As ListBox)
' /* Needs To Be In The Format: dd/mm/yyyy */
' /* You're Welcome Yomet! */
Dim I As Integer
Dim X As Integer
Dim IDay As String
Dim XDay As String
Dim IMonth As String
Dim XMonth As String
Dim IYear As String
Dim XYear As String

' /* Sort By Order Of Days*/
For I = 0 To UBound(DList)
    For X = 0 To UBound(DList)
        IDay = Left(DList(I), 2)
        XDay = Left(DList(X), 2)
        If XDay < IDay Then
            Call Swap(DList(), I, X)
        End If
    Next X
Next I

' /* Sort By Order Of Months */
For I = 0 To UBound(DList)
    For X = 0 To UBound(DList)
        IMonth = Mid(DList(I), 4, 2)
        XMonth = Mid(DList(X), 4, 2)
        If XMonth < IMonth Then
            Call Swap(DList(), I, X)
        End If
    Next X
Next I

' /* Sort By Order Of Years */
For I = 0 To UBound(DList)
    For X = 0 To UBound(DList)
        IYear = Right(DList(I), 4)
        XYear = Right(DList(X), 4)
        If XYear < IYear Then
            Call Swap(DList(), I, X)
        End If
    Next X
Next I




For Each tmp In DList
    LBox.AddItem tmp
Next tmp

End Sub

Public Sub S2B_BSort_Date(DList() As String, LBox As ListBox)
' /* Needs To Be In The Format: dd/mm/yyyy */
' /* You're Welcome Yomet! */
Dim I As Integer
Dim X As Integer
Dim IDay As String
Dim XDay As String
Dim IMonth As String
Dim XMonth As String
Dim IYear As String
Dim XYear As String

' /* Sort By Order Of Days*/
For I = 0 To UBound(DList)
    For X = 0 To UBound(DList)
        IDay = Left(DList(I), 2)
        XDay = Left(DList(X), 2)
        If XDay > IDay Then
            Call Swap(DList(), I, X)
        End If
    Next X
Next I

' /* Sort By Order Of Months */
For I = 0 To UBound(DList)
    For X = 0 To UBound(DList)
        IMonth = Mid(DList(I), 4, 2)
        XMonth = Mid(DList(X), 4, 2)
        If XMonth > IMonth Then
            Call Swap(DList(), I, X)
        End If
    Next X
Next I

' /* Sort By Order Of Years */
For I = 0 To UBound(DList)
    For X = 0 To UBound(DList)
        IYear = Right(DList(I), 4)
        XYear = Right(DList(X), 4)
        If XYear > IYear Then
            Call Swap(DList(), I, X)
        End If
    Next X
Next I




For Each tmp In DList
    LBox.AddItem tmp
Next tmp

End Sub

Public Sub Swap(ByRef AName() As String, Indice1 As Integer, Indice2 As Integer)
Dim tmp As String

tmp = AName(Indice1)
AName(Indice1) = AName(Indice2)
AName(Indice2) = tmp

End Sub
