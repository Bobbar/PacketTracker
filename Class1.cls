VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const WS_THICKFRAME = &H40000
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

Public Property Let Sizable(ByVal Value As Boolean)
   ' Set WS_THICKFRAME On or Off as requested.
   Call fFlipBit(WS_THICKFRAME, Value)
End Property

Public Property Get Sizable() As Boolean
   ' Return value of WS_THICKFRAME bit.
   Sizable = CBool(fStyle And WS_THICKFRAME)
End Property
