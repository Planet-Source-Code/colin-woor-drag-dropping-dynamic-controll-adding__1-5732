Attribute VB_Name = "modGridSnap"
Type POINTAPI
    x As Long
    y As Long
End Type


Declare Sub SetCursorPos Lib "User32" (ByVal x%, ByVal y%)
Declare Sub GetCursorPos Lib "User32" (lpoint As POINTAPI)

