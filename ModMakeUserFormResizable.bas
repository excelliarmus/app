Attribute VB_Name = "ModMakeUserFormResizable"
'Written: February 14, 2011
'Author:  Leith Ross
'
'NOTE:  This code should be executed within the UserForm_Activate() event.

Private Declare PtrSafe Function GetForegroundWindow Lib "User32.dll" () As Long


Private Declare PtrSafe Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
               
Private Declare PtrSafe Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WS_THICKFRAME As Long = &H40000
Private Const GWL_STYLE As Long = -16

Public Sub MakeFormResizable()

  Dim lStyle As Long
  Dim hWnd As Long
  Dim RetVal
  
    hWnd = GetForegroundWindow
  
    'Get the basic window style
     lStyle = GetWindowLong(hWnd, GWL_STYLE) Or WS_THICKFRAME

    'Set the basic window styles
     RetVal = SetWindowLong(hWnd, GWL_STYLE, lStyle)

End Sub

