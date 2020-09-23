Attribute VB_Name = "Module2"
Option Explicit
Private Const WM_NCDESTROY As Long = &H82
Private Const GWL_WNDPROC As Long = (-4)
Private Const OLDWNDPROC As String = "OldWndProc"

Private Declare Function GetProp Lib "user32" _
    Alias "GetPropA" _
   (ByVal hWnd As Long, _
    ByVal lpString As String) As Long
    
Private Declare Function SetProp Lib "user32" _
    Alias "SetPropA" _
    (ByVal hWnd As Long, _
    ByVal lpString As String, _
    ByVal hData As Long) As Long
    
Private Declare Function RemoveProp Lib "user32" _
    Alias "RemovePropA" _
    (ByVal hWnd As Long, _
    ByVal lpString As String) As Long

Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" _
    (ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long


Public Function SubClass(hWnd As Long) As Boolean

   Dim lpfnOld As Long
   Dim fSuccess As Boolean
  
   If (GetProp(hWnd, OLDWNDPROC) = 0) Then
   
      lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
    
      If lpfnOld Then
         fSuccess = SetProp(hWnd, OLDWNDPROC, lpfnOld)
      End If
      
   End If
  
   If fSuccess Then
      SubClass = True
   Else
      If lpfnOld Then Call UnSubClass(hWnd)
      MsgBox "Unable to successfully subclass &H" & Hex(hWnd), vbCritical
   End If
  
End Function


Public Function UnSubClass(hWnd As Long) As Boolean
  
   Dim lpfnOld As Long
  
   lpfnOld = GetProp(hWnd, OLDWNDPROC)
   
   If lpfnOld Then
      If RemoveProp(hWnd, OLDWNDPROC) Then
         UnSubClass = SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld)
      End If
   End If

End Function


Public Function WindowProc(ByVal hWnd As Long, _
                        ByVal uMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long
  
   Select Case uMsg
   
      Case WM_SHNOTIFY
         Call Form1.NotificationReceipt(wParam, lParam)
      
      Case WM_NCDESTROY
         Call UnSubClass(hWnd)
         MsgBox "Unsubclassed &H" & Hex(hWnd), vbCritical, "WindowProc Error"
   
   End Select
   
   WindowProc = CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
  
End Function

 
 
 
 
 


