VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFlashMe 
      Left            =   2160
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   8415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   7575
   End
   Begin VB.Timer Timer2 
      Left            =   7080
      Top             =   6240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

   If SubClass(hWnd) Then
   
      If IsIDE Then
      
      
      End If
      
      Call SHNotify_Register(hWnd)
      
   Else: Text1.Text = "Well, it's supposed to work."
   End If
   
  'position the window in the bottom corner
   Me.Move Screen.Width - Width, Screen.Height - Height
  
End Sub


Private Sub Form_Resize()

  On Error GoTo Out
  Text1.Move 0, 0, ScaleWidth, ScaleHeight
  
Out:
End Sub


Private Sub Form_Unload(Cancel As Integer)
  
   Call SHNotify_Unregister
   Call UnSubClass(hWnd)
  
End Sub


Private Function IsIDE() As Boolean

   On Error GoTo Out
   Debug.Print 1 / 0
  
Out:
   IsIDE = Err
End Function


Public Sub NotificationReceipt(wParam As Long, lParam As Long)

   Dim sOut As String
   Dim shns As SHNOTIFYSTRUCT
   
   sOut = SHNotify_GetEventStr(lParam) & vbCrLf
   
  'Fill the SHNOTIFYSTRUCT from it's pointer.
   CopyMemory shns, ByVal wParam, Len(shns)
       
  'lParam is the ID of the notification event,
  'one of the SHCN_EventIDs.
   Select Case lParam
      
     '----------------------------------------------------
     'For the SHCNE_FREESPACE event, dwItem1 points
     'to what looks like a 10 byte struct. The first
     'two bytes are the size of the struct, and the
     'next two members equate to SHChangeNotify's
     'dwItem1 and dwItem2 params.
    
     'The dwItem1 member is a bitfield indicating which
     'drive(s) had it's (their) free space changed.
     'The bitfield is identical to the bitfield returned
     'from a GetLogicalDrives call, i.e, bit 0 = A:\, bit
     '1 = B:\, 2, = C:\, etc. Since VB does DWORD alignment
     'when CopyMemory'ing to a struct, we'll extract the
     'bitfield directly from it's memory location.
    
      Case SHCNE_FREESPACE
      
         Dim dwDriveBits As Long
         Dim wHighBit As Integer
         Dim wBit As Integer
         
         CopyMemory dwDriveBits, ByVal shns.dwItem1 + 2, 4
   
        'Get the zero based position of the highest
        'bit set in the bitmask (essentially determining
        'the value's highest complete power of 2).
        'Use floating point division (we want the exact
        'values from the Logs) and remove the fractional
        'value (the fraction indicates the value of
        'the last incomplete power of 2, which means the
        'bit isn't set).
        
         wHighBit = Int(Log(dwDriveBits) / Log(2))
         
         For wBit = 0 To wHighBit
           
          'If the bit is set...
           If (2 ^ wBit) And dwDriveBits Then
             
            '... get it's drive string
             sOut = sOut & Chr$(vbKeyA + wBit) & ":\" & vbCrLf
   
           End If
         Next
      
     '----------------------------------------------------
     'shns.dwItem1 also points to a 10 byte struct. The
     'struct's second member (after the struct's first
     'WORD size member) points to the system imagelist
     'index of the image that was updated.
      Case SHCNE_UPDATEIMAGE
      
         Dim iImage As Long
      
         CopyMemory iImage, ByVal shns.dwItem1 + 2, 4
         sOut = sOut & "Index of image in system imagelist: " & iImage & vbCrLf
    
     '----------------------------------------------------
     'Everything else except SHCNE_ATTRIBUTES is the
     'pidl(s) of the changed item(s). For SHCNE_ATTRIBUTES,
     'neither item is used. See the description of the
     'values for the wEventId parameter of the
     'SHChangeNotify API function for more info.
      Case Else
         Dim sDisplayname As String
         
         If shns.dwItem1 Then
         
            sDisplayname = GetDisplayNameFromPIDL(shns.dwItem1)
            
            If Len(sDisplayname) Then
             sOut = sOut & "first item displayname: " & sDisplayname & vbCrLf
             sOut = sOut & "first item path: " & GetPathFromPIDL(shns.dwItem1) & vbCrLf
            Else
             sOut = sOut & "first item is invalid" & vbCrLf
            End If
            
         End If
         
         If shns.dwItem2 Then
         
            sDisplayname = GetDisplayNameFromPIDL(shns.dwItem2)
           
            If Len(sDisplayname) Then
               sOut = sOut & "second item displayname: " & sDisplayname & vbCrLf
               sOut = sOut & "second item path: " & GetPathFromPIDL(shns.dwItem2) & vbCrLf
            Else
               sOut = sOut & "second item is invalid" & vbCrLf
            End If
         End If
  
  End Select
  
 'update the text window and flash
 'the window title
  Text1.Text = Text1.Text & sOut & vbCrLf
  Text1.SelStart = Len(Text1.Text)
  tmrFlashMe = True

End Sub


Private Sub tmrFlashMe_Timer()

  'initial settings: Interval = 1, Enabled = False
  
   Static nCount As Integer
   
   If nCount = 0 Then tmrFlashMe.Interval = 200
   
   nCount = nCount + 1
   Call FlashWindow(hWnd, True)
   
  'Reset everything after 3 flash cycles
   If nCount = 6 Then
      nCount = 0
      tmrFlashMe.Interval = 1
      tmrFlashMe = False
   End If

End Sub


