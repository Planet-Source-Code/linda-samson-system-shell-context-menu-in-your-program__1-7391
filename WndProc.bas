Attribute VB_Name = "mWndProc"
Option Explicit

' Brad Martinez http://www.mvps.org/ccrp

' Code was written in and formatted for 8pt MS San Serif

Private Const WM_DRAWITEM = &H2B
Private Const WM_MEASUREITEM = &H2C
Private Const WM_INITMENUPOPUP = &H117

Public ICtxMenu2 As IContextMenu2

' =========================

Private Const WM_DESTROY = &H2

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Public Enum GWL_nIndex
  GWL_WNDPROC = (-4)
'  GWL_HWNDPARENT = (-8)
  GWL_ID = (-12)
  GWL_STYLE = (-16)
  GWL_EXSTYLE = (-20)
'  GWL_USERDATA = (-21)
End Enum

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex, ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const OLDWNDPROC = "OldWndProc"
Private Const OBJECTPTR = "ObjectPtr"

' Set to non-zero to prevent the IDE from freezing when subclassed and
' stepping through code. Requires the "Debug Object for AddressOf
' Subclassing" (Dbgwproc.dll), last found at:
' http://msdn.microsoft.com/vbasic/downloads/download.asp?ID=024
#Const DEBUGWINDOWPROC = 0

#If DEBUGWINDOWPROC Then
' maintains a WindowProcHook object reference for each subclassed window.
' The subclassed window's handle is used as the collection item's key string.
Private m_colWPHooks As New Collection
#End If
'

Public Function SubClass(hWnd As Long, _
                                         lpfnNew As Long, _
                                         Optional objNotify As Object = Nothing) As Boolean
  Dim lpfnOld As Long
  Dim fSuccess As Boolean
  On Error GoTo Out
  
  If GetProp(hWnd, OLDWNDPROC) Then
    SubClass = True
    Exit Function
  End If
  
#If (DEBUGWINDOWPROC = 0) Then
    lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, lpfnNew)

#Else
    Dim objWPHook As WindowProcHook
    
    Set objWPHook = CreateWindowProcHook
    m_colWPHooks.Add objWPHook, CStr(hWnd)
    
    With objWPHook
      Call .SetMainProc(lpfnNew)
      lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, .ProcAddress)
      Call .SetDebugProc(lpfnOld)
    End With

#End If
  
  If lpfnOld Then
    fSuccess = SetProp(hWnd, OLDWNDPROC, lpfnOld)
    If (objNotify Is Nothing) = False Then
      fSuccess = fSuccess And SetProp(hWnd, OBJECTPTR, ObjPtr(objNotify))
    End If
  End If
  
Out:
  If fSuccess Then
    SubClass = True
  
  Else
    If lpfnOld Then Call SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld)
    MsgBox "Error subclassing window &H" & Hex(hWnd) & vbCrLf & vbCrLf & _
                  "Err# " & Err.Number & ": " & Err.Description, vbExclamation
  End If
  
End Function

Public Function UnSubClass(hWnd As Long) As Boolean
  Dim lpfnOld As Long
  
  lpfnOld = GetProp(hWnd, OLDWNDPROC)
  If lpfnOld Then
    
    If SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld) Then
      Call RemoveProp(hWnd, OLDWNDPROC)
      Call RemoveProp(hWnd, OBJECTPTR)

#If DEBUGWINDOWPROC Then
      ' remove the WindowProcHook reference from the collection
      On Error Resume Next
      m_colWPHooks.Remove CStr(hWnd)
#End If
      
      UnSubClass = True
    
    End If   ' SetWindowLong
  End If   ' lpfnOld

End Function

' Returns the specified object reference stored in the subclassed
' window's OBJECTPTR window property.
' The object reference is valid for only as long as the calling proc holds it.

Public Function GetObj(hWnd As Long) As Object
  Dim Obj As Object
  Dim pObj As Long
  pObj = GetProp(hWnd, OBJECTPTR)
  If pObj Then
    MoveMemory Obj, pObj, 4
    Set GetObj = Obj
    MoveMemory Obj, 0&, 4
  End If
End Function

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
  Select Case uMsg
    
    ' ======================================================
    ' Handle owner-draw context menu messages (for the Send To submenu)
    
    Case WM_INITMENUPOPUP, WM_DRAWITEM, WM_MEASUREITEM
      If (ICtxMenu2 Is Nothing) = False Then
        Call ICtxMenu2.HandleMenuMsg(uMsg, wParam, lParam)
      End If
    
    ' ======================================================
    ' Unsubclass the window.
    
    Case WM_DESTROY
      ' OLDWNDPROC will be gone after UnSubClass is called!
      Call CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
      Call UnSubClass(hWnd)
      Exit Function
      
  End Select
  
  WndProc = CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
  
End Function
