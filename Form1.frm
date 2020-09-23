VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shell context menu demo"
   ClientHeight    =   4845
   ClientLeft      =   2310
   ClientTop       =   2475
   ClientWidth     =   5640
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4845
   ScaleWidth      =   5640
   Begin VB.CheckBox chkPrompt 
      Caption         =   "Prompt before executing selected context menu command."
      Height          =   285
      Left            =   300
      TabIndex        =   1
      Top             =   900
      Value           =   1  'Checked
      Width           =   4995
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Hidden          =   -1  'True
      Left            =   2940
      MultiSelect     =   2  'Extended
      System          =   -1  'True
      TabIndex        =   4
      Top             =   1350
      Width           =   2400
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   300
      TabIndex        =   3
      Top             =   1770
      Width           =   2400
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   300
      TabIndex        =   2
      Top             =   1350
      Width           =   2400
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   765
      Left            =   300
      TabIndex        =   0
      Top             =   120
      Width           =   4995
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Removed One Reference,
'changed with one that comes with VB.
'linda.69@mailcity.com

Option Explicit
'
' Brought to you by Brad Martinez
'   http://members.aol.com/btmtz/vb
'   http://www.mvps.org/ccrp
'
' Code was written in and formatted for 8pt MS San Serif
'
' Demonstrates how to show the shell context menu for any directory or
' group of files in the file system. Also demonstrates how to navigate pidls
' (pointers to item ID lists)
'
' Version 1.30
'
' Note that "IShellFolder Extended Type Library v1.2" (ISHF_Ex.tlb)
' included with this project, must be present and correctly registered
' on your system, and referenced by this project, to allow use of the
' IShellFolder, IContextMenu and IMalloc interfaces. ** Be aware that
' this type library is a newer version than one included with previous
' version of this demo project.**
'
' v1.10 update:
'   - Corrected FileListBox selection bug
'   - Shows how to retrieve the selected context menu command's string.
'   - Added the option to prompt before executing the selected command.
'
' v1.20 update:
'   - Corrected the "extended" FileListBox selection bug, again.
'   - Shows how insert and execute a user-defined menu command in the
'     context menu.
'   - Did a little more documenting...
'
' v1.30 update:
'   - Added IContextMenu2 support, allowing the "Send To" and "Open With"
'     submenus to be filled with their respective items.
'   - Added subclassing module to catch ownerdraw menu messages for above.
'   - Now inserts the focused FileListBox file as the first element of the array of
'     relative pidls passed to ShowShellContextMenu. This allows the context
'     menu to contain the commands for this file when multiple files are selected.
'   - Both listboxes now refresh after each context menu command is carried out.
'

Private Sub Form_Load()
'  File1.MultiSelect = 2
  Move (Screen.Width - Width) * 0.5, (Screen.Height - Height) * 0.5
End Sub

Private Sub Drive1_Change()
  On Error GoTo Out   ' covers invalid path selection
  Dir1.Path = Drive1.Drive
Out:
End Sub

Private Sub Dir1_Change()
  On Error GoTo Out   ' covers invalid path selection
  File1 = Dir1.Path
Out:
End Sub

Private Sub Dir1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If (Button = vbRightButton) Then
    Call ShellContextMenu(Dir1, x, y, Shift)
  End If
End Sub

' Selects all FileListBox items on a Ctrl+A keypress.

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyA) And (Shift = vbCtrlMask) Then
    Call SendMessage(File1.hWnd, LB_SETSEL, CTrue, ByVal -1)
  End If
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If (Button = vbRightButton) Then
    Call ShellContextMenu(File1, x, y, Shift)
  End If
End Sub

Private Sub ShellContextMenu(objLB As Control, _
                                                 x As Single, _
                                                 y As Single, _
                                                 Shift As Integer)
  
  Dim pt As POINTAPI               ' screen location of the cursor
  Dim iItem As Integer                ' listbox index of the selected item (item under the cursor)
  Dim cItems As Integer             ' count of selected items
  Dim i As Integer                       ' counter
  Dim asPaths() As String           ' array of selected items' paths (zero based)
  Dim apidlFQs() As Long           ' array of selected items' fully qualified pidls (zero based)
  Dim isfParent As IShellFolder   ' selected items' parent shell folder
  Dim apidlRels() As Long           ' array of selected items' relative pidls (zero based)
  
  ' ==================================================
  ' Get the listbox item under the cursor
  
  ' Convert the listbox's client twip coords to screen pixel coords.
  pt.x = x \ Screen.TwipsPerPixelX
  pt.y = y \ Screen.TwipsPerPixelY
  Call ClientToScreen(objLB.hWnd, pt)

  ' Get the zero-based index of the item under the cursor.
  ' If none exists, bail...
  iItem = LBItemFromPt(objLB.hWnd, pt.x, pt.y, False)
  If (iItem = LB_ERR) Then Exit Sub
    
  ' ==================================================
  ' Set listbox focus and selection
  
  objLB.SetFocus
  
  ' If neither the Control and/or Shift key are pressed...
  If (Shift And (vbCtrlMask Or vbShiftMask)) = False Then
    
    ' If Dir1 has the focus...
    If (TypeOf objLB Is DirListBox) Then
      ' Select the item under the cursor. The DirListBox
      ' doesn't have a Selected property, so we'll get forceful...
      Call SendMessage(Dir1.hWnd, LB_SETCURSEL, iItem, 0)
    
    Else
      ' File1 has the focus, duplicate Explorer listview selection functionality.
      
      ' If the right clicked item isn't selected....
      If (File1.Selected(iItem) = False) Then
        ' Deselect all of the items and select the right clicked item.
        Call SendMessage(File1.hWnd, LB_SETSEL, CFalse, ByVal -1)
        File1.Selected(iItem) = True
      Else
      ' The right clciked item is selected, give it the selection rectangle
      ' (or caret, does not deselect any other currently selected items).
      ' File1.Selected doesn't set the caret if the item is already selected.
        Call SendMessage(File1.hWnd, LB_SETCARETINDEX, iItem, ByVal 0&)
      End If
    
    End If   '  (TypeOf objLB Is DirListBox)
  End If   ' (Shift And (vbCtrlMask Or vbShiftMask)) = False
  
  ' ==================================================
  ' Load the path(s) of the selected listbox item(s) into the array.
  
  If (TypeOf objLB Is DirListBox) Then
    ' Only one directory can be selected in the DirLB
    cItems = 1
    ReDim asPaths(0)
    asPaths(0) = GetDirLBItemPath(Dir1, iItem)
  
  Else
    ' Put the focused (and selected) files's relative pidl in the
    ' first element of the array. This will be the file whose context
    ' menu will be shown if multiple files are selected.
    cItems = 1
    ReDim asPaths(0)
    asPaths(0) = GetFileLBItemPath(File1, iItem)
    
    ' Fill the array with the relative pidls of the rest of any selected
    ' files(s), making sure that we don't add the focused file again.
    For i = 0 To File1.ListCount - 1
      If (File1.Selected(i)) And (i <> iItem) Then
        cItems = cItems + 1
        ReDim Preserve asPaths(cItems - 1)
        asPaths(cItems - 1) = GetFileLBItemPath(File1, i)
      End If
    Next
  
  End If   ' (TypeOf objLB Is DirListBox)
  
  ' ==================================================
  ' Finally, get the IShellFolder of the selected directory, load the relative
  ' pidl(s) of the selected items into the array, and show the menu.
  ' This part won't be elaborated upon, as it is extensively involved.
  ' For more info on IShellFolder, pidls and the shell's context menu, see:
  ' http://msdn.microsoft.com/developer/sdk/inetsdk/help/itt/Shell/NameSpace.htm
  
  If Len(asPaths(0)) Then
    
    ' Get a copy of each selected item's fully qualified pidl from it's path.
    For i = 0 To cItems - 1
      ReDim Preserve apidlFQs(i)
      apidlFQs(i) = GetPIDLFromPath(hWnd, asPaths(i))
    Next
    
    If apidlFQs(0) Then
    
      ' Get the selected item's parent IShellFolder.
      Set isfParent = GetParentIShellFolder(apidlFQs(0))
      If (isfParent Is Nothing) = False Then
        
        ' Get a copy of each selected item's relative pidl (the last item ID)
        ' from each respective item's fully qualified pidl.
        For i = 0 To cItems - 1
          ReDim Preserve apidlRels(i)
          apidlRels(i) = GetItemID(apidlFQs(i), GIID_LAST)
        Next
        
        If apidlRels(0) Then
          ' Subclass the Form so we catch the menu's ownerdraw messages.
          Call SubClass(hWnd, AddressOf WndProc)
          ' Show the shell context menu for the selected items. If a
          ' menu command was executed, refresh the two listboxes.
          If ShowShellContextMenu(hWnd, isfParent, cItems, apidlRels(0), pt, chkPrompt) Then
            Dir1.Refresh
            Call RefreshListBox(File1)
          End If
          ' Finally, unsubclass the form.
          Call UnSubClass(hWnd)
        End If   ' apidlRels(0)
        
        ' Free each item's relative pidl.
        For i = 0 To cItems - 1
          Call MemAllocator.Free(ByVal apidlRels(i))
        Next
        
      End If   ' (isfParent Is Nothing) = False

      ' Free each item's fully qualified pidl.
      For i = 0 To cItems - 1
        Call MemAllocator.Free(ByVal apidlFQs(i))
      Next
      
    End If   ' apidlFQs(0)
  End If   ' Len(asPaths(0))
  
End Sub

Private Function GetFileLBItemPath(objFLB As FileListBox, iItem As Integer) As String
  Dim sPath As String
  
  sPath = objFLB.Path
  If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
  GetFileLBItemPath = sPath & objFLB.List(iItem)

End Function

' Returns the DirListBox Path from the specified listbox item index.

'   - the currently expanded directory (lowest in hierarchy) is ListIndex -1
'   - it's 1st parent directory's ListIndex is -2, if any (the parent indices get smaller)
'   - it's 1st child subdirectory's ListIndex is 0, if any (the child indices get larger)
'   - ListCount is the number of child subdirectories under the currently expanded directory.
'   - List(x) returns the full path of item whose index is x
'   - there is never more than one expanded directory on any directory hierachical level

' It's a little extra work getting the path of the selected DirListBox item...

Private Function GetDirLBItemPath(objDLB As DirListBox, iItem As Integer) As String
  Dim nItems As Integer
  
  ' Get the count of items in the DirLB
  nItems = SendMessage(objDLB.hWnd, LB_GETCOUNT, 0, 0)
  If (nItems > -1) Then   ' LB_ERR
        
    ' Subtract the actual number of LB items from the sum of:
    '   the DirLB's ListCount and
    '   the currently selected directory's real LB index value
    ' (nItems is a value of 1 greater than the last item's real LB index value)
    GetDirLBItemPath = objDLB.List((objDLB.ListCount + iItem) - nItems)

'Debug.Print "iItem: " & iItem & ", LiistIndex: " & (objDLB.ListCount + iItem) - nItems

  End If

End Function

Private Sub RefreshListBox(objLB As Control)
  Dim iFocusedItem As Integer
  Dim i As Integer
  Dim cItems As Integer
  Dim aiSelitems() As Integer
  
  ' Cache the focused item, if any.
  iFocusedItem = objLB.ListIndex
  
  ' Cache any selected items
  For i = 0 To objLB.ListCount - 1
    If objLB.Selected(i) Then
      cItems = cItems + 1
      ReDim Preserve aiSelitems(cItems - 1)
      aiSelitems(cItems - 1) = i
    End If
  Next
  
  ' Refresh the listbox, sets ListIndex = 0, and removes all selction.
  objLB.Refresh

  ' Restore focus and selection to the cached items.
'  objLB.ListIndex = iFocusedItem   ' this errs... (?)
  Call SendMessage(objLB.hWnd, LB_SETCARETINDEX, iFocusedItem, ByVal 0&)
  For i = 0 To cItems - 1
'    objLB.Selected(aiSelitems(i)) = True   ' may err...
    Call SendMessage(objLB.hWnd, LB_SETSEL, CTrue, ByVal aiSelitems(i))
  Next
  
End Sub
