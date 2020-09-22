VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{A22D979F-2684-11D2-8E21-10B404C10000}#1.4#0"; "CPOPMENU.OCX"
Begin VB.Form frmTest 
   Caption         =   "PopMenu Control Demonstration"
   ClientHeight    =   6255
   ClientLeft      =   4335
   ClientTop       =   2400
   ClientWidth     =   8340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "PMenuTst.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   8340
   Begin VB.PictureBox picbackround 
      Height          =   3495
      Left            =   840
      Picture         =   "PMenuTst.frx":030A
      ScaleHeight     =   3435
      ScaleWidth      =   4515
      TabIndex        =   5
      Top             =   480
      Width           =   4575
   End
   Begin VB.Frame fraSpecialEffects 
      Caption         =   "Special Effects/Styles"
      Height          =   675
      Left            =   60
      TabIndex        =   2
      Top             =   4800
      Width           =   2595
      Begin VB.CheckBox chkStyle 
         Caption         =   "Button &Select Style"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   420
         Width           =   2415
      End
      Begin VB.CheckBox chkBackground 
         Caption         =   "&Background Bitmap"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Value           =   1  'Checked
         Width           =   2295
      End
   End
   Begin cPopMenu.PopMenu ctlPopMenu 
      Left            =   7620
      Top             =   4980
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightStyle  =   1
   End
   Begin VB.CommandButton cmdUnload 
      Caption         =   "&Close"
      Height          =   435
      Left            =   5460
      TabIndex        =   1
      Top             =   5040
      Width           =   1155
   End
   Begin ComctlLib.ImageList ilsIcons 
      Left            =   6960
      Top             =   4980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":4082
            Key             =   "DOCUMENT"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":461E
            Key             =   "FOLDER"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   5940
      Width           =   8235
   End
   Begin VB.Menu mnuF0Main 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Print"
         Index           =   3
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Print Se&tup"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Test Invisible &1"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Test Invisible &2"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Test Invisible &3"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Test Invisible &4"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   11
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuE0MAIN 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "Cu&t"
         Index           =   0
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Copy"
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Paste"
         Index           =   2
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Search..."
         Index           =   4
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "&In Code"
      Begin VB.Menu mnuSub 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuD1Main 
      Caption         =   "&Directory1"
   End
   Begin VB.Menu mnuD2Main 
      Caption         =   "Direc&tory2"
      Begin VB.Menu mnuDir2 
         Caption         =   "<none>"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuH0MAIN 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Contents..."
         Index           =   0
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&On the Internet..."
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    left As Long
    tOp As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private m_lAboutId As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Private Sub pAddMultiColumnMenu()
Dim lParent As Long
Dim i As Long, iPerRow As Long
Dim s As String

   With ctlPopMenu
      lParent = .MenuIndex("mnuMultiColumn")
      .ClearSubMenusOfItem lParent
      iPerRow = ilsIcons.ListImages.Count \ 4
      For i = 2 To ilsIcons.ListImages.Count
         If ((i - 2) Mod iPerRow) = 0 Then
            s = "^" '& StrConv(ilsIcons.ListImages(i).Key, vbProperCase)
         Else
            s = "" 'StrConv(ilsIcons.ListImages(i).Key, vbProperCase)
         End If
         .AddItem s, "MultiColumn" & i, , , lParent, i - 1
      Next i
   End With
End Sub

Private Sub pCreateMenuItems()
Dim lParentIndex As Long
Dim lIndex As Long
Dim lThisIndex As Long
Dim sPath As String

    With ctlPopMenu
        
        lIndex = .MenuIndex("mnuSub(0)")
        .Caption(lIndex) = "&Back"
        .HelpText(lIndex) = "Move to the previous page"
            .AddItem "Test", "mnuNewTest(1)", , , lIndex
            .AddItem "Test2", "mnuNewTest(2)", , , lIndex
        sPath = "C:\windows\startm~1\"
        lParentIndex = .MenuIndex("mnuD1Main")
        pDirectoryAddItems lParentIndex, sPath
        
    End With
End Sub

Private Sub pDirectoryAddItems( _
        ByVal lParentIndex As Long, _
        ByVal sPath As String, _
        Optional ByVal bTop As Boolean = False _
    )
Dim sFiles() As String
Dim bDir() As Boolean
Dim iFileCOunt As Long
Dim iFile As Long
Dim lIndex As Long
Dim lFolderIcon As Long
Dim lDocIcon As Long
Dim sCaption As String
Dim iCount As Long
Dim X As String
Dim lsName As String
Static iTotal As Long

   If (bTop) Then
      iTotal = 0
   End If

   GetFilesInPath sPath, sFiles(), bDir(), iFileCOunt
   lFolderIcon = plGetIconIndex("FOLDER")
   lDocIcon = plGetIconIndex("DOCUMENT")
   For iFile = 1 To iFileCOunt
      sCaption = sFiles(iFile)
      iCount = iCount + 1
      iTotal = iTotal + 1

      If (iCount > ctlPopMenu.MenuItemsPerScreen) Then
         sCaption = "|" & sCaption
         iCount = 0
      End If
      If (bDir(iFile)) Then
         lIndex = ctlPopMenu.AddItem(sCaption, , , , lParentIndex, lFolderIcon)
      Else
         X = InStrRev(sCaption, ".", , vbTextCompare)
         If X <> 0 Then
         lsName = Trim$(left$(sCaption, X - 1))
         lIndex = ctlPopMenu.AddItem("            " + lsName, , , , lParentIndex, lDocIcon)
         End If
      End If
      If (bDir(iFile)) Then
         pDirectoryAddItems lIndex, sPath & "\" & sFiles(iFile)
      End If
   Next iFile
   If (iFileCOunt = 0) Then
      ctlPopMenu.AddItem "<empty>", , , , lParentIndex, , , False
   End If
   
End Sub
Private Sub GetFilesInPath( _
        ByVal sPath As String, _
        ByRef sFiles() As String, _
        ByRef bDir() As Boolean, _
        ByRef iFileCOunt As Long _
    )
Dim sDir As String
Dim bAdd As Boolean
Dim bIsDir As Boolean

    iFileCOunt = 0
    sDir = Dir(sPath & "\*.*", vbNormal Or vbDirectory)
    Do While Len(sDir) > 0
        If (sDir <> ".") And (sDir <> "..") Then
            bIsDir = ((GetAttr(sPath & "\" & sDir) And vbDirectory) = vbDirectory)
            bAdd = False
            If Not (bIsDir) Then
               bAdd = True
            Else
               bAdd = True
            End If
            If (bAdd) Then
                iFileCOunt = iFileCOunt + 1
                ReDim Preserve sFiles(1 To iFileCOunt) As String
                ReDim Preserve bDir(1 To iFileCOunt) As Boolean
                sFiles(iFileCOunt) = sDir
                bDir(iFileCOunt) = bIsDir
                If (iFileCOunt > ctlPopMenu.MenuItemsPerScreen * 3) Then
                  ' stop - too many...
                  Exit Do
               End If
            End If
        End If
        sDir = Dir
    Loop
End Sub
    
Private Sub chkBackground_Click()
'   If (chkBackground.Value = Checked) Then
 '     Set ctlPopMenu.BackgroundPicture = picBackground.Picture
 '  Else
 '     ctlPopMenu.ClearBackgroundPicture
 '  End If
End Sub

Private Sub chkDefault_Click()
  ' ctlPopMenu.MenuDefault("mnuEdit(4)") = chkDefault.Value * -1
End Sub

Private Sub chkEnable_Click()
  ' ctlPopMenu.Enabled("mnuEdit(2)") = chkEnable.Value * -1
End Sub

Private Sub chkENewest_Click()
  ' ctlPopMenu.Enabled("mnuSub(4)") = chkENewest.Value * -1
End Sub

Private Sub chkNewest_Click()
 '  ctlPopMenu.Checked("mnuSub(4)") = chkNewest.Value * -1
End Sub

Private Sub chkStyle_Click()
 '  If chkStyle.Value = Checked Then
'      ctlPopMenu.HighlightStyle = cspHighlightButton
'   Else
      ctlPopMenu.HighlightStyle = cspHighlightStandard
'   End If
End Sub

Private Sub cmdAdd_Click()
Dim lParent As Long
Dim lID As Long

   With ctlPopMenu
      If Not (.MenuExists("NewItem1")) Then
         ' We don't have this menu:
         lParent = .MenuIndex("mnuE0MAIN")
         .AddItem "-", "NewItem0", , , lParent
         .AddItem "Test Item 1", "NewItem1", , , lParent, 10
         lID = .AddItem("Test Item 2", "NewItem2", , , lParent, 11)
         .AddItem "Test Sub Item 2,1", "NewItem3", , , lID, 1
         .AddItem "Test Sub Item 2,2", "NewItem4", , , lID, 2
         .AddItem "Test Sub Item 2,3", "NewItem5", , , lID, 3
         .AddItem "Test Item 3", "NewItem6", , , lParent, 12
         
         Debug.Print "Add:AfterCount:" & .Count
      Else
         MsgBox "Menu items are already added.", vbInformation
      End If
   End With
 ' cmdRemove.Enabled = True
 ' cmdAdd.Enabled = False
End Sub


Private Sub cmdChangeCaption_Click()
   If (ctlPopMenu.Caption("mnuEdit(2)") = "&Paste") Then
      ctlPopMenu.Caption("mnuEdit(2)") = "Replacement Caption for &Paste"
      'ctlPopMenu.ReplaceItem "mnuEdit(2)", "Replacement Caption for &Paste"
   Else
      ctlPopMenu.Caption("mnuEdit(2)") = "&Paste"
      'ctlPopMenu.ReplaceItem "mnuEdit(2)", "&Paste"
   End If
 '  lblCaption.Caption = ctlPopMenu.Caption("mnuEdit(2)")
End Sub



Private Sub cmdFindHierarchy_Click()
Dim lH() As Long
Dim lCount As Long
Dim i As Long
Dim sIndex As String
Dim sI As String
Dim lIndex As Long

   sIndex = InputBox("Enter hierarchy to find item for: ", , "2,5")
   If (sIndex <> "") Then
      For i = 1 To Len(sIndex)
         If (Mid$(sIndex, i, 1) = ",") Then
            lCount = lCount + 1
            ReDim Preserve lH(1 To lCount) As Long
            lH(lCount) = CLng(sI)
            sI = ""
         Else
            sI = sI & Mid$(sIndex, i, 1)
         End If
      Next i
      If (sI <> "") Then
         lCount = lCount + 1
         ReDim Preserve lH(1 To lCount) As Long
         lH(lCount) = CLng(sI)
      End If
      
      lIndex = ctlPopMenu.IndexForMenuHierarchy(lH())
      If (lIndex > 0) Then
         MsgBox "Found at index " & lIndex & vbCrLf & "Caption: " & ctlPopMenu.Caption(lIndex) & vbCrLf & "Icon Index: " & ctlPopMenu.ItemIcon(lIndex), vbInformation
      End If
   End If
   Exit Sub
   
ErrorHandler:
   MsgBox "Couldn't interpret " & sIndex, vbInformation
   Exit Sub
End Sub

Private Sub cmdFindKey_Click()
Dim sI As String
Dim lIndex As Long
    sI = InputBox("Enter the key you wish to find", , "mnuEdit(2)")
    If (sI <> "") Then
        On Error Resume Next
        lIndex = ctlPopMenu.MenuIndex(sI)
        If (Err.Number <> 0) Then
            lIndex = -1
        End If
        If (lIndex > -1) Then
            MsgBox "Found at index " & lIndex & vbCrLf & "Caption: " & ctlPopMenu.Caption(lIndex) & vbCrLf & "Icon Index: " & ctlPopMenu.ItemIcon(lIndex), vbInformation
        Else
            MsgBox "No item with key '" & sI & "' was found.", vbInformation
        End If
    End If
End Sub

Private Sub cmdInsert_Click()
Static l As Long
   With ctlPopMenu
      .InsertItem "This has just been inserted", "mnuEdit(1)", "InsertItem" & l, , , Rnd * ilsIcons.ListImages.Count
      l = l + 1
   End With
End Sub

Private Sub cmdMDIDemo_Click()
 '   mfrmMDITest.Show
End Sub

Private Sub cmdMore_Click()
'Dim fB As frmBitmaps
'    Set fB = New frmBitmaps
'    fB.Show
End Sub

Private Sub cmdRemove_Click()
Dim iItem As Long
   With ctlPopMenu
      Debug.Print "Remove:BeforeCount:" & .Count
      ' NOTE: here we loop backwards because the menu
      ' item with key "NewItem2" has a sub menu.  When you delete
      ' "NewItem2" this automatically deletes the subitems (i.e.
      ' "NewItem3","NewItem4" and "NewItem5").  So if you delete
      ' "NewItem2" then "NewItem3" thru "NewItem5" no longer exist
      ' and you get a subscript out of range error.
      ' The alternative is to only delete the items with keys "NewItem0","NewItem1",
      ' "NewItem2" and "NewItem6"
      For iItem = 6 To 0 Step -1
          .RemoveItem "NewItem" & iItem
      Next iItem
      Debug.Print "Remove:AfterCount:" & .Count
   End With
  ' cmdRemove.Enabled = False
 '  cmdAdd.Enabled = True
End Sub

Private Sub cmdUnload_Click()
    mnuFile_Click 11
End Sub

Private Sub cmdVBPopup_Click()
   ' Me.PopupMenu mnuPop, , cmdVBPopup.Container.left + cmdVBPopup.left, cmdVBPopup.Container.tOp + cmdVBPopup.tOp + cmdVBPopup.Height
End Sub

Private Sub cmdGet_Click()
Dim lH() As Long
Dim lR As Long
Dim l As Long
Dim sOut As String

   ReDim lH(1 To 4) As Long
   lH(1) = 3
   lH(2) = 7
   lH(3) = 2
   lH(4) = 1
   lR = ctlPopMenu.IndexForMenuHierarchy(lH())
   If (lR > 0) Then
      For l = 1 To 4
         sOut = sOut & lH(l) & ","
      Next l
      MsgBox "Index for item at hierarchy position: " & vbCrLf & left$(sOut, Len(sOut) - 1) & vbCrLf & lR & " (" & ctlPopMenu.Caption(lR) & ")", vbInformation
   Else
      MsgBox "Index not found", vbExclamation
   End If
End Sub

Private Sub cmdAPIPopup_Click()
Dim lR As Long

   With ctlPopMenu
      ' Track popup menu now built into the control:
  '    lR = ctlPopMenu.ShowPopupMenu(cmdAPIPopup, "mnuSubSub(1)", 0, cmdAPIPopup.Height)
        
      ' How to do it with the API:
      'Dim lIndex As Long
      'Dim hMenu As Long
      'Dim tR As RECT
      'Dim tP As POINTAPI

      'lIndex = .MenuIndex("mnuSubSub(1)")
      'If (lIndex > 0) Then
      '    hMenu = .hPopupMenu(lIndex)
      '    If (hMenu > 0) Then
      '        tP.X = (cmdAPIPopup.Container.left + cmdAPIPopup.left) \ Screen.TwipsPerPixelX
      '        tP.Y = (cmdAPIPopup.Container.tOp + cmdAPIPopup.tOp + cmdAPIPopup.Height) \ Screen.TwipsPerPixelY
      '        ClientToScreen Me.hwnd, tP
      '        lR = TrackPopupMenu(hMenu, 0, tP.X, tP.Y, 0, Me.hwnd, tR)
      '    End If
      'End If
    End With
End Sub

Private Sub cmdChangeIcon_Click()
   If ctlPopMenu.ItemIcon("mnuEdit(2)") = 6 Then
      ctlPopMenu.ItemIcon("mnuEdit(2)") = Rnd * ilsIcons.ListImages.Count
   Else
      ctlPopMenu.ItemIcon("mnuEdit(2)") = 6
   End If
   
  ' picIcon.Cls
  ' ilsIcons.ListImages(ctlPopMenu.ItemIcon("mnuEdit(2)") + 1).Draw picIcon.hdc, 4 * Screen.TwipsPerPixelX, 4 * Screen.TwipsPerPixelY, imlTransparent
  ' picIcon.Refresh

End Sub

Private Sub cmdVisible_Click()
Static i As Long

   ' Make one of the invisible menu items
   ' visible again:
   If (i = 0) Then
      i = 5
   Else
      i = i + 1
   End If
   
   If (i = 9) Then
  '    cmdVisible.Enabled = False
   End If
      
   ' Make menu item visible:
   mnuFile(i).Visible = True
   ' Add it to the control.  When the control finds
   ' the menu you have just made Visible, it will
   ' raise a RequestNewMenuItem event, passing the
   ' caption.  You have to match this up to set
   ' the key property etc.
   ctlPopMenu.CheckForNewItems
   
End Sub

Private Sub ctlPopMenu_Click(ItemNumber As Long)
    Debug.Print "Clicked " & ItemNumber
    lblStatus = "Clicked: " & ctlPopMenu.Caption(ItemNumber)
End Sub

Private Sub ctlPopMenu_InitPopupMenu(ParentItemNumber As Long)
Dim lIndex As Long
Dim sPath As String
Dim sFiles() As String
Dim bDir() As Boolean
Dim sPrefix As String
Dim iCount As Long
Dim iFile As Integer
Dim lParent As Long
Dim lTop As Long
Dim iNumber As Long
Static bPopulated As Boolean

   With ctlPopMenu
      If (.MenuKey(ParentItemNumber) = "mnuMultiColumn") Then
         If Not (bPopulated) Then
            pAddMultiColumnMenu
            bPopulated = True
         End If
      Else
         lTop = .UltimateParent(ParentItemNumber)
         If (.MenuKey(lTop) = "mnuD2Main") Then
            'Debug.Print
            'Debug.Print .Caption(ParentItemNumber)
            'Debug.Print "InitPopupMenu:Start:" & .Count
            Screen.MousePointer = vbHourglass
            'Debug.Print "This is Directory 2!"
            If ParentItemNumber = lTop Then
               'Debug.Print "We are at the top"
               sPath = App.Path
            Else
               'Debug.Print "Doing a sub level"
               sPath = App.Path & "\" & .HierarchyPath(ParentItemNumber, 2, "\")
            End If
            lParent = .ClearSubMenusOfItem(ParentItemNumber)
            'Debug.Print "InitPopupMenu:AfterClear:" & .Count, ParentItemNumber
            GetFilesInPath sPath, sFiles(), bDir(), iCount
            If (iCount > 0) Then
            
               If (iCount > .MenuItemsPerScreen * 3) Then
                  MsgBox "There are too many menu items in the path '" & App.Path & "'." & vbCrLf & "Only the first " & .MenuItemsPerScreen * 3 & " items will be shown.", vbInformation
                  iCount = .MenuItemsPerScreen * 3
               End If
               
               For iFile = 1 To iCount
                  If (bDir(iFile)) Then
                     lIndex = .AddItem(sPrefix & sFiles(iFile), , , , lParent, plGetIconIndex("FOLDER"))
                     iNumber = iNumber + 1
                     If (iNumber >= .MenuItemsPerScreen) Then
                        sPrefix = "|"
                        iNumber = 0
                     Else
                        sPrefix = ""
                     End If
                     'Debug.Print lIndex, sFiles(iFile)
                     .AddItem "<none>", , , , lIndex, , , False
                  End If
               Next iFile
               
               For iFile = 1 To iCount
                  If Not (bDir(iFile)) Then
                     lIndex = .AddItem(sPrefix & sFiles(iFile), , , , lParent, plGetIconIndex("DOCUMENT"))
                     iNumber = iNumber + 1
                     If (iNumber >= .MenuItemsPerScreen) Then
                        sPrefix = "|"
                        iNumber = 0
                     Else
                        sPrefix = ""
                     End If
                     'Debug.Print lIndex, sFiles(iFile)
                  End If
               Next iFile
                
            Else
               .AddItem "<empty>", , , , ParentItemNumber, , , False
            End If
            Screen.MousePointer = vbNormal
            'Debug.Print "InitPopupMenu:End:" & .Count, ParentItemNumber
            
         End If
      End If
      
   End With
End Sub

Private Sub ctlPopMenu_ItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
Dim sText As String
    'Debug.Print "Highlight " & ItemNumber
    If Not (bSeparator) Then
        sText = ctlPopMenu.HelpText(ItemNumber)
        If (sText = "") Then
            sText = "Highlight " & ctlPopMenu.Caption(ItemNumber) & " (Help unavailable)"
        End If
        If (bEnabled) Then
            lblStatus = sText
        Else
            lblStatus = sText & " (Not available)"
        End If
    Else
        lblStatus = ""
    End If
End Sub

Private Sub ctlPopMenu_MenuExit()
    lblStatus = "Menu Exited."
End Sub

Private Sub ctlPopMenu_RequestNewMenuDetails(sCaption As String, sKey As String, iIcon As Long, lItemData As Long, sHelptext As String, sTag As String)
Dim iPos As Long
Dim i As Long

   ' This event is fired if the cPopMenu control
   ' detects a new menu item after the CheckForNewMenuItems
   ' method is called.  Unfortunately, the control can
   ' no longer automatically match up menu items to
   ' the VB menus, so this is your only opportunity
   ' to play with the captions.
   '
   '

   ' Only file menu items can be made invisible,
   ' so we search these to see if we get a
   ' caption match to set the correct key.
   Debug.Print "NEW MENU ITEM APPEARED: ", sCaption

   For i = 6 To 9
      If (sCaption = mnuFile(i).Caption) Then
         iPos = i
         Exit For
      End If
   Next i
   
   If (iPos > 0) Then
      sKey = "mnuFile(" & iPos & ")"
      iIcon = Rnd * ilsIcons.ListImages.Count - 1
   End If
   
End Sub

Private Sub ctlPopMenu_SystemMenuClick(ItemNumber As Long)
   
   ' This event is fired when a system menu
   ' item is clicked:
   
   Debug.Print ItemNumber, m_lAboutId
   Select Case ItemNumber
   Case SC_MOVE
      lblStatus = "Clicked on Move the Window"
   Case SC_MINIMIZE
      lblStatus = "Clicked on Minimise the window"
   Case SC_MAXIMIZE
      lblStatus = "Clicked on Maximise the window"
   Case SC_CLOSE
      ' Note this is removed in the demo
      lblStatus = "Clicked Close"
   Case SC_RESTORE
      lblStatus = "Clicked Restore"
   Case SC_SIZE
      lblStatus = "Clicked Size"
   Case m_lAboutId
      ' Clicked the customised about item
      ' added at run-time:
      Dim lMajor As Long, lMinor As Long, lRevision As Long
      ctlPopMenu.GetVersion lMajor, lMinor, lRevision
      MsgBox "vbAccelerator IconMenu control demonstration." & vbCrLf & "Visit vbAccelerator at http://vbaccelerator.com" & vbCrLf & "Copyright © 1998 Steve McMahon" & vbCrLf & vbCrLf & "Control Version: " & lMajor & "." & lMinor & "." & lRevision, vbInformation
   End Select
End Sub

Private Sub ctlPopMenu_SystemMenuItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
    Select Case ItemNumber
    Case SC_MOVE
        lblStatus = "Move the Window"
    Case SC_MINIMIZE
        lblStatus = "Minimise the window to an icon"
    Case SC_MAXIMIZE
        lblStatus = "Maximise the window's size"
    Case SC_CLOSE
        lblStatus = "Close this application"
    Case SC_RESTORE
        lblStatus = "Restore the window to its previous size"
    Case SC_SIZE
        lblStatus = "Change the size of the window"
    Case m_lAboutId
        lblStatus = "Find out about this program"
    Case Else
        lblStatus = ""
    End Select

End Sub

Private Sub ctlPopMenu_WinIniChange()
    
    ' This is not the place to put a message box - it will
    ' lock the system!
    Debug.Print "*******************************************"
    Debug.Print "GOT A WININICHANGE EVENT"
    Debug.Print "*******************************************"
    
    
End Sub

Private Sub Form_Click()
ctlPopMenu.ShowPopupMenu Me, "mnuD1Main", 10, 80, 0
End Sub

Private Sub Form_Load()
  
Dim l As Long
Dim lIndex As Long
Dim lC As Long

    With ctlPopMenu
        .ImageList = ilsIcons
        .SubClassMenu Me
    End With
    ' Add a whole new set of menu items and sub items to the last
    ' menu item:
    pCreateMenuItems
        
End Sub
Private Sub pSetIcon( _
        ByVal sIconKey As String, _
        ByVal sMenuKey As String _
    )
Dim lIconIndex As Long
    lIconIndex = plGetIconIndex(sIconKey)
    ctlPopMenu.ItemIcon(sMenuKey) = lIconIndex
End Sub
Private Function plGetIconIndex( _
        ByVal sKey As String _
    ) As Long
    plGetIconIndex = ilsIcons.ListImages.Item(sKey).Index - 1
End Function

Private Sub mnuEdit_Click(Index As Integer)
    MsgBox "Visual Basic Menu Edit Fired for Index:" & Index, vbInformation
End Sub

Private Sub mnuFile_Click(Index As Integer)
    If (Index = 11) Then
        If (vbYes = MsgBox("Are you sure you want to exit?", vbYesNo Or vbQuestion)) Then
            Unload Me
        End If
    Else
        MsgBox "Visual Basic Menu File Fired for Index:" & Index, vbInformation
    End If
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    If (Index < 3) Then
        MsgBox "Visual Basic Help Menu Fired for Index:" & Index, vbInformation
    Else
        ctlPopMenu_SystemMenuClick m_lAboutId
    End If
End Sub

Private Sub picIcon_Click()
Dim i As Long
Dim lIndex As Long
   For i = 0 To Controls.Count - 1
      Debug.Print Controls(i).Name,
      If TypeOf Controls(i) Is Menu Then
         Debug.Print Controls(i).Caption;
      Else
         Debug.Print
      End If
      On Error Resume Next
      lIndex = Controls(i).Index
      If (Err.Number = 0) Then
         Debug.Print lIndex
      End If
   Next i
End Sub
