VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMenuBMP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Bitmaps Example - By: Synthesize"
   ClientHeight    =   870
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Bitmaps From The Menus"
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   435
      Width           =   3135
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Bitmaps To The Menus"
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
   Begin ComctlLib.ImageList ilstView 
      Index           =   1
      Left            =   1500
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":0202
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":0404
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ilstView 
      Index           =   0
      Left            =   1440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":0606
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":0808
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":0A0A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ilstEdit 
      Index           =   1
      Left            =   780
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":0C0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":0E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":0FD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":122E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ilstEdit 
      Index           =   0
      Left            =   720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":1488
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":1692
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":1850
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":1AAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ilstFile 
      Index           =   1
      Left            =   60
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":1D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":1DFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":1EF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":1FF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":20EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ilstFile 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":21E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":22E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":23DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":24D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMenuBMP.frx":25CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolbox 
         Caption         =   "Tool&box"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMenuBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Delcare a variable that we can
' use as the type MENUITEMINFO
Dim MenuInfo As MENUITEMINFO

' Declare our functions that will
' be used to get and set some info
' about the menus
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

' Make a type that will hold the
' menu item info
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Sub cmdAdd_Click()
' If there is an error, continue
' on with the other stuff
On Error Resume Next

' Declare our variables that will
' be used to store info about the
' menus
Dim i, j
Dim lngMenu As Long
Dim lngSubMenu As Long
Dim lngMenuCount As Long
Dim lngSubMenuCount As Long
Dim lngMenuItemID As Long
Dim strMenuItemText As String
Dim lngCurrentItem As Long
Dim imgBitmapUCHK As ListImage
Dim imgBitmapCHK As ListImage

' Get the MAIN menu of our form
' that will have the bitmaps on
' the menus
lngMenu& = GetMenu(Me.hwnd)
' Get the number of MENUS (File,
' Edit, View, Window, ect.) that
' we have
lngMenuCount& = GetMenuItemCount(lngMenu&)

' Do a loop from 1 to the number
' of MENUS we have so that we can
' go through each menu and set
' each item of the menu with a
' bitmap
For i = 1 To lngMenuCount&
    ' Get a MENU (File, Edit,
    ' View, Window, ect.) of the
    ' MAIN menu
    lngSubMenu& = GetSubMenu(lngMenu&, i - 1)
    ' Get a count of items in the
    ' menu
    lngSubMenuCount& = GetMenuItemCount(lngSubMenu&)
    ' Set the current item (just
    ' a variable to count; has
    ' nothing to do with the menu
    ' info) to 1
    lngCurrentItem& = 1

    ' Do another loop from 1 to
    ' the number of items in the
    ' current menu so that we can
    ' set that item's bitmap
    For j = 1 To lngSubMenuCount&
        ' Get the ID of the
        ' current item (j) of the
        ' current menu
        lngMenuItemID& = GetMenuItemID(lngSubMenu&, j - 1)

        ' Set strMenuItemText$ to
        ' 256 null characters
        ' [Chr(0)]
        strMenuItemText$ = String(256, Chr(0))
        ' Get the menu item's text
        ' and set it to
        ' strMenuItemText$
        Call GetMenuString(lngMenu&, lngMenuItemID&, strMenuItemText$, 256, 0&)
        ' Trim all the left over
        ' null characters from
        ' strMenuItemText$
        strMenuItemText$ = Left(strMenuItemText$, InStr(strMenuItemText$, Chr(0)) - 1)

        ' If strMenuItemText$ is
        ' NOT null (contains
        ' nothing; separators are
        ' null) then do the
        ' process of adding the
        ' bitmaps
        If strMenuItemText$ <> vbNullString Then
            ' If i = 1 (the FIRST
            ' menu [File, in our
            ' program]) then set
            ' imgBitmapUCHK to the
            ' appropriate item of
            ' ilstFile(0) and set
            ' imgBitmapCHK to the
            ' appropriate item of
            ' ilstFile(1)
            If i = 1 Then Set imgBitmapUCHK = ilstFile(0).ListImages.Item(lngCurrentItem&): Set imgBitmapCHK = ilstFile(1).ListImages.Item(lngCurrentItem&)
            ' If i = 2 (the SECOND
            ' menu [Edit, in our
            ' program]) then set
            ' imgBitmapUCHK to the
            ' appropriate item of
            ' ilstEdit(0) and set
            ' imgBitmapCHK to the
            ' appropriate item of
            ' ilstEdit(1)
            If i = 2 Then Set imgBitmapUCHK = ilstEdit(0).ListImages.Item(lngCurrentItem&): Set imgBitmapCHK = ilstEdit(1).ListImages.Item(lngCurrentItem&)
            ' If i = 3 (the THIRD
            ' menu [View, in our
            ' program]) then set
            ' imgBitmapUCHK to the
            ' appropriate item of
            ' ilstView(0) and set
            ' imgBitmapCHK to the
            ' appropriate item of
            ' ilstView(1)
            If i = 3 Then Set imgBitmapUCHK = ilstView(0).ListImages.Item(lngCurrentItem&): Set imgBitmapCHK = ilstView(1).ListImages.Item(lngCurrentItem&)

            ' Set the bitmap to
            ' the current item
            Call SetMenuItemBitmaps(lngMenu&, lngMenuItemID&, 0&, imgBitmapUCHK.Picture, imgBitmapCHK.Picture)
            ' Move lngCurrentItem&
            ' up one so that we
            ' can bypass
            ' separators and move
            ' to the next item
            ' that IS NOT a
            ' separator
            lngCurrentItem& = lngCurrentItem& + 1
        End If
    ' Continue the loop of j
    Next j
' Continue the loop of i
Next i
End Sub

Private Sub cmdRemove_Click()
' If there is an error, continue
' on with the other stuff
On Error Resume Next

' Declare our variables that will
' be used to store info about the
' menus
Dim i, j
Dim lngMenu As Long
Dim lngSubMenu As Long
Dim lngMenuCount As Long
Dim lngSubMenuCount As Long
Dim lngMenuItemID As Long
Dim strMenuItemText As String
Dim lngCurrentItem As Long

' Get the MAIN menu of our form so
' that we can get the submenus
' (File, Edit, View, Window, ect.)
' of it
lngMenu& = GetMenu(Me.hwnd)
' Get the number of MENUS (File,
' Edit, View, Window, ect.) that
' we have
lngMenuCount& = GetMenuItemCount(lngMenu&)

' Do a loop from 1 to the number
' of MENUS we have so that we can
' go through each menu and remove
' each bitmap of each item of the
' menu
For i = 1 To lngMenuCount&
    ' Get a MENU (File, Edit,
    ' View, Window, ect.) of the
    ' MAIN menu
    lngSubMenu& = GetSubMenu(lngMenu&, i - 1)
    ' Get a count of items in the
    ' menu
    lngSubMenuCount& = GetMenuItemCount(lngSubMenu&)
    ' Set the current item (just
    ' a variable to count; has
    ' nothing to do with the menu
    ' info) to 1
    lngCurrentItem& = 1

    ' Do another loop from 1 to
    ' the number of items in the
    ' current menu so that we can
    ' remove that item's bitmap
    For j = 1 To lngSubMenuCount&
        ' Get the ID of the
        ' current item (j) of the
        ' current menu
        lngMenuItemID& = GetMenuItemID(lngSubMenu&, j - 1)

        ' Set strMenuItemText$ to
        ' 256 null characters
        ' [Chr(0)]
        strMenuItemText$ = String(256, Chr(0))
        ' Get the menu item's text
        ' and set it to
        ' strMenuItemText$
        Call GetMenuString(lngMenu&, lngMenuItemID&, strMenuItemText$, 256, 0&)
        ' Trim all the left over
        ' null characters from
        ' strMenuItemText$
        strMenuItemText$ = Left(strMenuItemText$, InStr(strMenuItemText$, Chr(0)) - 1)

        ' If strMenuItemText$ is
        ' NOT null (contains
        ' nothing; separators are
        ' null) then do the
        ' process of removing the
        ' bitmaps
        If strMenuItemText$ <> vbNullString Then
            ' Set the bitmap of
            ' the current item to
            ' nothing
            Call SetMenuItemBitmaps(lngMenu&, lngMenuItemID&, 0&, 0&, 0&)
            ' Move lngCurrentItem&
            ' up one so that we
            ' can bypass
            ' separators and move
            ' to the next item
            ' that IS NOT a
            ' separator
            lngCurrentItem& = lngCurrentItem& + 1
        End If
    ' Continue the loop of j
    Next j
' Continue the loop of i
Next i
End Sub

Private Sub Form_Load()
' If there is an error, continue
' on with the other stuff
On Error Resume Next

' Declare our variables that will
' be used to store info about the
' menus
Dim i, j
Dim lngMenu As Long
Dim lngSubMenu As Long
Dim lngMenuCount As Long
Dim lngSubMenuCount As Long
Dim lngMenuItemID As Long
Dim strMenuItemText As String
Dim lngCurrentItem As Long
Dim imgBitmapUCHK As ListImage
Dim imgBitmapCHK As ListImage

' Get the MAIN menu of our form
' that will have the bitmaps on
' the menus
lngMenu& = GetMenu(Me.hwnd)
' Get the number of MENUS (File,
' Edit, View, Window, ect.) that
' we have
lngMenuCount& = GetMenuItemCount(lngMenu&)

' Do a loop from 1 to the number
' of MENUS we have so that we can
' go through each menu and set
' each item of the menu with a
' bitmap
For i = 1 To lngMenuCount&
    ' Get a MENU (File, Edit,
    ' View, Window, ect.) of the
    ' MAIN menu
    lngSubMenu& = GetSubMenu(lngMenu&, i - 1)
    ' Get a count of items in the
    ' menu
    lngSubMenuCount& = GetMenuItemCount(lngSubMenu&)
    ' Set the current item (just
    ' a variable to count; has
    ' nothing to do with the menu
    ' info) to 1
    lngCurrentItem& = 1

    ' Do another loop from 1 to
    ' the number of items in the
    ' current menu so that we can
    ' set that item's bitmap
    For j = 1 To lngSubMenuCount&
        ' Get the ID of the
        ' current item (j) of the
        ' current menu
        lngMenuItemID& = GetMenuItemID(lngSubMenu&, j - 1)

        ' Set strMenuItemText$ to
        ' 256 null characters
        ' [Chr(0)]
        strMenuItemText$ = String(256, Chr(0))
        ' Get the menu item's text
        ' and set it to
        ' strMenuItemText$
        Call GetMenuString(lngMenu&, lngMenuItemID&, strMenuItemText$, 256, 0&)
        ' Trim all the left over
        ' null characters from
        ' strMenuItemText$
        strMenuItemText$ = Left(strMenuItemText$, InStr(strMenuItemText$, Chr(0)) - 1)

        ' If strMenuItemText$ is
        ' NOT null (contains
        ' nothing; separators are
        ' null) then do the
        ' process of adding the
        ' bitmaps
        If strMenuItemText$ <> vbNullString Then
            ' If i = 1 (the FIRST
            ' menu [File, in our
            ' program]) then set
            ' imgBitmapUCHK to the
            ' appropriate item of
            ' ilstFile(0) and set
            ' imgBitmapCHK to the
            ' appropriate item of
            ' ilstFile(1)
            If i = 1 Then Set imgBitmapUCHK = ilstFile(0).ListImages.Item(lngCurrentItem&): Set imgBitmapCHK = ilstFile(1).ListImages.Item(lngCurrentItem&)
            ' If i = 2 (the SECOND
            ' menu [Edit, in our
            ' program]) then set
            ' imgBitmapUCHK to the
            ' appropriate item of
            ' ilstEdit(0) and set
            ' imgBitmapCHK to the
            ' appropriate item of
            ' ilstEdit(1)
            If i = 2 Then Set imgBitmapUCHK = ilstEdit(0).ListImages.Item(lngCurrentItem&): Set imgBitmapCHK = ilstEdit(1).ListImages.Item(lngCurrentItem&)
            ' If i = 3 (the THIRD
            ' menu [View, in our
            ' program]) then set
            ' imgBitmapUCHK to the
            ' appropriate item of
            ' ilstView(0) and set
            ' imgBitmapCHK to the
            ' appropriate item of
            ' ilstView(1)
            If i = 3 Then Set imgBitmapUCHK = ilstView(0).ListImages.Item(lngCurrentItem&): Set imgBitmapCHK = ilstView(1).ListImages.Item(lngCurrentItem&)

            ' Set the bitmap to
            ' the current item
            Call SetMenuItemBitmaps(lngMenu&, lngMenuItemID&, 0&, imgBitmapUCHK.Picture, imgBitmapCHK.Picture)
            ' Move lngCurrentItem&
            ' up one so that we
            ' can bypass
            ' separators and move
            ' to the next item
            ' that IS NOT a
            ' separator
            lngCurrentItem& = lngCurrentItem& + 1
        End If
    ' Continue the loop of j
    Next j
' Continue the loop of i
Next i
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuStatus_Click()
    mnuStatus.Checked = Not mnuStatus.Checked
End Sub

Private Sub mnuToolbar_Click()
    mnuToolbar.Checked = Not mnuToolbar.Checked
End Sub

Private Sub mnuToolbox_Click()
    mnuToolbox.Checked = Not mnuToolbox.Checked
End Sub
