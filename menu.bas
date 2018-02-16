Attribute VB_Name = "menu"
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Const WM_COMMAND = &H111

Public Sub acionaMenu(ByVal MainWin As Long, menu As Long, itemmenu As Long)

        MainMenu = GetMenu(MainWin) 'Find Main Menu
        SubMenu = GetSubMenu(MainMenu, menu) 'Find Sub Menu (File)
        MenuID = GetMenuItemID(SubMenu, itemmenu) 'Find Import Menu Item
        PostMessage MainWin, WM_COMMAND, MenuID, 0& 'Select Import Menu Item, Import Dialog Box Comes Up
End Sub


