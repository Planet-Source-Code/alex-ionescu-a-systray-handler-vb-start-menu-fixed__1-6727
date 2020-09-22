Attribute VB_Name = "Module1"
Declare Function RegisterTray Lib "tray.dll" _
               (ByVal hInst As Long, _
               ByVal trayXpos As Long, _
               ByVal trayYpos As Long, _
               ByVal trayHeight As Long, _
               ByVal useEmptyTray As Long, _
               ByVal autoHiding As Long) As Long
  

