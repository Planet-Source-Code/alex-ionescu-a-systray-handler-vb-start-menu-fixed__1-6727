VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "The New Era Has Arrived"
   ClientHeight    =   960
   ClientLeft      =   240
   ClientTop       =   435
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   960
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Make it nicer!"
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' System Tray Handler
' Tray.dll By Serge.
' Frontend/Conversion By Alex Ionescu
' Start Menu by Alex Ionescu with the vbaccelerator PopMenu OCX
' If you decide to use any piece of code, please e-mail me
' at billyismycat@yahoo.com
' It is wrong to take credit for what others have done.
' You are not respecting Open Source if you do!











Private Sub Check1_Click()
'' If the check box is selected, put a backround picture ''
 If (Check1.Value = Checked) Then
      Set frmTest.ctlPopMenu.BackgroundPicture = frmTest.picbackround.Picture
   Else
      frmTest.ctlPopMenu.ClearBackgroundPicture
   End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
'' Special way of popping up this kind of menus ''
frmTest.ctlPopMenu.ShowPopupMenu Me, "mnuD1Main", 10, 80, 0
End Sub

Private Sub Form_Load()
'' Call RegisterTray from Module1 and tell it to go to positionY 640 and Y 480. Make its height 25. No autohide and don't make it empty ''
RegisterTray hInst, 500, 500, 25, 0, 0
End Sub
