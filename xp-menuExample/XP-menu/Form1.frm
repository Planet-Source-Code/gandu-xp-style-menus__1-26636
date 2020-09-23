VERSION 5.00
Object = "*\APopMenu.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList 
      Left            =   720
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":015A
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":02B4
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":040E
            Key             =   "move"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   0
      Picture         =   "Form1.frx":07E8
      ScaleHeight     =   435
      ScaleWidth      =   2235
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make menu XP"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   2640
      Width           =   1575
   End
   Begin cPopMenu.PopMenu xpMenu 
      Left            =   0
      Top             =   2520
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save     "
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete File     "
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuMove 
         Caption         =   "Move to Folder"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    xpMenu.SubClassMenu Me
    xpMenu.ImageList = ImageList
    
    xpMenu.HighlightStyle = cspHighlightButton
    
    Set xpMenu.BackgroundPicture = Picture1.Picture
    
    xpMenu.ItemIcon("mnuSave") = ImageList.ListImages.Item("save").Index - 1
    xpMenu.ItemIcon("mnuOpen") = ImageList.ListImages.Item("open").Index - 1
    xpMenu.ItemIcon("mnuDelete") = ImageList.ListImages.Item("delete").Index - 1
    xpMenu.ItemIcon("mnuMove") = ImageList.ListImages.Item("move").Index - 1
    
End Sub

Private Sub Form_Load()
    Picture1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    xpMenu.UnsubclassMenu
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub
