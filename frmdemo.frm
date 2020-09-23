VERSION 5.00
Begin VB.Form frmdemo 
   Caption         =   "Demo"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   390
      Left            =   1785
      TabIndex        =   5
      Top             =   4005
      Width           =   1425
   End
   Begin VB.PictureBox Picture1 
      Height          =   1875
      Left            =   2595
      ScaleHeight     =   1815
      ScaleWidth      =   1590
      TabIndex        =   2
      Top             =   525
      Width           =   1650
   End
   Begin Project1.devColorSelect devColorSelect1 
      Height          =   2295
      Left            =   165
      TabIndex        =   1
      Top             =   210
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   4048
   End
   Begin VB.CommandButton cmdFillList 
      Caption         =   "Show Colors"
      Height          =   390
      Left            =   180
      TabIndex        =   0
      Top             =   4005
      Width           =   1425
   End
   Begin VB.Label lblInfo 
      Height          =   795
      Left            =   165
      TabIndex        =   4
      Top             =   2775
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "Preview"
      Height          =   210
      Left            =   2625
      TabIndex        =   3
      Top             =   255
      Width           =   1695
   End
End
Attribute VB_Name = "frmdemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFillList_Click()
    devColorSelect1.Clear

    devColorSelect1.AddItem "Black", &H0
    devColorSelect1.AddItem "Red", &HFF
    devColorSelect1.AddItem "Green", vbGreen
    devColorSelect1.AddItem "Yellow", vbYellow
    devColorSelect1.AddItem "Blue", &HFF0000
    devColorSelect1.AddItem "Magenta", vbMagenta
    devColorSelect1.AddItem "Cyan", &HFFFF00
    devColorSelect1.AddItem "White", &HFFFFFF
    devColorSelect1.AddItem "Scroll Bar Color", &H80000000
    devColorSelect1.AddItem "Desktop", &H80000001
    devColorSelect1.AddItem "Active TitleBar", &H80000002
    devColorSelect1.AddItem "Inactive Title Bar", &H80000003
    devColorSelect1.AddItem "Menu Bar", &H80000004
    devColorSelect1.AddItem "Window Background", &H80000005
    devColorSelect1.AddItem "Window Frame", &H80000006
    devColorSelect1.AddItem "Menu Text", &H80000007
    devColorSelect1.AddItem "Window Text", &H80000008
    devColorSelect1.AddItem "Active Title Bar Text", &H80000009
    devColorSelect1.AddItem "Active Border", &H8000000A
    devColorSelect1.AddItem "Inactive Border", &H8000000B
    devColorSelect1.AddItem "Application Workspace", &H8000000C
    devColorSelect1.AddItem "Highlight", &H8000000D
    devColorSelect1.AddItem "Highlight Text", &H8000000E
    devColorSelect1.AddItem "Button Face", &H8000000F
    devColorSelect1.AddItem "Button Shadow", &H80000010
    devColorSelect1.AddItem "Disabled Text", &H80000011
    devColorSelect1.AddItem "Button Text", &H80000012
    devColorSelect1.AddItem "Inactive Title Bar Text", &H80000013
    devColorSelect1.AddItem "Button Highlight", &H80000014
    devColorSelect1.AddItem "Button Dark Shadow", &H80000015
    devColorSelect1.AddItem "Button Light Shadow", &H80000016
    devColorSelect1.AddItem "ToolTip Text", &H80000017
    devColorSelect1.AddItem "ToolTip", &H80000018
    
    devColorSelect1.PaintListItems ' setup the control
End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub devColorSelect1_DevColorSelectorMouseDown(button As Integer, ItemIndex As Integer, ItemCaption As String, ColorKey As Long)
    lblInfo.Caption = "Color Item " & ItemCaption _
    & vbCrLf & "Color Value : " & ColorKey _
    & vbCrLf & "Index Selected : " & devColorSelect1.ListIndex _
    & vbCrLf & "List Count : " & devColorSelect1.ListCount
End Sub

Private Sub devColorSelect1_DevColorSelectorMouseUp(button As Integer, ItemIndex As Integer, ItemCaption As String, ColorKey As Long)
    Picture1.BackColor = devColorSelect1.ColorKeyValue(ItemIndex)
End Sub

