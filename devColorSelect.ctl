VERSION 5.00
Begin VB.UserControl devColorSelect 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   161
   Begin VB.PictureBox PicBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   3
      Top             =   2835
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.PictureBox PicList 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1425
      Left            =   0
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   0
      Top             =   0
      Width           =   2400
      Begin VB.VScrollBar vBar 
         Height          =   360
         Left            =   1350
         TabIndex        =   2
         Top             =   60
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox PicItemDC 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   15
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   1
         Top             =   15
         Width           =   555
      End
   End
End
Attribute VB_Name = "devColorSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type TColorList
    Count As Integer
    ItemCaption() As String
    ItemColorKey() As Long
End Type

Private m_ItemSelectedPos As Integer
Private Const m_ItemHeight As Integer = 15
Private ColourList As TColorList

Event DevColorSelectorMouseDown(button As Integer, ItemIndex As Integer, ItemCaption As String, ColorKey As Long)
Event DevColorSelectorMouseUp(button As Integer, ItemIndex As Integer, ItemCaption As String, ColorKey As Long)

Private Sub PaintItem(index, Style As Integer)
    If ColourList.Count = 0 Then Exit Sub
    If Style = 1 Then
        DrawItem ColourList.ItemCaption(index), vbWhite, ColourList.ItemColorKey(index), vbHighlight, True
    Else
        DrawItem ColourList.ItemCaption(index), vbBlack, ColourList.ItemColorKey(index), vbWhite, False
    End If
    BitBlt PicItemDC.hDC, 0, index * m_ItemHeight, PicBuffer.Width, PicBuffer.Height, PicBuffer.hDC, 0, 0, vbSrcCopy
    PicItemDC.Refresh
End Sub

Public Sub PaintListItems()
Dim I As Integer, ItemPosition As Integer
    For I = 0 To ColourList.Count
       ItemPosition = (I * PicBuffer.Height)
       DrawItem ColourList.ItemCaption(I), vbBlack, ColourList.ItemColorKey(I), vbWhite, False
       BitBlt PicItemDC.hDC, 0, ItemPosition, PicBuffer.Width, PicBuffer.Height, PicBuffer.hDC, 0, 0, vbSrcCopy
       PicItemDC.Refresh
    Next
    ItemPosition = 0
    I = 0
End Sub

Private Sub DrawItem(ItemText As String, ItemTextColor As Long, ItemColorBk As Long, ListItemBackColor As Long, DrawSelection As Boolean, Optional YY As Integer)
    PicBuffer.Cls
    PicBuffer.DrawStyle = 0
    PicBuffer.BackColor = ListItemBackColor
    PicBuffer.Line (1, 1)-(10, 11), vbBlack, B ' Draw the outline border of the square
    PicBuffer.Line (2, 2)-(9, 10), ItemColorBk, BF ' draw a filled square with a filled color
    ' Add the color name
    PicBuffer.ForeColor = ItemTextColor
    PicBuffer.CurrentY = 1
    PicBuffer.CurrentX = 15
    PicBuffer.Print ItemText
    PicBuffer.Refresh
    
    If DrawSelection Then
        PicBuffer.DrawWidth = 1
        PicBuffer.DrawStyle = 2
        PicBuffer.Line (0, 0)-(UserControl.ScaleWidth - 20, 14), &H95DBF5, B ' draw a filled square with a filled color
    End If
    
End Sub

Public Sub AddItem(ColorItemCaption As String, ColorItem As Long, Optional nKey As Long)
    ColourList.Count = ColourList.Count + 1
    ReDim Preserve ColourList.ItemCaption(ColourList.Count)
    ReDim Preserve ColourList.ItemColorKey(ColourList.Count)

    ColourList.ItemCaption(ColourList.Count - 1) = ColorItemCaption
    ColourList.ItemColorKey(ColourList.Count - 1) = ColorItem
    
    PicItemDC.Height = ColourList.Count * m_ItemHeight
    vBar.Max = (ColourList.Count * m_ItemHeight - UserControl.ScaleHeight)
    vBar.Visible = vBar.Max > 0
    'PaintListItems
    vBar.Visible = vBar.Max > 0
End Sub

Private Sub PicItemDC_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    If ColourList.Count = 0 Then Exit Sub
    PaintItem m_ItemSelectedPos, 0
    m_ItemSelectedPos = Fix(Y \ m_ItemHeight)
    PaintItem m_ItemSelectedPos, 1
    RaiseEvent DevColorSelectorMouseDown(button, m_ItemSelectedPos, ColourList.ItemCaption(m_ItemSelectedPos), ColourList.ItemColorKey(m_ItemSelectedPos))
End Sub

Private Sub PicItemDC_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
    If ColourList.Count = 0 Then Exit Sub
    RaiseEvent DevColorSelectorMouseUp(button, m_ItemSelectedPos, ColourList.ItemCaption(m_ItemSelectedPos), ColourList.ItemColorKey(m_ItemSelectedPos))
End Sub

Private Sub vBar_Change()
    PicItemDC.Top = -vBar.Value
End Sub

Private Sub vBar_Scroll()
    vBar_Change
End Sub
Public Sub Clear()
    PicBuffer.Cls
    PicItemDC.Cls
    
    Set PicBuffer.Picture = Nothing
    Set PicItemDC = Nothing
    
    ColourList.Count = 0
    Erase ColourList.ItemCaption()
    Erase ColourList.ItemCaption
    vBar.Max = 0
    vBar.Value = 0
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    PicList.Width = UserControl.Width
    PicList.Height = UserControl.Height
    PicBuffer.Width = PicList.Width
    If Err Then UserControl.Size 100, 100
    vBar.Top = 0
    vBar.Left = (UserControl.ScaleWidth - vBar.Width)
    vBar.Height = (UserControl.ScaleHeight)
    PicItemDC.Width = vBar.Left - 2
End Sub

Private Sub UserControl_Show()
    ColourList.Count = 0
End Sub

Public Property Get ListCount() As Integer
    ListCount = ColourList.Count
End Property

Public Property Get ListIndex() As Integer
    ListIndex = m_ItemSelectedPos
End Property

Public Property Get ColorKeyValue(index) As Long
    ColorKeyValue = ColourList.ItemColorKey(index)
End Property

Public Property Get ItemCaption(index) As String
    ItemCaption = ColourList.ItemCaption(index)
End Property

