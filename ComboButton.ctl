VERSION 5.00
Begin VB.UserControl ComboButton 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   ScaleHeight     =   885
   ScaleWidth      =   4620
   ToolboxBitmap   =   "COMBOB~1.ctx":0000
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   3480
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   1
      Top             =   0
      Width           =   765
      Begin VB.PictureBox Picture1 
         Height          =   315
         Left            =   -15
         ScaleHeight     =   255
         ScaleWidth      =   240
         TabIndex        =   2
         Top             =   0
         Width           =   300
         Begin VB.CommandButton Command1 
            Caption         =   ".."
            Height          =   255
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   240
         End
      End
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   3525
   End
End
Attribute VB_Name = "ComboButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event ClickButton()
Public Event ClickCombo()
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event Change()



'==========================================
'// COMBOBOX & COMMAND BUTTON PROPERTIES
'==========================================

Public Property Let Enabled(blnEnabled As Boolean)
    Command1.Enabled = blnEnabled
    Combo1.Enabled = blnEnabled
End Property
Public Property Get Enabled() As Boolean
    Enabled = Command1.Enabled
End Property



'=================================
'// COMMAND BUTTON EVENTS
'=================================

Private Sub Command1_Click()
    RaiseEvent ClickButton
End Sub



'=================================
'// COMBO ONLY PROPERTIES
'=================================

Public Property Let TopIndex(Index As Long)
    Combo1.TopIndex = Index
End Property
Public Property Get TopIndex() As Long
    TopIndex = Combo1.TopIndex
End Property

Public Property Let Text(Text As String)
    Combo1.Text = Text
End Property
Public Property Get Text() As String
    Text = Combo1.Text
End Property

Public Property Set Font(Font As Object)
    Set Combo1.Font = Font
End Property
Public Property Get Font() As Object
    Set Font = Combo1.Font
End Property

Public Property Let FontSize(FontSize)
    Combo1.FontSize = FontSize
End Property
Public Property Get FontSize()
    FontSize = Combo1.FontSize
End Property

Public Property Let FontName(FontName As String)
    Combo1.FontName = FontName
End Property
Public Property Get FontName() As String
    FontName = Combo1.FontName
End Property

Public Property Let FontBold(FontBold As Boolean)
    Combo1.FontBold = FontBold
End Property
Public Property Get FontBold() As Boolean
    FontBold = Combo1.FontBold
End Property

Public Property Let FontItalic(FontItalic As Boolean)
    Combo1.FontItalic = FontItalic
End Property
Public Property Get FontItalic() As Boolean
    FontItalic = Combo1.FontItalic
End Property

Public Property Let ForeColor(ForeColor)
    Combo1.ForeColor = ForeColor
End Property
Public Property Get ForeColor()
    ForeColor = Combo1.ForeColor
End Property

Public Property Let BackColor(BackColor)
    Combo1.BackColor = BackColor
End Property
Public Property Get BackColor()
    BackColor = Combo1.BackColor
End Property

Public Property Get List(Index As Integer) As String
    List = Combo1.List(Index)
End Property

Public Property Get ListCount() As Long
    ListCount = Combo1.ListCount
End Property

Public Property Let SelLength(Number As Long)
    Combo1.SelLength = Number
End Property
Public Property Get SelLength() As Long
    SelLength = Combo1.SelLength
End Property

Public Property Let SelStart(Number As Long)
    Combo1.SelStart = Number
End Property
Public Property Get SelStart() As Long
    SelStart = Combo1.SelStart
End Property

Public Property Let SelText(Text As String)
    Combo1.SelText = Text
End Property
Public Property Get SelText() As String
    SelText = Combo1.SelText
End Property



'=================================
'// COMBOBOX SUBROUTINES
'=================================

Public Sub AddItem(Item As String, Optional Index)
    Combo1.AddItem Item, Index
End Sub

Public Sub RemoveItem(Index As Integer)
    Combo1.RemoveItem Index
End Sub

Public Sub Clear()
    Combo1.Clear
End Sub

Public Sub SetFocus()
    Combo1.SetFocus
End Sub



'=================================
'// COMBOBOX EVENTS
'=================================

Private Sub Combo1_Change()
    RaiseEvent Change
End Sub

Private Sub Combo1_Click()
    RaiseEvent ClickCombo
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub



Private Sub UserControl_Resize()
    
    If UserControl.Width < 660 Then
        UserControl.Width = 660
        Exit Sub
    End If

    UserControl.Combo1.Width = UserControl.Width - 240
    UserControl.Picture2.Left = UserControl.Combo1.Width - 45
    UserControl.Height = 315
        
End Sub

