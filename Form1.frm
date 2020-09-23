VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Using Flex Grid"
   ClientHeight    =   5040
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCell 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid fg 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6800
      _Version        =   393216
      WordWrap        =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   1
   End
   Begin VB.Menu mpopup 
      Caption         =   "PopupMenu"
      Begin VB.Menu copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnucut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnudel 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnupaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnudelrow 
         Caption         =   "&Delete Row"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rhold, chold
Dim r1, c1, r2, c2
Dim vr1, vc1, vr2, vc2
Dim mychar
Dim x, y
Dim txt
Dim txtarr()
Dim mcop, mpas, mdel, mcut, dselc

      
Private Sub MoveTextBox()

      txtCell.Visible = True
      txtCell.Left = fg.Left + fg.CellLeft
      txtCell.Top = fg.Top + fg.CellTop
      txtCell.Height = fg.CellHeight
      txtCell.Width = fg.CellWidth
      txtCell.Text = fg.Text
      
      If (mychar >= 48 And mychar <= 57) Or (mychar >= 65 And mychar <= 90) Then
      txtCell.Text = txtCell.Text & Chr(mychar)
      txtCell.SelStart = Len(txtCell.Text)
      txtCell.SelLength = 0
      End If
      
      txtCell.SetFocus
          
      txtCell.ZOrder (0)
      mychar = 0

End Sub



Private Sub copy_Click()
On Error Resume Next
dselc = True
txt = ""

      vr1 = fg.Row
      vc1 = fg.Col
      vr2 = fg.RowSel
      vc2 = fg.ColSel

ReDim txtarr(vc1 To vc2, vr1 To vr2)
For y = vr1 To vr2
    For x = vc1 To vc2
    txtarr(x, y) = fg.TextMatrix(y, x)
    txt = txt & ":" & txtarr(x, y)
    Next x
txt = txt & vbNewLine
Next y

End Sub





Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)

'48 to 57 keycode for numeris char
mychar = KeyCode

If mychar <> 9 And mychar <> 16 And mychar <> 17 _
   And mychar <> 18 And mychar <> 8 And mychar <> 46 And mychar <> 13 Then
    MoveTextBox
ElseIf mychar = 46 Then
Call mnudel_Click
End If

End Sub

Private Sub fg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
Call PopupMenu(mpopup)
End If
End Sub


Private Sub fg_Scroll()
    fg.Text = txtCell.Text
    txtCell.Visible = False
End Sub

Private Sub fg_SelChange()

If dselc = False Then

      r1 = fg.Row
      c1 = fg.Col
     
      r2 = fg.RowSel
      c2 = fg.ColSel
      
      fg.Row = rhold
      fg.Col = chold
      
      txtCell.Visible = False
      fg.Text = txtCell.Text
       
      fg.Row = r1
      fg.Col = c1
      fg.ColSel = c2
      fg.RowSel = r2
End If

End Sub

Private Sub Form_Load()
fg.Cols = 10
fg.Rows = 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "BY: Zeff O. Calilung" & vbNewLine _
        & "Email: Zeff@programmer.net" & vbNewLine _
        & "Thanks for downloading my program hope this help you" & vbNewLine _
        & vbNewLine _
        & "Do you have an old laptop or PC? coz i iwant to have a PC but i could not afford to buy" & vbNewLine _
        & "Please if ever you are willing to donate an old PC i would be glad to have it from you" & vbNewLine _
& "you can send it to #9 Coca-Cola Village Matina, Davao city 8000 Philippines" & vbNewLine
End Sub

Private Sub mnucut_Click()
On Error Resume Next
dselc = True
txtCell.Text = ""
txt = ""

      rhold = fg.Row
      chold = fg.Col
      
      vr1 = fg.Row
      vc1 = fg.Col
      vr2 = fg.RowSel
      vc2 = fg.ColSel

ReDim txtarr(vc1 To vc2, vr1 To vr2)
For y = vr1 To vr2
    For x = vc1 To vc2
    txtarr(x, y) = fg.TextMatrix(y, x)
    txt = txt & ":" & txtarr(x, y)
    fg.TextMatrix(y, x) = ""
    Next x
txt = txt & vbNewLine
Next y

End Sub

Private Sub mnudel_Click()
On Error Resume Next
dselc = True
txtCell.Text = ""

      rhold = fg.Row
      chold = fg.Col
      
      r1 = fg.Row
      c1 = fg.Col
      r2 = fg.RowSel
      c2 = fg.ColSel
      
     rhold = r1
     chold = c1
For y = r1 To r2
    For x = c1 To c2

    fg.TextMatrix(y, x) = ""
    
    Next x
    txt = txt & vbNewLine
Next y


End Sub







Private Sub mnudelrow_Click()
dselc = True
fg.RemoveItem (fg.Row)
End Sub

Private Sub mnuInsert_Click()
dselc = True

End Sub

Private Sub mnupaste_Click()
dselc = True
On Error Resume Next
Dim nc1, nr1, dumX, dumY
nc1 = fg.Col
nr1 = fg.Row

rhold = fg.Row
chold = fg.Col

dumY = nr1

For y = vr1 To vr2
    dumX = nc1
    For x = vc1 To vc2
        fg.TextMatrix(dumY, dumX) = txtarr(x, y)
        dumX = dumX + 1
    Next x
dumY = dumY + 1
Next y

End Sub


Private Sub txtCell_GotFocus()
dselc = False
rhold = fg.Row
chold = fg.Col
End Sub

Private Sub txtCell_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = 13 Then
    fg.Text = txtCell.Text
    txtCell.Visible = False
    
    If fg.Col <> fg.Cols - 1 Then
        fg.Col = fg.Col + 1
    Else
        If fg.Row <> fg.Rows - 1 Then
            fg.Row = fg.Row + 1
            fg.Col = 1
        End If
    End If

ElseIf KeyCode = 37 Then

    fg.Text = txtCell.Text
    txtCell.Visible = False
    
    If fg.Col <> 0 Then
        fg.Col = fg.Col - 1
    End If
    
ElseIf KeyCode = 39 Then
    fg.Text = txtCell.Text
    txtCell.Visible = False
    
    If fg.Col <> fg.Cols - 1 Then
        fg.Col = fg.Col + 1
    End If

ElseIf KeyCode = 38 Then
    fg.Text = txtCell.Text
    txtCell.Visible = False
    If fg.Row <> 1 Then
       fg.Row = fg.Row - 1
    End If
ElseIf KeyCode = 40 Then
    fg.Text = txtCell.Text
    txtCell.Visible = False
    If fg.Row <> fg.Rows - 1 Then
       fg.Row = fg.Row + 1
    End If

End If

End Sub

