VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   Caption         =   "Calli-Art"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   DrawWidth       =   5
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Pen Values"
      ForeColor       =   &H80000008&
      Height          =   2820
      Left            =   6720
      TabIndex        =   31
      Top             =   3015
      Width           =   3105
      Begin Project1.ccXPButton cmdDelete 
         Height          =   390
         Left            =   1935
         TabIndex        =   41
         Top             =   1860
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   688
         Caption         =   "Delete"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin Project1.ccXPButton cmdAdd 
         Height          =   390
         Left            =   90
         TabIndex        =   40
         Top             =   1860
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   688
         Caption         =   "Add"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin VB.ListBox lst1 
         Height          =   1620
         Left            =   45
         TabIndex        =   32
         Top             =   195
         Width           =   3015
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Click and drag to re-arrange list"
         Height          =   285
         Left            =   420
         TabIndex        =   34
         Top             =   2535
         Width           =   2340
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Dbl click selection to apply values"
         Height          =   255
         Left            =   420
         TabIndex        =   33
         Top             =   2280
         Width           =   2430
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Control Panel"
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   1560
      TabIndex        =   10
      Top             =   3015
      Width           =   5085
      Begin Project1.ccXPButton cmdSaveBitmap 
         Height          =   315
         Left            =   3525
         TabIndex        =   39
         Top             =   2130
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         Caption         =   "Save as Bitmap"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin Project1.ccXPButton cmdSaveJpeg 
         Height          =   360
         Left            =   3525
         TabIndex        =   38
         Top             =   1650
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   635
         Caption         =   "Save as Jpeg"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin Project1.ccXPButton cmdClear 
         Height          =   465
         Left            =   3690
         TabIndex        =   37
         Top             =   1035
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   820
         Caption         =   "Clear"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin Project1.ccXPButton cmdUndo 
         Height          =   645
         Left            =   3690
         TabIndex        =   36
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1138
         Caption         =   "Undo Last"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin VB.HScrollBar HS5 
         Height          =   225
         Left            =   1020
         Max             =   10
         Min             =   -10
         TabIndex        =   28
         Top             =   1095
         Value           =   -3
         Width           =   1320
      End
      Begin VB.CommandButton cmdDef 
         Caption         =   "Default"
         Height          =   330
         Left            =   2805
         TabIndex        =   27
         Top             =   1095
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1290
         TabIndex        =   14
         Top             =   2145
         Width           =   2145
      End
      Begin VB.HScrollBar HS1 
         Height          =   210
         Left            =   1020
         Max             =   10
         Min             =   1
         TabIndex        =   13
         Top             =   270
         Value           =   3
         Width           =   2085
      End
      Begin VB.HScrollBar HS2 
         Height          =   225
         Left            =   1020
         Max             =   10
         Min             =   -10
         TabIndex        =   12
         Top             =   540
         Value           =   -3
         Width           =   2085
      End
      Begin VB.HScrollBar HS3 
         Height          =   210
         Left            =   1020
         Max             =   10
         Min             =   -10
         TabIndex        =   11
         Top             =   825
         Value           =   -3
         Width           =   2070
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2370
         TabIndex        =   30
         Top             =   1095
         Width           =   315
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Angle"
         Height          =   255
         Left            =   525
         TabIndex        =   29
         Top             =   1095
         Width           =   465
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   25
         Top             =   1725
         Width           =   885
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "FileName"
         Height          =   225
         Left            =   570
         TabIndex        =   24
         Top             =   2175
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Draw Color"
         Height          =   225
         Left            =   285
         TabIndex        =   23
         Top             =   1440
         Width           =   870
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   22
         Top             =   1395
         Width           =   885
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Shadow Color"
         Height          =   210
         Left            =   75
         TabIndex        =   21
         Top             =   1740
         Width           =   1035
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3120
         TabIndex        =   20
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3120
         TabIndex        =   19
         Top             =   540
         Width           =   390
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3120
         TabIndex        =   18
         Top             =   825
         Width           =   390
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Draw Width"
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   270
         Width           =   870
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Y Position"
         Height          =   225
         Left            =   195
         TabIndex        =   16
         Top             =   795
         Width           =   780
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "X Position"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   555
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Colors"
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   90
      TabIndex        =   1
      Top             =   3015
      Width           =   1365
      Begin Project1.ccXPButton cmdShowColor 
         Height          =   645
         Left            =   60
         TabIndex        =   35
         Top             =   225
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1138
         Caption         =   "Select Color"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin VB.OptionButton optColorSel 
         BackColor       =   &H00008000&
         Caption         =   "Shadow"
         Height          =   210
         Index           =   2
         Left            =   75
         TabIndex        =   9
         Top             =   1260
         Width           =   900
      End
      Begin VB.OptionButton optColorSel 
         BackColor       =   &H00008000&
         Caption         =   "Background"
         Height          =   210
         Index           =   1
         Left            =   75
         TabIndex        =   8
         Top             =   1500
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton optColorSel 
         BackColor       =   &H00008000&
         Caption         =   "Font"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   7
         Top             =   1005
         Width           =   630
      End
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawWidth       =   4
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   30
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   636
      TabIndex        =   0
      Top             =   15
      Width           =   9570
   End
   Begin VB.PictureBox picundo 
      BackColor       =   &H00FFFFFF&
      Height          =   1140
      Index           =   0
      Left            =   30
      ScaleHeight     =   1080
      ScaleWidth      =   1230
      TabIndex        =   3
      Top             =   30
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.PictureBox picundo 
      BackColor       =   &H00FFFFFF&
      Height          =   1080
      Index           =   1
      Left            =   1440
      ScaleHeight     =   1020
      ScaleWidth      =   1185
      TabIndex        =   4
      Top             =   45
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.PictureBox picundo 
      BackColor       =   &H00FFFFFF&
      Height          =   1065
      Index           =   2
      Left            =   2880
      ScaleHeight     =   1005
      ScaleWidth      =   1200
      TabIndex        =   5
      Top             =   75
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.PictureBox picundo 
      BackColor       =   &H00FFFFFF&
      Height          =   1155
      Index           =   3
      Left            =   4260
      ScaleHeight     =   1095
      ScaleWidth      =   1020
      TabIndex        =   6
      Top             =   60
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Note:Select Back ground Color first, if different than default."
      ForeColor       =   &H0000FFFF&
      Height          =   885
      Left            =   105
      TabIndex        =   26
      Top             =   4815
      Width           =   1350
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   9420
      MousePointer    =   15  'Size All
      TabIndex        =   2
      ToolTipText     =   "Resize Canvas"
      Top             =   1410
      Width           =   150
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fontColor As Long
Dim undoCt As Integer   'undo counter
Dim sColor As Long        'shadow color
Dim bColor As Long        'background color
Dim ary1() As String   'stores line points
Dim aryValues() As String  'stores slider values
Dim aStatic1 As String  'temp storage
Private sCurrentLine As String
Private iFileNumber As Integer

'used for sorting listbox
Dim thing1 As String
Dim thing2 As String
Dim ind As Integer

Private Sub Form_Load()
   fontColor = vbRed
   undoCt = -1
   Fillit
   Load_Color
   Label7.Caption = HS1.Value
   Label8.Caption = HS2.Value
   Label9.Caption = HS3.Value
   Label18.Caption = HS5.Value
   Label12.ForeColor = Label5.BackColor
   Label14.ForeColor = Label5.BackColor
   LoadList
End Sub

Private Sub Form_Resize()
   If Form1.Height < 6420 Then Form1.Height = 6420
   If Form1.Width < 8565 Then Form1.Width = 8565
End Sub

Private Sub Form_Terminate()
   Save_Color
   Unload Me
End Sub

Private Sub cmdAdd_Click()
   lst1.AddItem HS1.Value & "," & HS2.Value & "," & HS3.Value & "," & HS5.Value & "," & Label3.BackColor & "," & Label5.BackColor
   SaveList
End Sub

Private Sub cmdDef_Click()
   HS1.Value = 3
   HS2.Value = -3
   HS3.Value = -3
   HS5.Value = -3
   Label3.BackColor = &H0&
   Label5.BackColor = &HFF&
   sColor = Label3.BackColor
   fontColor = Label5.BackColor
End Sub

Private Sub cmdDelete_Click()
   Dim ListCount As Long
   
   ListCount& = lst1.ListCount
   Do While ListCount& > 0&
      ListCount& = ListCount& - 1
      If lst1.Selected(ListCount&) = True Then
         lst1.RemoveItem (ListCount&)
      End If
      Loop
      SaveList
   End Sub

Private Sub LoadList()
   
   iFileNumber = FreeFile
   
   lst1.Clear
   Open App.Path & "\SavedValues.txt" For Input As #iFileNumber
   While Not EOF(iFileNumber)
      Line Input #iFileNumber, sCurrentLine
      lst1.AddItem sCurrentLine
      Wend
      Close #iFileNumber
   End Sub

Private Sub SaveList()
   Dim i As Integer
   
   iFileNumber = FreeFile
   
   Open App.Path & "\SavedValues.txt" For Output As #iFileNumber
   For i = 0 To lst1.ListCount - 1
      Print #iFileNumber, lst1.List(i)
   Next i
   Close #iFileNumber
   
End Sub

Private Sub HS1_Change()
   HS1_Scroll
End Sub

Private Sub HS1_Scroll()
   Label7.Caption = HS1.Value
End Sub

Private Sub HS2_Change()
   HS2_Scroll
End Sub

Private Sub HS2_Scroll()
   Label8.Caption = HS2.Value
End Sub

Private Sub HS3_Change()
   HS3_Scroll
End Sub

Private Sub HS3_Scroll()
   Label9.Caption = HS3.Value
End Sub

Private Sub HS5_Change()
   HS5_Scroll
End Sub

Private Sub HS5_Scroll()
   Label18.Caption = HS5.Value
End Sub

Private Sub lst1_DblClick()
   aryValues = Split(lst1.Text, ",")
   HS1.Value = aryValues(0)
   HS2.Value = aryValues(1)
   HS3.Value = aryValues(2)
   HS5.Value = aryValues(3)
   Label3.BackColor = aryValues(4)
   Label5.BackColor = aryValues(5)
   sColor = Label3.BackColor
   fontColor = Label5.BackColor
   If Label5.BackColor = &H8000& Then
      Label12.ForeColor = vbBlack
      Label14.ForeColor = vbBlack
   Else
      Label12.ForeColor = Label5.BackColor
      Label14.ForeColor = Label5.BackColor
   End If
End Sub

Private Sub lst1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 1 Then 'left mousebutton is down
   thing1 = lst1.Text
   ind = lst1.ListIndex 'the index is Set
End If
End Sub

Private Sub lst1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   
   If thing1 = lst1.Text Then Exit Sub
   If thing1 = "" Then Exit Sub
   If Button = 1 Then
      thing2 = lst1.Text
      lst1.List(ind) = thing2
      ind = lst1.ListIndex
      lst1.List(ind) = thing1
   End If
End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   undoCt = undoCt + 1
   If undoCt > 3 Then   'shift undo pictureboxes to the left
      picundo(0).Picture = picundo(1).Picture
      picundo(1).Picture = picundo(2).Picture
      picundo(2).Picture = picundo(3).Picture
      undoCt = 3
   End If
   picCanvas.Picture = picCanvas.Image   'render picture
   picundo(undoCt).Picture = picCanvas.Picture
   cmdUndo.Caption = "Undo Last " & undoCt + 1
   picCanvas.ForeColor = sColor
   picCanvas.MousePointer = 99
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   picCanvas.CurrentX = x
   picCanvas.CurrentY = y
   'draw shadow
   If Button = 1 Then
      picCanvas.ForeColor = sColor
      picCanvas.Line -(x + HS1.Value, y + HS5.Value)
      aStatic1 = aStatic1 & picCanvas.CurrentX & "," & picCanvas.CurrentY & "," & x & "," & y & ","
   End If
End Sub

Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim XX As Integer
   ary1 = Split(aStatic1, ",")
   'draw outline -- this can be improved and expanded
   For XX = 0 To UBound(ary1) - 3 Step 4
      picCanvas.Line (ary1(XX) + HS2.Value, ary1(XX + 1) - 1 + HS3.Value)-(ary1(XX + 2) + HS2.Value, ary1(XX + 3) + HS3.Value), sColor
      picCanvas.Line (ary1(XX) - 1 + HS2.Value, ary1(XX + 1) - 1 + HS3.Value)-(ary1(XX + 2) - 1 + HS2.Value, ary1(XX + 3) + HS3.Value), sColor
      picCanvas.Line (ary1(XX) + HS2.Value, ary1(XX + 1) + 1 + HS3.Value)-(ary1(XX + 2) + 1 + HS2.Value, ary1(XX + 3) + HS3.Value), sColor
   Next XX
   'draw main
   For XX = 0 To UBound(ary1) - 3 Step 4
      picCanvas.ForeColor = fontColor
      picCanvas.Line (ary1(XX) + HS2.Value, ary1(XX + 1) + HS3.Value)-(ary1(XX + 2) + HS2.Value, ary1(XX + 3) + HS3.Value)
   Next XX
   aStatic1 = ""
End Sub

Private Sub cmdClear_Click()
   Dim x As Integer
   
   picCanvas.Cls
   picCanvas.Picture = LoadPicture()
   For x = 0 To 3
      picundo(x).Picture = LoadPicture()
   Next x
   undoCt = -1
   cmdUndo.Caption = "Undo Last 0"
End Sub

Private Sub cmdSaveBitmap_Click()
   Dim rsp As String
   If Text1.Text = "" Then
      MsgBox "Enter a Filename", , "No filename"
      Exit Sub
   End If
   picCanvas.Picture = picCanvas.Image  'render picture
   'check if file exists already
   If Dir(App.Path & "\" & Text1.Text & ".bmp") = "" Then
      SavePicture picCanvas.Picture, App.Path & "\" & Text1.Text & ".bmp"
      MsgBox "Picture saved at " & App.Path & "\" & Text1.Text & ".bmp", , "Save a Bitmap"
   Else
      rsp = MsgBox("File exists. Do you want to overwrite?", vbYesNo)
      If rsp = vbNo Then GoTo here
      SavePicture picCanvas.Picture, App.Path & "\" & Text1.Text & ".bmp"
      MsgBox "Picture saved at " & App.Path & "\" & Text1.Text & ".bmp", , "Picture Saved"
here:
   End If
   
   Text1.Text = ""
End Sub

Private Sub cmdSaveJPEG_Click()
   Dim rsp As String
   If Text1.Text = "" Then
      MsgBox "Enter a Filename", , "No filename"
      Exit Sub
   End If
   picCanvas.Picture = picCanvas.Image  'render picture
   'check if file exists already
   If Dir(App.Path & "\" & Text1.Text & ".jpg") = "" Then
      SaveJPEG App.Path & "\" & Text1.Text & ".jpg", picCanvas, Me, True, 90
      MsgBox "Picture saved at " & App.Path & "\" & Text1.Text & ".jpg", , "Save as jpg"
   Else
      rsp = MsgBox("File exists. Do you want to overwrite?", vbYesNo)
      If rsp = vbNo Then GoTo here
      SaveJPEG App.Path & "\" & Text1.Text & ".jpg", picCanvas, Me, True, 90
      MsgBox "Picture saved at " & App.Path & "\" & Text1.Text & ".jpg", , "Picture Saved"
here:
   End If
   Text1.Text = ""
End Sub

Private Sub cmdShowColor_Click()
   Dim sure As Long
   sure = ShowColor
   If sure = -1 Then Exit Sub
   If optColorSel(0).Value = True Then
      fontColor = sure
      Label5.BackColor = sure
      Label12.ForeColor = Label5.BackColor
      Label14.ForeColor = Label5.BackColor
   End If
   If optColorSel(1).Value = True Then
      picCanvas.BackColor = sure
   End If
   If optColorSel(2).Value = True Then
      sColor = sure
      Label3.BackColor = sure
   End If
End Sub

Private Sub cmdUndo_Click()
   picCanvas.Cls
   If cmdUndo.Caption = "Undo Last 0" Then Exit Sub   'no undo's left
   If undoCt < 0 Then undoCt = 3
   picCanvas.Picture = picundo(undoCt).Picture
   undoCt = undoCt - 1   'countdown undo counter
   picundo(undoCt + 1).Picture = LoadPicture()   'clear picturebox
   cmdUndo.Caption = "Undo Last " & undoCt + 1   'show undo count
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 1 Then  'resize main picturebox...click and drag
   Label1.Left = Label1.Left + Label1.Width / 12 + x - 95
   picCanvas.Width = picCanvas.Left + Label1.Left + 105
   Label1.Top = Label1.Top + y / 2 - 35
   picCanvas.Height = picCanvas.Top + Label1.Top - 100
   If picCanvas.Height > 2730 Then
      picCanvas.Height = 2730
      Label1.Top = picCanvas.Height + 100
   End If
   picCanvas.Picture = picCanvas.Image
End If
End Sub

Private Function SaveJPEG(ByVal Filename As String, Pic As PictureBox, PForm As Form, Optional ByVal Overwrite As Boolean = True, Optional ByVal Quality As Byte = 90) As Boolean
   Dim JPEGclass As cJpeg
   Dim m_Picture As IPictureDisp
   Dim m_DC As Long
   Dim m_Millimeter As Single
   m_Millimeter = PForm.ScaleX(100, vbPixels, vbMillimeters)
   Set m_Picture = Pic
   m_DC = Pic.hdc
   'this is not my code....from PSC
   'initialize class
   Set JPEGclass = New cJpeg
   'check there is image to save and the filename string is not empty
   If m_DC <> 0 And LenB(Filename) > 0 Then
      'check for valid quality
      If Quality < 1 Then Quality = 1
      If Quality > 100 Then Quality = 100
      'set quality
      JPEGclass.Quality = Quality
      'save in full color
      JPEGclass.SetSamplingFrequencies 1, 1, 1, 1, 1, 1
      'copy image from hDC
      If JPEGclass.SampleHDC(m_DC, CLng(m_Picture.Width / m_Millimeter), CLng(m_Picture.Height / m_Millimeter)) = 0 Then
         'if overwrite is set and file exists, delete the file
         If Overwrite And LenB(Dir$(Filename)) > 0 Then Kill Filename
         'save file and return True if success
         SaveJPEG = JPEGclass.SaveFile(Filename) = 0
      End If
   End If
   'clear memory
   Set JPEGclass = Nothing
End Function
