Attribute VB_Name = "modColorOnly"
Option Explicit
   
'---------------------------------------------------------------------------
'Form1 Code Example

'Private Sub Command1_Click()
'Dim sure As Long
'sure = ShowColor
'If sure = -1 Then Exit Sub
'Label1.BackColor = sure

'End Sub

'Private Sub Form_Load()
'Fillit
'Load_Color
'End Sub


'Private Sub Form_Terminate()
'Save_Color
'Unload Me
'End Sub
'-------------------------------------------------------------------------------
   Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
   
   Private Type CHOOSECOLOR
   lStructSize As Long
   hwndOwner As Long
   hInstance As Long
   rgbResult As Long
   lpCustColors As String
   flags As Long
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type

Public CustomColors() As Byte
Dim cc As CHOOSECOLOR

Public Function ShowColor() As Long
   
   'set the structure size
   cc.lStructSize = Len(cc)
   'Set the owner
   cc.hwndOwner = Form1.hWnd
   'set the application's instance
   cc.hInstance = App.hInstance
   'set the custom colors (converted to Unicode)
   cc.lpCustColors = StrConv(CustomColors, vbUnicode)
   'no extra flags
   cc.flags = 0  'set to 0 = define custom colors unselected. 2= define custom colors selected
   
   'Show the 'Select Color'-dialog
   If CHOOSECOLOR(cc) <> 0 Then
      ShowColor = (cc.rgbResult)
      CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
   Else
      ShowColor = -1
   End If
   
End Function

Public Sub Fillit()
   Dim i As Integer
   
   ReDim CustomColors(0 To 16 * 4 - 1) As Byte
  
   For i = LBound(CustomColors) To UBound(CustomColors)
      CustomColors(i) = 0
   Next i
End Sub

Public Sub Save_Color()
   Dim FileName As String
   Dim Free As Long

   FileName = App.Path & "\" & "ColorPal.txt"    'name of file to be saved to

   If FileName <> "" Then

      Free = FreeFile

      Open FileName For Binary As #Free
      Put #Free, , CustomColors   'this is the array to be saved
      Close #Free
   End If

End Sub

'----------------------------------------------------------------------------------------------------

'load array example

Public Sub Load_Color()
   Dim FileName As String
   Dim Y As Integer
   Dim X As Long
   
   FileName = App.Path & "\" & "ColorPal.txt"  'file where array is saved

   If FileName <> "" Then
      Dim Free As Long

      Free = FreeFile

      Open FileName For Binary As #Free
      Get #Free, , CustomColors   'the array to load
      Close #Free
   End If

   'put values in array
   For Y = 0 To 15
      X = CustomColors(Y)
   Next Y

 End Sub
