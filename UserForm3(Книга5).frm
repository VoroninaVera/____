VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   8310.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10350
   OleObjectBlob   =   "UserForm3(Книга5).frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a() As Integer, val As Integer
Private Sub ListBox1_Click()

End Sub

Private Sub UserForm3_Initialize()
Dim i As Integer, j As Integer
val = 100
ReDim a(val) As Integer
End Sub

Private Sub CommandButton1_Click()
ListBox1.Clear
Dim i As Integer, j As Integer
ReDim a(val) As Integer
For i = 1 To val
a(i) = Rnd * 60 + 1
Next i
For i = 1 To val
    ListBox1.AddItem a(i)
Next i
End Sub


Private Sub OptionButton1_Click()
val = 10
End Sub

Private Sub OptionButton2_Click()
val = 20
End Sub

Private Sub OptionButton3_Click()
val = 30
End Sub

Private Sub CommandButton2_Click()
Dim sum As Integer
sum = 0
For i = 0 To val
sum = sum + a(i)
Next i
MsgBox ("Сумма всех элементов массива равна " & sum)
End Sub

Private Sub CommandButton3_Click()
Dim minus As Integer
count = 0
For i = 0 To val
    If a(i) Mod 2 = 0 Then
        count = count + 1
    End If
Next i
MsgBox ("Количество чётных элементов в массиве равно " & count)
End Sub
