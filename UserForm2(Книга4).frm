VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7980
   OleObjectBlob   =   "UserForm2(Книга4).frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim text As String
    Dim length As Integer
    text = TextBox1.text
    length = Len(text) ' функция LEN возвращает длину строки
    TextBox2.text = "Количество букв в предложении = " & length
End Sub
Private Sub CommandButton2_Click()
    Dim text As String
    Dim count As Integer
    text = TextBox1.text
    count = UBound(Split(text, "А")) - LBound(Split(text, "А"))
    TextBox2.text = "Количество букв 'А' в предложении = " & count
End Sub
Private Sub CommandButton3_Click()
    Dim text As String
    Dim words() As String
    text = TextBox1.text
    words = Split(text, " ") ' разбиваем строку по пробелам
    TextBox2.text = "Количество слов в предложении = " & UBound(words) + 1
End Sub
Private Sub CommandButton4_Click()
    Dim text As String
    text = TextBox1.text
    TextBox2.text = UCase(text) ' UCase преобразует в верхний регистр
End Sub
Private Sub CommandButton5_Click()
    Dim text As String
    Dim reversedText As String
    text = TextBox1.text
    For i = Len(text) To 1 Step -1
        reversedText = reversedText & Mid(text, i, 1)
    Next i
    TextBox2.text = reversedText
End Sub
Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub
