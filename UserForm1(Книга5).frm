VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8370.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9945.001
   OleObjectBlob   =   "UserForm1(Книга5).frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private A(19) As Double ' Массив из 20 элементов (индексы 0–19)
Private MinVal As Double, MaxVal As Double ' Границы диапазона
Private Sub CommandButton1_Click()
    ' Проверяем, заполнены ли поля ввода
    If TextBox1.Text = "" Or TextBox2.Text = "" Then
        MsgBox "Введите диапазон случайных чисел!", vbExclamation
        Exit Sub
    End If
    
    ' Считываем границы диапазона
    MinVal = CDbl(TextBox1.Text)
    MaxVal = CDbl(TextBox2.Text)
    
    ' Проверяем корректность диапазона
    If MinVal >= MaxVal Then
        MsgBox "Минимальное значение должно быть меньше максимального!", vbExclamation
        Exit Sub
    End If
    
    ' Инициализируем генератор случайных чисел
    Randomize
    
    ' Заполняем массив случайными числами
    Dim i As Integer
    For i = 0 To 19
        A(i) = Rnd() * (MaxVal - MinVal) + MinVal
    Next i
    
    MsgBox "Массив заполнен!", vbInformation
End Sub
Private Sub CommandButton2_Click()
    Dim OutputText As String
    Dim i As Integer
    
    ' Очищаем предыдущее содержимое
    TextBox3.Text = ""
    
    ' Формируем строку для вывода
    For i = 0 To 19
        OutputText = OutputText & Format(A(i), "0.000") & vbCrLf
    Next i
    
    ' Выводим результат в TextBox3 (область «Массив A(20)»)
    TextBox3.Text = OutputText
End Sub
Private Sub TextBox1_Change()
    If Not IsNumeric(TextBox1.Text) Then
        TextBox1.Text = "" ' Очищаем некорректный ввод
        MsgBox "Введите числовое значение!", vbExclamation
    End If
End Sub

Private Sub TextBox2_Change()
    If Not IsNumeric(TextBox2.Text) Then
        TextBox2.Text = "" ' Очищаем некорректный ввод
        MsgBox "Введите числовое значение!", vbExclamation
    End If
End Sub

