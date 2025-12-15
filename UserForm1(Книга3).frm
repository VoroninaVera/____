VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5925
   OleObjectBlob   =   "UserForm1(Книга3).frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub ToggleButton1_Click()
    Dim a As Double, b As Double, sum As Double
    a = Val(TextBox1.Value)
    b = Val(TextBox2.Value)
    sum = a + b
    MsgBox "Сумма чисел: " & sum, vbInformation, "Результат"
End Sub

Private Sub ToggleButton2_Click()
    If Val(TextBox1.Value) = Val(TextBox2.Value) Then
        MsgBox "Числа равны", vbInformation
    Else
        MsgBox "Числа не равны", vbExclamation
    End If
End Sub

Private Sub ToggleButton3_Click()
    If Val(TextBox1.Value) > Val(TextBox2.Value) Then
        MsgBox "Большее число: " & TextBox1.Value, vbInformation
    Else
        MsgBox "Большее число: " & TextBox2.Value, vbInformation
    End If
End Sub

Private Sub ToggleButton4_Click()
    Dim a As Double, b As Double, result As Double
    a = Val(TextBox1.Value)
    b = Val(TextBox2.Value)
    result = a - b ' или a * b, a / b и т.д.
    MsgBox "Результат: " & result, vbInformation
End Sub
