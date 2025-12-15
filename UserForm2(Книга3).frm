VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   OleObjectBlob   =   "UserForm2(Книга3).frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ToggleButton1_Click()
    Image1.Left = Image1.Left + 100 ' Сдвиг вправо на 100 точек
End Sub

Private Sub ToggleButton2_Click()
    Image1.Top = Image1.Top + 100 ' Сдвиг вниз на 100 точек
End Sub
Private Sub ToggleButton3_Click()
    Randomize ' Инициализация генератора случайных чисел
    Image1.Left = Int(Rnd * (UserForm2.Width - Image1.Width)) ' Случайная X-координата
    Image1.Top = Int(Rnd * (UserForm2.Height - Image1.Height)) ' Случайная Y-координата
End Sub
Private Sub ToggleButton4_Click()

End Sub
Private Sub ToggleButton5_Click()
    Unload UserForm2 ' Закрывает форму
End Sub



