VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1(Книга2).frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
Label2.Caption = "Выбран студент: " & ComboBox1.Text
End Sub

Private Sub UserForm_Initialize()
    ComboBox1.AddItem "Иванов"
    ComboBox1.AddItem "Петров"
    ComboBox1.AddItem "Сидоров"
    ComboBox1.AddItem "Смирнов"
End Sub


