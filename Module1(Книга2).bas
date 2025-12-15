Attribute VB_Name = "Module1"
Option Explicit

Private Sub form_initialize()
    cmbStudent.AddItem "Иванов"
    cmbStudent.AddItem "Петров"
    cmbStudent.AddItem "Сидоров"
    cmbStudent.AddItem "Смирнов"
End Sub

Private Sub vibor_Click()
    Label1.Caption = "Выбран студент" & cmbStudent.Text
End Sub
