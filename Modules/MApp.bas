Attribute VB_Name = "MApp"
Option Explicit
Private m_Persons As CCollection '(Of Person)

Sub Main()
    Set m_Persons = MNew.CCollection(True)
    FMain.ShowDialog m_Persons
End Sub
