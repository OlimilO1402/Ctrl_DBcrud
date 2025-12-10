Attribute VB_Name = "MNew"
Option Explicit

Public Function CCollection(ByVal IsHashed As Boolean, Optional Col As Collection = Nothing) As CCollection
    Set CCollection = New CCollection: CCollection.New_ IsHashed, Col
End Function

Public Function ModalDialog(aDialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton) As ModalDialog
    Set ModalDialog = New ModalDialog: ModalDialog.New_ aDialog, BtnOK, BtnCancel
End Function

Public Function DBcrud(Col As CCollection, ByVal IsHashed As Boolean, ListBox, _
                       BtnAdd As CommandButton, Optional BtnAddClone, Optional BtnInsert, Optional BtnInsertClone, Optional BtnEdit, _
                       Optional BtnDelete, Optional BtnMoveUp, Optional BtnMoveDown, Optional BtnSortUp, Optional BtnSortDown, Optional BtnSearch) As DBcrud
    Set DBcrud = New DBcrud: DBcrud.New_ Col, ListBox, BtnAdd, BtnAddClone, BtnInsert, BtnInsertClone, BtnEdit, BtnDelete, BtnMoveUp, BtnMoveDown, BtnSortUp, BtnSortDown, BtnSearch
End Function

Public Function Person(ByVal Name As String, ByVal BirthDate As Date, ByVal BirthCity As String, ByVal EyeColor As ENaturalEyeColor, ByVal HairColor As ENaturalHairColor) As Person
    Set Person = New Person: Person.New_ Name, BirthDate, BirthCity, EyeColor, HairColor
End Function

Public Function PersonDefault() As Person
    Set PersonDefault = MNew.Person("Max Mustermann", "01.01.1980", "Musterhausen", ENaturalEyeColor.Brown, ENaturalHairColor.Gray)
End Function

' v ' ############################## ' v '    ENaturalEyeColor     ' v ' ############################## ' v '
Public Function ENaturalEyeColor_ToStr(ByVal e As ENaturalEyeColor) As String
    Dim s As String
    Select Case e
    Case ENaturalEyeColor.Gray:       s = "Gray"
    Case ENaturalEyeColor.GreenGray:  s = "GreenGray"
    Case ENaturalEyeColor.Green:      s = "Green"
    Case ENaturalEyeColor.GreenBrown: s = "GreenBrown"
    Case ENaturalEyeColor.Brown:      s = "Brown"
    Case ENaturalEyeColor.Blue:       s = "Blue"
    Case ENaturalEyeColor.GrayBlue:   s = "GrayBlue"
    Case ENaturalEyeColor.BlueGreen:  s = "BlueGreen"
    End Select
    ENaturalEyeColor_ToStr = s
End Function

Public Function ENaturalEyeColor_Parse(ByVal s As String) As ENaturalEyeColor
    Dim e As ENaturalEyeColor
    Select Case s
    Case "Gray":       e = ENaturalEyeColor.Gray
    Case "GreenGray":  e = ENaturalEyeColor.GreenGray
    Case "Green":      e = ENaturalEyeColor.Green
    Case "GreenBrown": e = ENaturalEyeColor.GreenBrown
    Case "Brown":      e = ENaturalEyeColor.Brown
    Case "Blue":       e = ENaturalEyeColor.Blue
    Case "GrayBlue":   e = ENaturalEyeColor.GrayBlue
    Case "BlueGreen":  e = ENaturalEyeColor.BlueGreen
    End Select
    ENaturalEyeColor_Parse = e
End Function

Public Sub ENaturalEyeColor_ToLB(ListBoxOrComboBox)
    Dim s As String, i As Long
    For i = 1 To 10
        s = ENaturalEyeColor_ToStr(2 ^ i)
        If Len(s) Then ListBoxOrComboBox.AddItem s
    Next
End Sub

Public Function ENaturalEyeColor_ToIndex(e As ENaturalEyeColor) As Long
    ENaturalEyeColor_ToIndex = BitToIndex(e)
End Function
' ^ ' ############################## ' ^ '    ENaturalEyeColor    ' ^ ' ############################## ' ^ '

' v ' ############################## ' v '    ENaturalHairColor    ' v ' ############################## ' v '
Public Function ENaturalHairColor_ToStr(e As ENaturalHairColor) As String
    Dim s As String
    Select Case e
    Case ENaturalHairColor.White:     s = "White"
    Case ENaturalHairColor.Gray:      s = "Gray"
    Case ENaturalHairColor.Black:     s = "Black"
    Case ENaturalHairColor.Brown:     s = "Brown"
    Case ENaturalHairColor.Brunett:   s = "Brunett"
    Case ENaturalHairColor.Blond:     s = "Blond"
    Case ENaturalHairColor.Gingerred: s = "Gingerred"
    Case ENaturalHairColor.RedBlond:  s = "RedBlond"
    End Select
    ENaturalHairColor_ToStr = s
End Function

Public Function ENaturalHairColor_Parse(ByVal s As String) As ENaturalHairColor
    Dim e As ENaturalHairColor
    Select Case s
    Case "White":     e = ENaturalHairColor.White
    Case "Gray":      e = ENaturalHairColor.Gray
    Case "Black":     e = ENaturalHairColor.Black
    Case "Brown":     e = ENaturalHairColor.Brown
    Case "Brunett":   e = ENaturalHairColor.Brunett
    Case "Blond":     e = ENaturalHairColor.Blond
    Case "Gingerred": e = ENaturalHairColor.Gingerred
    Case "RedBlond":  e = ENaturalHairColor.RedBlond
    End Select
    ENaturalHairColor_Parse = e
End Function

Public Function ENaturalHairColor_ToIndex(e As ENaturalHairColor) As Long
    ENaturalHairColor_ToIndex = BitToIndex(e)
End Function

Public Sub ENaturalHairColor_ToLB(ListBoxOrComboBox)
    Dim s As String, i As Long
    For i = 1 To 10
        s = ENaturalHairColor_ToStr(2 ^ i)
        If Len(s) Then ListBoxOrComboBox.AddItem s
    Next
End Sub
' ^ ' ############################## ' ^ '    ENaturalHairColor    ' ^ ' ############################## ' ^ '

Public Function BitToIndex(ByVal e As Long) As Long
    For BitToIndex = 0 To 30
        If e = 2 ^ (BitToIndex + 1) Then Exit Function
    Next
End Function
