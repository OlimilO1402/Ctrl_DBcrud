VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ENaturalHairColor
    White = &H1
    Gray = &H2
    Black = &H4
    Brown = &H8
    Brunett = &H10
    Blond = &H20
    Gingerred = &H40
    RedBlond = &H80
End Enum

Public Enum ENaturalEyeColor
    Gray = &H1
    GreenGray = &H2
    Green = &H4
    GreenBrown = &H8
    Brown = &H10
    Blue = &H20
    GrayBlue = &H40
    BlueGreen = &H80
End Enum

Private m_Name      As String
Private m_BirthD    As Date
Private m_BirthCity As String
Private m_EyeColor  As ENaturalEyeColor  'does normally stay the same throughout life
Private m_HairColor As ENaturalHairColor 'could change over the years

Friend Sub New_(Name As String, BirthDate As Date, BirthCity As String, EyeColor As ENaturalEyeColor, HairColor As ENaturalHairColor)
    m_Name = Name: m_BirthD = BirthDate: m_BirthCity = BirthCity: m_EyeColor = EyeColor: m_HairColor = HairColor
End Sub

Public Sub NewC(other As Person)
    With other: New_ .Name, .BirthDay, .BirthCity, .EyeColor, .HairColor: End With
End Sub
'
'Public Sub NewV(other)
'    Dim Obj As Person: Set Obj = other: NewC Obj
'End Sub

Public Function Clone() As Person
    Set Clone = New Person: Clone.NewC Me
End Function

Public Function Compare(other As Person) As Long
    Dim c1 As String: c1 = m_Name & " " & Date_Format(m_BirthD) & " " & m_BirthCity
    Dim c2 As String: c2 = other.Name & " " & Date_Format(other.BirthDay) & " " & other.BirthCity
    Compare = StrComp(c1, c2, VbCompareMethod.vbTextCompare)
End Function

Private Function Date_Format(d As Date) As String
    Date_Format = Year(d) & "-" & Month(d) & "-" & Day(d)
End Function
    
Public Sub SetParams(Name, BirthDay, BirthCity, NaturalEyeColor, NaturalHairColor)
    If Not IsMissing(Name) Then m_Name = Name
    If Not IsMissing(BirthDay) Then m_BirthD = BirthDay
    If Not IsMissing(BirthCity) Then m_BirthCity = BirthCity
    If Not IsMissing(NaturalEyeColor) Then m_EyeColor = NaturalEyeColor
    If Not IsMissing(NaturalHairColor) Then m_HairColor = NaturalHairColor
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get BirthDay() As Date
    BirthDay = m_BirthD
End Property

Public Property Get BirthCity() As String
    BirthCity = m_BirthCity
End Property

Public Property Get EyeColor() As ENaturalEyeColor
    EyeColor = m_EyeColor
End Property

Public Property Get HairColor() As ENaturalHairColor
    HairColor = m_HairColor
End Property

Public Function Ptr()
    Ptr = ObjPtr(Me)
End Function

Public Property Get Key() As String
    Key = Me.Ptr ' Me.ToStr
End Property

Public Function ToStr() As String
    ToStr = m_Name & " " & Me.BirthDay & " " & Me.BirthCity
End Function
