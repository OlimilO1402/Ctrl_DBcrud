VERSION 5.00
Begin VB.Form FPerson 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Edit Person"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FPerson.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox Combo2 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   2040
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox TxtBirthCity 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox TxtBirthday 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox TxtName 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Hair Color"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Eye Color"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Birth City"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Birthdate"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   525
   End
End
Attribute VB_Name = "FPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Person As Person

Private Sub Form_Load()
    MNew.ENaturalEyeColor_ToLB Me.Combo1
    MNew.ENaturalHairColor_ToLB Me.Combo2
End Sub

Public Sub UpdateView(Obj)
    Set m_Person = Obj
    If m_Person Is Nothing Then MsgBox "The Person does not exist": Exit Sub
    TxtName.Text = m_Person.Name
    TxtBirthday.Text = m_Person.BirthDay
    TxtBirthCity.Text = m_Person.BirthCity
    Combo1.ListIndex = MNew.ENaturalEyeColor_ToIndex(m_Person.EyeColor)
    Combo2.ListIndex = MNew.ENaturalHairColor_ToIndex(m_Person.HairColor)
End Sub

Public Function UpdateData(Obj) As Boolean
Try: On Error GoTo Catch
    
    'Dim bIsOK As Boolean
    'Dim bd As Date:     bd = m_Person.BirthDay: TxtBirthday.Text = MString.Date_TryParseValidate(TxtBirthday.Text, "Birthday", "", bIsOK, bd): If Not bIsOK Then Exit Function
    
    Dim nm As String:   nm = TxtName.Text
    Dim bd As Date:     bd = TxtBirthday.Text ' m_Person.BirthDay: TxtBirthday.Text = MString.Date_TryParseValidate(TxtBirthday.Text, "Birthday", "", bIsOK, bd): If Not bIsOK Then Exit Function
    Dim ct As String:   ct = TxtBirthCity.Text
    Dim ec As ENaturalEyeColor:  ec = MNew.ENaturalEyeColor_Parse(Combo1.Text)
    Dim hc As ENaturalHairColor: hc = MNew.ENaturalHairColor_Parse(Combo2.Text)
    m_Person.SetParams nm, bd, ct, ec, hc
    'Set Obj = m_Person
    UpdateData = True
Catch:
End Function

