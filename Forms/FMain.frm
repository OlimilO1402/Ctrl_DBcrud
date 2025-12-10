VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox LstPersons 
      Height          =   3885
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   4815
   End
   Begin VB.TextBox TxtPersonSearch 
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton BtnPersonSearch 
      Caption         =   "°\"
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      ToolTipText     =   "Search"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton BtnPersonSortDwn 
      Caption         =   "|v|"
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      ToolTipText     =   "Sort Down"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton BtnPersonSortUp 
      Caption         =   "|^|"
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      ToolTipText     =   "Sort Up"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton BtnPersonMoveDown 
      Caption         =   "v"
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      ToolTipText     =   "Move Down"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton BtnPersonMoveUp 
      Caption         =   "^"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      ToolTipText     =   "Move Up"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton BtnPersonDelete 
      Caption         =   "-"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      ToolTipText     =   "Delete"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton BtnPersonEdit 
      Caption         =   "/"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "Edit"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton BtnPersonInsertClone 
      Caption         =   "i++"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      ToolTipText     =   "Insert Clone"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton BtnPersonInsert 
      Caption         =   "i"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Insert"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton BtnPersonAddClone 
      Caption         =   "++"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   "Add Clone"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton BtnPersonAdd 
      Caption         =   "+"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Add"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'what exactly means "crud"
'crud is an acronym and it stands for
'c = Create
'r = Read
'u = Update
'd = delete
'
'this are obiously commands in an DB-application-ui
'in fact it can mean much more like:
'Alle erdenklichen Datenbank-CRUD-Funktionen (Create, Read, Update, Delete)
' 1. [ + ]  Hinzufügen,  ein neues       Objekt Erstellen, Editieren, und am Ende der Liste anfügen
' 2. [ ++]  Klonen,      das ausgewählte Objekt Kopieren , Editieren, und am Ende der Liste anfügen
' 3. [ i ]  Einfügen,    ein neues       Objekt Erstellen, Editieren, und oberhalb dem aktuell ausgewählten Objekt einfügen
' 4. [i++]  Kop+Einf,    das ausgewählte Objekt Kopieren , Editieren, und oberhalb dem aktuell ausgewählten Objekt einfügen
' 5. [ / ]  Edit,        das ausgewählte Objekt Editieren, einen Modalen Dialog mit den Eigenschaften des Objekts anzeigen, OK oder Abbrechen
' 6. [ - ]  Löschen,     das ausgewählte Objekt Löschen vorher nachfragen
' 7. [ ^ ]  Nach Oben,   das ausgewählte Objekt um eine Stelle nach oben  schieben
' 8. [ v ]  Nach Unten,  das ausgewählte Objekt um eine Stelle nach unten schieben
' 9. [ |v]  Aufsteigend, alle Objekte aufsteigend Sortieren
'10. [ |^]  Absteigend,  alle Objekte absteigend  Sortieren
'11. [°\ ]  Suchen,      ein Objekt suchen, einen Suchdialog anzeigen, Ergebnisliste anzeigen
'
Private WithEvents PersonCRUD As DBcrud
Attribute PersonCRUD.VB_VarHelpID = -1

Private Sub Form_Load()
    Set PersonCRUD = MNew.DBcrud(MApp.Persons, False, Me.LstPersons, Me.BtnPersonAdd, Me.BtnPersonAddClone, Me.BtnPersonInsert, Me.BtnPersonInsertClone, Me.BtnPersonEdit, Me.BtnPersonDelete, Me.BtnPersonMoveUp, Me.BtnPersonMoveDown, Me.BtnPersonSortUp, Me.BtnPersonSortDwn, Me.BtnPersonSearch)
End Sub

Private Sub LstPersons_Click()
    Dim li As Long: li = LstPersons.ListIndex
    Label1.Caption = LstPersons.ItemData(li)
End Sub

Private Sub PersonCRUD_OnEdit(Obj_inout As Object)
    Dim p As Person: Set p = IIf(Obj_inout Is Nothing, MNew.PersonDefault, Obj_inout)
    If MNew.ModalDialog(FPerson, FPerson.BtnOK, FPerson.BtnCancel).ShowDialog(p, Me) = vbCancel Then
        Set Obj_inout = Nothing
        Exit Sub
    End If
    Set Obj_inout = p
End Sub
