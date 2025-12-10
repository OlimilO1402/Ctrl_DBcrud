Attribute VB_Name = "MCol"
Option Explicit

' v ' ############################## ' v '    ListBox Functions    ' v ' ############################## ' v '
Public Sub ListBox_Add(aLB As ListBox, ByVal Object As Object)
    aLB.AddItem Object.ToStr
    aLB.ItemData(aLB.ListCount - 1) = Object.Key
End Sub

Public Sub ListBox_Swap(aLB As ListBox, ByVal i1 As Long, ByVal i2 As Long)
    Dim lc As Long: lc = aLB.ListCount
    If i1 < 0 Or lc - 1 < i1 Then Exit Sub
    If i2 < 0 Or lc - 1 < i2 Then Exit Sub
    With aLB
        Dim tmp As String, tid As Long
              tmp = .List(i1):           tid = .ItemData(i1)
        .List(i1) = .List(i2): .ItemData(i1) = .ItemData(i2)
        .List(i2) = tmp:       .ItemData(i2) = tid
    End With
End Sub

Public Function Listbox_IsOutOfBounds(aLB As ListBox, ByVal i As Long) As Boolean
    'returns true if i is out of bounds
    Dim lc As Long: lc = aLB.ListCount
    Listbox_IsOutOfBounds = i < 0 Or lc - 1 < i
End Function

Public Function Listbox_IsOutOfBounds2(aLB As ListBox, ByVal i1 As Long, ByVal i2 As Long) As Boolean
    'returns true  if one  of i1, i2 is  out of bounds
    'returns false if both of i1, i2 are inside bounds
    Dim lc As Long: lc = aLB.ListCount
    If (0 <= i1 And i1 < lc) And (0 <= i2 And i2 < lc) Then Exit Function
    Listbox_IsOutOfBounds2 = True
End Function

Public Sub ListBox_Remove(aLB As ListBox, ByVal i As Long)
    With aLB
        If 0 <= i And i < .ListCount Then
            .RemoveItem i
            .ListIndex = i
        End If
    End With
End Sub

Public Sub ListBox_MoveUp(aLB As ListBox) ', ByVal i As Long)
    Dim i1 As Long: i1 = aLB.ListIndex
    Dim i2 As Long: i2 = i1 - 1
    If Listbox_IsOutOfBounds2(aLB, i1, i2) Then Exit Sub
    ListBox_Swap aLB, i1, i2
    aLB.ListIndex = i2
End Sub

Public Sub ListBox_MoveDown(aLB As ListBox) ', ByVal i As Long)
    Dim i1 As Long: i1 = aLB.ListIndex
    Dim i2 As Long: i2 = i1 + 1
    If Listbox_IsOutOfBounds2(aLB, i1, i2) Then Exit Sub
    ListBox_Swap aLB, i1, i2
    aLB.ListIndex = i2
End Sub
' ^ ' ############################## ' ^ '    ListBox Functions    ' ^ ' ############################## ' ^ '

' v ############################## v '    Collection Functions    ' v ############################## v '
Public Function Col_Contains(Col As Collection, Key As String) As Boolean
    'for this function all credits go to the incredible www.vb-tec.de alias Jost Schwider
    'you can find the original version of the function "IsInCollection" here: https://vb-tec.de/collctns.htm
    On Error Resume Next
    '"Extras->Optionen->Allgemein->Unterbrechen bei Fehlern->Bei nicht verarbeiteten Fehlern"
    If IsEmpty(Col(Key)) Then: 'DoNothing
    Col_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function Col_Add(Col As Collection, Obj As Object) As Object
    Set Col_Add = Obj:  Col.Add Obj
End Function

'Public Function Col_Add(Col As Collection, Value)
'   'Nope just use eiter "Col.Add Value" or Col.Add Value, CStr(Value)
'    Col_AddV = Value:   Col.Add Value
'End Function

Public Function Col_AddKey(Col As Collection, Obj As Object) As Object
    Set Col_AddKey = Obj:  Col.Add Obj, Obj.Key ' the object needs to have a Public Function/PropertyGet Key As String
End Function

Public Function Col_AddOrGet(Col As Collection, Obj As Object) As Object
    Dim Key As String: Key = Obj.Key ' the object needs to have a Public Function/PropertyGet Key As String
    If Col_Contains(Col, Key) Then
        Set Col_AddOrGet = Col.Item(Key)
    Else
        Set Col_AddOrGet = Obj
        Col.Add Obj, Key
    End If
End Function

Public Function Col_TryAddObject(Col As Collection, Obj As Object, ByVal Key As String) As Boolean
Try: On Error GoTo Catch
    Col.Add Obj, Key
    Col_TryAddObject = True
Catch: On Error GoTo 0
End Function

Public Sub Col_Remove(Col As Collection, Obj As Object)
    Dim o As Object
    For Each o In Col
        If o.IsSame(Obj) Then 'Obj needs Public Function IsSame(other) As Boolean
            If Col_Contains(Col, Obj.Key) Then Col.Remove Obj.Key 'Obj needs Public Property Key As String
        End If
    Next
End Sub

Public Function Col_IndexFromObject(Col As Collection, Obj As Object) As Long
    Dim i As Long, v, o As Object
    For Each v In Col
        Set o = v
        If o.Key = Obj.Key Then
            Col_IndexFromObject = i
            Exit Function
        End If
        i = i + 1
    Next
End Function

Public Sub Col_ChangeKey(Col As Collection, OldIndexKey, NewKey As String)
    'for this function all credits go to the incredible www.vb-tec.de alias Jost Schwider
    'you can find the original version of the function "CollectionChangeKey" here: https://vb-tec.de/collctns.htm
    Dim Value As Variant
    If IsObject(Col(OldIndexKey)) Then
        Set Value = Col.Item(OldIndexKey)
    Else
        Value = Col.Item(OldIndexKey)
    End If
    Col.Add Value, NewKey  'first add
    Col.Remove OldIndexKey 'then remove
End Sub

Public Sub Col_SwapItems(Col As Collection, ByVal i1 As Long, ByVal i2 As Long)
    Dim c As Long: c = Col.Count
    If c = 0 Then Exit Sub
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    If i1 <= 0 Or c <= i1 Then Exit Sub
    If i2 <= 0 Or c < i2 Then Exit Sub
    If i1 = i2 Then Exit Sub
    Dim Obj1: If IsObject(Col.Item(i1)) Then Set Obj1 = Col.Item(i1) Else Obj1 = Col.Item(i1)
    Dim Obj2: If IsObject(Col.Item(i2)) Then Set Obj2 = Col.Item(i2) Else Obj2 = Col.Item(i2)
    Col.Remove i1: Col.Add Obj2, , i1:     Col.Remove i2
    If i2 < c Then Col.Add Obj1, , i2 Else Col.Add Obj1
End Sub

Public Sub Col_SwapItemsKey(Col As Collection, ByVal i1 As Long, ByVal Key1 As String, _
                                               ByVal i2 As Long, ByVal Key2 As String)
    Dim c As Long: c = Col.Count
    If c = 0 Then Exit Sub
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    If i1 <= 0 Or c <= i1 Then Exit Sub
    If i2 <= 0 Or c < i2 Then Exit Sub
    If i1 = i2 Then Exit Sub
    Dim Obj1: If IsObject(Col.Item(Key1)) Then Set Obj1 = Col.Item(Key1) Else Obj1 = Col.Item(Key1)
    Dim Obj2: If IsObject(Col.Item(Key2)) Then Set Obj2 = Col.Item(Key2) Else Obj2 = Col.Item(Key2)
    Col.Remove Key2: Col.Add Obj2, Key2, i1
    Col.Remove Key1: Col.Add Obj1, Key1, i1
End Sub

Public Sub Col_MoveUp(Col As Collection, ByVal i As Long)
    Dim c As Long: c = Col.Count
    If i <= 1 Or c < i Then Exit Sub
    Col_SwapItems Col, i, i - 1
End Sub

Public Sub Col_MoveUpKey(Col As Collection, ByVal i As Long)
    Dim c As Long: c = Col.Count
    If i <= 1 Or c < i Then Exit Sub
    Dim i1 As Long: i1 = i
    Dim i2 As Long: i2 = i - 1
    Dim Obj1 As Object: Set Obj1 = Col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = Col.Item(i2)
    Dim Key1 As String: Key1 = Obj1.Key
    Dim Key2 As String: Key2 = Obj2.Key
    Col_SwapItemsKey Col, i1, Key1, i2, Key2
End Sub

Public Sub Col_MoveDown(Col As Collection, ByVal i As Long)
    Dim c As Long: c = Col.Count
    If i < 1 Or c <= i Then Exit Sub
    Col_SwapItems Col, i, i + 1
End Sub

Public Sub Col_MoveDownKey(Col As Collection, ByVal i As Long)
    Dim c As Long: c = Col.Count
    If i < 1 Or c <= i Then Exit Sub
    Dim i1 As Long: i1 = i
    Dim i2 As Long: i2 = i + 1
    Dim Obj1 As Object: Set Obj1 = Col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = Col.Item(i2)
    Dim Key1 As String: Key1 = Obj1.Key
    Dim Key2 As String: Key2 = Obj2.Key
    Col_SwapItemsKey Col, i1, Key1, i2, Key2
End Sub

Public Sub Col_ToListBox(Col As Collection, aLB As ListBox, Optional ByVal addEmptyLineFirst As Boolean = False, Optional ByVal doPtrToItemData As Boolean = False)
    Col_ToListCtrl Col, aLB, addEmptyLineFirst
End Sub

Public Sub Col_ToComboBox(Col As Collection, aCB As ComboBox, Optional ByVal addEmptyLineFirst As Boolean = False, Optional ByVal doPtrToItemData As Boolean = False)
    Col_ToListCtrl Col, aCB, addEmptyLineFirst
End Sub

Public Sub Col_ToListCtrl(Col As Collection, ComboBoxOrListBox, Optional ByVal addEmptyLineFirst As Boolean = False, Optional ByVal doPtrToItemData As Boolean = False)
    If Col Is Nothing Then Exit Sub
    Dim i As Long, c As Long: c = Col.Count: If c = 0 Then Exit Sub
    Dim vt As VbVarType: vt = VarType(Col.Item(1))
    Dim v, Obj As Object
    With ComboBoxOrListBox
        If .ListCount Then .Clear
        If addEmptyLineFirst Then .AddItem vbNullString
        Select Case vt
        Case vbByte, vbInteger, vbLong, vbCurrency, vbDate, vbSingle, vbDouble, vbDecimal, vbString
            For i = 1 To c
                .AddItem Col.Item(i)
            Next
        Case vbObject
            For i = 1 To c
                Set Obj = Col.Item(i)
                .AddItem Obj.ToStr ' the object needs to have a Public Function ToStr As String
                If doPtrToItemData Then .ItemData(i - 1) = Obj.Ptr ': Debug.Print Obj.Ptr ' and a Public Function Ptr As LongPtr
                ' ItemData can only be of type Long no String
            Next
        End Select
    End With
End Sub

Public Property Get Col_ObjectFromListCtrl(Col As Collection, ComboBoxOrListBox, i_out As Long) As Object
    Dim li As Long: li = ComboBoxOrListBox.ListIndex
    If i_out < 0 Then i_out = li
    'i_out = IIf(li < 0, i_out, li)
    If i_out < 0 Then Exit Property
    Dim Key As String: Key = ComboBoxOrListBox.ItemData(i_out)
    If Col_Contains(Col, Key) Then Set Col_ObjectFromListCtrl = Col.Item(Key)
End Property

Public Sub Col_Sort(Col As Collection)
    Set m_Col = Col
    Dim c As Long: c = m_Col.Count
    If c = 0 Then: Set m_Col = Nothing: Exit Sub
    Dim vt As VbVarType: vt = VarType(m_Col.Item(1))
    Select Case vt
    Case vbByte, vbInteger, vbLong, vbCurrency, vbDate, vbSingle, vbDouble, vbDecimal
        Col_QuickSortVar 1, c
    Case vbString
        Col_QuickSortStr 1, c
    Case vbObject
        Col_QuickSortObj 1, c
    End Select
    Set m_Col = Nothing
End Sub

Public Function Col_ToStr(Col As Collection) As String
    Dim s As String, v, o As Object
    For Each v In Col
        If IsObject(v) Then
            Set o = v
            s = s & o.ToStr & vbCrLf
        Else
            s = s & CStr(v) & vbCrLf
        End If
    Next
    Col_ToStr = s
End Function
    
' The recursive data-independent QuickSort for primitive data-variables
Private Sub Col_QuickSortVar(ByVal i1 As Long, ByVal i2 As Long)
    Dim t As Long
    If i2 > i1 Then
        t = Col_DivideVar(i1, i2)
        Col_QuickSortVar i1, t - 1
        Col_QuickSortVar t + 1, i2
    End If
End Sub

Private Function Col_DivideVar(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (Col_CompareVar(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (Col_CompareVar(p, j) < 0))
        If i < j Then Col_SwapVar i, j
    Loop While (i < j)
    Col_SwapVar i, p
    Col_DivideVar = i
End Function

Private Function Col_CompareVar(ByVal i1 As Long, ByVal i2 As Long) As Variant
    Col_CompareVar = m_Col.Item(i1) - m_Col.Item(i2)
End Function

Private Sub Col_SwapVar(ByVal i1 As Long, ByVal i2 As Long)
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_Col.Count
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    Dim Var1: Var1 = m_Col.Item(i1)
    Dim Var2: Var2 = m_Col.Item(i2)
    m_Col.Remove i1: m_Col.Add Var2, , i1:   m_Col.Remove i2
    If i2 < c Then m_Col.Add Var1, , i2 Else m_Col.Add Var1
End Sub

' The recursive data-independent QuickSort for strings
Private Sub Col_QuickSortStr(ByVal i1 As Long, ByVal i2 As Long)
    Dim t As Long
    If i1 < i2 Then
        t = Col_DivideStr(i1, i2)
        Col_QuickSortStr i1, t - 1
        Col_QuickSortStr t + 1, i2
    End If
End Sub

Private Function Col_DivideStr(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (Col_CompareStr(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (Col_CompareStr(p, j) < 0))
        If i < j Then Col_SwapStr i, j
    Loop While (i < j)
    Col_SwapStr i, p
    Col_DivideStr = i
End Function

Private Function Col_CompareStr(ByVal i1 As Long, ByVal i2 As Long)
    Col_CompareStr = StrComp(m_Col.Item(i1), m_Col.Item(i2))
    'Dim Str1 As String: Str1 = m_col.Item(i1)
    'Dim Str2 As String: Str2 = m_col.Item(i2)
    'CompareStr = StrComp(Str1, Str2)
End Function

Private Sub Col_SwapStr(ByVal i1 As Long, ByVal i2 As Long)
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_Col.Count
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    Dim Str1 As String: Str1 = m_Col.Item(i1)
    Dim Str2 As String: Str2 = m_Col.Item(i2)
    m_Col.Remove i1: m_Col.Add Str2, , i1:   m_Col.Remove i2
    If i2 < c Then m_Col.Add Str1, , i2 Else m_Col.Add Str1
End Sub

' The recursive data-independent QuickSort for objects
Private Sub Col_QuickSortObj(ByVal i1 As Long, ByVal i2 As Long)
    Dim t As Long
    If i2 > i1 Then
        t = Col_DivideObj(i1, i2)
        Col_QuickSortObj i1, t - 1
        Col_QuickSortObj t + 1, i2
    End If
End Sub

Private Function Col_DivideObj(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (Col_CompareObj(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (Col_CompareObj(p, j) < 0))
        If i < j Then Col_SwapObj i, j
    Loop While (i < j)
    Col_SwapObj i, p
    Col_DivideObj = i
End Function

Private Function Col_CompareObj(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim Obj1 As Object: Set Obj1 = m_Col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = m_Col.Item(i2)
    Col_CompareObj = Obj1.compare(Obj2)
End Function

Private Sub Col_SwapObj(ByVal i1 As Long, ByVal i2 As Long)
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_Col.Count
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    Dim Obj1 As Object: Set Obj1 = m_Col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = m_Col.Item(i2)
    m_Col.Remove i1: m_Col.Add Obj2, , i1:   m_Col.Remove i2
    If i2 < c Then m_Col.Add Obj1, , i2 Else m_Col.Add Obj1
End Sub

' ^ ############################## ^ '    Collection Functions    ' ^ ############################## ^ '


