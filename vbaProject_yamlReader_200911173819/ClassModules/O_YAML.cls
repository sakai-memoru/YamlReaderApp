VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "O_YAML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private jObj As cJobject
Private jagaryKeys As Variant
Private dictRow As Dictionary
Private dictElement As Dictionary
Private dictFullkey As Dictionary
Private structureOfJObj As String

Public Enum enumNotationType
'''' ******************************************
''
key
KeyValue
json
Sequence
SequenceKey
SequenceValue
SequenceKeyValue
SequenceJson
End Enum

'//--------------------------------------------------------
'// Class Event
'//
Private Sub Class_Initialize()
'''' /********************************************************
''
Set dictRow = New Dictionary
'Set dictRowFullKey = New Dictionary
Set dictFullkey = New Dictionary
Set dictElement = New Dictionary
Set jObj = New cJobject
Call jObj.init(Nothing)
''
End Sub

Private Sub Class_Terminate()
'''' ********************************************************
''
Set dictRow = Nothing
'Set dictRowFullKey = Nothing
Set dictFullkey = Nothing
Set dictElement = Nothing
Call jObj.tearDown
End Sub

Public Function GetNotationType(Value) As String
'''' ********************************************************
'''' @function GetNotationType
'''' @param {Variant} value
'''' @return {String}
''
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
Dim dict As Dictionary
Set dict = New Dictionary
''
Call dict.Add("Key", enumNotationType.key)
Call dict.Add("KeyValue", enumNotationType.KeyValue)
Call dict.Add("Json", enumNotationType.json)
''
Call dict.Add("Sequence", enumNotationType.Sequence)
Call dict.Add("SequenceKey", enumNotationType.SequenceKey)
Call dict.Add("SequenceValue", enumNotationType.SequenceValue)
Call dict.Add("SequenceKeyValue", enumNotationType.SequenceKeyValue)
Call dict.Add("SequenceJson", enumNotationType.SequenceJson)
''
Let GetNotationType = C_Dictionary.GetKeyByItem(dict, Value)
''
End Function

Public Function GetElement(ByVal str As String) As Dictionary
'''' ******************************************
'' parse a parameter of str and extract value and judge data type
'' @param {String} str
'' @return {Dictionary}
''
Dim dictRtn As Dictionary
Set dictRtn = New Dictionary
''
Dim C_String As C_String
Set C_String = New C_String
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim CONS_COMMENT As String
Let CONS_COMMENT = "Comment"
Dim CONS_KEY As String
Let CONS_KEY = "Key"
Dim CONS_VAL As String
Let CONS_VAL = "Value"
Dim CONS_ITEM As String
Let CONS_ITEM = "Item"
Dim CONS_SEQUENCE As String
Let CONS_SEQUENCE = "Sequence"
Dim CONS_JSON As String
Let CONS_JSON = "Json"
''
Dim elm As String
Dim cmt As String
Dim typ As String
Let str = VBA.Trim(str)
Dim aryTemp As Variant
''
If C_String.Contain(str, "#") Then
    If C_String.StartsWith(str, "#") Then
        Let typ = CONS_COMMENT
        Call dictRtn.Add("Type", typ)
        Call dictRtn.Add("Comment", str)
        Set GetElement = dictRtn
        Exit Function
    Else
        Let aryTemp = VBA.Split(str, "#")
        Let elm = VBA.Trim(aryTemp((LBound(aryTemp))))
        Call C_Array.shift(aryTemp)
        Let cmt = "# " & Join(aryTemp, " # ")
    End If
Else
    Let elm = VBA.Trim(str)
End If
''
If Not C_String.StartsWith(elm, "-") Then
    ' key - value item
    If C_String.IsJsonString(elm) Then
    ' JSON String
        Let typ = CONS_JSON
    ''
    ElseIf Not C_String.Contain(elm, ":") Then
        Let typ = CONS_VAL
    Else
        If C_String.EndsWith(elm, ":") Then
            Let aryTemp = VBA.Split(elm, ":")
            Let elm = VBA.Trim(aryTemp((LBound(aryTemp))))
            Let typ = CONS_KEY
        Else
            '' FIXME
            Let typ = CONS_KEY & CONS_VAL
        End If
    End If
Else
    ' sequence item
    ''Console.log "str: " & str
    ''Console.log "elm: " & elm
        Let elm = VBA.Trim(C_String.Chomp(elm, 1))
    If elm = "" Then
        Console.info "format error!"
        Set GetElement = dictRtn
        Exit Function
    End If
    If C_String.EndsWith(elm, ";") Then
        Let elm = C_String.Chop(elm, 1)
        Let typ = CONS_SEQUENCE
    ElseIf C_String.IsJsonString(elm) Then
        ' JSON String
        Let typ = CONS_SEQUENCE & CONS_JSON
        ''
    ElseIf Not C_String.Contain(elm, ":") Then
        Let typ = CONS_SEQUENCE & CONS_VAL
    Else
        If C_String.EndsWith(elm, ":") Then
            Let aryTemp = VBA.Split(elm, ":")
            Let elm = VBA.Trim(aryTemp(LBound(aryTemp)))
            Let typ = CONS_SEQUENCE & CONS_KEY
        Else
            Let typ = CONS_SEQUENCE & CONS_KEY & CONS_VAL
        End If
    End If
End If
''
Call dictRtn.Add("Type", typ)
Call dictRtn.Add("Element", elm)
If Not cmt = "" Then
    Call dictRtn.Add("Comment", cmt)
End If
Set GetElement = dictRtn
''
End Function

Public Sub PopulateKeyValueInJObject(ByVal fullkey As String, ByVal Value As Variant, Optional ByVal isSeqKey As Boolean = False)
'''' ******************************************
''
Dim C_String As C_String
Set C_String = New C_String
''
Dim keyTemp As String
Dim aryKey As Variant
Let aryKey = VBA.Split(fullkey, ".")
'Console.log "fullkey : " & fullkey
'Console.Dump aryKey
''
Dim i As Long
Dim preKeyTemp As String
''Let preKeyTemp = "1"
For i = LBound(aryKey) To UBound(aryKey)
    If i = LBound(aryKey) Then ' aryKey(0)
        Let keyTemp = aryKey(i)
        If C_String.IsNumeric(aryKey(i)) Then
            Let jObj.root.isArrayRoot = True
            If jObj.root.children.Count < CLng(aryKey(i)) Then
                'Call jObj.root.add.addArray
                Call jObj.root.Add
            End If
        Else
            Call jObj.root.Add(keyTemp)
        End If
    ElseIf i = LBound(aryKey) + 1 Then ' aryKey(1)
        Let keyTemp = keyTemp & "." & aryKey(i)
        'Console.log "i = " & i
        'Console.log "aryKey  : " & aryKey(i)
        'Console.log "keyTemp : " & keyTemp
        'Console.log "fullkey : " & fullkey
        'Console.log "value   : " & value
        If Not C_String.IsNumeric(aryKey(i)) Then
            Call jObj.Add(keyTemp)
        Else
            If Not C_String.IsNumeric(preKeyTemp) Then
                If Not jObj.find(preKeyTemp).isArrayRoot Then
                    jObj.find(preKeyTemp).addArray
                    If jObj.find(preKeyTemp).children.Count < CLng(aryKey(i)) Then
                        Call jObj.find(preKeyTemp).Add
                    End If
                Else
                    If Not jObj.find(preKeyTemp).hasChildren Then
                        Call jObj.find(preKeyTemp).Add
                    Else
                        If jObj.find(preKeyTemp).children.Count < CLng(aryKey(i)) Then
                            Call jObj.find(preKeyTemp).Add
                        Else
                            'Let jObj.child(preKeyTemp).isArrayRoot = True
                            Call jObj.Add(keyTemp)
                        End If
                    End If
                End If
            Else
                Let jObj.child(preKeyTemp).isArrayRoot = True
                Call jObj.child(preKeyTemp).Add
            End If
        End If
    Else
        Let keyTemp = keyTemp & "." & aryKey(i)
        'Console.log "i = " & i
        'Console.log "aryKey  : " & aryKey(i)
        'Console.log "keyTemp : " & keyTemp
        'Console.log "fullkey : " & fullkey
        'Console.log "value   : " & value
        If Not C_String.IsNumeric(aryKey(i)) Then
            If Not C_String.IsNumeric(preKeyTemp) Then
                Call jObj.Add(keyTemp)
            Else
                Let jObj.child(preKeyTemp).isArrayRoot = True
                Call jObj.Add(keyTemp)
            End If
        Else
            If Not jObj.find(preKeyTemp).isArrayRoot Then
                jObj.find(preKeyTemp).addArray
                If jObj.find(preKeyTemp).children.Count < CLng(aryKey(i)) Then
                    Call jObj.find(preKeyTemp).Add
                End If
            End If
            If Not jObj.find(preKeyTemp).hasChildren Then
                Call jObj.find(preKeyTemp).Add
    Else
                If jObj.find(preKeyTemp).children.Count < CLng(aryKey(i)) Then
                    Call jObj.find(preKeyTemp).Add
                End If
                Call jObj.Add(keyTemp)
            End If
        End If
    End If
    'Console.log "format : " & vbCrLf & jObj.formatData
    Let preKeyTemp = keyTemp
Next i
''
Dim jo As cJobject
Set jo = New cJobject
Dim json_str As String
If isSeqKey Then
    'Call jObj.find(fullkey).parent.addArray
Else
    If C_String.IsJsonString(Value) Then
        If Not C_String.IsEmpty(Value) Then
            Let json_str = Value
            Set jo = usefulcJobject.JSONParse(json_str)
            If jo.isArrayRoot Then
            'If isSeqKey Then
                'Call jObj.find(fullkey)
                Call jObj.find(fullkey).addArray.arrayAppend(jo)
            Else
                Call jObj.find(fullkey).merge(jo)
            End If
        End If
    Else
    Call jObj.find(fullkey).setValue(Value)
    End If
End If

'Console.log jObj.formatData
''
End Sub

Public Sub SetValueIntoJObject()
'''' ******************************************
''
Dim C_String As C_String
Set C_String = New C_String
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim CONS_OBJECT As String
Let CONS_OBJECT = "Object"
Dim CONS_ARRAY As String
Let CONS_ARRAY = "Array"
''
Let structureOfJObj = ""
Dim elmType As String
Dim gid As Variant
Dim d As Dictionary
Set d = New Dictionary
Dim fullkey As String
Dim val As String
Dim jo As cJobject
Set jo = New cJobject
''
For Each gid In dictFullkey.keys
    Set d = dictElement.Item(gid)
    Let elmType = d.Item("TYPE")
    Let val = d.Item("VALUE")
    Let fullkey = dictFullkey.Item(gid)
    'Console.Dump d
    'Console.log "fullkey : " & fullkey
    If C_String.IsEmpty(structureOfJObj) Then
        If C_String.MatchWith(elmType, "Sequence") Then
            Let structureOfJObj = CONS_ARRAY
            If Not jObj.hasChildren Then
                Call jObj.init(Nothing).addArray
            End If
        ElseIf Not C_String.EqualIgnore(elmType, "Comment") Then
            Let structureOfJObj = CONS_OBJECT
            If Not jObj.hasChildren Then
                Call jObj.init(Nothing)
            End If
        End If
        ''
        'Console.log structureOfJObj
    End If
        ''
    If C_String.EqualIgnore(elmType, "SequenceKey") Then
        Call Me.PopulateKeyValueInJObject(fullkey, "", True)
    ElseIf C_String.EqualIgnore(elmType, "SequenceValue") Then
        Call Me.PopulateKeyValueInJObject(fullkey, val)
    ElseIf C_String.EqualIgnore(elmType, "SequenceKeyValue") Then
        Call Me.PopulateKeyValueInJObject(fullkey, val)
    ElseIf C_String.EqualIgnore(elmType, "Value") Then
        Call Me.PopulateKeyValueInJObject(fullkey, val)
    ElseIf C_String.EqualIgnore(elmType, "KeyValue") Then
        Call Me.PopulateKeyValueInJObject(fullkey, val)
    ElseIf C_String.EqualIgnore(elmType, "Sequence") Then
        Call Me.PopulateKeyValueInJObject(fullkey, val, True)
    ElseIf C_String.EqualIgnore(elmType, "SequenceJson") Then
        Call Me.PopulateKeyValueInJObject(fullkey, val, True)
    ElseIf C_String.EqualIgnore(elmType, "Json") Then
        Call Me.PopulateKeyValueInJObject(fullkey, val)
        Else
        ''
        End If
Next gid
''
'Console.log jObj.formatData
''
End Sub

Public Function CreateFullKeyDictionay(aryaryKeys) As Variant
'''' ******************************************
''
Dim C_String As C_String
Set C_String = New C_String
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim aryarySeq() As Variant
ReDim aryarySeq(LBound(aryaryKeys, 1) To UBound(aryaryKeys, 1), LBound(aryaryKeys, 2) To UBound(aryaryKeys, 2) - 1)
Dim aryaryTemp() As Variant
ReDim aryaryTemp(LBound(aryaryKeys, 1) To UBound(aryaryKeys, 1), LBound(aryaryKeys, 2) To UBound(aryaryKeys, 2) - 1)
''
Dim r As Long
Dim c As Long
Dim rowAry As Variant
Dim gid As String
Dim indentLevel As Long
Dim elmType As String
Dim elm As Variant
''
'''' process: Generate Dummy Full Key Array
For r = LBound(aryaryKeys, 1) To UBound(aryaryKeys, 1)
    Let gid = aryaryKeys(r, UBound(aryaryKeys, 2))
    Let rowAry = dictRow.Item(gid)
    If C_Array.IsEmptyArray(rowAry) Then
        ''
            Else
        Let indentLevel = rowAry(LBound(rowAry))
        For c = LBound(aryaryKeys, 2) To indentLevel + 1
            Let elm = aryaryKeys(r, c)
            If c = LBound(aryaryKeys, 2) Then
                Let aryaryTemp(r, c) = elm
        Else
                Let aryaryTemp(r, c) = aryaryTemp(r, c - 1) & "." & elm
        End If
    Next c
    End If
Next r
'Console.Dump aryaryTemp
''
'''' process: Count Sequence No of Sequence Element
Dim seq As Long
Dim preElm As Variant
Dim elmParent As Variant
Dim preElmParent As Variant
''
For c = LBound(aryaryTemp, 2) To UBound(aryaryTemp, 2)
    For r = LBound(aryaryTemp, 1) To UBound(aryaryTemp, 1)
        Let gid = aryaryKeys(r, UBound(aryaryKeys, 2))
        Let rowAry = dictRow.Item(gid)
        If C_Array.IsEmptyArray(rowAry) Then
            ''
        Else
            Let elmType = rowAry(UBound(rowAry))
            Let elm = aryaryTemp(r, c)
            If c = LBound(aryaryTemp, 2) Then
                Let elmParent = ""
            Else
                Let elmParent = aryaryTemp(r, c - 1)
            End If
            'Console.log elm
            'Console.log elmType
            If elm = "" Then
                Let seq = 0
            Else
                If C_String.MatchWith(elmType, "Sequence") Then
                    If Not preElm = elm Then
                        Let seq = seq + 1
                    End If
                End If
            End If
            Let aryarySeq(r, c) = seq
            Let preElm = elm
            Let preElmParent = elmParent
        End If
    Next r
    Let seq = 0
    Let preElm = ""
    Let preElmParent = ""
Next c
'Console.Dump aryarySeq

'''' process: Create fullkey
'Dim elmAry As Variant
Dim elmDict As Dictionary
Set elmDict = New Dictionary
Dim seqNo As Long
Dim fullkey As String
Dim cmt As String
Dim aryTemp As Variant
For r = LBound(aryaryKeys, 1) To UBound(aryaryKeys, 1)
    Let gid = aryaryKeys(r, UBound(aryaryKeys, 2))
    Let rowAry = dictRow.Item(gid)
    If C_Array.IsEmptyArray(rowAry) Then
    ''
        Else
        Let indentLevel = rowAry(LBound(rowAry))
        For c = LBound(aryaryKeys, 2) To indentLevel + 1
            Let seqNo = aryarySeq(r, c)
            Set elmDict = Me.GetElement(aryaryKeys(r, c))
            ''Let elmAry = Me.GetElement(aryaryKeys(r, c))
            Let elm = elmDict.Item("Element")
            Let elmType = elmDict.Item("Type")
            If elmDict.Exists("Comment") Then
                Let cmt = elmDict.Item("Comment")
            End If
            If C_String.MatchWith(elmType, "Sequence") Then
                If c = LBound(aryaryKeys, 2) Then
                    Let fullkey = seqNo
                Else
                    Let fullkey = fullkey & "." & seqNo
                End If
                If C_String.EqualIgnore(elmType, "SequenceKey") Then
                    Let fullkey = fullkey & "." & elm
                End If
            ElseIf C_String.EqualIgnore(elmType, "Key") Then
                If c = LBound(aryaryKeys, 2) Then
                    Let fullkey = elm
                Else
                    Let fullkey = fullkey & "." & elm
                End If
            ElseIf C_String.EqualIgnore(elmType, "KeyValue") Then
                Let aryTemp = VBA.Split(elm, ":")
                If c = LBound(aryaryKeys, 2) Then
                    Let fullkey = Trim(aryTemp(LBound(aryTemp)))
                Else
                    Let fullkey = fullkey & "." & Trim(aryTemp(LBound(aryTemp)))
                End If
            ElseIf C_String.EqualIgnore(elmType, "Json") Then
                ''
        End If
    Next c
    End If
    'Console.log fullkey
    Call dictFullkey.Add(gid, fullkey)
    Let fullkey = ""
Next r
''
End Function

Public Sub CreateElementDictionary(aryaryKeys)
'''' ******************************************
''
Dim C_String As C_String
Set C_String = New C_String
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
''
Dim r As Long
Dim c As Long
Dim ary As Variant
Dim elm As String
Dim elmType As String
Dim cmt As String
Dim gid As String
Dim jOb As cJobject
Dim d As Dictionary
Set d = New Dictionary
Dim commentFirstFlag As Variant
Let commentFirstFlag = ""
Dim preElmType As String
Dim elmDict As Dictionary
Set elmDict = New Dictionary
Dim aryTemp As Variant
Dim elmTemp As String
''
For r = LBound(aryaryKeys, 1) To UBound(aryaryKeys, 1) ''FIXME
    For c = LBound(aryaryKeys, 2) To UBound(aryaryKeys, 2) - 1
        If Not aryaryKeys(r, c) = "" Then
            Set elmDict = Me.GetElement(aryaryKeys(r, c))
            Let elm = elmDict.Item("Element")
            Let elmType = elmDict.Item("Type")
            If elmDict.Exists("Comment") Then
                Let cmt = elmDict.Item("Comment")
            End If
            If elmType = "Comment" Then
                If commentFirstFlag = "" Then
                    Let commentFirstFlag = True
                End If
                    Exit For
            Else
                If commentFirstFlag = "" Then
                    Let commentFirstFlag = False
                End If
                If C_String.MatchWith(elmType, "Sequence") Then
                    If elmType = "SequenceKey" Then
                        'Call d.add("KEY", elm)
                    ElseIf elmType = "SequenceValue" Then
                        Call d.Add("VALUE", elm)
                    ElseIf elmType = "SequenceKeyValue" Then
                        Let aryTemp = VBA.Split(elm, ":")
                        Call C_Array.shift(aryTemp)
                        Let elmTemp = VBA.Trim(VBA.Join(aryTemp, ":"))
                        If C_String.IsJsonString(elmTemp) Then
                            Call d.Add("VALUE", elmTemp)
                        Else
                            Call d.Add("VALUE", elmTemp)
                        End If
                    ElseIf elmType = "SequenceJson" Then
                        Call d.Add("VALUE", elmTemp)
                    ElseIf elmType = "Sequence" Then
                        ''
                    Else
                        ''
                    End If
                Else
                    If elmType = "Key" Then
                        'Call d.add("KEY", elm)
                    ElseIf elmType = "Value" Then
                        Call d.Add("VALUE", elm)
                    ElseIf elmType = "KeyValue" Then
                        Let aryTemp = VBA.Split(elm, ":")
                        Call C_Array.shift(aryTemp)
                        Let elmTemp = VBA.Trim(VBA.Join(aryTemp, ":"))
                        If C_String.IsJsonString(elmTemp) Then
                            Call d.Add("VALUE", elmTemp)
                        Else
                            Call d.Add("VALUE", elmTemp)
                        End If
                        ''
                    ElseIf elmType = "Json" Then
                        Call d.Add("VALUE", elm)
                        ''
                    Else
                        '' FIXME
                    End If
                    'Console.Dump ary
                End If
                Let preElmType = elmType
            End If
            Let ary = Array()
        Else
            '' FIXME
        End If
    Next c
    Let gid = aryaryKeys(r, UBound(aryaryKeys, 2))
    If commentFirstFlag Then
        Call d.Add("TYPE", "Comment")
        Call d.Add("COMMENT", cmt)
    Else
        Call d.Add("TYPE", preElmType)
        Call d.Add("COMMENT", cmt)
    End If
    Let commentFirstFlag = ""
    Call dictElement.Add(gid, C_Dictionary.DeepCopy(d))
    Call d.RemoveAll
Next r
''Console.Dump dictElement
''
End Sub

Public Function GetIndentLevel(ary As Variant) As Long
'''' ******************************************
''
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim indentLevel As Long
''
Dim c As Long
Dim s As String
For c = LBound(ary) To UBound(ary)
    Let s = VBA.Trim(ary(c))
    If s = "" Then
        Let indentLevel = indentLevel + 1
        Else
        Exit For
            End If
Next c
If indentLevel = C_Array.Count(ary) Then
    Let indentLevel = -1
End If
Let GetIndentLevel = indentLevel
''
End Function

Public Function GetLineTrimmed(ByVal line As String, Optional ByVal indentNum As Variant) As Variant
'''' ******************************************
''
Dim C_String As C_String
Set C_String = New C_String
Dim C_Array As C_Array
Set C_Array = New C_Array
''
If VBA.IsMissing(indentNum) Then
    Let indentNum = 4
End If
''
Dim aryLine(1 To 6) As Variant
Dim ary() As Variant
Dim aryTemp As Variant
Dim Indented As String
Dim notation As String
Dim comment As String
Dim objectType As String
Dim notationType As String
Let line = C_String.TrimRightSpace(line)
''
Let ary = C_String.ToCharArray(line)
Dim cnt As Long
''Console.dump ary
Dim i As Long
For i = LBound(ary) To UBound(ary)
    If ary(i) = " " Then
        Let Indented = Indented + ary(i)
        Let cnt = cnt + 1
    Else
        Exit For
    End If
Next i
''
If C_String.Contain(line, "#") Then
    If C_String.StartsWith(line, "#") Then
        Let aryLine(1) = cnt / indentNum
        Let aryLine(2) = Indented
        Let aryLine(3) = ""
        Let aryLine(4) = line
        Let aryLine(5) = "Comment"
        Let aryLine(6) = "Comment"
        Let GetLineTrimmed = aryLine
        Exit Function
    Else
        Let aryTemp = VBA.Split(line, "#")
        Let notation = VBA.Trim(aryTemp((LBound(aryTemp))))
        Call C_Array.shift(aryTemp)
        Let comment = "# " & VBA.Join(aryTemp, " # ")
    End If
Else
    Let notation = VBA.Trim(line)
End If
Let line = VBA.Trim(line)
If Not C_String.StartsWith(line, "-") Then
    If Not C_String.Contain(line, ":") Then
        Let GetLineTrimmed = False
        Exit Function
    Else
        Let objectType = "Mapping"
        If C_String.IsJsonString(line) Then
            Let notationType = Me.GetNotationType(enumNotationType.json)
        ElseIf C_String.EndsWith(line, ":") Then
            Let notationType = Me.GetNotationType(enumNotationType.key)
        ElseIf C_String.Contain(line, ":") Then
            Let notationType = Me.GetNotationType(enumNotationType.KeyValue)
        Else
            ' N/A
        End If
    End If
Else
    Let objectType = "Sequence"
    If C_String.IsJsonString(line) Then
        Let notationType = Me.GetNotationType(enumNotationType.SequenceJson)
    ElseIf C_String.EndsWith(line, ":") Then
        Let notationType = Me.GetNotationType(enumNotationType.SequenceKey)
    Else
        If C_String.Contain(line, ":") Then
            Let notationType = Me.GetNotationType(enumNotationType.SequenceKeyValue)
        Else
            If line = "-" Then
                Let notationType = Me.GetNotationType(enumNotationType.Sequence)
            Else
                Let notationType = Me.GetNotationType(enumNotationType.SequenceValue)
            End If
        End If
    End If
End If
''
Let aryLine(1) = cnt / indentNum
Let aryLine(2) = Indented
Let aryLine(3) = notation
Let aryLine(4) = comment
Let aryLine(5) = objectType
Let aryLine(6) = notationType
Let GetLineTrimmed = aryLine
''
End Function

Public Sub SetData(ByRef aryary)
'''' ******************************************
''
Dim C_Array As C_Array
Set C_Array = New C_Array
''
''
'''' process: put a dummy key on Sequence string, "-"
Dim seqItemNo As Long
Dim r As Long
Dim c As Long
For c = LBound(aryary, 2) To UBound(aryary, 2)
    For r = LBound(aryary, 1) To UBound(aryary, 1)
        If VBA.Trim(aryary(r, c)) = "-" Then
            Let seqItemNo = seqItemNo + 1
            Let aryary(r, c) = VBA.Trim(aryary(r, c)) & " " & seqItemNo & ";"
        End If
    Next r
Next c
''Console.Dump aryary
''
Dim jagary() As Variant
Let jagary = C_Array.ArrayArrayToJagArray(aryary)
''
Dim ary As Variant
Dim aryKey As Variant
ReDim aryKey(LBound(jagary(1)) To UBound(jagary(1)) + 1)
ReDim jagaryKeys(LBound(jagary) To UBound(jagary))
''
Dim gid As String
Dim indentLevel As Long
Dim line As String
''
For r = LBound(jagary) To UBound(jagary)
    Let gid = C_Commons.CreateGUID
    'Console.Dump jagary(r)
    Let aryKey(UBound(jagary(r)) + 1) = gid
    Let indentLevel = Me.GetIndentLevel(jagary(r))
    ''Console.log indentLevel
    If indentLevel > -1 Then
        For c = LBound(jagary(r)) To UBound(jagary(r))
            If Not VBA.IsEmpty(jagary(r)(c)) Then
                Let aryKey(c) = jagary(r)(c)
            Else
                'Console.log "r = " & r
                'Console.log "c = " & c
                'Console.log "indentLevel = " & indentLevel
                If c <= indentLevel Then
                    Let aryKey(c) = jagaryKeys(r - 1)(c)
                Else
                    Let aryKey(c) = ""
                End If
            End If
        Next c
        ''Console.log "aryKey"
        ''Console.Dump aryKey
        Let jagaryKeys(r) = aryKey
        Let line = C_Array.ToLine(jagary(r))
        Let ary = Me.GetLineTrimmed(line)
        ''Console.Dump ary
        If VBA.IsArray(ary) Then
            Call dictRow.Add(gid, ary)
        End If
    End If
Next r
''
Call C_Array.DeleteEmptyElement(jagaryKeys)
'Console.Dump dictRow
'Console.Dump jagaryKeys
''
End Sub

Public Function ConverttoJObject(ByVal targetParam As String, ByVal shtName As String, Optional ByVal wbName As Variant, Optional opt As Variant = False) As cJobject
'''' ********************************************************
'''' @function ConverttoJObject
'''' @description
'''' @param targetparam {String} - target range parameter
'''' @param shtName {String} - sheet name
'''' @param wbName {Variant<String>} Optional - workbook name
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_Range As C_Range
Set C_Range = New C_Range
''
If VBA.IsMissing(wbName) Then
    Let wbName = ThisWorkbook.Name
End If
''
Dim aryary() As Variant
Let aryary = C_Range.GetCurrentRegionByKeyword(targetParam, shtName, bookName:=wbName)

If opt Then
    Let aryary = C_Range.GetCurrentRegionByKeyword(targetParam, shtName, bookName:=wbName)
Else
    Let aryary = C_Range.GetCurrentRegionByKeyword(targetParam, shtName, bookName:=wbName)
    Call C_Array.DeleteRowForArrayArray(aryary, 1)
    Let aryary = C_Array.FilterRowForArrayArray(aryary, 1)
    Call C_Array.DeleteColumnForArrayArray(aryary, 1)
End If
''
Call Me.SetData(aryary)
''
Dim aryaryKeys As Variant
Let aryaryKeys = Me.GetArrayArrayKeys
'Console.Dump aryaryKeys
''
Call Me.CreateFullKeyDictionay(aryaryKeys)
Call Me.CreateElementDictionary(aryaryKeys)
''
Call Me.SetValueIntoJObject
''
Set ConverttoJObject = Me.GetJObject
''
End Function

Public Function GetYamlLine(ByVal targetParam As String, ByVal shtName As String, Optional ByVal wbName As Variant, Optional opt As Variant) As Variant
'''' ********************************************************
''
''
Dim C_String As C_String
Set C_String = New C_String
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_Range As C_Range
Set C_Range = New C_Range
''
If VBA.IsMissing(wbName) Then
    Let wbName = ThisWorkbook.Name
End If
If VBA.IsMissing(opt) Then
    Let opt = False
End If
''
Dim aryary() As Variant
If opt Then
    Let aryary = C_Range.GetCurrentRegionByKeyword(targetParam, shtName, bookName:=wbName)
Else
    Let aryary = C_Range.GetCurrentRegionByKeyword(targetParam, shtName, bookName:=wbName)
    Call C_Array.DeleteRowForArrayArray(aryary, 1)
    Let aryary = C_Array.FilterRowForArrayArray(aryary, 1)
    Call C_Array.DeleteColumnForArrayArray(aryary, 1)
End If
''
Call Me.SetData(aryary)
''
'Console.Dump dictRow
Dim gid As Variant
Dim ary As Variant
Dim line As String
Dim lines() As Variant
For Each gid In dictRow.keys
    Let ary = dictRow.Item(gid)
    If C_String.IsEmpty(ary(4)) Then
        Let line = ary(2) & ary(3)
    ElseIf C_String.IsEmpty(ary(3)) Then
        Let line = ary(2) & ary(4)
    Else
        Let line = ary(2) & ary(3) & "    " & ary(4)
    End If
    Call C_Array.Add(lines, line)
Next gid

Let GetYamlLine = lines
''
End Function

Public Function YamlFileToJObject(file_name As String, Optional base_folder As Variant) As cJobject
'''' *********************************************
''
Dim C_String As C_String
Set C_String = New C_String
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_Range As C_Range
Set C_Range = New C_Range
''
Dim aryLines As Variant
Let aryLines = Me.ReadYamlFile(file_name)
'Console.Dir aryLines
Dim line As Variant 'String
Dim jagary() As Variant
Dim ary As Variant
For Each line In aryLines
    If Not C_String.IsEmpty(VBA.Trim(line)) Then
        Let ary = C_String.GetIndentAndElement(line, 4)
        Call C_Array.AddJagArray(jagary, ary)
    End If
Next line
''
Dim peak As Long
For Each ary In jagary
    If peak < ary(LBound(ary)) Then
        Let peak = ary(LBound(ary))
    End If
Next ary
Dim aryary() As Variant
ReDim aryary(LBound(jagary) To UBound(jagary), 1 To peak + 1)
Dim r As Long
For r = LBound(jagary) To UBound(jagary)
    Let aryary(r, jagary(r)(0) + 1) = jagary(r)(2)
Next r
''Console.Dir aryary
''Console.info C_Array.TypeNameEx(aryary)
Call Me.SetData(aryary)
Dim aryaryKeys As Variant
Let aryaryKeys = Me.GetArrayArrayKeys
Call Me.CreateFullKeyDictionay(aryaryKeys)
Call Me.CreateElementDictionary(aryaryKeys)
Call Me.SetValueIntoJObject
Dim jObj As cJobject
Set jObj = New cJobject
Set jObj = Me.GetJObject
''
Set YamlFileToJObject = jObj
''
End Function

Public Sub Dump()
'''' ********************************************************
''
Console.Dump jagaryKeys
'Console.Dump dictRowFullKey
Console.info jObj.formatData
Console.info JSONStringify(jObj)
''
End Sub

Public Function GetJObject() As cJobject
'''' ******************************************
''
Set GetJObject = jObj
End Function

Public Function GetArrayArrayKeys() As Variant
'''' ******************************************
''
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Let GetArrayArrayKeys = C_Array.JagArrayToArrayArray(jagaryKeys)
''
End Function

Public Function GetDictRowFullKey() As Dictionary
'''' ******************************************
''
Set GetDictRowFullKey = dictRowFullKey
''
End Function

Public Function GetDictRow() As Dictionary
'''' ******************************************
''
Set GetDictRow = dictRow
''
End Function

Public Function GetDictFullKey() As Dictionary
'''' ******************************************
''
Set GetDictFullKey = dictFullkey
''
End Function

Public Function GetDictElement() As Dictionary
'''' ******************************************
''
Set GetDictElement = dictElement
''
End Function

Public Function ReadYamlFile(file_name As String, Optional base_folder As Variant) As Variant
'''' ********************************************************
''
Dim C_File As C_File
Set C_File = New C_File
Dim C_FileIO As C_FileIO
Set C_FileIO = New C_FileIO
''
If VBA.IsMissing(base_folder) Then
    Dim book_path As String
    Let book_path = ThisWorkbook.FullName
    Let base_folder = C_File.GetParentFolder(C_File.GetLocalName(book_path))
End If
Dim file_path As String
Let file_path = C_File.BuildPath(base_folder, file_name)
Console.log "file path : " & file_path
''
Dim aryLines As Variant
Let aryLines = C_FileIO.ReadTextAllAsUTF8ToArray(file_path)
''
Let ReadYamlFile = aryLines
''
End Function

Public Function YamlArrayToArrayArray(aryLines) As Variant
'''' ********************************************************
''
Dim C_String As C_String
Set C_String = New C_String
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim line As Variant 'String
Dim jagary() As Variant
Dim ary As Variant
For Each line In aryLines
    If Not C_String.IsEmpty(VBA.Trim(line)) Then
        Let ary = C_String.GetIndentAndElement(line, 4)
        Call C_Array.AddJagArray(jagary, ary)
    End If
Next line
''
Dim peak As Long
For Each ary In jagary
    If peak < ary(LBound(ary)) Then
        Let peak = ary(LBound(ary))
    End If
Next ary
Dim aryary() As Variant
ReDim aryary(LBound(jagary) To UBound(jagary), 1 To peak + 1)
Dim r As Long
For r = LBound(jagary) To UBound(jagary)
    Let aryary(r, jagary(r)(0) + 1) = jagary(r)(2)
Next r
''
Let YamlArrayToArrayArray = aryary
''
End Function
