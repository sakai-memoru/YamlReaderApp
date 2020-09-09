VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_JObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private C_JObj As cJobject
Private aryFullkey() As Variant
Private C_Array As C_Array

'//--------------------------------------------------------
'// Class Event
'//
Private Sub Class_Initialize()
'''' /********************************************************
''
Set C_JObj = New cJobject
Set C_Array = New C_Array
Call C_Array.Initialize(aryFullkey)
''
End Sub

Private Sub Class_Terminate()
'''' ********************************************************
''
Call C_JObj.tearDown
End Sub

Function SetFullkeyArray(jOb As Variant, isExcludeRoot As Boolean)
'''' *************************************************
''
Dim jo As Variant 'cJObject
''
If jOb.hasChildren Then
    Console.log jOb.fullkey
    Call C_Array.Add(aryFullkey, jOb.fullkey(isExcludeRoot))
    For Each jo In jOb.children
        If jo.hasChildren Then
            Call SetFullkeyArray(jo, isExcludeRoot)
        Else
            Call C_Array.Add(aryFullkey, jo.fullkey(isExcludeRoot))
            Console.log jo.fullkey
        End If
    Next jo
Else
    Call C_Array.Add(aryFullkey, jOb.fullkey(isExcludeRoot))
    Console.log jOb.fullkey
End If
''
End Function

Function GetFullkeyArray(jOb As Variant, Optional excludeRoot As Variant) As Variant
'''' *************************************************
''
Dim isIncludeRoot As Boolean
If VBA.IsMissing(excludeRoot) Then
    Let isIncludeRoot = False
Else
    Let isIncludeRoot = True
End If
Call Me.SetFullkeyArray(jOb, isIncludeRoot)
Let GetFullkeyArray = aryFullkey
''
End Function

Function CreateFullkeyAndProperties(ByRef ary As Variant, jOb As cJobject) As Dictionary
'''' *************************************************
''
Dim C_String As C_String
Set C_String = New C_String
Dim C_Dictionary As C_Dictionary
Set C_Dictionary = New C_Dictionary
''
Dim dictRtn As Dictionary
Set dictRtn = New Dictionary
''
Dim d As Dictionary
Set d = New Dictionary
Dim jo As cJobject
Set jo = New cJobject
''
Dim fullkey As String
Dim val As Variant
Dim ky As Variant
Dim childIdx As Long
Dim childCnt As Long
Dim isRoot As Boolean
Dim isArrayRoot As Boolean
Dim isArrayMember As Boolean
Dim isVoid As Boolean
Dim hasKey As Boolean
Dim hasChildren As Boolean
Dim aryKey As Variant
Dim lastPos As Long
Dim i As Long
For i = LBound(ary) To UBound(ary)
    Let fullkey = ary(i)
    Console.info fullkey
    Let aryKey = VBA.Split(fullkey, ".")
    Let lastPos = UBound(aryKey) + 1
    'Set jo = job.child(fullkey)
    ''
    If Not C_String.IsEmpty(fullkey) Then
        Let val = C_String.DefaultString(jOb.child(fullkey).Value, "")
        Let ky = jOb.child(fullkey).key
        Let childIdx = jOb.child(fullkey).childIndex
        Let isRoot = jOb.child(fullkey).isRoot
        Let isArrayRoot = jOb.child(fullkey).isArrayRoot
        Let isArrayMember = jOb.child(fullkey).isArrayMember
        Let isVoid = jOb.child(fullkey).isVoid
        Let hasKey = jOb.child(fullkey).hasKey
        Let hasChildren = jOb.child(fullkey).hasChildren
        If jOb.child(fullkey).hasChildren Then
            Let childCnt = jOb.child(fullkey).children.Count
        Else
            Let childCnt = 0
        End If
        ''
    Else
        Let fullkey = "_ROOT"
        Let val = C_String.DefaultString(jOb.root.Value, "")
        Let ky = jOb.root.key
        Let childIdx = jOb.root.childIndex
        Let isRoot = jOb.root.isRoot
        Let isArrayRoot = jOb.root.isArrayRoot
        Let isArrayMember = jOb.root.isArrayMember
        Let isVoid = jOb.root.isVoid
        Let hasKey = jOb.root.hasKey
        Let hasChildren = jOb.root.hasChildren
        If jOb.root.hasChildren Then
            Let childCnt = jOb.root.children.Count
        Else
            Let childCnt = 0
        End If
    End If
    Call d.Add("lastPosition", lastPos)
    Call d.Add("value", val)
    Call d.Add("key", ky)
    Call d.Add("childIndex", childIdx)
    Call d.Add("isRoot", isRoot)
    Call d.Add("isArrayRoot", isArrayRoot)
    Call d.Add("isArrayMember", isArrayMember)
    Call d.Add("isVoid", isVoid)
    Call d.Add("hasKey", hasKey)
    Call d.Add("hasChildren", hasChildren)
    Call d.Add("childrenCount", childCnt)
    Call dictRtn.Add(fullkey, C_Dictionary.DeepCopy(d))
    Call d.RemoveAll
Next i
Set CreateFullkeyAndProperties = dictRtn
''
End Function

Function CreateFullkeyJagArray(ary As Variant) As Variant
'''' *************************************************
''
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim arKeys As Variant
Dim i As Long
Dim peak As Long
For i = LBound(ary) To UBound(ary)
    Let arKeys = VBA.Split(ary(i), ".")
    If peak < UBound(arKeys) Then
        Let peak = UBound(arKeys)
    End If
    Console.Dump arKeys
Next i
Console.log "peak = " & peak
Dim ar As Variant
ReDim ar(peak + 1)
Dim jagary() As Variant
Dim j As Long
Dim fullkey As String
For i = LBound(ary) To UBound(ary)
    Let fullkey = ary(i)
    Let arKeys = VBA.Split(fullkey, ".")
    For j = LBound(arKeys) To UBound(arKeys)
        Let ar(j) = arKeys(j)
    Next j
    Let ar(UBound(ar)) = fullkey
    Call C_Array.AddJagArray(jagary, C_Array.ChangeIndexOfBegin(ar))
Next i
Let CreateFullkeyJagArray = jagary
''
End Function

Function JagArrayToYamlArrayArray(jagary As Variant) As Variant
'''' *************************************************
''
Dim C_String As C_String
Set C_String = New C_String
''
Dim aryaryRtn As Variant
ReDim aryaryRtn(LBound(jagary) To UBound(jagary), LBound(jagary(LBound(jagary))) To UBound(jagary(LBound(jagary))))
'Console.Dir jagary
Dim r As Long
Dim c As Long
Dim fullkey As String
For r = LBound(jagary) To UBound(jagary)
    Let fullkey = jagary(r)(UBound(jagary(r)))
    For c = LBound(jagary(LBound(jagary))) To UBound(jagary(LBound(jagary))) - 1
        If r = LBound(jagary) Then
            Let aryaryRtn(r, c) = C_String.DefaultString(jagary(r)(c), "")
        Else
            If jagary(r - 1)(c) = jagary(r)(c) Then
                Let aryaryRtn(r, c) = ""
            Else
                Let aryaryRtn(r, c) = C_String.DefaultString(jagary(r)(c), "")
            End If
        End If
    Next c
    Let aryaryRtn(r, UBound(jagary(r))) = C_String.DefaultString(fullkey, "")
Next r
Let JagArrayToYamlArrayArray = aryaryRtn
''
End Function

Public Sub FillValueInYamlArrayArray(ByRef aryary, ByRef dictFullkey As Dictionary)
'''' *************************************************
''
Dim C_String As C_String
Set C_String = New C_String
''
Dim r As Long
Dim c As Long
Dim fullkey As String
Dim d As Dictionary
Set d = New Dictionary
Dim ky As String
Dim val As String
Dim lastPos As Long
For r = LBound(aryary, 1) To UBound(aryary, 1)
    Let fullkey = aryary(r, UBound(aryary, 2))
    If fullkey = "" Then
        Set d = dictFullkey.Item("_ROOT")
    Else
        Set d = dictFullkey.Item(fullkey)
    End If
    'Console.Dir d
    Let ky = d.Item("key")
    Let val = d.Item("value")
    Let lastPos = d.Item("lastPosition")
    If lastPos > 0 Then
        If C_String.IsNumeric(ky) Then
            Let aryary(r, lastPos) = "- " & val
        Else
            Let aryary(r, lastPos) = ky & " : " & val
        End If
    Else
        Let aryary(r, 1) = "# YAML"
    End If
    Let aryary(r, UBound(aryary, 2)) = "# " & aryary(r, UBound(aryary, 2))
Next r
''
End Sub

Public Function YamlArrayArrayToLines(ByRef yamlaryary, Optional ByVal num = 4) As Variant
'''' *************************************************
''
Dim C_String As C_String
Set C_String = New C_String
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim indent As String
Let indent = C_String.Spaces(num)
Dim r As Long
Dim c As Long
For r = LBound(yamlaryary, 1) To UBound(yamlaryary, 1)
    For c = LBound(yamlaryary, 2) To UBound(yamlaryary, 2)
        If C_String.IsEmpty(yamlaryary(r, c)) Then
            Let yamlaryary(r, c) = indent
        End If
    Next c
Next r
''
Dim lines() As Variant
Dim line As String
For r = LBound(yamlaryary, 1) To UBound(yamlaryary, 1)
    For c = LBound(yamlaryary, 2) To UBound(yamlaryary, 2)
        Let line = line + yamlaryary(r, c)
    Next c
    Let line = C_String.TrimRightSpace(line)
    Call C_Array.Add(lines, line)
    Let line = ""
Next r
Let YamlArrayArrayToLines = lines
''
End Function

Public Function ConverttoYamlArray(ByRef jOb As cJobject) As Variant
'''' *************************************************
''
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim fullkeyary As Variant
Let fullkeyary = Me.GetFullkeyArray(jOb)
Console.Dir fullkeyary
''
Dim dictFullkey As Dictionary
Set dictFullkey = New Dictionary
Set dictFullkey = Me.CreateFullkeyAndProperties(fullkeyary, jOb)
Console.Dump dictFullkey
''
Dim jagary As Variant
Let jagary = Me.CreateFullkeyJagArray(fullkeyary)
Console.Dump jagary
''
Dim aryary As Variant
Let aryary = Me.JagArrayToYamlArrayArray(jagary)
Console.Dump aryary
''
Call Me.FillValueInYamlArrayArray(aryary, dictFullkey)
Console.Dump aryary
''
Dim colNum As Long
Let colNum = UBound(aryary, 2)
Call C_Array.DeleteColumnForArrayArray(aryary, colNum)
''
Dim lines As Variant
Let lines = Me.YamlArrayArrayToLines(aryary)
Console.Dir lines
''
Let ConverttoYamlArray = lines
''
End Function

'''' ///////////////////////////////////////////////////////////////
'''' ///////////////////////////////////////////////////////////////
Public Function ConvertfromObject(ByRef obj As Variant) As cJobject
'''' *************************************************
''
Dim rtnJObj As cJobject
Set rtnJObj = New cJobject
''
Dim json_str As String
Let json_str = JsonConverter.ConvertToJson(obj)
'Console.info json_str
''
Set rtnJObj = C_JObj.parse(json_str)
Set ConvertfromObject = rtnJObj
''
End Function

Public Function ConvertfromArray(ByRef ary As Variant) As cJobject
'''' *************************************************
''
Dim C_Commons As C_Commons
Set C_Commons = New C_Commons
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim rtnJObj As cJobject
Set rtnJObj = New cJobject
Call rtnJObj.init(Nothing)
''
If Not VBA.IsArray(ary) Then
    Console.info "Can not parse because ary is not an Array. : function C_JObject.ConverttoArray."
    Set ConvertfromArray = rtnJObj
    Exit Function
End If
Dim json_str As String
Let json_str = C_Array.ToJson(ary)
'Console.info json_str
''
Set rtnJObj = C_JObj.parse(json_str)
Set ConvertfromArray = rtnJObj
''
End Function

Public Function ConvertfromArrayArrayToArrayDictionary(ByRef aryary As Variant, _
                    Optional aryHeader As Variant) As cJobject
'''' *************************************************
''
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim rtnJObj As cJobject
Set rtnJObj = New cJobject
Call rtnJObj.init(Nothing)
''
If Not VBA.IsArray(aryary) Then
    Console.info "Can not parse because ary is not an Array."  ''FIXME
    Set ConvertfromArrayArrayToArrayDictionary = rtnJObj
    Exit Function
End If
''
Dim jagary As Variant
Let jagary = C_Array.ArrayArrayToJagArray(aryary)
''
If IsMissing(aryHeader) Then
    Let aryHeader = jagary(LBound(jagary))
    Call C_Array.DeleteElementAt(jagary, LBound(jagary))
End If
If LBound(jagary) = 0 Then
    Let jagary = C_Array.ChangeIndexOfBegin(jagary)
End If
''
Dim ary As Variant
Let ary = C_Array.JagArrayToArrayDictionary(jagary, aryHeader)
Dim json_str As String
Let json_str = JsonConverter.ConvertToJson(ary)
'Console.info json_str
''
Set rtnJObj = C_JObj.parse(json_str)
Set ConvertfromArrayArrayToArrayDictionary = rtnJObj
''
End Function

Public Function ConvertfromArrayArrayToDictionary(ByRef aryary As Variant, _
                    Optional colKeyNum As Variant, Optional aryHeader As Variant) As cJobject
'''' *************************************************
''
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim rtnJObj As cJobject
Set rtnJObj = New cJobject
Call rtnJObj.init(Nothing)
''
If Not VBA.IsArray(aryary) Then
    Console.info "Can not parse because ary is not an Array."  ''FIXME
    Set ConvertfromArrayArrayToDictionary = rtnJObj
    Exit Function
End If
''
Dim jagary As Variant
Let jagary = C_Array.ArrayArrayToJagArray(aryary)
''
Dim colKey As String
If IsMissing(colKeyNum) Then
    Let colKeyNum = 1
End If
If IsMissing(aryHeader) Then
    Let aryHeader = jagary(LBound(jagary))
    Call C_Array.DeleteElementAt(jagary, LBound(jagary))
End If
If LBound(jagary) = 0 Then
    Let jagary = C_Array.ChangeIndexOfBegin(jagary)
End If
''
Dim dict As Dictionary
Set dict = New Dictionary
Set dict = C_Array.JagArrayToDictionary(jagary, colKeyNum, aryHeader)
Dim json_str As String
Let json_str = JsonConverter.ConvertToJson(dict)
'Console.info json_str
''
Set rtnJObj = C_JObj.parse(json_str)
Set ConvertfromArrayArrayToDictionary = rtnJObj
''
End Function

