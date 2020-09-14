VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Commons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Commons.cls
'''  *<pre>Common Utility
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  */
'
''' /********************************************************
'''  * @class C_Commons
'''  * @classdesc Common Utility Class
'''  * @require C_String, C_Array, JsonConverter
'''  */
'

Dim C_String As New C_String

'//--------------------------------------------------------
'// Public Method (Function)
'//
Public Function GetElement(Item As Variant, key As Variant) As Variant
''' /********************************************************
'''  * @function getElement
'''  *</br>文字列要素の取得　（あまり利用価値はない：配列なら普通にidxでゲットできる）
'''  * @param {Variant} item 配列か、連想配列か、文字列か、数値か、・・・
'''  * @param {Variant} key 要素index もしくは、key
'''  * @return {Variant} getElement 要素
'''  *
'''  * XXX□□Not Test
'''  */
    If IsArray(Item) Then
        If (LBound(Item) <= key) And (UBound(Item) >= key) Then
            GetElement = Item(key)
        Else
            GetElement = False
        End If
    ElseIf IsObject(Item) Then
        If TypeName(Item) = "Dictionary" Then
                GetElement = Item(key)
        ElseIf TypeName(Item) = "Collection" Then
                GetElement = Item(key)
        Else
                GetElement = False '''FIXME□□他のオブジェクトをどこまでサポートするか？
        End If
    Else
        GetElement = Item '''配列、オブジェクトでないものを渡された場合は、そのまま返却。
    End If
End Function

Public Sub Bind(target, Value)
''' /********************************************************
'''  * @function Bind
'''  *<br/>　target変数に値(value)を設定する。
'''  * @param {Variant} target 【ByRef】
'''  * @param {Variant} value
'''  */
    If IsObject(Value) Then
        Set target = Value
    Else
        target = Value
    End If
End Sub

Public Sub BindAt(dicTarget, key, Value)
''' /********************************************************
'''  * @function Bind
'''  *<br/>　target連想配列変数（もしくは配列）に、
'''  *<br/>　key , value(ValueがObjectにも対応)
'''  *<br/>　を設定する。
'''  * @param {Variant} dicTarget 【ByRef】
'''  * @param {Variant} key
'''  * @param {Variant} value
'''  */
    If IsObject(Value) Then
        Set dicTarget(key) = Value 'Dictionary
    Else
        dicTarget(key) = Value     'Array (keyはindex)
    End If
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'  Common Class Method for ToString
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public Function toString(ByRef Value)
''' /********************************************************
'''  *@function ToString
'''  *<br/> 引数で渡された値を、JSON形式にする。
'''  * 注意：内部でのオブジェクト可視化用。外部Interfaseとやり取りする
'''  *       場合は、外部ライブラリを利用する。
'''  * @param {Variant} value
'''  * @return {Variant} ToString
'''  */
    Dim rtn
    Err.clear
    '
On Error Resume Next
    If VarType(Value) = vbString Then
        If IsObject(Value) Then
            rtn = ShowObject(Value)
        Else
            'strings
            rtn = ShowString(Value)
        End If
    Else
        If IsObject(Value) Then
        'Object
            rtn = ShowObject(Value)
        ElseIf IsArray(Value) Then
        'Array
            rtn = ShowArray(Value)
        ElseIf IsEmpty(Value) Then
        'Variant Empyt
            rtn = "<empty>"
        ElseIf IsNull(Value) Then
        'Variant Null
            rtn = "<null>"
        Else
        'Except strings （Interger, long, Date, Boolean）
            rtn = ShowOther(Value)
        End If
    End If
    '
    If Err.Number <> 0 Then
        'unkown
        Err.clear
        rtn = ShowUnknown(Value)
    End If
    '
    toString = rtn
    Exit Function
End Function

Public Function ShowString(Value)
''' /********************************************************
'''  * @function ShowString
'''  * @param {Variant} value
'''  * @return {Variant} ShowString
'''  */
  ShowString = """" & Value & """"
End Function

Public Function ShowArray(Value)
''' /********************************************************
'''  * @function ShowArray
'''  * @param {Variant} value
'''  * @return {Variant} ShowArray
'''  */
''FIXME□□RangeをArrayと判断した場合に、対応できていない。
''         Arrayの次元数をチェックして、動作を変える。
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowArray"
    Dim i
    Dim clt As Collection
    Set clt = New Collection
    Dim str As String
    Dim C_Array As C_Array
    Set C_Array = New C_Array
    Dim arytype As String
'
'On Error GoTo EXCEPTION
On Error Resume Next
    Let arytype = C_Array.TypeArrayPattern(Value)
    '配列と判定されたvalueを処理
    If arytype = "ArrayNotInit" Then
        ShowArray = "<" & arytype & "> []"
    ElseIf arytype = "ArrayEmpty" Then
        ShowArray = "<" & arytype & "> []"
    ElseIf arytype = "Array" Then
        For Each i In Value
            clt.Add toString(i)
        Next
        For i = 1 To clt.Count
            str = str & clt.Item(i) & ","
        Next i
        ShowArray = "<" & arytype & "> [" & left(str, Len(str) - 1) & "]"
    ElseIf arytype = "JagArray" Then
        str = ShowJagArray(Value)
        ShowArray = "<" & arytype & "> " & str
    ElseIf arytype = "ArrayArray" Then
        str = ShowArrayArray(Value)
        ShowArray = "<" & arytype & "> " & str
    Else
        ShowArray = "<" & arytype & "> " & toString(Value.Name)
    End If
    
    '
    'ToStringでエラー
    If Err.Number <> 0 Then
        If IsObject(Value) Then
            ShowArray = toString(Value.Name)
        End If
        Err.clear
    End If
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    ' Object開放
    If Err.Number <> 0 Then
        Console.info TypeName(Value)
        Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    End If
    Err.clear
    Set clt = Nothing
    Return
    '
'EXCEPTION:
'    GoSub FINALLY
'    Console.info ("Err.Number = " & Err.Number)
'    Console.info ("Err.Source = " & Err.source)
'    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
'    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
'    Exit Function
    '
End Function

Public Function ShowObject(Value)
''' /********************************************************
'''  * @function ShowObject
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    Dim rtn
    Err.clear
    '
On Error Resume Next
'    Console.info TypeName(value)
'
    'Objectと判定されたvalueを処理
    Select Case TypeName(Value)
        Case "Dictionary"
            rtn = ShowDictionary(Value)
        Case "Collection"
            rtn = ShowCollection(Value)
        Case "ArrayList"
            rtn = ShowArrayList(Value)
        Case "Names"
            rtn = ShowNames(Value)
        Case "Name"
            rtn = ShowName(Value)
        Case "Range"
            rtn = ShowRange(Value)
        Case "ListObjects"
            rtn = ShowListObjects(Value)
        Case "ListObject"
            rtn = ShowListObject(Value)
        Case "ListRows"
            rtn = ShowListRows(Value)
        Case "Workbooks"
            rtn = ShowWorkbooks(Value)
        Case "Workbook"
            rtn = ShowWorkbook(Value)
        Case "Sheets"
            rtn = ShowWorksheets(Value)
        Case "Worksheets"
            rtn = ShowWorksheets(Value)
        Case "Worksheet"
            rtn = ShowWorksheet(Value)
        Case "Shapes"
            rtn = ShowShapes(Value)
        Case "Shape"
            rtn = ShowShape(Value)
        Case "cJobject"
            rtn = ShowCJObject(Value)
        Case Else
            If left(TypeName(Value), 2) = "C_" Then
                rtn = ShowClass(Value)
            ElseIf left(TypeName(Value), 2) = "O_" Then
                rtn = ShowClass(Value)
            ElseIf left(TypeName(Value), 3) = "VO_" Then
                rtn = ShowClass(Value)
            Else
                If Right(TypeName(Value), 10) = "Properties" Then
                    rtn = ShowProperties(Value)
                Else
                    rtn = ShowArray(Value)
                End If
            End If
    End Select
    '
    If Err.Number <> 0 Then
        Err.clear
        '''valueをForEachで回せるObjectとして処理
        rtn = ShowArray(Value)
    End If
    If Err.Number <> 0 Then
        Err.clear
        '''value.ItemsをForEachで回せるObjectとして処理
        rtn = ShowArray(Value.Items)
    End If
    If Err.Number <> 0 Then
        Err.clear
        rtn = "<" & TypeName(Value) & ">"
    End If
    ShowObject = rtn
End Function

Public Function ShowOther(Value)
''' /********************************************************
'''  * @function ShowObject
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    Dim rtn
    Err.clear
On Error Resume Next
    '数値,日付,ブール値と判定されたvalueを処理
    rtn = CStr(Value)
    If Err.Number <> 0 Then
        Err.clear
        rtn = ShowUnknown(Value)
    End If
    ShowOther = rtn
End Function

Public Function ShowUnknown(Value)
''' /********************************************************
'''  * @function ShowObject
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    '''初期化忘れ等で、発生。
    ShowUnknown = "<unknown:" & VarType(Value) & " " & TypeName(Value) & ">"
End Function

Public Function ShowDictionary(Value)
''' /********************************************************
'''  * @function ShowObject
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    'Dictionaryと判定されたvalueを処理
    Dim i
    Dim str
    Dim clt
    Set clt = New Collection
    '
    For Each i In Value.keys
        clt.Add toString(i) & ":" & toString(Value(i))
    Next
    For i = 1 To clt.Count
        str = str & clt.Item(i) & ","
    Next i
    '
    ShowDictionary = "{" & left(str, Len(str) - 1) & "}"
    '
    Set clt = Nothing
End Function

Public Function ShowNames(Value)
''' /********************************************************
'''  * @function ShowNames
'''  * @param {Variant} value
'''  * @return {Variant} ShowNames
'''  */
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowNames"
    '
On Error GoTo EXCEPTION
    Dim nm As Name
    Dim ary() As Variant
    Dim i As Long
    i = -1
    Dim C_String As C_String
    Set C_String = New C_String
    '
    For Each nm In Value
        i = i + 1
        ReDim Preserve ary(i)
        Set ary(i) = createObject("Scripting.Dictionary")
        Call ary(i).Add("Name", nm.Name)
        Call ary(i).Add("RefersTo", nm.RefersTo)
        If Not C_String.ContainWord(nm.RefersTo, "#REF!") Then
            Call ary(i).Add("Value", nm.RefersToRange)
        End If
        Call ary(i).Add("SheetName", C_String.Chop(C_String.SubStringBefore(nm.RefersTo, "!", 2), 1))
    Next nm
    ShowNames = Me.toString(ary)
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    ' Object開放
    Return
    '
EXCEPTION:
    GoSub FINALLY
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Source = " & Err.source)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function
End Function

Public Function ShowName(Value)
''' /********************************************************
'''  * @function ShowName
'''  * @param {Variant} value
'''  * @return {Variant} ShowName
'''  */
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowName"
    '
On Error GoTo EXCEPTION
    Dim C_String As C_String
    Set C_String = New C_String
    Dim dict As Object
    Set dict = createObject("Scripting.Dictionary")
    Call dict.Add("Name", Value.Name)
    Call dict.Add("RefersTo", Value.RefersTo)
    If Not C_String.ContainWord(Value.RefersTo, "#REF!") Then
        Call dict.Add("Value", Value.RefersToRange)
    End If
    Call dict.Add("SheetName", C_String.Chop(C_String.SubStringBefore(Value.RefersTo, "!", 2), 1))
    ShowName = Me.toString(dict)
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    ' Object開放
    Set dict = Nothing
    Return
    '
EXCEPTION:
    GoSub FINALLY
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Source = " & Err.source)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function
End Function

Public Function ShowListObjects(Value)
''' /********************************************************
'''  * @function ShowListObjects
'''  * @param {Variant} value
'''  * @return {Variant} ShowListObjects
'''  */
    'ListObjectsと判定されたvalueを処理 (Names内に参照#REF等がないこと）
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowListObjects"
    '
On Error GoTo EXCEPTION
    '
    Dim itm
    Dim str
    Dim i
    If Value.Count > 0 Then
        For Each itm In Value
            str = str & toString(itm.Name) & ":" & toString(itm.Range.Address) & ","
        Next
        If Err.Number <> 0 Then
            Err.clear
            ShowListObjects = ShowUnknown(Value)
        End If
    Else
        str = ","
    End If
    '
    ShowListObjects = "[" & left(str, Len(str) - 1) & "]"
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    ' Object開放
    Return
    '
EXCEPTION:
    GoSub FINALLY
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Source = " & Err.source)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function
End Function

Public Function ShowListRows(Value)
''' /********************************************************
'''  * @function ShowListRows
'''  * @param {Variant} value
'''  * @return {Variant} ShowListRows
'''  */
    'ShowListRowsと判定されたvalueを処理 (Names内に参照#REF等がないこと）
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowListRows"
    '
On Error GoTo EXCEPTION
    '
    Dim itm
    Dim str
    Dim i
    If Value.Count > 0 Then
        For Each itm In Value
            str = str & toString(itm) & ","
        Next
        If Err.Number <> 0 Then
            Err.clear
            ShowListRows = ShowUnknown(Value)
        End If
    Else
        str = ","
    End If
    '
    ShowListRows = "[" & left(str, Len(str) - 1) & "]"
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    ' Object開放
    Return
    '
EXCEPTION:
    GoSub FINALLY
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Source = " & Err.source)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function
End Function

Public Function ShowProperties(Value)
''' /********************************************************
'''  * @function ShowProperties
'''  * @param {Variant} value
'''  * @return {Variant} ShowProperties
'''  */
    'Propertiesと判定されたvalueを処理
    Dim itm
    Dim str
    Dim i
On Error Resume Next ''FIXME□□itemでValueを取得する際にErrorとなる場合がある
    For Each itm In Value
        str = str & toString(itm.Name) & ":" & toString(itm.Value) & ","
    Next
    '
    If Err.Number <> 0 Then
        Err.clear
'        ShowProperties = ShowUnknown(value)
    End If
    ShowProperties = "[" & left(str, Len(str) - 1) & "]"
End Function

Public Function ShowCollection(Value)
''' /********************************************************
'''  * @function ShowCollection
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
'Collectionと判定されたvalueを処理。Collectionのidxは1から開始。
' keyはCollection Objectから取得できない。より、Arrayと同等の出力。
''
Dim ary() As Variant
ReDim ary(1 To Value.Count)
Dim idx As Long
For idx = 1 To Value.Count
    ary(idx) = C_Commons.toString(Value(idx))
Next idx
ShowCollection = "<Collection> " & VBA.Join(ary)
''
End Function

Public Function ShowArrayList(Value)
''' /********************************************************
'''  * @function ShowArrayList
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
'Collectionと判定されたvalueを処理。Collectionのidxは1から開始。
' keyはCollection Objectから取得できない。より、Arrayと同等の出力。
''
Dim ary() As Variant
ReDim ary(0 To Value.Count - 1)
Dim idx As Long
For idx = 0 To Value.Count - 1
    ary(idx) = C_Commons.toString(Value.Item(idx))
Next idx
ShowArrayList = "<ArrayList> " & VBA.Join(ary, ", ")
''
End Function

Public Function ShowWorkbooks(Value)
''' /********************************************************
'''  * @function ShowWorkbooks
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    Dim bk As Workbook
    Dim ary() As Variant
    Dim i As Long
    i = -1
    For Each bk In Value
        i = i + 1
        ReDim Preserve ary(i)
        Set ary(i) = createObject("Scripting.Dictionary")
        Call ary(i).Add("CodeName", bk.codeName)
        Call ary(i).Add("Name", bk.Name)
        Call ary(i).Add("Path", bk.path)
    Next bk
    ShowWorkbooks = Me.toString(ary)
End Function

Public Function ShowWorkbook(Value)
''' /********************************************************
'''  * @function ShowWorkbook
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    Dim dict As Object
    Set dict = createObject("Scripting.Dictionary")
    Call dict.Add("CodeName", Value.codeName)
    Call dict.Add("Name", Value.Name)
    Call dict.Add("Path", Value.path)
    Call dict.Add("BuiltinDocumentProperties", Value.BuiltinDocumentProperties)
    ShowWorkbook = Me.toString(dict)
    Set dict = Nothing
End Function

Public Function ShowWorksheets(Value)
''' /********************************************************
'''  * @function ShowWorksheets
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    Dim sht As Worksheet
    Dim ary() As Variant
    Dim i As Long
    i = -1
    For Each sht In Value
        i = i + 1
        ReDim Preserve ary(i)
        Set ary(i) = createObject("Scripting.Dictionary")
        Call ary(i).Add("Name", sht.Name)
        Call ary(i).Add("Index", sht.Index)
    Next sht
    ShowWorksheets = "<Worksheets> " & VBA.Right(Me.toString(ary), Len(Me.toString(ary)) - 8)
End Function

Public Function ShowWorksheet(Value)
''' /********************************************************
'''  * @function ShowWorksheet
'''  * @param {Variant} value
'''  * @return {Variant} ShowObject
'''  */
    Dim dict As Object
    Set dict = createObject("Scripting.Dictionary")
    Call dict.Add("Name", Value.Name)
    Call dict.Add("Index", Value.Index)
    Call dict.Add("UsedRange", Value.UsedRange.Address)
    Call dict.Add("Parent", Value.Parent.Name)
    Call dict.Add("CustomProperties", Value.CustomProperties)
    ShowWorksheet = Me.toString(dict)
    Set dict = Nothing
End Function

Public Function ShowShapes(Value)
''' /********************************************************
'''  * @function ShowShapes
'''  * @param {Variant} value
'''  * @return {Variant} ShowShapes
'''  */
    Dim shp As Shape
    Dim ary() As Variant
    Dim i As Long
    i = -1
    For Each shp In Value
        i = i + 1
        ReDim Preserve ary(i)
        Set ary(i) = createObject("Scripting.Dictionary")
        Call ary(i).Add("Name", shp.Name)
        Call ary(i).Add("SheetName", shp.Parent.Name)
        Call ary(i).Add("Text", shp.TextFrame.Characters.text)
'        Call ary(i).add("Type", shp.Type)
'        Call ary(i).add("TopLeftCell", shp.TopLeftCell.Address)
'        Call ary(i).add("BottomRightCell", shp.BottomRightCell.Address)
'        Call ary(i).add("FilBackColor", shp.Fill.BackColor.RGB)
    Next shp
    ShowShapes = Me.toString(ary)
End Function

Public Function ShowShape(Value)
''' /********************************************************
'''  * @function ShowShape
'''  * @param {Variant} value
'''  * @return {Variant} ShowShape
'''  */
    Dim dict As Object
    Set dict = createObject("Scripting.Dictionary")
    Call dict.Add("Name", Value.Name)
    Call dict.Add("SheetName", Value.Parent.Name)
    Call dict.Add("Text", Value.TextFrame.Characters.text)
    Call dict.Add("Type", Value.Type)
    Call dict.Add("Id", Value.id)
    Call dict.Add("TopLeftCell", Value.topLeftCell.Address)
    Call dict.Add("BottomRightCell", Value.bottomRightCell.Address)
    Call dict.Add("FillForeColor", Me.RGBToHexStr(Value.Fill.ForeColor.RGB))
    Call dict.Add("FillBackColor", Me.RGBToHexStr(Value.Fill.BackColor.RGB))
    Call dict.Add("Top", Value.top)
    Call dict.Add("Left", Value.left)
    Call dict.Add("Height", Value.Height)
    Call dict.Add("Width", Value.Width)
    ShowShape = Me.toString(dict)
End Function

Public Function ShowClass(Value)
''' /********************************************************
'''  * @function ShowClass
'''  * @param {Variant} value
'''  * @return {Variant} ShowClass
'''  */
    '独自クラスと判定されたvalueを処理。toSourceを必ず実装していること
    ShowClass = Value.ToSource
End Function

Public Function ShowRange(ByRef Value)
''' /********************************************************
'''  * @function ShowRange
'''  * @param {Variant} value As Range
'''  * @return {Variant} ShowRange
'''  */
    '' On evaluting as a Range
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowRange"
    ''
On Error GoTo EXCEPTION
    ''
    Dim aryary() As Variant
    Dim ary As Variant
    ''
    Dim ary2 As Variant
    Dim ary3() As Variant
    Dim rng As Range
    Dim k As Long
    If Value.Areas.Count > 1 Then
        k = 0
        For Each rng In Value.Areas
            ReDim Preserve ary3(k)
            ary3(k) = Me.toString(rng)
            k = k + 1
        Next rng
        ShowRange = "<Range> [" & Join(ary3, ",") & "]"
        Exit Function
    End If
    ''
    Dim rowCnt As Long
    rowCnt = Value.Rows.Count
    Dim colCnt As Long
    colCnt = Value.Columns.Count
    '
    ''a Cell
    If rowCnt = 1 And colCnt = 1 Then
        ShowRange = "<Range> " & C_Commons.toString(Value.Value)
        Exit Function
    End If
    '
    ''some Cells
    Let aryary = Value.Value
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    Dim temp As String
    temp = ShowArrayArray(aryary)
    temp = C_String.Chomp(C_String.Chop(temp, 1), 1)
    ShowRange = "<Range> [" & temp & "]"
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Source = " & Err.source)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function
    '
End Function

Public Function ShowListObject(ByRef Value)
''' /********************************************************
'''  * @function ShowListObject
'''  * @param {Variant} value
'''  * @return {Variant} ShowListObject
'''  */
    'C_xxx（独自クラス）と判定されたvalueを処理。toSourceを必ず実装していること
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowListObject"
    '
On Error GoTo EXCEPTION
    '
    Dim rng As Range
    Set rng = Value.Range
    '
    Dim jagary As Variant
    Dim ary As Variant
    '
    Dim rowCnt As Long
    rowCnt = rng.Rows.Count
    Dim colCnt As Long
    colCnt = rng.Columns.Count
    '
    ReDim jagary(rowCnt - 1)
    ReDim ary(colCnt - 1)
    '
    Dim i
    Dim j
    '
    For i = 0 To rowCnt - 1
        jagary(i) = ary
        For j = 0 To colCnt - 1
            Let jagary(i)(j) = rng(i + 1, j + 1)
        Next
    Next
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    ShowListObject = ShowJagArray(jagary)
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Source = " & Err.source)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    Exit Function
    '
End Function

Public Function ShowJagArray(jagary) As Variant
''' /********************************************************
'''  * @function ShowJagArray
'''  * @param {Variant} value
'''  * @return {Variant} ShowJagArray
'''  *　注意：Range/ListObjectのJagArrayしたものを処理する専用
'''  *　注意：文字列化（dump）のための実装。JSON形式ではない。
'''  */
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowJagArray"
    Dim lbjagFirst As Long
    Dim ubjagFirst As Long
    Dim clt As Collection
    Set clt = New Collection
    Dim elm As Variant
    '
    Dim str As String
    Dim strTemp As String
    '
On Error GoTo EXCEPTION
    'jagArrayを処理する
    If UBound(jagary) <> -1 Then
        lbjagFirst = LBound(jagary(LBound(jagary))) '''jagAryの最初のArray
        ubjagFirst = UBound(jagary(LBound(jagary)))
        '
        ''各値を取得
        Dim r As Long
        Dim c As Long
        For r = LBound(jagary) To UBound(jagary)
            If LBound(jagary(r)) <> lbjagFirst Or UBound(jagary(r)) <> ubjagFirst Then
                Err.Raise 5
            End If
            For c = LBound(jagary(r)) To UBound(jagary(r))
                Let elm = jagary(r)(c)
                If C_String.IsJsonString(elm) Then
                    str = str & elm & ","
                Else
                    str = str & Me.toString(elm) & ","
                End If
            Next c
            clt.Add "[" & left(str, Len(str) - 1) & "]"
            str = ""
        Next r
        ''
        Dim Item As Variant
        For Each Item In clt
            strTemp = strTemp & Item & ","
        Next Item
        ''
        'Console.log strTemp
        ShowJagArray = "[" & left(strTemp, Len(strTemp) - 1) & "]"
    Else
    '配列が空
        ShowJagArray = "[]"
    End If
    '
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    '
    Return
    '
EXCEPTION:   'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
End Function

Public Function ShowArrayArray(aryary As Variant) As String
''' /********************************************************
''''  * @function ShowArrayArray
''''  * @param {Variant} value
''''  * @return {Variant}
''''  * @note : for arrayarray of primitive values
''''  */
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.ShowArrayArray"
    ''
    Dim r As Long
    Dim c As Long
    Dim strRtn As String
On Error GoTo EXCEPTION
    ''
    For r = LBound(aryary, 1) To UBound(aryary, 1)
        For c = LBound(aryary, 2) To UBound(aryary, 2)
            'Console.log (aryary(r, c))
            strRtn = strRtn & toString(aryary(r, c)) & ","
        Next c
        strRtn = C_String.Chop(strRtn, 1) & ";"
    Next r
    strRtn = "[{" & C_String.Chop(strRtn, 1) & "}]"
    ShowArrayArray = strRtn
    ''
    GoSub FINALLY
    Exit Function
    ''
FINALLY:    ''Label
    ''
    Return
    ''
EXCEPTION:   ''Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
End Function

Public Function ShowCJObject(Value)
''' /********************************************************
'''  * @function ShowCJObject
'''  * @param {Variant} value
'''  * @return {Variant} ShowShape
'''  */
Dim json_str As String
Dim json_obj As Variant
Let json_str = Value.stringify(True)
''Console.log json_str
Let ShowCJObject = "<cJobject> " & vbCrLf & JsonConverter.ConvertToJson(JsonConverter.ParseJson(json_str))
''
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'  jsonStringをObject化する共通クラスメソッド
'  ＜注意＞
'  外部と連携しないで、内部で配列とオブジェクトのJSON形式
'  を利用する場合のみ、使用する。
'    JSON形式文字列の完全なるparseは、外部ライブラリを利用する。
'    ※JsonLib
'      https://github.com/VBA-tools/VBA-JSON
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public Function ParseToJson(ByVal str As String) As Variant
'''/********************************************************
''' * @function parseToString
''' * @param {String} str jsonString
''' * @return {Variant} 生成オブジェクト（配列もしくは連想配列）
''' *
''' */
ParseToJson = JsonConverter.ParseJson(str)
'
End Function

Public Function ParseToString(ByVal str As String) As Variant
'''/********************************************************
''' * @function parseToString
''' * @param {String} str jsonString
''' * @return {Variant} 生成オブジェクト（配列もしくは連想配列）
''' *
''' */
Dim rtn As Variant
'
If (left(Trim(str), 1) = "[") Then
        ParseToString = ParseToArray(str)
ElseIf (left(Trim(str), 1) = "{") Then
        Set ParseToString = ParseToDictionary(str)
Else
        ParseToString = str
End If
'
End Function

Public Function ParseToArray(ByVal str As String, Optional delimiter As Variant = ",") As Variant
'''/********************************************************
''' * @function parseToArray
''' * @param {String} str jsonString
''' * @param {Variant} delimiter 【Optional】
''' * @return {Variant} parseToArray 配列オブジェクト
''' *
''' */
    Dim ary As Variant
    Dim aryTemp As Variant
    Dim strTemp As String
    Dim strTemp2 As String
    Dim i As Long
    If C_String.StartWith(Trim(str)) = "[" And C_String.EndWith(Trim(str)) = "]" Then
        strTemp = C_String.ChopDouble(Trim(str))
    Else
        strTemp = Trim(str)
    End If
    If delimiter = " " Then
        strTemp2 = C_String.SqueezeSpace(strTemp)
    ElseIf delimiter = vbLf Then
        strTemp2 = C_String.DeleteSpace(strTemp)
    Else
        strTemp2 = C_String.DeleteWhiteSpace(strTemp)
    End If
    aryTemp = Split(strTemp2, delimiter)
    '
    For i = LBound(aryTemp) To UBound(aryTemp)
        If C_String.StartWith(aryTemp(i)) = """" And C_String.EndWith(aryTemp(i)) = """" Then
            aryTemp(i) = C_String.TrimSpace(C_String.ChopDouble(aryTemp(i)))
        Else
            aryTemp(i) = C_String.TrimSpace(aryTemp(i))
        End If
    Next i
    '
    ParseToArray = aryTemp
End Function

Public Function ParseToDictionary(ByVal str As String) As Variant
'''/********************************************************
''' * @function parseToDictionary
''' * @param {String} str jsonString
''' * @return {Variant} parseToDictionary 連想配列オブジェクト
''' */
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Commons.parseToDictionary"
    Dim dict As Object
    Set dict = createObject("Scripting.Dictionary")
    Dim strTemp As String
    Dim flag As Boolean
    Dim ary  As Variant
    Dim aryTemp As Variant
    Dim i As Long
    Dim key As Variant
    Dim Value As Variant
    Dim keyTemp As Variant
    Dim valueTemp As Variant
    '
On Error GoTo EXCEPTION
    '//--------------------------------------
    '// 処理
    flag = True
    '
    If C_String.StartWith(Trim(str)) = "{" And C_String.EndWith(Trim(str)) = "}" Then
        strTemp = C_String.ChopDouble(Trim(str))
    Else
        flag = False
    End If
    ary = Split(strTemp, ",")
    For i = 0 To UBound(ary)
        If C_String.Contain(ary(i), ":") Then
            aryTemp = Split(ary(i), ":")
            keyTemp = aryTemp(0)
            valueTemp = aryTemp(1)
            Erase aryTemp
            keyTemp = C_String.Strip(keyTemp, """")
            valueTemp = C_String.Strip(valueTemp, """")
            key = Trim(C_String.TrimSpace(keyTemp))
            Value = Trim(valueTemp)
            dict.Add Trim(key), Trim(Value)
        Else
            dict.Add Trim(ary(i)), ""
        End If
    Next i
    '
    If flag Then
        Set ParseToDictionary = dict
    Else
        ParseToDictionary = Nothing
    End If
    '
    GoSub FINALLY
    Exit Function
FINALLY: 'Label
    '//--------------------------------------
    '// 終了処理
    Set dict = Nothing
    Return
    '
EXCEPTION: 'Label
    '//--------------------------------------
    ' 例外処理
    GoSub FINALLY
    Console.info ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    '
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'  jsonStringをObject化する共通クラスメソッド（Eval利用）
'  ＜注意＞
'  外部と連携しないで、内部で配列とオブジェクトのJSON形式
'  を利用する場合のみ、使用する。
'  　JSON形式文字列の完全なるparseは、外部ライブラリを利用する。
'  　※JsonLib
'  　https://code.google.com/p/vba-json/
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Public Function ToJson(ByRef vnt As Variant) As String
''' FIXME
ToJson = Me.toString(vnt)
End Function

Public Function parse(ByVal jsonString As Variant) As Object
'''/********************************************************
''' * @function parse
''' * @param {String} str jsonString
''' * @return {Object} parse オブジェクト
''' */
Dim objSC As Object
Dim objJson As Object
'
Set objSC = createObject("ScriptControl")
objSC.Language = "JScript"
'
objSC.AddCode "function toObject(str){return eval('(' + str + ')' ); } "
'
Set objJson = objSC.CodeObject.ToObject(jsonString)
'
Set parse = objJson
'
Set objSC = Nothing
Set objJson = Nothing
'
End Function

Private Function EvalScript(ByVal expr As String, ByVal lang As String) As String
    Dim doc As Object: Set doc = createObject("HtmlFile")
    doc.parentWindow.execScript "document.write(" & expr & ")", lang
    If Not doc.body Is Nothing Then EvalScript = doc.body.innerHTML
End Function

''' @param vbsExpr As String
''' @return As Variant
Public Function EvalVBScipt(ByVal vbsExpr As String) As String
    EvalVBScipt = EvalScript(vbsExpr, "VBScript")
End Function

''' @param jsExpr As String
''' @return As Variant
Public Function EvalJScript(ByVal jsExpr As String) As Variant
'''/********************************************************
''' * @function EvalJScript
''' * @param {String} str jsonString
''' * @return {Variant} EvalJScript オブジェクト
''' */
    EvalJScript = EvalScript(jsExpr, "JScript")
End Function

Public Function EvalJScriptToObject(ByVal jsExpr As String) As Object
'''/********************************************************
''' * @function EvalJScriptToObject
''' * @param {String} str jsonString
''' * @return {Object} EvalJScript オブジェクト
''' */
    Set EvalJScriptToObject = EvalJScript(jsExpr)
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'  Other Common Class Method
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

Public Function RGBToHexStr(lngRGB As Long) As String
'''/********************************************************
''' * @function RGBToHexStr
''' * @param {Long} lngRGB
''' * @return {String} RGBTo16
''' */
''
Dim strR As String
Dim strG As String
Dim strB As String
strR = Right$("00" & Hex$((lngRGB And "&H" & "0000FF")), 2)
strG = Right$("00" & Hex$((lngRGB And "&H" & "00FF00") / 256), 2)
strB = Right$("00" & Hex$((lngRGB And "&H" & "FF0000") / 256 ^ 2), 2)
RGBToHexStr = "#" & strR & strG & strB
''
End Function

Public Function HexStrToRGB(strRGB As String) As Long
'''/********************************************************
''' * @function HexStrToRGB
''' * @param {String} strRGB
''' * @return {Long}
''' */
' FIXME
Dim lngR As String
Dim lngG As String
Dim lngB As String
Dim strR As String
Dim strG As String
Dim strB As String
strR = VBA.Mid(strRGB, 2, 2)
strG = VBA.Mid(strRGB, 4, 2)
strB = VBA.Mid(strRGB, 6, 2)
lngR = VBA.CLng("&H" & strR)
lngG = VBA.CLng("&H" & strG)
lngB = VBA.CLng("&H" & strB)
HexStrToRGB = RGB(lngR, lngG, lngB)
''
End Function

Public Function CreateGUID() As String
''' *************************************************
Do While Len(CreateGUID) < 32
    If Len(CreateGUID) = 16 Then
        '17th character holds version information
        CreateGUID = CreateGUID & Hex$(8 + CInt(Rnd * 3))
    End If
    CreateGUID = CreateGUID & Hex$(CInt(Rnd * 15))
Loop
CreateGUID = Mid(CreateGUID, 1, 8) & "-" & Mid(CreateGUID, 9, 4) & "-" & Mid(CreateGUID, 13, 4) & "-" & Mid(CreateGUID, 17, 4) & "-" & Mid(CreateGUID, 21, 12)
'
End Function

Public Function TypeDataInCell(rng As Range) As String
''' *************************************************
Application.Volatile
Select Case True
    Case VBA.IsEmpty(rng): TypeDataInCell = "Empty"
    Case Application.IsText(rng): TypeDataInCell = "String"
    Case Application.IsLogical(rng): TypeDataInCell = "Boolean"
    Case Application.IsErr(rng): TypeDataInCell = "Error"
    Case VBA.IsDate(rng): TypeDataInCell = "Date"
    Case VBA.InStr(1, rng.text, ":") <> 0: TypeDataInCell = "Time"
    Case VBA.IsNumeric(rng): TypeDataInCell = "Number"
End Select
'
End Function

'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Sub UnitTest()
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    Dim ary1 As Variant
    Dim ary2 As Variant
    Dim ary3 As Variant
    str1 = "[sak , mit , 　saki , miss , ask , ao ]"
    str2 = " sak   mit   saki　   miss   ask   ao  "
'    str3 = Range(ThisWorkbook.Names("TEST_STRING").RefersTo).text
    ary1 = Me.ParseToArray(str1)
    ary2 = Me.ParseToArray(str2, " ")
'    ary3 = Me.parseToArray(str3, vbLf)
    Console.Dump ary1
    Console.Dump ary2
    Console.Dump ary3
    '
    Dim strr1 As String
    Dim obj1 As Variant
    strr1 = "{saka : mits , 　aki : miss , sak2 : ao }"
    Set obj1 = Me.ParseToDictionary(strr1)
    Console.Dump obj1
    '
    Dim strr2 As String
    Dim obj2 As Variant
    strr2 = "{""sak1"" : mit ,     　""sass ai"" : miss , ""sak2"" : ao ""}"
    Set obj2 = Me.ParseToDictionary(strr2)
    Console.Dump obj2
    '
    Dim strr3 As String
    Dim obj3 As Variant
    Dim dict As Object
    Set dict = createObject("Scripting.Dictionary")
    dict.Add "sak", "mit"
    strr3 = C_Commons.toString(dict)
    '
    Set obj3 = Me.ParseToDictionary(strr3) ''両方に"があるものは対応可能
    Console.Dump obj3
    '
    Console.Dump Workbooks
    Console.Dump ThisWorkbook.Worksheets
    Console.Dump ThisWorkbook.Names
    Console.Dump ActiveSheet.ListObjects
    '
    Console.Dump Me.ParseToString("saksak")
    '
''''RGBToHexStr,HexStrToRGB
'Dim lng As Long
'Dim str As String
'str = C_Commons.RGBToHexStr(16777215)
'Console.log VBA.IIf(str = "#FFFFFF", "OK", "NG :->" & str)
'str = C_Commons.RGBToHexStr(RGB(255, 255, 255))
'Console.log VBA.IIf(str = "#FFFFFF", "OK", "NG :->" & str)
'lng = C_Commons.HexStrToRGB("#FFFF00")
'Console.log VBA.IIf(lng = 65535, "OK", "NG :->" & lng)
'lng = C_Commons.HexStrToRGB("#FFFFFF")
''Console.log VBA.IIf(lng <> 16777215, "OK", "NG :->" & lng)
'Console.log VBA.IIf(lng = 16777215, "OK", "NG :->" & lng)

''' CreateGUID
'str = C_Commons.CreateGUID()
'Console.log str
End Sub

