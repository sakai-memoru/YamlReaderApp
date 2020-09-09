VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "O_DataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file O_DataSet.cls
'''  * Common Data Set Object for this Framework
'''  * @version 1.00
'''  * @since 2020/04/01
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  *
'''  * @Require C_String,C_Collection,C_Range
'''  */
'
''' /********************************************************
'''  * @class O_DataSet
'''  * @classdesc <pre>Common Data Set Object for this Framework
'''  * Common Variables Object (Data Set) in sheets
'''  * You may set key strings with data
'''  *</pre>
'''  */
'
'//--------------------------------------------------------
'// Class Property
'//
''' /** @property @private @name clt_ @type {Object<Collection>} @desc object  */
Private clt_ As Collection
''' /** @property @private @name key_ @type {Object<Collection>} @desc key  */
Private key_ As Collection

'//--------------------------------------------------------
'// Class Event
'//
Private Sub Class_Initialize()
'''' /********************************************************
''''  * @event Class_Initialize
''''  */
''Constructor
Set clt_ = New Collection
Set key_ = New Collection
''
End Sub

Private Sub Class_Terminated()
'''' /********************************************************
''''  *Class_Terminate
''''  *<br/>Terminate 関数
''''  *<br/>Class廃棄時に、実行される。
''''  */
''Deconstructor
Set clt_ = Nothing
Set key_ = Nothing
''
End Sub

Public Function toString() As String
'''' /********************************************************
''''  *@function ToString
''''  *@return {String}
''''  */
Dim C_String As C_String
Set C_String = New C_String
''
Dim str As String
Dim strTempKey As String
Dim i As Long
For i = 1 To clt_.Count
    If C_String.IsNumeric(key_.Item(i)) Then
        str = str & C_Commons.toString(clt_.Item(i)) & ","
    Else
        str = str & C_Commons.toString(key_.Item(i)) & ":" & C_Commons.toString(clt_.Item(i)) & ","
    End If
Next i
toString = left(str, Len(str) - 1)
Set C_String = Nothing
''
End Function

Public Function ToSource() As String
'''' /********************************************************
''''  *@function ToString
''''  *@return {String}
''''  */
Dim i As Long
Dim str As String
str = Me.toString()
ToSource = "O_DataSet Count= " & clt_.Count & " [" & str & "]"
''
End Function

Public Function Count() As Long
'''' /********************************************************
''''  *@function Count
''''  *@return {Long} Count
''''  */
Count = clt_.Count
''
End Function

Public Function Exists(key_or_idx As Variant) As Boolean
'''' /********************************************************
''''  *@function Exists
''''  * If key or idx exists
''''  *@return {Variant} key_or_idx
''''  *@return {Boolean} Exists
''''  * FIXME□No Test
''''  */
Dim blnRtn As Boolean
Dim C_Collection As C_Collection
Set C_Collection = New C_Collection
''
If IsNumeric(key_or_idx) Then
    If CLng(key_or_idx) > clt_.Count Then
        blnRtn = False
    Else
        blnRtn = C_Collection.Exists(key_, key_or_idx)
    End If
Else
    blnRtn = C_Collection.Exists(key_, key_or_idx)
End If
Exists = blnRtn
Set C_Collection = Nothing
''
End Function

Public Function GetValue(Optional key_or_idx As Variant) As Variant
'''' /********************************************************
''''  *@function getValue
''''  *@param {Variant<String/Long>} key_or_idx
''''  *@return {Variant<String/Long> or Variant(i) or Variant<Dictionary>}
''''  */
If IsMissing(key_or_idx) Then
    If IsObject(clt_.Item(clt_.Count)) Then
        Set GetValue = clt_.Item(clt_.Count)
    Else
        GetValue = clt_.Item(clt_.Count)
    End If
Else
    If IsObject(clt_.Item(key_or_idx)) Then
        Set GetValue = clt_.Item(key_or_idx)
    Else
        GetValue = clt_.Item(key_or_idx)
    End If
End If
''
End Function

Public Function ToDictionary() As Dictionary
'''' /********************************************************
''''  *@function ToDictionary
''''  *@return {Dictionary}
''''  */
Dim dict As Dictionary
Set dict = New Dictionary
Dim ky As Variant
For Each ky In key_
    Call dict.Add(key_(ky), clt_(ky))
Next ky
Set ToDictionary = dict
Set dict = Nothing
''
End Function

Public Function Extract(aryHeader As Variant) As Dictionary
'''' /********************************************************
''''  *@function Extract
''''  *@param {Variant<String or Long>()} aryHeader
''''  *@return {Dictionary} Extract
''''  * FIXME□No Test
''''  */
Dim dict As Dictionary
Set dict = New Dictionary
''
Dim i As Long
For i = LBound(aryHeader) To UBound(aryHeader)
    If Me.Exists(aryHeader(i)) Then
        Call dict.Add(aryHeader(i), clt_.Item(aryHeader(i)))
    Else
        Call dict.Add(aryHeader(i), "")
    End If
Next i
Set Extract = dict
Set dict = Nothing
''
End Function

Public Function ExtractToArray(aryHeader As Variant) As Dictionary
'''' /********************************************************
''''  *@function ExtractToArray
''''  *@param {Variant<String or Long>()} aryHeader
''''  *@return {Variant<String or Long>()} ExtractToArray
''''  * FIXME□No Test
''''  */
''
Dim dict As Dictionary
Set dict = Me.Extract(aryHeader)
Set ExtractToArray = dict.Items
Set dict = Nothing
''
End Function

Public Function PutValue(Value As Variant, Optional key As Variant) As O_DataSet
'''' /********************************************************
''''  *@function putValue
''''  *@param value
''''  *@param key
''''  *@return {Object<O_DataSet>}
''''  */
Dim C_Collection As C_Collection
Set C_Collection = New C_Collection
If IsMissing(key) Then
    key = CStr(clt_.Count + 1)
End If
If C_Collection.Exists(clt_, key) Then
    Call clt_.remove(key)
    Call key_.remove(key)
End If

Call clt_.Add(Value, key)
Call key_.Add(key, key)
Set PutValue = Me
''
End Function

Public Function PutByKeyword(ByVal sheetName As Variant _
                    , ByVal searchWord As Variant _
                    , Optional ByVal key As Variant _
                    , Optional Row As Variant = 0 _
                    , Optional column As Variant = 1 _
                    , Optional wbName As Variant _
                    ) As O_DataSet
'''' /********************************************************
''''  *@function putByKeyword
''''  *@param sheetname
''''  *@param searchWord
''''  *@param key
''''  *@param offset row
''''  *@param offset column
''''  *@return {Object<O_DataSet>}
''''  */
Dim C_String As C_String
Set C_String = New C_String
''
If IsMissing(wbName) Then
    Let wbName = ThisWorkbook.Name
End If
Dim wb As Workbook
Set wb = Workbooks(wbName)
''
Dim rngFindResult As Range
Dim rngFindFirst  As Range
Dim rngUsed As Range
Set rngUsed = wb.Worksheets(sheetName).UsedRange
Set rngFindResult = rngUsed.find(what:=searchWord, Lookat:=xlWhole)
Set rngFindFirst = rngFindResult
Dim strTemp As String
Dim ary() As Variant
Dim i As Long
Let i = -1

If IsMissing(key) Then
    key = CStr(clt_.Count + 1)
End If

If rngFindResult Is Nothing Then
    i = i + 1
    ReDim ary(i)
    ary(i) = ""
Else
    Do
        i = i + 1
        ReDim Preserve ary(i)
        Let strTemp = C_String.TrimSpace(rngFindResult.Offset(Row, column).Value)
        If C_String.StartWith(strTemp) = "{" And C_String.EndWith(strTemp) = "}" Then
            Set ary(i) = C_Commons.ParseToJson(strTemp)
        Else
            Let ary(i) = strTemp
        End If
        
        Set rngFindResult = rngUsed.FindNext(rngFindResult)
    Loop While (rngFindResult.Address <> rngFindFirst.Address)
End If

If (UBound(ary) - LBound(ary) + 1) = 1 Then
    Call Me.PutValue(ary(LBound(ary)), key)
Else
    Call Me.PutValue(ary, key)
End If
Set PutByKeyword = Me
''
End Function

Public Function PutByName(ByVal sheetName As Variant _
                    , ByVal Name As Variant _
                    , Optional ByVal key As Variant _
                    , Optional Row As Long = 0 _
                    , Optional column As Long = 0 _
                    , Optional wbName As Variant _
                    ) As O_DataSet
'''' /********************************************************
''''  *@function putByName
''''  *@param sheetname
''''  *@param name
''''  *@param key
''''  *@param row
''''  *@param column
''''  *@return {Object<O_DataSet>} 文字列
''''  */
Dim C_Range As C_Range
Set C_Range = New C_Range
Dim C_String As C_String
Set C_String = New C_String
''
If IsMissing(wbName) Then
    Let wbName = ThisWorkbook.Name
End If
Dim wb As Workbook
Set wb = Workbooks(wbName)
''
If IsMissing(key) Then
    key = Name
End If
Dim vnt As Variant
Dim rng As Range
Set rng = wb.Names(Name).RefersToRange.Offset(Row, column)
Dim val As String
Let val = rng.Value
''
If C_String.IsJsonString(val) Then
    Set vnt = JsonConverter.ParseJson(val)
Else
    Let vnt = val
End If
''
Call Me.PutValue(vnt, key)
Set PutByName = Me
''
End Function

Public Function PutByCurrentTable( _
                      ByVal sheetName As Variant _
                    , ByVal searchWord As Variant _
                    , Optional ByVal key As Variant _
                    , Optional Row As Long = 0 _
                    , Optional column As Long = 1 _
                    , Optional wbName As Variant _
                    ) As O_DataSet
'''' ********************************************************
Dim C_Range As C_Range
Set C_Range = New C_Range
Dim C_Array As C_Array
Set C_Array = New C_Array
''
If IsMissing(wbName) Then
    Let wbName = ThisWorkbook.Name
End If
''
If IsMissing(key) Then
    key = CStr(clt_.Count + 1)
End If
Dim rng As Range
Dim ary() As Variant
Set rng = C_Range.GetCurrentRegionByKeyword(searchWord, sheetName, Row, column, bookName:=wbName)
Dim aryary() As Variant
Let aryary = C_Array.FilterRowForArrayArray(rng.Value, 1, 2)
Let ary = C_Array.ArrayArrayToArrayDictionary(aryary)

Call Me.PutValue(ary, key)
Set PutByCurrentTable = Me
End Function

Public Function PutByRange(ByRef rng As Range _
                    , Optional ByVal key As Variant _
                    , Optional Row As Long = 0 _
                    , Optional column As Long = 0) As O_DataSet
'''' /********************************************************
''''  *@function putByRange
''''  *@param {Range} rng
''''  *@param {Variant<string>} key
''''  *@param {Variant<Long>} row
''''  *@param {Variant<Long>} column
''''  *@return {Object<O_DataSet>} 文字列
''''  * FIXME
''''  */
Dim C_Range As C_Range
Set C_Range = New C_Range
If IsMissing(key) Then
    key = CStr(clt_.Count + 1)
End If
Call Me.PutValue(C_Range.ToJsonObject(rng.Offset(Row, column)), key)
Set PutByRange = Me
''
End Function

'Public Function appendArray(Optional ByVal key As Variant _
'                        , Optional cnt As Variant) As Variant
'''' /********************************************************
''''  *@function appendArray
''''  *<pre>配列をappendする
''''  *</pre>
''''  *@param {Variant<string>} key
''''  *@param {Variant<Long>} cnt
''''  *@return {Variant()} appendArray　配列
''''  *FIXME□□生成したArrayに値を追加するのが手間。いちいちReDim
''''  */
'    Dim ary() As Variant
'    If IsMissing(cnt) Then
'        ary = Array()
'    Else
'        ReDim ary(cnt - 1)
'    End If
'    If IsMissing(key) Then
'        key = CStr(clt_.Count + 1)
'    End If
'    Call Me.putValue(ary, key)
'    appendArray = Me.getValue(key)
'End Function

Public Sub AppendDictionary(Optional ByVal key As Variant)
'''' /********************************************************
''''  *@function appendDictionary
''''  *<pre>連想配列をappendする
''''  *</pre>
''''  *@param {Variant<string>} key
''''  *@return {Variant<Dictionary>} appendDictionary　連想配列
''''  */
If IsMissing(key) Then
    key = CStr(clt_.Count + 1)
End If
Call Me.PutValue(New Dictionary, key)
''
End Sub

Public Sub AppendCollection(Optional ByVal key As Variant)
'''' /********************************************************
''''  *@function appendDictionary
''''  *<pre>連想配列をappendする
''''  *</pre>
''''  *@param {Variant<string>} key
''''  *@return {Variant<Dictionary>} appendDictionary　連想配列
''''  */
If IsMissing(key) Then
    key = CStr(clt_.Count + 1)
End If
Call Me.PutValue(New Collection, key)
''
End Sub

Public Function IsObjectEx(Optional key_or_idx As Variant) As Boolean
'''' /********************************************************
''''  *@function IsObjectEx
''''  *@param {Variant<String/Long>} key_or_idx
''''  *@return {Boolean} IsObjectEx
''''  */
If IsMissing(key_or_idx) Then
    key_or_idx = Me.Count
End If
IsObjectEx = IsObject(clt_.Item(key_or_idx))
''
End Function

Public Function IsArrayEx(Optional key_or_idx As Variant) As Boolean
''' /********************************************************
'''  *@function IsObjectEx
'''  *@param {Variant<String/Long>} key_or_idx
'''  *@return {Boolean} IsObjectEx
'''  */
If IsMissing(key_or_idx) Then
    key_or_idx = Me.Count
End If
IsArrayEx = IsArray(clt_.Item(key_or_idx))
''
End Function

Public Function keys() As Variant
'''' /********************************************************
''''  *@function Keys
''''  *@return {Variant<String>(i)} Keys
''''  */
Dim C_Collection As C_Collection
Set C_Collection = New C_Collection
keys = C_Collection.ToArray(key_)
Set C_Collection = Nothing
''
End Function

Public Function Items() As Variant
'''' /********************************************************
''''  *@function Items
''''  *@return {Variant<String/Long/ArrayO/bject>(i)} Items
''''  */
Dim C_Collection As C_Collection
Set C_Collection = New C_Collection
Items = C_Collection.ToArray(clt_)
Set C_Collection = Nothing
''
End Function

'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Sub UnitTest()
    Dim dict As Dictionary
    Set dict = New Dictionary
    Dim ary As Variant
    Dim vnt As Variant
''putValue
'    Call Me.putValue("sakai.mitsuru")
'    Call Me.putValue("sakai.mitsuru", "sakai1")
'    Call Me.putValue("sakai.mitsuru1", "sakai1")
''ToString
'    Console.info (Me.toString)
''getValue
'    Console.info Me.getValue
'    Console.info Me.getValue(1)
'    Console.info Me.getValue(2)
'    Console.info Me.getValue("sakai1")
''putByKeyword
'    vnt = Me.putByKeyword("sheet1", "Name", "name", 1, 0).getValue("name") 'Object型のJsonを取得する場合は、Set
'    Console.Dump vnt
'    vnt = Me.putByKeyword("sheet1", "佐藤", "name2", 0, 1).getValue("name2") 'Object型のJsonを取得する場合は、Set
'    Console.Dump vnt
'    ary = Me.putByKeyword("sheet1", "ARRAY", "ARRAY", 0, 1).getValue("ARRAY")
'    Console.Dump ary
''putByName
'    Set dict = Me.putByName("sheet1", "DICTIONARY").getValue("DICTIONARY")
'    Console.Dump dict
'    Console.info dict("sakai")
'    Console.Dump Me.putByName("sheet1", "TEMP").getValue("TEMP")
'    Console.Dump Me.putByName("sheet1", "EmployeeName").getValue("EmployeeName")
'    Console.Dump Me.putByName("sheet1", "Employee").getValue("Employee")
''putByRange
'    vnt = Me.putByRange(Range("b5")).getValue()
'    Console.info vnt
'    Console.info Me.getValue(1)
'     ''FIXME□□Variant型にDictionaryが設定された場合のDumpの動作
'    Console.Dump Me.putByRange(Range("A9:B13"))
'    Console.Dump Me.putByRange(Range("A9:C13"))
'    Console.Dump Me.putByRange(Range("A9:A13"))
'    Console.Dump Me.putByRange(Range("A9:F9"))
'    Console.Dump Me.putByRange(Range("DICTIONARY"))
'    Set dict = Me.getValue(Me.Count)  '''対象変数を意識して、取得する
'    ary = Me.getValue(5) '''対象変数を意識して、取得する
'    Console.Dump dict
'    Console.Dump ary
''IsObjectEx / IsArrayEx
'    Console.info Me.IsObjectEx(Me.Count)
'    Console.info Me.IsArrayEx(Me.Count)
'    Console.info Me.IsObjectEx(5)
'    Console.info Me.IsArrayEx(5)
'
''Keys
'    Call Me.putValue("sakai", "misa")
'    Console.Dump Me.Keys
'''Items
'    Console.Dump Me.Items
'''Exists
'    Console.info Me.Exists("misa")
'    Console.info Me.Exists("sakamoto")
'    Console.info Me.Exists("sakai")
'    Console.info Me.Exists(1)
'    Console.info Me.Exists(6)
'    Console.info Me.Exists(7)
'    Console.info Me.Exists(8)
''Count
'    Console.info Me.getValue(Me.Count)
''ToSource
'    Console.info Me.ToSource
''AppendArray　　''FIXME□□ByRefで操作ができないため、意味なし
'    Call Me.appendArray("sakai", 2)
'    Me.getValue("sakai") = Array(1, 2)
'    Console.Dump (Me.getValue)
    
' ''AppendArray
'    Call Me.appendDictionary("sakai2")
'    Call Me.getValue("sakai2").add("aiueo", "kaki")
'    Call Me.getValue("sakai2").add("sasi", "suse")
'    Console.Dump (Me.getValue)

End Sub



