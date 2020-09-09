VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Array"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Array.cls
'''  *<pre>Array Utility Class
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  *
'''  */
'
''' /********************************************************
'''  * @class C_Array
'''  * @classdesc 配列用ユーティリティ・クラス
'''  * 注意：引数で渡す配列は、配列として定義されていることを前提とする。
'''  * 　より、IsArray判定はしない。
'''  *
'''  * @require C_String, C_Commons
'''  */
'
Dim CONS_ARRAY_TYPE As Variant
Private Enum ARRAY_TYPE
    ARRAY_NORMAL = 0
    ARRAY_ARRAY = 1
    JagArray = 2
    ARRAY_EMPTY = 3
End Enum

'//--------------------------------------------------------
'// Class Event
'//
Private Sub Class_Initialize()
''' /********************************************************
'''  * @event Class_Initialize
'''  *<br/>Initialize 関数
'''  *<br/>Class生成時に、実行される。
'''  */
CONS_ARRAY_TYPE = Array("Array", "ArrayArray", "JagArray", "ArrayEmpty")
''
End Sub


'//--------------------------------------------------------
'// Class Methods
'//
Public Function TypeNameEx(ary As Variant) As String
'''' /********************************************************
''''  *@function TypeNameEx
''''  *<pre>配列の形式を取得する（ary/aryary/jagaryを判定する）
''''  *</pre>
''''  *  @param {Variant()} ary
''''  *  @return {String} TypeNameEx
''''  */
''
Dim str As String
Dim C_String As C_String
Set C_String = New C_String
''
If Not IsArray(ary) Then
    Err.Raise 13
End If
''
str = TypeName(ary)
str = C_String.StripEx(str, "[\(\)]")
''
If Me.IsInit(ary) Then
    If Me.Rank(ary) = 1 Then
        If UBound(ary) >= 0 Then
            If IsArray(ary(LBound(ary))) Then
                'JagArray(最初のrのaryにて判定）
                str = str & "(" & LBound(ary) & " To " & UBound(ary) & ")" & "(" & LBound(ary(LBound(ary))) & " To " & UBound(ary(LBound(ary))) & ")"
            Else
                'Array
                str = str & "(" & LBound(ary) & " To " & UBound(ary) & ")"
            End If
        ElseIf UBound(ary) = -1 Then
            '空のArray
            str = str & "()"
        End If
    ElseIf Me.Rank(ary) = 2 Then
        'ArrayArray
        str = str & "(" & LBound(ary, 1) & " To " & UBound(ary, 1) & "," & LBound(ary, 2) & " To " & UBound(ary, 2) & ")"
    Else
        '３次元以上のArrayArray
        str = str & "( Rank = " & Me.Rank(ary) & ")"
    End If
Else
    '初期化していないArray
    str = str & "()"
End If
TypeNameEx = str
''
End Function

Public Function TypeArrayPattern(ary As Variant) As String
'''' /********************************************************
''''  *@function TypeArray
''''  *<pre>配列の型を取得する（ary/aryary/jagaryを判定する）
''''  *</pre>
''''  *  @param {Variant()} ary
''''  *  @return {String} TypeArray
''''  */
''
Dim str As String
If Not IsArray(ary) Then
    Err.Raise 13
End If
If Me.IsInit(ary) Then
    If Me.Rank(ary) = 1 Then
        If UBound(ary) >= 0 Then
            If IsArray(ary(LBound(ary))) Then
                ''JagArray(最初のrのaryにて判定）
                str = CONS_ARRAY_TYPE(ARRAY_TYPE.JagArray) ''JagArray
            Else
                ''Array
                str = CONS_ARRAY_TYPE(ARRAY_TYPE.ARRAY_NORMAL) ''Array
            End If
        ElseIf UBound(ary) = -1 Then
            ''空のArray
            str = CONS_ARRAY_TYPE(ARRAY_TYPE.ARRAY_EMPTY) ''ArrayEmpty
        End If
    ElseIf Me.Rank(ary) = 2 Then
        ''ArrayArray
        str = CONS_ARRAY_TYPE(ARRAY_TYPE.ARRAY_ARRAY) ''ArrayArray
    Else
        ''３次元以上のArrayArray
        str = "ArrayArray(Rank= " & Me.Rank(ary) & ")"
    End If
Else
    ''初期化していないArray
    str = "ArrayNotInit"
End If
TypeArrayPattern = str
''
End Function

Public Function Rank(ary As Variant) As Long
'''' /********************************************************
''''  *@function Rank
''''  *<pre>配列のランク数（一次元配列=1、二次元配列=2）
''''  *</pre>
''''  *  @param {Variant()} ary
''''  *  @return {Long} Rank
''''  */
''
Dim i As Long
Dim x As Long
If Not VBA.IsArray(ary) Then
    Err.Raise 13
End If
''
On Error Resume Next
    i = 0
    While Err.Number = 0
        i = i + 1
        x = UBound(ary, i)
    Wend
    Rank = i - 1
End Function

Public Function IsArrayEx(ByRef ary As Variant) As Long
'''' /********************************************************
''''  *@function IsArrayEx
''''  *<pre>IsArrayの拡張
''''  *</pre>
''''  *  @param {Variant()} ary
''''  *  @return {Long} IsArrayEx
''''  *      0 : 配列ではない
''''  *    , 1 : 配列
''''  *    ,-1 : 空の配列
''''  *    ,-2 : 未初期化な動的配列
''''  */
''
Dim lng As Long
Dim ub  As Long
''
On Error Resume Next
    If IsEmpty(ary) Then
        IsArrayEx = 0 ' 配列でない
        Exit Function
    End If
    
    If Not IsArray(ary) Then
        IsArrayEx = 0 ' 配列でない
        Exit Function
    Else
        ub = UBound(ary)
        If Err.Number = 0 Then
            If ub < 0 Then
                IsArrayEx = -1 '---> 空の動的配列
                Exit Function
            Else
                IsArrayEx = 1  '---> 配列
                Exit Function
            End If
        ElseIf Err.Number = 9 Then
            If Not Not ary Then   '''TIPS:なぜNot Notでチェックできるかは不明
                IsArrayEx = -2 '未初期化の動的配列
                Exit Function
            Else
                IsArrayEx = -99 '---> 発生しない
                Exit Function
            End If
        Else
            IsArrayEx = -98 '---> 発生しない
            Exit Function
        End If
    End If
End Function

Public Function IsInit(ByRef ary As Variant) As Boolean
'''' /********************************************************
''''  *@function IsInit
''''  *<pre>動的配列が初期化されているかの判定
''''  *</pre>
''''  *  @param {Variant()} ary
''''  *  @return {Boolean} IsInit
''''  */
''
Dim lng As Long
lng = Me.IsArrayEx(ary)
If lng = -2 Or lng = -3 Then
        IsInit = False
ElseIf lng = -1 Or lng = 1 Then
        IsInit = True
Else
        Err.Raise 12001, "C_Array.IsInit", "配列ではありません"
End If
End Function

Public Function IsEmptyArray(ary As Variant) As Boolean
'''' /********************************************************
''''  *@function IsEmptyArray for String()
''''  *<br/>
''''  *  @param {Variant()} ary
''''  */
'' This function returns true if array is empty
Dim Length As Long
''
On Error Resume Next
    Length = Len(Join(ary, ""))
    If Length = 0 Then
      IsEmptyArray = True
    Else
      IsEmptyArray = False
    End If
    
    If Err.Number > 0 Then
        IsEmptyArray = True
    End If
''
On Error GoTo 0
''
End Function

Public Function Count(ary As Variant) As Long
''' /********************************************************
'''  *@function Count
'''  *<pre>配列の要素の個数を返却
'''  *</pre>
'''  *  @param {Variant()} ary
'''  *  @return {Long} Count
'''  */
'
Dim aryTemp As Variant
If IsArray(ary) Then
    If Me.Rank(ary) = 2 Then
        aryTemp = Me.ArrayArrayToArray(ary)
    Else
        aryTemp = ary
    End If
    If Me.IsInit(aryTemp) Then
        Count = UBound(aryTemp) - LBound(aryTemp) + 1
    Else
        Count = -1
    End If
Else
    Err.Raise 12001, "C_Array.Count", "配列変数ではありません"
End If
''
End Function

Public Function Initialize(ByRef ary As Variant)
''' /********************************************************
'''  *@function Initialize
'''  *<pre> initialize array variable, assing empty array 'Array()' into param 'ary'
'''  *</pre>
'''  *  @param {Variant} ary
'''  */
'
If IsArray(ary) Then
        ary = Array()
End If
''
End Function

Public Function GetElementLast(ary As Variant) As Variant
''' /********************************************************
'''  *@function getElementLast
'''  *<br/>配列の最後の要素を取得する。
'''  *  @param {Variant} ary 【ByRef】 配列
'''  *  @return {Variant} 要素の値
'''  */
''
GetElementLast = ary(UBound(ary))
''
End Function

Public Sub Add(ByRef ary As Variant, Value As Variant)
'''' /********************************************************
''''  *@function add
''''  *</br> 配列をReDimして、渡された要素を追加する。
''''  * @param {Variant} ary  ByRef
''''  * @param {Variant} value
''''  *
''''  * 注意：配列でないものをary引数で渡すとエラーとなる
''''  */
Dim lng As Long
Let lng = Me.IsArrayEx(ary)
Select Case lng
        Case 0
            Err.Raise 5, "C_Array.add", "引数aryが配列でありません。"
        Case 1, -1
        If LBound(ary) = 0 Then
            ReDim Preserve ary(LBound(ary) To UBound(ary) - LBound(ary) + 1) 'lng=-1 空の配列
        ElseIf LBound(ary) = 1 Then
            ReDim Preserve ary(LBound(ary) To UBound(ary) - LBound(ary) + 2)
        Else
            ReDim Preserve ary(LBound(ary) To UBound(ary) - LBound(ary) + 2)
        End If
        Case -2, -3 'lng=-2,-3 未初期化配列
        Let ary = Array()
            ReDim Preserve ary(UBound(ary) - LBound(ary) + 1)
End Select
If IsObject(Value) Then
        Set ary(UBound(ary)) = Value
Else
        Let ary(UBound(ary)) = Value
End If
''
End Sub

Public Sub addArray(ByRef aryTo As Variant, ary As Variant)
'''' /********************************************************
''''  *@function add
''''  *</br> 配列をReDimして、渡された要素を追加する。
''''  * @param {Variant} ary  ByRef
''''  * @param {Variant} ary2
''''  *
''''  * 注意：配列でないものをary引数で渡すとエラーとなる
''''  */
Dim i As Long
For i = LBound(ary) To UBound(ary)
    Me.Add aryTo, ary(i)
Next i
End Sub

Public Sub AddJagArray(ByRef jagary As Variant, ary As Variant)
''' /********************************************************
'''  *@function addJagArray
'''  *</br> ジャグ配列に、配列を追加する。
'''  * @param {Variant} jagAry
'''  * @param {Variant} ary
'''  * @return {Variant} addJagArray jag配列
'''  *
'''  * 注意：最初にJagArrayを作成する際は、Array関数で可能。
'''  *       Array(ary,ary2)
'''  */
Dim lng As Long
Dim aryTemp As Variant
Dim i As Long
Dim c As Long
lng = Me.IsArrayEx(jagary)  '' FIXME -1 : 空のvnt, -2 : ary()で宣言した初期化が未な動的配列

Select Case lng
        Case 0
            Err.Raise 13
        Case 1, -1
        If LBound(jagary) = 0 Then
            ReDim Preserve jagary(LBound(jagary) To UBound(jagary) - LBound(jagary) + 1)
        ElseIf LBound(jagary) = 1 Then
            ReDim Preserve jagary(LBound(jagary) To UBound(jagary) - LBound(jagary) + 2)
        Else
            ReDim Preserve jagary(LBound(jagary) To UBound(jagary) - LBound(jagary) + 2)
        End If
        Case -2, -3
        jagary = Array()
        ReDim Preserve jagary(1 To 1)
End Select
''
If (lng > -1) Then
    'Console.log (LBound(jagAry))
    aryTemp = jagary(LBound(jagary)) '最初の配列 (aryは同じ大きさの配列である前提）
    ReDim aryTemp(LBound(aryTemp) To UBound(aryTemp))
    Me.FillEmptySpaceForArray aryTemp
    Let c = LBound(aryTemp)
        For i = LBound(ary) To UBound(ary)
            aryTemp(c) = ary(i)
            c = c + 1
        If c > Me.Count(aryTemp) Then
            Exit For
        End If
        Next i
Else
        aryTemp = ary '空の配列
End If
''
jagary(UBound(jagary)) = aryTemp
''
End Sub

Public Sub AddParamArray(ByRef ary As Variant, ParamArray pAry() As Variant)
'''' /********************************************************
''''  *@function addParamArray
''''  *<pre> ParamArrayで渡された値を、配列をReDimして、１つ１つ追加する。
''''  *</pre>
''''  * @param {Variant} ary
''''  * @param {Variant} pAry
''''  * @return {Variant} addParamArray
''''  *
''''  */
Dim i As Long
For i = LBound(pAry) To UBound(pAry)
    Me.Add ary, pAry(i)
Next i
''
End Sub

Public Function append(ByRef ary As Variant _
                     , ByRef Value As Variant) As Object
'''' /********************************************************
''''  *@function Append
''''  *<pre>
''''  *値を追加して、その値を取得する。
''''  *</pre>
''''  *@param {Variant()} ary ByRef
''''  *@param {Object} value value
''''  *@return {Variant} Append
''''  *@note 注意：New Objectして、そのObjectを操作するのに利用することを想定。
''''  */
''
Call Me.Add(ary, Value)
If IsObject(Value) Then
        Set append = ary(UBound(ary))
Else
        Let append = ary(UBound(ary))
End If
''
End Function

Public Function AppendDictionary(ByRef ary As Variant) As Dictionary
'''' /********************************************************
''''  *@function appendDictionary
''''  *<pre> 連想配列をNewして追加して、そのオブジェクトを返却する
''''  *</pre>
''''  * @param {Variant} ary
''''  * @return {Dictionary} appendDictionary
''''  *
''''  */
''
Set AppendDictionary = Me.append(ary, New Dictionary)
''
End Function

Public Sub DeleteElementAt(ByRef ary As Variant, idx As Long)
'''' /********************************************************
''''  *@function deleteElementAt
''''  *<br/>指定位置の配列要素を削除する。
''''  *  @param {Variant} ary 【ByRef】 配列
''''  *  @param {Long} idx 要素添字
''''  */
''
Dim i As Long
If Me.Count(ary) = 1 Then
    Me.Initialize ary
Else
    For i = idx + 1 To UBound(ary)
        If IsObject(ary(i)) Then
            Set ary(i - 1) = ary(i)
        Else
            Let ary(i - 1) = ary(i)
        End If
    Next
    ReDim Preserve ary(UBound(ary) - LBound(ary) - 1)
End If
''
End Sub

Public Sub InsertElementAt(ByRef ary As Variant, idx As Long, elm As Variant)
''' /********************************************************
'''  *@function insertElementAt
'''  *<br/>指定位置の配列要素を挿入する。
'''  *  @param {Variant()} ary 【ByRef】 配列
'''  *  @param {Long} idx 要素添字
'''  *  @param {Variant} elm 要素
'''  */
''
If idx > UBound(ary) Then
        Err.Raise 9
End If
ReDim Preserve ary(UBound(ary) - LBound(ary) + 1)
Dim i As Long
For i = UBound(ary) To idx + 1 Step -1
        If IsObject(ary(i - 1)) Then
            Set ary(i) = ary(i - 1)
        Else
            Let ary(i) = ary(i - 1)
        End If
Next
If IsObject(elm) Then
        Set ary(idx) = elm
Else
        Let ary(idx) = elm
End If
''
End Sub

Public Function Transpose(ByRef ary As Variant) As Variant
'''' /********************************************************
''''  *@function Transpose
Dim aryaryTemp
Dim arytype As String
Let arytype = Me.TypeArrayPattern(ary)
If arytype = "ArrayArray" Then
    Transpose = WorksheetFunction.Transpose(ary)
ElseIf arytype = "Array" Then
    Transpose = Me.ArrayToArrayArrayVertical(ary)
ElseIf arytype = "JagArray" Then
    Let aryaryTemp = Me.JagArrayToArrayArray(ary)
    Let aryaryTemp = WorksheetFunction.Transpose(aryaryTemp)
    Transpose = Me.ArrayArrayToJagArray(aryaryTemp)
Else
    Transpose = ary
End If
''
End Function

Public Function Chop(ByRef ary As Variant)
'''' /********************************************************
''''  *@function chomp
''''  *<pre> delete last element from param of 'ary'
''''  *</pre>
''''  *  @param {Variant} ary
''''  */
''
If IsArray(ary) Then
        Me.DeleteElementAt ary, UBound(ary) - LBound(ary)
End If
''
End Function

Public Function Chomp(ByRef ary As Variant)
'''' /********************************************************
''''  *@function chomp
''''  *<pre> delete last element from param of 'ary'
''''  *</pre>
''''  *  @param {Variant} ary
''''  */
''
If IsArray(ary) Then
        Me.DeleteElementAt ary, LBound(ary)
End If
''
End Function

Public Function Pop(ByRef ary As Variant) As Variant
'''' /********************************************************
''''  *@function Pop
''''  */
''
If IsArray(ary) Then
    Let Pop = ary(UBound(ary))
    Me.DeleteElementAt ary, UBound(ary)
End If
''
End Function

Public Sub Push(ByRef ary As Variant, ParamArray pAry() As Variant)
'''' /********************************************************
''''  *@function Pop
''''  */
''
If IsArray(ary) Then
    Dim i As Long
    For i = LBound(pAry) To UBound(pAry)
        Me.Add ary, pAry(i)
    Next i
End If
''
End Sub

Public Function shift(ByRef ary As Variant) As Variant
'''' /********************************************************
''''  *@function Pop
''''  */
''
If IsArray(ary) Then
    Let shift = ary(LBound(ary))
    Me.DeleteElementAt ary, LBound(ary)
End If
''
End Function

Public Sub Unshift(ByRef ary As Variant, ParamArray pAry() As Variant)
'''' /********************************************************
''''  *@function Pop
''''  */
''
If IsArray(ary) Then
    Dim i As Long
    For i = UBound(pAry) To LBound(pAry) Step -1
        Me.InsertElementAt ary, LBound(ary), pAry(i)
    Next i
End If
''
End Sub

Public Sub Concat(ByRef ary As Variant, anAry As Variant)
'''' /********************************************************
''''  *@function concat
''''  *</br> 配列をReDimして、渡された配列の要素を１つ１つ追加する。
''''  * @param {Variant} ary
''''  * @param {Variant} anAry
''''  *
''''  */
Dim i As Long
For i = LBound(anAry) To UBound(anAry)
        Call Me.Add(ary, anAry(i))
Next i
''
End Sub

Public Sub MergeArrayArray(ByRef aryary As Variant, aryaryDist As Variant)
'''' /********************************************************
''''  *@function MergeArrayArray
''''  *@constraint : Both have same columns.
''''  */
Dim jagary() As Variant
Let jagary = Me.ArrayArrayToJagArray(aryary)
Dim jagarydist() As Variant
Let jagarydist = Me.ArrayArrayToJagArray(aryaryDist)
Call Me.Concat(jagary, jagarydist)
Let aryary = Me.JagArrayToArrayArray(jagary)
''
End Sub

Public Function ChangeIndexOfBegin(ByVal ary As Variant) As Variant
'''' /********************************************************
''''  *@function changeIndex
''''  *
''''  */
Dim aryTemp() As Variant
Dim i As Long
''
If LBound(ary) = 0 Then
    ReDim aryTemp(LBound(ary) + 1 To UBound(ary) + 1)
    For i = LBound(ary) To UBound(ary)
        Let aryTemp(i + 1) = ary(i)
    Next i
ElseIf LBound(ary) = 1 Then
    ReDim aryTemp(LBound(ary) - 1 To UBound(ary) - 1)
    For i = LBound(ary) To UBound(ary)
        Let aryTemp(i - 1) = ary(i)
    Next i
Else
    ChangeIndexOfBegin = ary
    Exit Function
End If
''
Let ChangeIndexOfBegin = aryTemp
''
End Function


Public Function Slice(ByRef ary As Variant, idx As Long _
                    , howMany As Long) As Variant
''' /********************************************************
'''  *@function slice
'''  *<br/>要素を取り除く
'''  *  @param {Variant} ary 【ByRef】 配列 取り除いた配列要素
'''  *  @param {Long} idx 要素添字
'''  *  @param {Long} howMany 要素数。0の場合はどの要素も取り除かない。
'''  *                -1以下の場合はindex以降を全て取り除く
'''  *  @return {Variant} 取り除かれた配列要素
'''  *
'''  */
Dim aryRet() As Variant
Dim n As Long
''
If howMany > -1 Then
    n = howMany
Else
    n = Me.Count(ary) - idx
End If
ReDim aryRet(n - 1)
Dim ix As Long
If UBound(ary) < idx Then
    ix = UBound(ary)
Else
    If 0 > idx Then
        ix = UBound(ary) + idx
        Else
        ix = idx
    End If
End If
''
Dim l As Long
Dim m As Long
m = 0
For l = ix To ix + n - 1
    aryRet(m) = ary(l)
    ary(l) = ""
    m = m + 1
Next l
''
For l = UBound(ary) To LBound(ary) Step -1
    'Console.log ary(l)
    If ary(l) = "" Then
        Call DeleteElementAt(ary, l)
    End If
Next l
Slice = aryRet
''
End Function

Public Function ToArrayList(ByRef ary As Variant) As Object
'''' /********************************************************
''''  *@function toArrayList
''''  *<br/>配列を内部で、ArrayListに変換する。
''''  *  @param {Variant} ary
''''  *  @return {Object} System.Collections.ArrayList
''''  */
''
Dim list As Object
Set list = createObject("System.Collections.ArrayList")
Dim i As Integer
For i = LBound(ary) To UBound(ary)
    list.Add (ary(i))
Next i
Set ToArrayList = list
''
End Function

Public Function ToCollection(ByRef ary As Variant) As Collection
'''' /********************************************************
''''  *@function toArrayList
''''  *<br/>配列を内部で、ArrayListに変換する。
''''  *  @param {Variant} ary
''''  *  @return {Object} System.Collections.ArrayList
''''  */
''
Dim clt As Collection
Set clt = New Collection
Dim i As Integer
For i = LBound(ary) To UBound(ary)
    clt.Add (ary(i))
Next i
Set ToCollection = clt
''
End Function

Public Function SortArrayArray(ByRef aryary As Variant, ByVal colNo As Long, Optional asc As Variant = 1, Optional colSort As Variant = False) As Variant
'''' /********************************************************
''''  *@function SortArrayArray
''''  */
''
Let SortArrayArray = WorksheetFunction.Sort(aryary, colNo, asc, colSort)
''
End Function

Public Function SortbyArrayArray(ByRef aryary As Variant, ByVal col1 As Long, asc1 As Variant, ByVal col2 As Long, asc2 As Variant) As Variant
'''' /********************************************************
''''  *@function SortbyArrayArray
''''  */
Dim aryaryCol1 As Variant
Dim aryaryCol2 As Variant
Let aryaryCol1 = Me.GetColumnForArrayArray(aryary, col1)
Let aryaryCol2 = Me.GetColumnForArrayArray(aryary, col2)
''
SortbyArrayArray = WorksheetFunction.SortBy(aryary, aryaryCol1, asc1, aryaryCol2, asc2)
''
End Function

Public Function Vlookup(ByVal ky As Variant, ByRef aryary As Variant, ByVal colNo As Long, Optional searchType As Variant = False)
'''' *************************************************
Dim C_String As C_String
Set C_String = New C_String
''
On Error Resume Next
Err.clear
Dim rtn As Variant
Let rtn = WorksheetFunction.Vlookup(ky, aryary, colNo, False)
If Err.Number = 0 Then
    ''
Else
    Console.log "Could not find value: "
End If
''
Let Vlookup = C_String.DefaultString(rtn)
''
End Function

Public Function Unique(ByRef aryary As Variant, Optional ByVal rowDir As Boolean = True) As Variant
'''' *************************************************
''
Let Unique = WorksheetFunction.Unique(aryary, rowDir)
''
End Function

Public Function UniqueArrayArray(aryary As Variant) As Variant
''' /********************************************************
'''  * @function Unique
'''  */
'' FIXME
Dim ary() As Variant
ReDim ary(LBound(aryary, 1) To UBound(aryary, 1))
Dim jagary() As Variant
Let jagary = Me.ArrayArrayToJagArray(aryary)
Dim r As Long
For r = LBound(ary) To UBound(ary)
    ary(r) = Join(jagary(r), vbTab)
Next r
Dim aryTemp() As Variant
Let aryTemp = Me.Unique(ary)
Dim i As Long
For i = LBound(aryTemp) To UBound(aryTemp)
    Let aryTemp(i) = Split(aryTemp(i), vbTab)
Next i
''
Let UniqueArrayArray = Me.JagArrayToArrayArray(aryTemp)
''
End Function

Public Function Index(ByRef aryary As Variant, ByVal rowNo As Long, ByVal colNo As Long)
'''' *************************************************
''
Dim C_String As C_String
Set C_String = New C_String
''
On Error Resume Next
Err.clear
Dim rtn As Variant
Let rtn = WorksheetFunction.Index(aryary, rowNo, colNo)
If Err.Number = 0 Then
    ''
Else
    Console.log "Could not find value: "
End If
''
Let Index = C_String.DefaultString(rtn)
''
End Function

Public Function match(ByVal ky As Variant, ByRef aryary As Variant)
'''' *************************************************
''
Dim C_String As C_String
Set C_String = New C_String
''
On Error Resume Next
Err.clear
Dim rtn As Variant
Let rtn = WorksheetFunction.match(ky, aryary, 0)
If Err.Number = 0 Then
    ''
Else
    Console.log "Could not find value: "
End If
''
Let match = C_String.DefaultString(rtn)
''
End Function

Public Function GetValueWithMatchIndex(ByVal ky As Variant, ByRef aryary As Variant, ByVal colTarget As Long, ByVal colNo As Long)
'''' *************************************************
''
Dim rowNo As Variant
Dim aryColTarget As Variant
Let aryColTarget = Me.GetColumnForArrayArray(aryary, colTarget)
Let rowNo = Me.match(ky, aryColTarget)
If rowNo = "N/A" Then
    Let GetValueWithMatchIndex = rowNo
    Exit Function
End If
Let GetValueWithMatchIndex = Me.Index(aryary, rowNo, colNo)
''
End Function

Public Function Large(ByRef aryary As Variant, ByVal k As Double) As Double
'''' *************************************************
''
Dim aryaryTemp As Variant
''
If Me.TypeArrayPattern(aryary) = "Array" Then
    Let aryaryTemp = Me.ArrayToArrayArray(aryary)
    Let Large = WorksheetFunction.Large(aryaryTemp, k)
ElseIf Me.TypeArrayPattern(aryary) = "ArrayArray" Then
    Let Large = WorksheetFunction.Large(aryary, k)
Else
    Err.Raise 13
End If
''
End Function

Public Function Small(ByRef aryary As Variant, ByVal k As Double) As Double
'''' *************************************************
''
Dim aryaryTemp As Variant
''
If Me.TypeArrayPattern(aryary) = "Array" Then
    Let aryaryTemp = Me.ArrayToArrayArray(aryary)
    Let Small = WorksheetFunction.Small(aryaryTemp, k)
ElseIf Me.TypeArrayPattern(aryary) = "ArrayArray" Then
    Let Small = WorksheetFunction.Small(aryary, k)
Else
    Err.Raise 13
End If
''
End Function

Public Function Includes(ByRef aryary As Variant, ByVal vnt As Variant) As Boolean
'''' *************************************************
''
Dim C_String As C_String
Set C_String = New C_String
Dim sheetName As String
Let sheetName = C_String.GetTemporaryStr
''
Application.DisplayAlerts = False
''
Dim sht As Worksheet
Set sht = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
Let sht.Name = sheetName
Dim rng As Range
Set rng = sht.Range("A1")
''
''
Dim aryaryTemp As Variant
Dim cnt As Long
''
''
If Me.TypeArrayPattern(aryary) = "Array" Then
    Set rng = rng.Resize(1, UBound(aryary) - LBound(aryary) + 1)
    Let rng.Value = aryary
    Let cnt = WorksheetFunction.CountIf(rng, vnt) ''param is a rng
ElseIf Me.TypeArrayPattern(aryary) = "ArrayArray" Then
    Set rng = rng.Resize(UBound(aryary, 1) - LBound(aryary, 1) + 1, UBound(aryary, 2) - LBound(aryary, 2) + 1)
    Let rng.Value = aryary
    Let cnt = WorksheetFunction.CountIf(rng, vnt)
Else
    Err.Raise 13
End If
If cnt = 0 Then
    Let Includes = False
Else
    Let Includes = True
End If
''
ThisWorkbook.Worksheets(sheetName).Delete
Application.DisplayAlerts = True
''
End Function

Public Function Sort(ByRef ary As Variant)
'''' /********************************************************
''''  *@function sort
''''  *<br/>対象配列を、ArrayList#Sortでソートする
''''  *  @param {Variant} array1 ByRef 配列
''''  *
''''  *  注意：文字列、数値である前提（並び順は、ArrayList.Sortに準ずる）
''''  *　　　　文字列と数値が混合してはソートできない。（ArrayList.Sortの仕様）
''''  */
''配列をソートする。
Dim aryList As Object
Set aryList = ToArrayList(ary)
Call aryList.Sort
Erase ary
ary = aryList.ToArray()
Set aryList = Nothing
''
End Function

Public Function Reverse(ByRef ary As Variant)
'''' /********************************************************
''''  *@function reverse
''''  *<br/>対象配列を、ArrayList#reverseで逆順する
''''  *  @param {Variant} ary ByRef 配列
''''  *
''''  */
''配列を逆にする。
Dim aryList As Object
Set aryList = ToArrayList(ary)
Call aryList.Reverse
Erase ary
ary = aryList.ToArray()
Set aryList = Nothing
''
End Function

Public Function RemoveDuplicated(ByRef ary As Variant) As Variant
'''' /********************************************************
''''  *@function deleteDuplicated
''''  *<pre>
''''  *配列内で値が重複している場合に削除する
''''  *</pre>
''''  *  @param {Variant()} ary ByRef 配列
''''  *  @return {Variant()} deleteDuplicated 変換後の配列
''''  */
''
If Not Me.IsDuplicated(ary) Then
    Exit Function
End If
''
Dim dict As Dictionary ''重複チェック用
Set dict = New Dictionary
Dim aryRtn As Variant
Dim i As Long
''
For i = LBound(ary) To UBound(ary)
    If Not dict.Exists(ary(i)) Then
        Call dict.Add(ary(i), "")
        End If
Next
RemoveDuplicated = dict.keys
''
End Function

Public Sub DeleteEmptyElement(ByRef ary As Variant)
'''' /********************************************************
''''  *@function deleteEmptyElement
''''  *<br/>空の配列要素を削除する。
''''  *@param {Variant} ary 【ByRef】 配列
''''  *@note 注意：対象配列はary or jagaryとする
''''  *@note 注意：jagaryは、最初の行のArrayが正しく設定されていることとする。
''''  */
Dim i As Long
Dim aryNumber() As Long
Dim j As Long
j = 0
For i = LBound(ary) To UBound(ary)
        'Console.dump ary(i)
        If IsArray(ary(i)) Then
            If (ary(i)(LBound(ary(i))) = "") Or IsEmpty(ary(i)(LBound(ary(i)))) Then
                ReDim Preserve aryNumber(j)
                aryNumber(j) = i
                j = j + 1
            End If
        Else
            If (ary(i) = "") Or IsEmpty(ary(i)) Or IsNull(ary(i)) Then
                ReDim Preserve aryNumber(j)
                aryNumber(j) = i
                j = j + 1
            End If
        End If
Next i
''
If Me.Count(aryNumber) > 0 Then
    For j = UBound(aryNumber) To LBound(aryNumber) Step -1
    Call Me.DeleteElementAt(ary, aryNumber(j))
    Next j
End If
''
End Sub

Public Function Diff(ByVal array1 As Variant, ByVal array2 As Variant) As Variant
'''' /********************************************************
''''  *@function diff
''''  *<br/>配列の差分を返却
''''  *  @param {Variant} array1 ByVal 配列
''''  *  @param {Variant} array2 ByVal 配列
''''  *  @return {Variant} 差分配列　（array2に無いarray1要素の配列）
''''  *  @note FIXME 数字は効かない！？
''''  */
''配列の差分を返します
Dim listDel As Object 'ArrayList
Dim list1 As Object 'ArrayList
Dim list2 As Object 'ArrayList
Dim Item As Variant
Dim i As Long
Dim cnt2 As Long
''
Set listDel = createObject("System.Collections.ArrayList")
''
Set list1 = ToArrayList(array1)
Set list2 = ToArrayList(array2)
''
cnt2 = list2.Count - 1
For i = 0 To cnt2
    Item = list2(i)
    Console.log (Item)
    If list1.Contains(Item) Then
        list1.remove (Item)
        End If
Next i
Call listDel.AddRange(list1)
''
Diff = listDel.ToArray
''
Set listDel = Nothing
Set list1 = Nothing
Set list2 = Nothing
''
End Function

Public Function DiffBoth(ByVal array1 As Variant, ByVal array2 As Variant) As Variant
'''' /********************************************************
''''  *@function diff_both
''''  *<br/>配列の差分を返却（array1の不一致、array2の不一致の順で両方を返却）
''''  *  @param {Variant} array1 ByVal 配列１
''''  *  @param {Variant} array2 ByVal 配列１
''''  *  @return {Variant} diffBoth ジャグ配列
''''  *  　array2に無いarray1要素の配列、array2にあってarray1要素に無い配列
''''  */
''
Dim listAdd As Object 'ArrayList
Dim listDel As Object 'ArrayList
Dim list1 As Object 'ArrayList
Dim list2 As Object 'ArrayList
Dim Item As Variant
Dim i As Long
Dim cnt2 As Long
Dim aryResultAdd As Variant 'Array
Dim aryResultDel As Variant 'Array
''
Set listAdd = createObject("System.Collections.ArrayList")
Set listDel = createObject("System.Collections.ArrayList")
''
Set list1 = ToArrayList(array1)
Set list2 = ToArrayList(array2)
''
cnt2 = list2.Count - 1
For i = 0 To cnt2
    Item = list2(i)
    'Console.log (item)
    If list1.Contains(Item) Then
        list1.remove (Item)
    Else
        listAdd.Add (Item)
    End If
Next i
Call listDel.AddRange(list1)
''
aryResultDel = listDel.ToArray
aryResultAdd = listAdd.ToArray
''
DiffBoth = Array(aryResultDel, aryResultAdd)
''
Me.Initialize aryResultDel
Me.Initialize aryResultAdd
Set listAdd = Nothing
Set listDel = Nothing
Set list1 = Nothing
Set list2 = Nothing
''
End Function

Public Function IndexOf(ByRef ary As Variant _
                    , Value As Variant _
                    , Optional idxStart As Variant _
                    , Optional cnt As Variant) As Long
'''' /********************************************************
''''  *@function indexOf
''''  *<pre>
''''  *値がヒットしたidxを返却する。ヒットしない場合は、-1
''''  *</pre>
''''  *  @param {Variant} ary ByRef 配列
''''  *  @param {Variant} value 値
''''  *  @param {Variant<Long>} idxStart 開始位置
''''  *  @param {Variant<Long>} cnt 終了位置（何個か）
''''  *  @return {Long} indexOf
''''  */
If Not IsArray(ary) Then
    Err.Raise 13
End If
''
Dim i As Long
Dim lb As Long
Dim ub As Long
lb = LBound(ary)
ub = UBound(ary)
''
If IsMissing(idxStart) Then
    idxStart = lb
End If
If IsNumeric(idxStart) Then
    idxStart = CLng(idxStart)
Else
    Err.Raise 13
End If
If idxStart < lb Then
    Err.Raise 5
End If
If IsMissing(cnt) Then
    cnt = ub - lb + 1
End If
If IsNumeric(cnt) Then
    cnt = CLng(cnt)
Else
    Err.Raise 13
End If
''
If idxStart + cnt > ub - lb + 1 Then
    cnt = ub - lb + 1 - idxStart
End If
For i = idxStart To idxStart + cnt
    If ary(i) = Value Then
        IndexOf = i
        Exit Function
    End If
Next
''
IndexOf = -1
''
End Function

Public Function Exists(astrItems As Variant, strSearch As String) As Boolean
''' /********************************************************
'''  *@function exists
'''  *<br/>配列に完全に一致する文字列が含まれているか、否かを評価する
'''  *  @param {Variant} astrItems 【ByRef】 配列
'''  *  @param {Variant} strSearch 文字列
'''  */
Dim ary As Variant
''
ary = FilterMatchExist(astrItems, strSearch)
''
If UBound(ary) = -1 Then
        Exists = False
Else
        Exists = True
End If
''
End Function

Public Function FilterMatchExist(astrItems As Variant, _
                          strSearch As String, Optional compareMode As Variant = vbBinaryCompare) As Variant
'''' /********************************************************
''''  * @function FilterMatchExist
''''  * ＝＞機能としては利用価値無し
''''  *</br>検索文字列と完全一致する配列の要素を、抽出して返却
''''  *</br>する。
''''  * @param {Variant} astrItems 文字列配列
''''  * @param {String} strSearch 検索文字列
''''  * @param {Variant} compareMode 比較モード vbTextCompare=0,vbBinaryCompare=1
''''  * @return {Variant} FilterMatchExist 抽出した文字の配列
''''  *
''''  * 注意：配列の中は文字列を前提とする。FILTER関数の仕様による。
''''  */      ''FIXME
'' 検索文字列に完全一致する要素を、文字列型の配列内で検索。
Dim astrFilter() As String
Dim astrTemp() As String
Dim ub As Long
Dim lb As Long
Dim i     As Long
Dim cnt   As Long
cnt = 0
''
astrFilter = Filter(astrItems, strSearch)
ub = UBound(astrFilter)
lb = LBound(astrFilter)
''
If ub >= 0 Then
        ' 一時配列を生成
        ReDim astrTemp(lb To ub)
        ' フィルタが適用された配列の各要素をループ
        For i = lb To ub
            ' 要素が検索文字列に完全一致することを確認
            If compareMode = vbBinaryCompare Then '
                If astrFilter(i) = strSearch Then
                    astrTemp(cnt) = astrFilter(i)
                    cnt = cnt + 1
                End If
            Else
                If LCase(astrFilter(i)) = LCase(strSearch) Then
                    astrTemp(cnt) = astrFilter(i)
                    cnt = cnt + 1
                End If
            End If
        Next i
Else
        FilterMatchExist = Array() '''長さゼロの配列を返却
        Exit Function
End If
''
If cnt = 0 Then
        Erase astrTemp
        FilterMatchExist = Array() '''長さゼロの配列を返却
Else
        ' 完全一致する要素が含まれる配列を返却
        ReDim Preserve astrTemp(lb To cnt - 1)
        FilterMatchExist = astrTemp
End If
''
End Function

Public Function FilterMatchPrefix(astrItems As Variant, _
                          strPrefix As String, _
                          Optional compareMode As Variant = vbBinaryCompare) As Variant
'''' /********************************************************
''''  * @function FilterMatchPrefix
''''  *</br>検索文字列と前方一致する配列の要素を、抽出して返却
''''  *</br>する。
''''  * @param {Variant} astrItems 文字列配列
''''  * @param {String} strPrefix 検索文字列
''''  * @param {Variant} compareMode 比較モード vbTextCompare=0,vbBinaryCompare=1
''''  * @return {Variant} FilterMatchPrefix 抽出した文字の配列
''''  */
'' 検索文字列に前方一致する要素を、文字列型の配列内で検索。
Dim astrFilter() As String
Dim astrTemp() As String
Dim ub As Long
Dim lb As Long
Dim i As Long
Dim cnt As Long
cnt = 0
''
astrFilter = Filter(astrItems, strPrefix)
ub = UBound(astrFilter)
lb = LBound(astrFilter)
''
If ub >= 0 Then
    '' 一時配列を生成
        ReDim astrTemp(lb To ub)
    ''
        For i = lb To ub
            If compareMode = vbBinaryCompare Then '
                If left(astrFilter(i), Len(strPrefix)) = strPrefix Then
                    astrTemp(cnt) = astrFilter(i)
                    cnt = cnt + 1
                End If
            Else
                If LCase(left(astrFilter(i), Len(strPrefix))) = LCase(strPrefix) Then
                    astrTemp(cnt) = astrFilter(i)
                    cnt = cnt + 1
                End If
            End If
        Next i
Else
        FilterMatchPrefix = Array() '''長さゼロの配列を返却
        Exit Function
End If
''
If cnt = 0 Then
        Erase astrTemp
        FilterMatchPrefix = Array() '''長さゼロの配列を返却
Else
        ' 完全一致する要素が含まれる配列を返却
        ReDim Preserve astrTemp(lb To cnt - 1)
        FilterMatchPrefix = astrTemp
End If
''
End Function

Public Function FilterMatchSuffix(astrItems As Variant, _
                          strSufix As String, Optional compareMode As Variant = vbBinaryCompare) As Variant
'''' /********************************************************
''''  * @function FilterMatchSufix
''''  *</br>検索文字列と後方一致する配列の要素を、抽出して返却
''''  *</br>する。
''''  * @param {Variant} astrItems 文字列配列
''''  * @param {String} strSufix 検索文字列
''''  * @param {Variant} compareMode 比較モード vbTextCompare=0,vbBinaryCompare=1
''''  * @return {Variant} FilterMatchSufix 抽出した文字の配列
''''  * FIXME□□Like文でWildcardも実装する
''''  */
'' 検索文字列に後方一致する要素を、文字列型の配列内で検索。
Dim astrFilter() As String
Dim astrTemp() As String
Dim ub As Long
Dim lb As Long
Dim i As Long
Dim cnt As Long
cnt = 0
''
astrFilter = Filter(astrItems, strSufix)
ub = UBound(astrFilter)
lb = LBound(astrFilter)
''
If ub >= 0 Then
        ' 一時配列を生成
        ReDim astrTemp(lb To ub)
        For i = lb To ub
            ' 要素が検索文字列に後方一致することを確認
            If compareMode = vbBinaryCompare Then '
                If Right(astrFilter(i), Len(strSufix)) = strSufix Then
                    astrTemp(cnt) = astrFilter(i)
                    cnt = cnt + 1
                End If
            Else
                If LCase(Right(astrFilter(i), Len(strSufix))) = LCase(strSufix) Then
                    astrTemp(cnt) = astrFilter(i)
                    cnt = cnt + 1
                End If
            End If
        Next i
Else
        FilterMatchSuffix = Array() '''長さゼロの配列を返却
        Exit Function
End If
If cnt = 0 Then
        Erase astrTemp
        FilterMatchSuffix = Array() '''長さゼロの配列を返却
Else
        ' 完全一致する要素が含まれる配列を返却
        ReDim Preserve astrTemp(lb To cnt - 1)
        FilterMatchSuffix = astrTemp
End If
''
End Function

Public Function Contains(ary As Variant, str As Variant) As Boolean
''' /********************************************************
'''  *@function contains
'''  *<pre>配列に文字列（Like)が含まれているか、否かを、Filter関数を
'''  *利用して評価する
'''  *</pre>
'''  *  @param {Variant} ary 【ByRef】 配列
'''  *  @param {Variant} str 文字列
'''  */
Dim aryTemp As Variant
    
aryTemp = Filter(ary, str)
    
If UBound(aryTemp) = -1 Then
        Contains = False
Else
        Contains = True
End If
''
End Function

Public Function ContainsMatchPrefix(astrItems As Variant, strSearch As String, Optional compareMode As Variant = vbBinaryCompare) As Boolean
''' /********************************************************
'''  *@function containsMatchPrefix
'''  *<br/>配列に完全に一致する文字列が含まれているか、否かを評価する
'''  *  @param {Variant} ary 【ByRef】 配列
'''  *  @param {Variant} str 文字列
'''  */
Dim ary As Variant
ary = FilterMatchPrefix(astrItems, strSearch, compareMode)
''
If UBound(ary) = -1 Then
        ContainsMatchPrefix = False
Else
        ContainsMatchPrefix = True
End If
''
End Function

Public Function ContainsMatchSuffix(astrItems As Variant, strSearch As String, Optional compareMode As Variant = vbBinaryCompare) As Boolean
''' /********************************************************
'''  *@function containsMatchSuffix
'''  *<br/>配列に完全に一致する文字列が含まれているか、否かを評価する
'''  *  @param {Variant} ary 【ByRef】 配列
'''  *  @param {Variant} str 文字列
'''  */
Dim ary As Variant
ary = FilterMatchSuffix(astrItems, strSearch, compareMode)
''
If UBound(ary) = -1 Then
        ContainsMatchSuffix = False
Else
        ContainsMatchSuffix = True
End If
''
End Function

Public Function CountMatches(ByRef ary As Variant _
                    , Value As Variant) As Long
''' /********************************************************
'''  *@function countMatchies
'''  *<pre>
'''  *値がヒットした個数を返却する。ヒットしない場合は、0
'''  *</pre>
'''  *  @param {Variant} ary ByRef 配列
'''  *  @param {Variant} value 値
'''  *  @return {Long} countMatchies
'''  *
'''  */
If Not IsArray(ary) Then
        Err.Raise 13
End If
Dim cnt As Long
cnt = 0
Dim i As Long
For i = LBound(ary) To UBound(ary)
        If ary(i) = Value Then
            cnt = cnt + 1
                End If
Next
CountMatches = cnt
''
End Function

Public Function IsDuplicated(ByRef ary As Variant) As Boolean
''' /********************************************************
'''  *@function IsDuplicated
'''  *<pre>
'''  *配列内で値が重複しているか、否かを返却
'''  *</pre>
'''  *  @param {Variant} ary ByRef 配列
'''  *  @return {Boolean} IsDuplicated
'''  *
'''  */
If Not IsArray(ary) Then
        Err.Raise 13
End If
Dim cnt As Long
Dim i As Long
For i = LBound(ary) To UBound(ary)
        cnt = Me.CountMatches(ary, ary(i))
        If cnt > 1 Then
            IsDuplicated = True
            Exit Function
                End If
Next i
IsDuplicated = False
''
End Function

Public Function Equals(ary, ary2) As Boolean
''' /********************************************************
'''  *@function equals
'''  *<br/>一次配列の要素が同じかを判定
'''  *  @param {Variant} ary
'''  *  @param {Variant} ary2
'''  *  @return {Boolean} equals
'''  */
'
Dim blnRtn As Boolean
If Not (IsArray(ary) And IsArray(ary2)) Then
    Equals = False
    Exit Function
End If
Dim iMax As Long
iMax = UBound(ary) - LBound(ary) + 1
Dim i2Max As Long
i2Max = UBound(ary2) - LBound(ary2) + 1
If iMax <> i2Max Then
    Equals = False
    Exit Function
End If
''
Dim i As Long
For i = LBound(ary) To UBound(ary)
    If ary(i) <> ary2(i) Then
            blnRtn = False
            Exit Function
        End If
Next
blnRtn = True
Equals = blnRtn
''
End Function

Public Function EqualsArrayArray(aryaryA, aryaryB) As Boolean
''' /********************************************************
'''  *@function equals
'''  *<br/>二次元配列の要素が同じかを判定
'''  *  @param {Variant} aryary
'''  *  @param {Variant} aryary2
'''  *  @return {Boolean} equals
'''  */
'
Dim blnRtn As Boolean
If Not (IsArray(aryaryA) And IsArray(aryaryB)) Then
    EqualsArrayArray = False
    Exit Function
End If
Dim cntA1 As Long
Dim cntA2 As Long
Dim cntB1 As Long
Dim cntB2 As Long
''
Dim aryTempA As Variant
Dim aryTempB As Variant
Dim jagaryA As Variant
Dim jagaryB As Variant
''
If Not (Me.Rank(aryaryA) = Me.Rank(aryaryB)) Then
        EqualsArrayArray = False
        Exit Function
End If
''
If Me.Rank(aryaryA) = 1 Then
        cntA1 = UBound(aryaryA) - LBound(aryaryA) + 1
        cntB1 = UBound(aryaryB) - LBound(aryaryB) + 1
        aryTempA = aryaryA
        aryTempB = aryaryB
        If cntA1 = cntB1 Then
            EqualsArrayArray = Me.Equals(aryTempA, aryTempB)
        Else
            EqualsArrayArray = False
        Exit Function
        End If
ElseIf Me.Rank(aryaryA) = 2 Then
        cntA1 = UBound(aryaryA, 1) - LBound(aryaryA, 1) + 1
        cntB1 = UBound(aryaryB, 1) - LBound(aryaryB, 1) + 1
        cntA2 = UBound(aryaryA, 2) - LBound(aryaryA, 2) + 1
        cntB2 = UBound(aryaryB, 2) - LBound(aryaryB, 2) + 1
    jagaryA = Me.ArrayArrayToJagArray(aryaryA)
    jagaryB = Me.ArrayArrayToJagArray(aryaryB)
    Dim r As Long
        If (cntA1 = cntB1) And (cntA2 = cntB2) Then
        For r = LBound(jagaryA) To UBound(jagaryA)
            aryTempA = jagaryA(r)
            aryTempB = jagaryB(r)
            If Not (Me.Equals(aryTempA, aryTempB)) Then
                EqualsArrayArray = False
                Exit Function
            End If
        Next r
        EqualsArrayArray = True
        Exit Function
        Else
            EqualsArrayArray = False
        Exit Function
        End If
Else
    EqualsArrayArray = False
    Exit Function
End If
''
End Function

Public Function FilterRowForArrayArray(ByRef aryary As Variant _
                                    , ByVal columnNum As Long _
                                    , Optional rowStart As Variant _
                                    , Optional Instance As Variant _
                                    , Optional methodname As Variant _
                                    ) As Variant
''' /********************************************************
'''  * @function FilterRowForArrayArray
'''  *<pre>行ごとに、methodで判定して、行をフィルタした
'''  *ArrayArrayを返却する。
'''  *</pre>
'''  * @param {Variant} aryary
'''  * @param {Long} columnNum 判定カラム
'''  * @param {Long} rowStart 開始行
'''  * @param {Variant<Object>} instance
'''  * @return {Variant<String>} methodname
'''  */
Dim r As Long
Dim bln As Boolean
Dim jagary As Variant
jagary = Me.ArrayArrayToJagArray(aryary)
    
If IsMissing(rowStart) Then
        rowStart = 1
End If
If IsMissing(Instance) Then
        jagary = Me.FilterRowForJagArray(jagary, columnNum, rowStart)
Else
        jagary = Me.FilterRowForJagArray(jagary, columnNum, rowStart, Instance, methodname)
End If
''
FilterRowForArrayArray = Me.JagArrayToArrayArray(jagary)
''
End Function

Public Function FilterRowForJagArray(ByRef jagary As Variant _
                                    , ByVal columnNum As Long _
                                    , Optional rowStart As Variant _
                                    , Optional Instance As Variant _
                                    , Optional methodname As Variant _
                                    ) As Variant
''' /********************************************************
'''  * @function FilterRowForJagArray
'''  *<pre>行ごとに、methodで判定して、行をフィルタした
'''  *JagArrayを返却する。
'''  *</pre>
'''  * @param {Variant} jagary
'''  * @param {Long} columnNum 判定カラム
'''  * @param {Long} rowStart 開始行
'''  * @param {Variant<Object>} instance
'''  * @return {Variant<String>} methodname
'''  */
Dim r As Long
Dim bln As Boolean
If IsMissing(Instance) Then
        For r = rowStart To UBound(jagary)
            bln = CallByName(Me, "checkNumberic", VbMethod, CStr(jagary(r)(columnNum)))
            If Not bln Then
                jagary(r) = ""
            End If
        Next r
Else
        For r = rowStart To UBound(jagary)
            bln = CallByName(Instance, methodname, VbMethod, C_Commons.ToJson(jagary(r)))
            If Not bln Then
                jagary(r)(LBound(jagary(r))) = ""
            End If
        Next r
End If
Call Me.DeleteEmptyElement(jagary)
FilterRowForJagArray = jagary
''
End Function
    
Public Function checkNumberic(Value As String) As Boolean
    checkNumberic = IsNumeric(Value)
End Function

Public Function SplitEx(ByVal str As String, _
                    Optional ByVal delimiter As Variant = " ", _
                    Optional ByVal limit As Long = -1, _
                    Optional ByVal Compare As Integer = 0) As Variant
'''' /********************************************************
''''  *@function SplitArray
''''  *<br/> Delimiterを複数指定可能なSplit拡張関数
''''  *  @param {String} str ByVal
''''  *  @param {Variant()} delimiter ByVal 文字セット（配列）
''''  *  @param {Integer} limit Splitのlimit引数
''''  *  @param {Integer} Compare SplitのCompare引数
''''  *  @return {Variant<string>()} SplitEx 配列
''''  */
''
Dim delTemp As String
Dim strTemp As String
''
strTemp = str
Dim dl As Variant
If IsArray(delimiter) Then
    delTemp = delimiter(LBound(delimiter))
    For Each dl In delimiter
        strTemp = VBA.Replace(strTemp, dl, delTemp)
    Next dl
Else
        delTemp = delimiter
End If
SplitEx = VBA.Split(strTemp, delTemp, limit, Compare)
''
End Function

Public Function ArrayToDictionary(ByVal ary As Variant, aryHeader As Variant) As Variant
'''' /********************************************************
''''  *@function ArrayToDictionary
''''  *<br/>配列（Array(i)を、headerをもとに連想配列に変換する
''''  *  @param {Variant<String>(i)} ary 配列
''''  *  @param {Variant<string>(i)} aryHeader ヘッダ配列
''''  *  @return {Variant<Dictionary>} ArrayToDictionary 連想配列
''''  */
''
Dim C_String As C_String
Set C_String = New C_String
''
Dim dict As Variant
Set dict = New Dictionary
Dim i As Long
Dim j As Long
Dim k As Long
Dim ubh As Long
Dim lbh As Long
ubh = UBound(aryHeader)
lbh = LBound(aryHeader)
''
If (UBound(ary) - LBound(ary) + 1) > (ubh - lbh + 1) Then 'aryの方が個数が多い
    j = LBound(ary)
        For i = lbh To ubh
            If C_String.IsJsonString(ary(i)) Then
                Call dict.Add(aryHeader(i), JsonConverter.ParseJson(ary(j)))
            Else
            Call dict.Add(aryHeader(i), ary(j))
            End If
            j = j + 1
        Next i
Else  'aryHeaderの方が個数が多い もしくは同じ
    j = lbh
    For i = LBound(ary) To UBound(ary)
        If C_String.IsJsonString(ary(i)) Then
            Call dict.Add(aryHeader(i), JsonConverter.ParseJson(ary(j)))
        Else
            Call dict.Add(aryHeader(i), ary(j))
        End If
        'Call dict.add(aryHeader(j), ary(i))
        j = j + 1
    Next i
    If j > ubh Then
        ''
    Else
        For k = j To ubh
            Call dict.Add(aryHeader(k), "")
        Next k
    End If
End If
Set ArrayToDictionary = dict
''
End Function

Public Function JagArrayToDictionary(ByVal jagary As Variant, _
        Optional colKeyNum As Variant, Optional aryHeader As Variant) As Variant
''''  ********************************************************
''''  * @function JagArrayToArrayDictionary
''
Dim dictRtn As Dictionary
Set dictRtn = New Dictionary
Dim cnt As Long
Dim r As Long
Dim rowStart As Long
Dim colKey As String
If IsMissing(colKeyNum) Then
    Let colKeyNum = 1
End If
''
If IsMissing(aryHeader) Then
    aryHeader = jagary(LBound(jagary))
    rowStart = LBound(jagary) + 1
Else
    rowStart = LBound(jagary)
End If
Dim i As Long
For i = LBound(aryHeader) To UBound(aryHeader)
    If VBA.IsEmpty(aryHeader(i)) Then
        aryHeader(i) = "col" & i
    End If
Next i

If Me.IsDuplicated(aryHeader) Then
        Err.Raise 5, "JagArrayToArrayDictionary", "プロシージャの呼び出し、または引数が不正です。引数の配列に重複があります。"
End If
cnt = -1
For r = rowStart To UBound(jagary)
    Let colKey = jagary(r)(colKeyNum)
    Call dictRtn.Add(colKey, Me.ArrayToDictionary(jagary(r), aryHeader))
Next r
Set JagArrayToDictionary = dictRtn
''
End Function

Public Function JagArrayToArrayDictionary(ByVal jagary As Variant, Optional aryHeader As Variant) As Variant
'''' /********************************************************
''''  *@function JagArrayToArrayDictionary
''''  *<pre>Jag配列（Array(r)(c)を、headerをもとに連想配列の配列形式にする。
''''  *返却形式 Array<Dictionary>(r)
''''  *</pre>
''''  *  @param {Variant<String>(r)(c)} jagary Jag配列
''''  *  @param {Variant<string>(c)} aryHeader ヘッダ配列
''''  *  @return {Variant<Dictionary>(r)} JagArrayToArrayDictionary 連想配列の配列
''''  */
''
Dim aryRtn() As Variant
Dim cnt As Long
Dim r As Long
Dim rowStart As Long
If IsMissing(aryHeader) Then
    aryHeader = jagary(LBound(jagary))
    rowStart = LBound(jagary) + 1
Else
    rowStart = LBound(jagary)
End If
Dim i As Long
For i = LBound(aryHeader) To UBound(aryHeader)
    If VBA.IsEmpty(aryHeader(i)) Then
        aryHeader(i) = "col" & i
    End If
Next i

If Me.IsDuplicated(aryHeader) Then
        Err.Raise 5, "JagArrayToArrayDictionary", "プロシージャの呼び出し、または引数が不正です。引数の配列に重複があります。"
End If
cnt = -1
For r = rowStart To UBound(jagary)
        cnt = cnt + 1
        ReDim Preserve aryRtn(cnt)
    Set aryRtn(cnt) = Me.ArrayToDictionary(jagary(r), aryHeader)
Next r
JagArrayToArrayDictionary = aryRtn
''
End Function


Public Function ArrayArrayToDictionary(ByVal aryary As Variant _
                    , Optional colKeyNum As Variant, Optional aryHeader As Variant) As Dictionary
'''' ********************************************************
''''  *@function ArrayArrayToDictionary
''
Dim jagary As Variant
jagary = Me.ArrayArrayToJagArray(aryary)
''
Dim colKey As String
If IsMissing(colKeyNum) Then
    Let colKeyNum = 1
End If
If IsMissing(aryHeader) Then
    aryHeader = jagary(LBound(jagary))
    Call Me.DeleteElementAt(jagary, LBound(jagary))
End If
If LBound(jagary) = 0 Then
    Let jagary = Me.ChangeIndexOfBegin(jagary)
End If
Set ArrayArrayToDictionary = Me.JagArrayToDictionary(jagary, colKeyNum, aryHeader)
End Function

Public Function ArrayArrayToArrayDictionary(ByVal aryary As Variant _
                                , Optional aryHeader As Variant) As Variant
'''' /********************************************************
''''  *@function ArraygArrayToArrayDictionary
''''  *<pre>Jag配列 Array(r,c)を、ヘッダ配列に合わせて
''''  *Dictionary()に変換する
''''  *  @param {Variant(r,c)} aryary ByVal 配列
''''  *  @return {Variant<Dictionary>(r)} ArrayArrayToArrayDictionary 連想配列の配列
''''  *
''''  */
''
Dim jagary As Variant
jagary = Me.ArrayArrayToJagArray(aryary)
''
If IsMissing(aryHeader) Then
    aryHeader = jagary(LBound(jagary))
    Call Me.DeleteElementAt(jagary, LBound(jagary))
End If
If LBound(jagary) = 0 Then
    Let jagary = Me.ChangeIndexOfBegin(jagary)
End If
ArrayArrayToArrayDictionary = Me.JagArrayToArrayDictionary(jagary, aryHeader)
End Function

Public Function ArrayToArrayArray(ByVal ary As Variant) As Variant
'''' /********************************************************
''''  *@function ArrayToArrayArray
''''  *<br/>配列（Array(i)を、二次元配列(Array(LBound(ary),c)に変換する
''''  *  @param {Variant} ary ByVal 配列
''''  *  @return {Variant} ArrayToArrayArray 二次元配列
''''  *
''''  */
''
Dim aryary As Variant
ReDim aryary(1 To 1, LBound(ary) To UBound(ary))
''
Dim c As Long
For c = LBound(ary) To UBound(ary)
    aryary(1, c) = ary(c)
Next c
ArrayToArrayArray = aryary
''
End Function

Public Function ArrayToArrayArrayVertical(ByVal ary As Variant) As Variant
'''' /********************************************************
''''  *@function ArrayToArrayArrayVertical
''''  *<br/>配列（Array(i)を、二次元配列(Array(r,LBound(ary))に変換する
''''  *  @param {Variant} ary ByVal 配列
''''  *  @return {Variant} ArrayToArrayArrayVertical 二次元配列
''''  *
''''  */
''
Dim aryary As Variant
ReDim aryary(LBound(ary) To UBound(ary), 1 To 1)
''
Dim r As Long
For r = LBound(ary) To UBound(ary)
    aryary(r, 1) = ary(r)
Next r
ArrayToArrayArrayVertical = aryary
''
End Function

Public Function JagArrayToArrayArray(ByVal jagary As Variant) As Variant
'''' /********************************************************
''''  *@function JagAryayToArrayArray
''''  *<br/>ジャグ配列（Array(r)(c)を、二次元配列(Array(r,c)に変換する
''''  *  @param {Variant} aryary ByVal 二次元配列
''''  *  @return {Variant} JagAryayToArrayArray ジャグ配列
''''  *
''''  *  FIXME□□注意：ジャグ配列内の配列に長さに違いがあると変換できない。
''''  */
''
Dim aryary As Variant
aryary = Array()
''
Dim r As Long
Dim c As Long
''
Dim lbInnFst As Long
Dim ubInnFst As Long
''
If Not IsArray(jagary) Then
        Err.Raise 13, "jagAry param is not Array"
End If
''
' 配列の判定
If Me.IsArrayEx(jagary) < 0 Then
        GoTo FINALLY  '長さゼロ
End If
If Me.CountRowForJagArray(jagary) = 1 Then
    ReDim aryary(1 To 1, LBound(jagary(LBound(jagary))) To UBound(jagary(LBound(jagary))))
    For c = LBound(jagary(LBound(jagary))) To UBound(jagary(LBound(jagary)))
        Let aryary(1, c) = jagary(LBound(jagary))(c)
    Next c
    GoTo FINALLY  'only one line
End If
''
' 最初のAry
If Not IsArray(jagary(LBound(jagary))) Then Err.Raise 13
''
lbInnFst = LBound(jagary(LBound(jagary)))
ubInnFst = UBound(jagary(LBound(jagary)))
If ubInnFst - lbInnFst < 0 Then
    For r = LBound(jagary) + 1 To UBound(jagary)
        If (UBound(jagary(r)) - LBound(jagary(r)) + 1) > 0 Then Err.Raise 5
    Next r
    GoTo FINALLY
End If
''
ReDim aryary(LBound(jagary) To UBound(jagary), lbInnFst To ubInnFst)
''
If IsObject(jagary(LBound(jagary))(lbInnFst)) Then
    For r = LBound(jagary) To UBound(jagary)
        If LBound(jagary(r)) <> lbInnFst Or UBound(jagary(r)) <> ubInnFst Then Err.Raise 5
        For c = LBound(jagary(r)) To UBound(jagary(r))
            Set aryary(r, c) = jagary(r)(c)
        Next c
    Next r
Else
    For r = LBound(jagary) To UBound(jagary)
        If LBound(jagary(r)) <> lbInnFst Or UBound(jagary(r)) <> ubInnFst Then Err.Raise 5
        For c = LBound(jagary(r)) To UBound(jagary(r))
            Let aryary(r, c) = jagary(r)(c)
        Next c
    Next r
End If
''
FINALLY:
    JagArrayToArrayArray = aryary
''
End Function

Public Function ArrayArrayToJagArray(ByVal aryary As Variant) As Variant
'''' /********************************************************
''''  *@function ArrayArrayToJagArray
''''  *<br/>二次元配列(Array(r,c)を、ジャグ配列（Array(r)(c)に変換する
''''  *  @param {Variant} aryary ByVal 二次元配列
''''  *  @return {Variant} ArrayArrayToJagArray ジャグ配列
''''  *
''''  */
''
Dim jagary As Variant
jagary = Array()
Dim lb1 As Long
Dim ub1 As Long
lb1 = LBound(aryary, 1) '一次元目の配列
ub1 = UBound(aryary, 1)
If ub1 - lb1 < 0 Then GoTo FINALLY
ReDim jagary(lb1 To ub1)
''
Dim lb2 As Long
Dim ub2 As Long
lb2 = LBound(aryary, 2) '二次元目の配列
ub2 = UBound(aryary, 2)
Dim ary As Variant
ReDim ary(lb2 To ub2)
''
Dim r As Long
Dim c As Long
''
If IsObject(aryary(lb1, lb2)) Then
    For r = lb1 To ub1
        jagary(r) = ary
        For c = lb2 To ub2
            Set jagary(r)(c) = aryary(r, c)
            Next
        Next
Else
    For r = lb1 To ub1
        jagary(r) = ary
        For c = lb2 To ub2
            Let jagary(r)(c) = aryary(r, c)
            Next
        Next
End If
''
FINALLY:
    '返却
    ArrayArrayToJagArray = jagary
End Function

Public Function ArrayArrayToArray(ByVal aryary As Variant) As Variant
'''' /********************************************************
''''  *@function ArrayArrayToArray
''''  *<pre>二次元配列(Array(r,c)を、配列（Array(i)に変換する
''''  * c（左→右）方向を優先し、r方向（上→下）に降りていく
''''  *</pre>
''''  *  @param {Variant} aryary ByVal 二次元配列
''''  *  @return {Variant} ArrayArrayToArray 一次元配列
''''  *
''''  */
''
Dim ary As Variant
ary = Array()
Dim ub As Long
''
Dim lb1 As Long
Dim ub1 As Long
lb1 = LBound(aryary, 1) '一次元目の配列
ub1 = UBound(aryary, 1)
If ub1 - lb1 < 0 Then GoTo FINALLY
''
Dim lb2 As Long
Dim ub2 As Long
lb2 = LBound(aryary, 2) '二次元目の配列
ub2 = UBound(aryary, 2)
''
If lb2 = 0 Then
    ub = (ub1 - lb1 + 1) * (ub2 - lb2 + 1) - 1 ''FIXME 二次元目のLBoundに合わせる
Else
    ub = (ub1 - lb1 + 1) * (ub2 - lb2 + 1)
End If
''
ReDim ary(lb2 To ub)
''
Dim i As Long
i = lb2           ''FIXME
Dim r As Long
Dim c As Long
''
For r = lb1 To ub1
    For c = lb2 To ub2
            If IsObject(aryary(lb1, lb2)) Then
            Set ary(i) = aryary(r, c)
                i = i + 1
            Else
            ary(i) = aryary(r, c)
                i = i + 1
            End If
        Next
Next
''
FINALLY:
    ArrayArrayToArray = ary
''
End Function

Public Function ForEach(ByRef ary As Variant _
                      , ByRef Instance As Object _
                      , ByVal methodname As Variant _
                      , ParamArray pAry() As Variant)
''' /********************************************************
'''  *@function ForEach
'''  *<pre>
'''  *一次元配列に対して、各々のアイテムに、処理を実行する。
'''  *</pre>
'''  *@param {Variant()} ary ByRef
'''  *@param {Object} instance
'''  *@param {Variant} methodname
'''  *@param {Variant()} pAry ParamArray
'''  *※呼び出したmethodの戻り値を、対象Itemにセットする。
'''  *　戻り値は必ず設定する必要がある。
'''  *※pAryは、呼び出すmethodにはそのまま渡す。
'''  *　呼び出される側は、item(文字列を想定）,anAryを引数とする。
'''  */
'
Dim i As Long
For i = LBound(ary) To UBound(ary)
        ary(i) = CallByName(Instance, methodname, VbMethod, CStr(ary(i)), pAry)
Next i
End Function

Public Function ForEachArrayArray(ByRef aryary As Variant _
                      , ByRef Instance As Object _
                      , ByVal methodname As Variant _
                      , ParamArray pAry() As Variant)
''' /********************************************************
'''  *@function ForEachArrayArray
'''  *<pre>
'''  *二次元配列に対して、各々のアイテムに、処理を実行する。
'''  *</pre>
'''  *@param {Variant(r,c)} aryary ByRef
'''  *@param {Object} instance
'''  *@param {Variant} methodname
'''  *@param {Variant()} pAry ParamArray
'''  *※呼び出したmethodの戻り値を、対象Itemにセットする。
'''  *　戻り値は必ず設定する必要がある。
'''  *※pAryは、呼び出すmethodにはそのまま渡す。
'''  *　呼び出される側は、item(文字列を想定）,anAryを引数とする。
'''  */
'
Dim r As Long
Dim c As Long
For r = LBound(aryary, 1) To UBound(aryary, 1)
        For c = LBound(aryary, 2) To UBound(aryary, 2)
            aryary(r, c) = CallByName(Instance, methodname, VbMethod, CStr(aryary(r, c)), pAry)
        Next c
Next r
''
End Function

Public Function ForEachArrayArrayRow(ByRef aryary As Variant _
                      , ByRef Instance As Object _
                      , ByVal methodname As Variant _
                      , ParamArray pAry() As Variant)
''' /********************************************************
'''  *@function ForEachArrayArrayRow
'''  *<pre>
'''  *二次元配列に対して、各々の行に、処理を実行する。
'''  *</pre>
'''  *@param {Variant(r,c)} aryary ByRef
'''  *@param {Object} instance
'''  *@param {Variant} methodname
'''  *@param {Variant()} pAry ParamArray
'''  *※呼び出したmethodの戻り値を、対象Itemにセットする。
'''  *　戻り値は必ず設定する必要がある。
'''  *※pAryは、呼び出すmethodにはそのまま渡す。
'''  *　呼び出される側は、anAry,pAryを引数とする。
'''  *FIXME□□r=1,c=1始まりの場合の動作確認。
'''  *         aryaryがr=0,c=0始まりに変換されないかを確認
'''  */
'
Dim r As Long
Dim json_str As String
Dim jagary As Variant
jagary = Me.ArrayArrayToJagArray(aryary)
For r = LBound(jagary) To UBound(jagary)
    json_str = CallByName(Instance, methodname, VbMethod, C_Commons.ToJson(jagary(r)), pAry)
    jagary(r) = C_Commons.ParseToArray(json_str)
Next r
''
aryary = Me.JagArrayToArrayArray(jagary)
''
End Function

Public Function ForEachJagArray(ByRef jagary As Variant _
                      , ByRef Instance As Object _
                      , ByVal methodname As Variant _
                      , ParamArray pAry() As Variant)
''' /********************************************************
'''  *@function ForEachJagArray
'''  *<pre>
'''  *ジャグ配列に対して、各々のアイテムに、処理を実行する。
'''  *</pre>
'''  *@param {Variant(r)(c)} jagary ByRef
'''  *@param {Object} instance
'''  *@param {Variant} methodname
'''  *@param {Variant()} pAry ParamArray
'''  *※呼び出したmethodの戻り値を、対象Itemにセットする。
'''  *　戻り値は必ず設定する必要がある。
'''  *※pAryは、呼び出すmethodにはそのまま渡す。
'''  *　呼び出される側は、item(文字列を想定）,pAryを引数とする。
'''  */
'
Dim r As Long
Dim c As Long
''
For r = LBound(jagary) To UBound(jagary)
        For c = LBound(jagary(r)) To UBound(jagary(r))
            jagary(r)(c) = CallByName(Instance, methodname, VbMethod, CStr(jagary(r)(c)), pAry)
        Next c
Next r
''
End Function

Public Function ForEachJagArrayRow(ByRef jagary As Variant _
                      , ByRef Instance As Object _
                      , ByVal methodname As Variant _
                      , ParamArray pAry() As Variant)
''' /********************************************************
'''  *@function ForEachJagArrayRow
'''  *<pre>
'''  *ジャグ配列に対して、各々の行に、処理を実行する。
'''  *</pre>
'''  *@param {Variant(r)(c)} jagary ByRef
'''  *@param {Object} instance
'''  *@param {Variant} methodname
'''  *@param {Variant()} pAry ParamArray
'''  *※呼び出したmethodの戻り値を、対象Itemにセットする。
'''  *　戻り値は必ず設定する必要がある。
'''  *※pAryは、呼び出すmethodにはそのまま渡す。
'''  *　呼び出される側は、anAry,pAryを引数とする。
'''  */
'
Dim r As Long
Dim json_str As String
For r = LBound(jagary) To UBound(jagary)
    json_str = CallByName(Instance, methodname, VbMethod, C_Commons.ToJson(jagary(r)), pAry)
    jagary(r) = C_Commons.ParseToArray(json_str)
Next r
End Function

Public Function AddRowForArrayArray(ByRef aryary As Variant _
                        , ByRef ary As Variant) As Variant
'''' /********************************************************
''''  *@function addRowForArrayArray
''''  *<pre>
''''  *二次元配列に対して、行配列を追加する。
''''  *</pre>
''''  *@param {Variant(r,c)} aryary ByRef
''''  *@param {Variant(c)} ary
''''  *※注意：配列の列数は同じであることが前提
''''  */
''
Dim aryaryRtn As Variant
Dim jagary As Variant
jagary = Me.ArrayArrayToJagArray(aryary)
Call Me.AddJagArray(jagary, ary)
'Console.Dump jagAry
aryaryRtn = Me.JagArrayToArrayArray(jagary)
''
AddRowForArrayArray = aryaryRtn
''
End Function

Public Function DeleteRowForArrayArray(ByRef aryary As Variant _
                        , ByVal rowNum As Long)
'''' /********************************************************
''''  *@function deleteRowForArrayArray
''''  *<pre>
''''  *二次元配列に対して、行配列を削除する。
''''  *</pre>
''''  *@param {Variant(r,c)} aryary ByRef
''''  *@param {Long} rowNum aryaryのrを指定
''''  */
''
Dim jagary As Variant
jagary = Me.ArrayArrayToJagArray(aryary)
Call Me.DeleteElementAt(jagary, rowNum)
''
aryary = Me.JagArrayToArrayArray(jagary)
''
End Function

Public Function DeleteColumnForArrayArray( _
    ByRef aryary As Variant, _
    ByVal colNum As Long)
'''' ********************************************************
''
Dim aryaryTransposed() As Variant
Let aryaryTransposed = Me.Transpose(aryary)
Dim jagary As Variant
Let jagary = Me.ArrayArrayToJagArray(aryaryTransposed)
Call Me.DeleteElementAt(jagary, colNum)
''
Let aryary = Me.Transpose(Me.JagArrayToArrayArray(jagary))
''
End Function

Public Function InsertRowForArrayArray(ByRef aryary As Variant _
                        , ByVal rowNum As Long, anAry As Variant)
'''' /********************************************************
''''  *@function deleteRowForArrayArray
''''  *<pre>
''''  *二次元配列に対して、行配列を削除する。
''''  *</pre>
''''  *@param {Variant(r,c)} aryary ByRef
''''  *@param {Long} rowNum aryaryのrを指定
''''  */
''
If Not IsArray(anAry) Then
        Err.Raise 5
End If
If Not IsArray(aryary) Then
        Err.Raise 5
End If
''
Dim aryTemp() As Variant
ReDim aryTemp(LBound(aryary, 2) To UBound(aryary, 2))
Dim i As Long
Dim cnt As Long
i = LBound(anAry)
cnt = LBound(aryTemp)
If UBound(aryTemp) - LBound(aryTemp) >= UBound(anAry) - LBound(anAry) Then
    For i = LBound(anAry) To UBound(anAry)
            aryTemp(cnt) = anAry(i)
            cnt = cnt + 1
        Next i
Else
        For cnt = LBound(aryTemp) To UBound(aryTemp)
            aryTemp(cnt) = anAry(i)
            i = i + 1
        Next cnt
End If
''
Dim jagary As Variant
jagary = Me.ArrayArrayToJagArray(aryary)
Call Me.InsertElementAt(jagary, rowNum, aryTemp)
''
aryary = Me.JagArrayToArrayArray(jagary)
''
End Function

Public Function GetRowForArrayArray(ByRef aryary As Variant _
                        , ByVal rowNum As Long) As Variant
'''' /********************************************************
''''  *@function getRowForArrayArray
''''  *<pre>
''''  *二次元配列に対して、行配列をrowNumで取得する。
''''  *</pre>
''''  *@param {Variant(r,c)} aryary ByRef
''''  *@param {Long} rowNum aryaryのrを指定
''''  *@return {Variant()} getRowForArrayArray
''''  */
''
Dim jagary As Variant
jagary = Me.ArrayArrayToJagArray(aryary)
''
GetRowForArrayArray = jagary(rowNum)
''
End Function

Public Function GetColumnArrayForArrayArray(ByRef aryary As Variant _
                        , ByVal columnNum As Long) As Variant
'''' /********************************************************
''''  *@function getColumnForArrayArray
''''  *<pre>
''''  *二次元配列に対して、列配列をcolumnNumで取得する。
''''  *</pre>
''''  *@param {Variant<String or Long>(r,c)} aryary ByRef
''''  *@param {Long} columnNum aryaryのcを指定
''''  *@return {Variant(r)} getColumnForArrayArray
''''  */
''
Dim aryRtn() As Variant
ReDim aryRtn(LBound(aryary, 1) To UBound(aryary, 1))
''
Dim r As Long
For r = LBound(aryary, 1) To UBound(aryary, 1)
    aryRtn(r) = aryary(r, columnNum)
Next r
''
GetColumnArrayForArrayArray = aryRtn
''
End Function
    
Public Function GetColumnForArrayArray(ByRef aryary As Variant _
                        , ByVal columnNum As Long) As Variant
'''' /********************************************************
''''  *@function getColumnForArrayArray
''''  *<pre>
''''  *二次元配列に対して、列配列をcolumnNumで取得する。
''''  *</pre>
''''  *@param {Variant<String or Long>(r,c)} aryary ByRef
''''  *@param {Long} columnNum aryaryのcを指定
''''  *@return {Variant(r,1)} getColumnForArrayArray
''''  */
''
Dim aryRtn() As Variant
ReDim aryRtn(LBound(aryary, 1) To UBound(aryary, 1), 1 To 1)
''
Dim r As Long
For r = LBound(aryary, 1) To UBound(aryary, 1)
    aryRtn(r, 1) = aryary(r, columnNum)
Next r
''
GetColumnForArrayArray = aryRtn
''
End Function

Public Function GetColumnArrayForJagArray(ByRef jagary As Variant _
                        , ByVal columnNum As Long) As Variant
'''' /********************************************************
''''  *@function getColumnForJagArray
''''  *<pre>
''''  *ジャグ配列に対して、列配列をcolumnNumで取得する。
''''  *</pre>
''''  *@param {Variant<String or Long>(r)(c)} jagary ByRef
''''  *@param {Long} columnNum jagaryのcを指定
''''  *@return {Variant()} getColumnForJagArray
''''  */
''
Dim aryRtn() As Variant
ReDim aryRtn(LBound(jagary) To UBound(jagary))
''
Dim r As Long
For r = LBound(jagary) To UBound(jagary)
    aryRtn(r) = jagary(r)(columnNum)
Next r
''
GetColumnArrayForJagArray = aryRtn
''
End Function

Public Function CountColumnForJagArray(jagary As Variant) As Long
'''' /********************************************************
''''  *@function countColumnForJagArray
''''  *<pre>JagArrayの最大Column数を返却
''''  *</pre>
''''  *  @param {Variant} jagary
''''  *  @return {Long} countColumnForJagArray
''''  */
''
If Not IsArray(jagary) Then
    Err.Raise 13
End If
Dim cntMax As Long
cntMax = 0
Dim r As Long
For r = LBound(jagary) To UBound(jagary)
    If cntMax < (UBound(jagary(r)) - LBound(jagary(r)) + 1) Then
        cntMax = UBound(jagary(r)) - LBound(jagary(r)) + 1
    End If
Next r
''
CountColumnForJagArray = cntMax
''
End Function

Public Function CountRowForJagArray(jagary As Variant) As Long
'''' /********************************************************
''''  *@function countColumnForJagArray
''''  *<pre>JagArrayの最大Column数を返却
''''  *</pre>
''''  *  @param {Variant} jagary
''''  *  @return {Long} countColumnForJagArray
''''  */
''
If Not IsArray(jagary) Then
    Err.Raise 13
End If
''
CountRowForJagArray = UBound(jagary) - LBound(jagary) + 1
''
End Function

Public Function CountColumnForArrayArray(aryary As Variant) As Long
'''' /********************************************************
''''  *@function CountColumnForArrayArray
''''  *<pre>JagArrayの最大Column数を返却
''''  *</pre>
''''  *  @param {Variant} jagary
''''  *  @return {Long} countColumnForJagArray
''''  */
''
If Not IsArray(aryary) Then
    Err.Raise 13
End If
Dim cnt As Long
Let cnt = UBound(aryary, 2) - LBound(aryary, 2) + 1
''
CountColumnForArrayArray = cnt
''
End Function

Public Function CountRowForArrayArray(aryary As Variant) As Long
'''' /********************************************************
''''  *@function CountRowForArrayArray
''''  *<pre>JagArrayの最大Column数を返却
''''  *</pre>
''''  *  @param {Variant} jagary
''''  *  @return {Long} countColumnForJagArray
''''  */
''
If Not IsArray(aryary) Then
    Err.Raise 13
End If
Dim cnt As Long
Let cnt = UBound(aryary, 1) - LBound(aryary, 1) + 1
''
CountRowForArrayArray = cnt
''
End Function

Public Sub FillEmptySpaceForArray(ByRef ary As Variant, Optional ByVal str As String = "", Optional ByVal trimOn As Boolean = False)
'''' /********************************************************
''''  *@function FillEmptySpaceForArray
''''  *<br/>空の配列要素に、""を代入する。
''''  */
Dim C_String As C_String
Set C_String = New C_String
Dim i As Long
For i = LBound(ary) To UBound(ary)
    If (ary(i) = "") Or IsEmpty(ary(i)) Or IsNull(ary(i)) Then
        ary(i) = str
    Else
        If trimOn Then
            ary(i) = C_String.TrimSpace(ary(i))
        End If
    End If
Next i
''
End Sub

Public Sub FillEmptySpaceForArrayArray(ByRef aryary As Variant, Optional ByVal str As String = "", Optional ByVal trimOn As Boolean = False)
'''' /********************************************************
''''  *@function FillEmptySpaceForArrayArray
''''  *<br/>空の配列要素に、""を代入する。
''''  */
Dim C_String As C_String
Set C_String = New C_String
Dim r As Long
Dim c As Long
For r = LBound(aryary, 1) To UBound(aryary, 1)
    For c = LBound(aryary, 2) To UBound(aryary, 2)
        If (aryary(r, c) = "") Or IsEmpty(aryary(r, c)) Or IsNull(aryary(r, c)) Then
            aryary(r, c) = str
        Else
            If trimOn Then
                aryary(r, c) = C_String.TrimSpace(aryary(r, c))
            End If
        End If
    Next c
Next r
''
End Sub

Public Sub FillEmptySpaceForJagArrary(ByRef ary As Variant, Optional ByVal str As String = "", Optional ByVal trimOn As Boolean = False)
'''' /********************************************************
''''  *@function fillEmptyElementForJagArrary
''''  *<br/>空の配列要素に、""を代入する。
''''  */
Dim C_String As C_String
Set C_String = New C_String
Dim r As Long
Dim c As Long
For r = LBound(ary) To UBound(ary)
    For c = LBound(ary(r)) To UBound(ary(r))
        If (ary(r)(c) = "") Or IsEmpty(ary(r)(c)) Or IsNull(ary(r)(c)) Then
            ary(r)(c) = str
        Else
            If trimOn Then
                ary(r)(c) = C_String.TrimSpace(ary(r)(c))
            End If
        End If
    Next c
Next r
''
End Sub

Public Function ToJson(Value)
''' /********************************************************
'''  * @function ToJson
'''  * @param {Variant} value
'''  * @return {Variant} ToJson
'''  */
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Array.ToJson"
    Dim clt As Collection
    Set clt = New Collection
    Dim elm As Variant
    Dim i As Long
    Dim str As String
    Dim C_Commons As C_Commons
    Set C_Commons = New C_Commons
    Dim C_String As C_String
    Set C_String = New C_String
    Dim arytype As String
    Dim jagary As Variant
'
'On Error GoTo EXCEPTION
On Error Resume Next
    Let arytype = Me.TypeArrayPattern(Value)
    '配列と判定されたvalueを処理
    If arytype = "ArrayNotInit" Then
        ToJson = "[]"
    ElseIf arytype = "ArrayEmpty" Then
        ToJson = "[]"
    ElseIf arytype = "Array" Then
        For Each elm In Value
            If C_String.IsJsonString(elm) Then
                clt.Add elm
            Else
                clt.Add C_Commons.toString(elm)
            End If
        Next
        For i = 1 To clt.Count
            str = str & clt.Item(i) & ","
        Next i
        ToJson = "[" & left(str, Len(str) - 1) & "]"
    ElseIf arytype = "JagArray" Then
        str = C_Commons.ShowJagArray(Value)
        ToJson = str
    ElseIf arytype = "ArrayArray" Then
        Let jagary = Me.ArrayArrayToJagArray(Value)
        str = C_Commons.ShowJagArray(jagary)
        ToJson = str
    Else
        ToJson = C_Commons.toString(Value.Name)
    End If
    
    '
    'ToStringでエラー
    If Err.Number <> 0 Then
        If IsObject(Value) Then
            ToJson = C_Commons.toString(Value.Name)
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
End Function

Public Function ToCsvRow(ByRef ary As Variant, Optional delimiter As Variant = vbTab) As String
''' /********************************************************
'''  *@function ToCsvRow
'''  *<pre>
'''  *配列<String>をCSV形式で返却する
'''  *</pre>
'''  *@param {Variant<String>()} ary ByRef
'''  *@param {Variant<String>} delimiter
'''  *@return {String} ToCsvRow
'''  */
'
Dim aryTemp() As Variant
ReDim aryTemp(LBound(ary) To UBound(ary))
Dim i As Long
For i = LBound(ary) To UBound(ary)
        aryTemp(i) = C_Commons.ToJson(ary(i))
Next i
ToCsvRow = Join(aryTemp, delimiter)
''
End Function

Public Function ToLine(ByRef ary As Variant, Optional ByVal padString As Variant) As String
'''' ********************************************************
''
Dim C_String As C_String
Set C_String = New C_String
''
If VBA.IsMissing(padString) Then
    Let padString = C_String.Rept(" ", 4)
End If
''
Dim i As Long
Dim line As String
Dim s As String
For i = LBound(ary) To UBound(ary)
    Let s = VBA.Trim(ary(i))
    Let line = line + C_String.DefaultString(s, padString)
Next i
''
Let ToLine = line
''
End Function


Public Function ToCsv(ByRef aryAllType As Variant, Optional delimiter As Variant = vbTab) As String
''' /********************************************************
'''  *@function ToCsv
'''  *<pre>
'''  *配列をCSV形式で返却する
'''  *</pre>
'''  *@param {Variant<String>()} aryAllType ByRef
'''  *@param {Variant<String>} delimiter
'''  *@return {String} ToCsv
'''  */
'
If Not IsArray(aryAllType) Then
        Err.Raise 13
End If
Dim aryTemp() As Variant
Dim aryBuff() As Variant
Dim strTemp As String
strTemp = Me.TypeArrayPattern(aryAllType)
''
Dim r As Long
''
Select Case strTemp
        Case CONS_ARRAY_TYPE(ARRAY_TYPE.ARRAY_EMPTY)
            ToCsv = ""
            Exit Function
        Case CONS_ARRAY_TYPE(ARRAY_TYPE.ARRAY_NORMAL)
            aryTemp = Array(aryAllType)
        Case CONS_ARRAY_TYPE(ARRAY_TYPE.JagArray)
            aryTemp = aryAllType
        Case CONS_ARRAY_TYPE(ARRAY_TYPE.ARRAY_ARRAY)
            aryTemp = Me.ArrayArrayToJagArray(aryAllType)
        Case Else
            Err.Raise 13
End Select
''
ReDim aryBuff(LBound(aryTemp) To UBound(aryTemp))
''
For r = LBound(aryTemp) To UBound(aryTemp)
        aryBuff(r) = Me.ToCsvRow(aryTemp(r), delimiter)
Next r
''
ToCsv = Join(aryBuff, vbCrLf)
''
End Function
    
Public Function ToHtmlTable(ByRef aryAllType As Variant, Optional aryHeader As Variant, Optional templ As Variant) As String
''' /********************************************************
'''  *@function ToHtmlTable
'''  */
'
Dim C_String As C_String
Set C_String = New C_String
''
If Not IsArray(aryAllType) Then
    Err.Raise 13
End If
Dim aryTemp() As Variant
Dim aryBuff() As Variant
Dim strAryType As String
Let strAryType = Me.TypeArrayPattern(aryAllType)
''
Dim r As Long
''
Select Case strAryType
    Case CONS_ARRAY_TYPE(ARRAY_TYPE.ARRAY_EMPTY)
        ToHtmlTable = ""
        Exit Function
    Case CONS_ARRAY_TYPE(ARRAY_TYPE.ARRAY_NORMAL)
        aryTemp = Array(aryAllType)
    Case CONS_ARRAY_TYPE(ARRAY_TYPE.JagArray)
        aryTemp = aryAllType
    Case CONS_ARRAY_TYPE(ARRAY_TYPE.ARRAY_ARRAY)
        aryTemp = Me.ArrayArrayToJagArray(aryAllType)
    Case Else
        Err.Raise 13
End Select
''
If IsMissing(aryHeader) Then
    Let aryHeader = aryTemp(LBound(aryTemp))
    Call Me.DeleteElementAt(aryTemp, LBound(aryTemp))
End If
If IsMissing(templ) Then
    Let templ = "<table class='table table-striped'>${thead}${tbody}</table><link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css' integrity='sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T' crossorigin='anonymous'>"
End If
''
Dim thead As String
Dim tbody As String
Dim dict As Dictionary
Set dict = New Dictionary
Dim c As Long
Dim aryHeaderBuff() As Variant
ReDim aryHeaderBuff(LBound(aryHeader) To UBound(aryHeader))
For c = LBound(aryHeader) To UBound(aryHeader)
    Let aryHeaderBuff(c) = "<th scope='col'>" & C_String.DefaultString(aryHeader(c)) & "</th>"
Next c
ReDim aryBuff(LBound(aryTemp) To UBound(aryTemp))
For r = LBound(aryTemp) To UBound(aryTemp)
    Let aryBuff(r) = Me.ToHtmlRow(aryTemp(r))
Next r
Let thead = "<thead class='thead-dark'><tr>" & Join(aryHeaderBuff, "") & "</tr></thead>"
Let tbody = "<tbody class='thead-dark'>" & Join(aryBuff, vbCrLf) & "</tbody>"
Call dict.Add("thead", thead)
Call dict.Add("tbody", tbody)
ToHtmlTable = C_String.RenderTemplate(templ, dict)
''
End Function

Public Function ToHtmlRow(ByRef ary As Variant) As String
''' /********************************************************
'''  *@function ToHtmlRow
'''  *@param {Variant<String>()} ary ByRef
'''  *@param {Variant<String>} delimiter
'''  *@return {String} ToHtmlRow
'''  */
'
Dim C_String As C_String
Set C_String = New C_String
''
Dim aryTemp() As Variant
ReDim aryTemp(LBound(ary) To UBound(ary))
Dim i As Long
For i = LBound(ary) To UBound(ary)
        aryTemp(i) = "<td>" & C_String.DefaultString(ary(i)) & "</td>"
Next i
''
ToHtmlRow = "<tr>" & Join(aryTemp, "") & "</tr>"
End Function

'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Sub UnitTest()
    Dim CONS_MODULE_NAME As String
    Let CONS_MODULE_NAME = "C_Array.unittest"
On Error GoTo EXCEPTION
    
    ''
    'Call debugC_Array.unitTest
    'Call debugC_Array.unittest2
    'Call debugC_Array.unittest3
    ''
    GoSub FINALLY
    Exit Sub
    
FINALLY: 'Label
    '//--------------------------------------
    '// 終了処理
    Return

EXCEPTION: 'Label
    '//--------------------------------------
    ' 例外処理
    GoSub FINALLY
    Console.info ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    
End Sub
