VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Dictionary.cls
'''  *<pre> Dictionary Utility Class
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  */
'
''' /********************************************************
'''  * @class C_Dictionary
'''  * @classdesc
'''  * @note
'''  *
'''  *
'''  * @require C_Array, C_String
'''  *
'''  * @require must refer to Microsoft.Scripting.Runtime
'''  */
'
'//-----------------------------------------
'// Class Method
Public Function IsDictionary(vnt As Variant) As Boolean
''' /********************************************************
'''  *@function  IsDictionary
'''  *<pre>Dictionary型かを判定する
'''  *</pre>
'''  * @param {Variant} vnt
'''  * @return {Boolean} IsDictionary
'''  */
'
Dim blnRtn As Boolean
If IsObject(vnt) Then
    If TypeName(vnt) = "Dictionary" Then
        blnRtn = True
    Else
        blnRtn = False
    End If
Else
    blnRtn = False
End If
IsDictionary = blnRtn
End Function

Public Function AppendDictionary(ByRef dict As Dictionary _
                                , ByVal key As Variant) As Dictionary
''' /********************************************************
'''  *@function  AppendDictionary
'''  *<pre>新しいDictinary型の変数を、dictに追加(Append)する
'''  *使い道は不明。変数に代入せざるえない・・
'''  *</pre>
'''  * @param {Variant} vnt
'''  * @return {Dictionary}
'''  */
'
    Call Me.PutItem(dict, key, New Dictionary)
    Set AppendDictionary = dict(key)
End Function

Public Sub PutItem(ByRef dict As Dictionary _
                , ky As Variant _
                , itm As Variant)
''' /********************************************************
'''  *@function  putItem
'''  *<br/> set item to key and
'''  *<br/> if key not exists, create new key and item
'''  *</pre>
'''  * @param {Dictionary} dict
'''  * @param {Variant<String or Long>} ky
'''  * @param {Variant} itm
'''  */
'
If dict.Exists(ky) Then
    If IsObject(itm) Then
        Set dict(ky) = itm
    Else
        Let dict(ky) = itm
    End If
Else
    Call dict.Add(ky, itm)
End If
End Sub

Public Sub PutParamArray(ByRef dict As Dictionary, ParamArray pAry())
''' /********************************************************
'''  * @function  putParamArray
'''  *<br/> put items with ParamArray key1,value1,key2,value2,key3,value3
'''  * @param {ParamArray String[]} pAry()
'''  */
Dim i As Long
Dim cnt As Long
Dim ub As Long
Dim lb As Long
ub = UBound(pAry)
lb = LBound(pAry)
cnt = lb
For i = lb To ub Step 2
    Call Me.PutItem(dict, pAry(i), pAry(i + 1))
    cnt = cnt + 2
    If ub - lb + 1 - cnt = 1 Then
        Exit For
    End If
Next i
'
End Sub

Public Sub MergeDictionary(ByRef dictsrc As Dictionary, ByRef dictdst As Dictionary)
''' /********************************************************
'''  *@function  mergeDictionary
'''  *<br/>渡されたdictdstを、dictsrcにマージする。
'''  *<br/>
'''  * @param {Dictionary} dictsrc  As Dictionary
'''  * @param {Dictionary} dictdst  As Dictionary
'''  *
'''  */
Dim keyVal As Variant
For Each keyVal In dictdst
        If dictsrc.Exists(keyVal) Then
            If IsObject(dictdst.Item(keyVal)) Then
                Set dictsrc.Item(keyVal) = dictdst.Item(keyVal)
            Else
                dictsrc.Item(keyVal) = dictdst.Item(keyVal)
            End If
        Else
            dictsrc.Add keyVal, dictdst.Item(keyVal)
        End If
Next keyVal
''
End Sub

Public Sub PutDictionaryByKey(ByRef dict As Dictionary _
                        , ByRef dictdst As Dictionary _
                        , ParamArray pAry())
''' /********************************************************
'''  *@function  putDictionaryByComplexKey
'''  *<br/>渡されたdictdstを、dictにputする。
'''  *<br/>
'''  * @param {Dictionary} dict  As Dictionary
'''  * @param {Dictionary} dictdst  As Dictionary
'''  * @param {ParamArray} pAry キーとするアイテム名
'''  *
'''  */
    Dim i As Long
    Dim cnt As Long
    cnt = 0
    Dim aryKeyVal() As Variant
    Dim strKeyVal As String
    
    For i = LBound(pAry) To UBound(pAry)
        If dictdst.Exists(pAry(i)) Then
            ReDim Preserve aryKeyVal(cnt)
            aryKeyVal(cnt) = dictdst(pAry(i))
            cnt = cnt + 1
        Else
            Console.log "Not Exists key=" & pAry(i)
        End If
    Next i
    strKeyVal = Join(aryKeyVal, "|") ''|でつないでKeyとする
    Call Me.PutItem(dict, strKeyVal, dictdst)
End Sub

Public Sub PutDictionaryByKeyArray(ByRef dict As Dictionary _
                        , ByRef dictdst As Dictionary _
                        , keyAry As Variant)
''' /********************************************************
'''  *@function  putDictionaryByKeyArray
'''  *<br/>渡されたdictdstからkeyArrayでExtractしてdictにputする。
'''  *<br/>
'''  * @param {Dictionary} dict  As Dictionary
'''  * @param {Dictionary} dictdst  As Dictionary
'''  * @param {ParamArray} keyAry キーとするアイテム名
'''  *
'''  */
    Dim i As Long
    Dim cnt As Long
    cnt = 0
    Dim aryKeyVal() As Variant
    Dim strKeyVal As String
    
    For i = LBound(keyAry) To UBound(keyAry)
        If dictdst.Exists(keyAry(i)) Then
            ReDim Preserve aryKeyVal(cnt)
            aryKeyVal(cnt) = dictdst(keyAry(i))
            cnt = cnt + 1
        Else
            Console.log "Not Exists key=" & keyAry(i)
        End If
    Next i
    strKeyVal = Join(aryKeyVal, "|") ''|でつないでKeyとする
    Call Me.PutItem(dict, strKeyVal, dictdst)
End Sub

Public Function ExistsItem(ByRef dict As Dictionary, ByVal str As String) As Boolean
''' /********************************************************
'''  *@function ExistsItem
'''  *<br/>指定した項目値の存在を確認する
'''  *<br/>
'''  * @return {Dictionary} dict  As Dictionary
'''  * @return {String} str  As String
'''  *
'''  * @return {Boolean} ExistsItem
'''  *
'''  * FIXME□□未実装
'''  */
    '指定した項目がディクショナリに含まれているかどうかを返却
    ' Itemsを利用して、対応する
    Dim bln As Boolean
    bln = False
    Dim k As Variant

    Dim dictTemp As Dictionary
    Set dictTemp = Me.FilterMatchItem(dict, str)
    
    If dictTemp.Count = 0 Then
        bln = False
    Else
        For Each k In dictTemp
            If str = dictTemp(k) Then
                bln = True
                Exit For
            Else
                bln = False
            End If
        Next k
    End If
    ExistsItem = bln
End Function

Public Function GetKeyByItem(ByRef dict As Dictionary, ByVal itm As String) As String
'''' /********************************************************
''''  *@function GetKeyByItem
''''  *<br/>指定した項目値の存在を確認する
''''  *<br/>
''''  * @return {Dictionary} dict  As Dictionary
''''  * @return {String} itm  As String
''''  *
''''  * @return {String}
''''  *
''''  */
'' 指定した項目がディクショナリに含まれているかどうかを返却
'' Itemsを利用して、対応する
Dim keyRtn As String
Let keyRtn = ""
Dim ky As Variant
''
If Me.ExistsItem(dict, itm) Then
    For Each ky In dict
        If itm = dict(ky) Then
            keyRtn = ky
            Exit For
        End If
    Next ky
End If
GetKeyByItem = keyRtn
''
End Function

Public Function Filter(ByRef dict As Variant, ByVal likestr As String, Optional ByVal Include As Variant) As Dictionary
''' /********************************************************
'''  *@function Filter
'''  *<br/> Filer keys of dictionary  with VBA.Filter and
'''  *<br/> make subset dictioary object
'''  *<br/>
'''  * @param {Object} dict  As Dictionary
'''  * @param {String} likestr  As String
'''  * @param {Variant} include  As Variant (Optional Boolean)
'''  * @return {Dictionary}
'''  */
'ディクショナリのkey,itemを追加する。
Dim dictRet As Dictionary
Set dictRet = New Dictionary
Dim ary As Variant
Dim aryKeyTemp As Variant
Dim i As Long
'
If IsMissing(Include) Then
    Include = True
End If
ary = dict.keys
aryKeyTemp = VBA.Filter(ary, likestr, Include)
'
For i = LBound(aryKeyTemp) To UBound(aryKeyTemp)
    dictRet.Add aryKeyTemp(i), dict(aryKeyTemp(i))
Next i
Set Filter = dictRet
'
End Function

Public Function FilterUnmatch(ByRef dict As Dictionary, ByRef aryExcluded As Variant) As Dictionary
''' /********************************************************
'''  *@function FilterUnmatch
'''  *<br/>
'''  *<br/>
'''  *<br/>
'''  * @param {Object} dict  As Dictionary
'''  * @param {String[]} aryExcluded  As Variant
'''  * @return {Dictionary}
'''  */
Dim dictTemp As Dictionary
Set dictTemp = Me.DeepCopy(dict)

Dim i As Integer
For i = LBound(aryExcluded) To UBound(aryExcluded)
    Set dictTemp = Me.Filter(dictTemp, aryExcluded(i), False)
Next i
'
Set FilterUnmatch = dictTemp
'
End Function

Public Function DeepCopy(ByRef dict As Variant) As Dictionary
''' /********************************************************
'''  *@function deepCopy
'''  *<br/>
'''  *<br/>
'''  *<br/>
'''  * @param {Object} dict  As Dictionary
'''  * @return {Dictionary}
'''  *
'''  * @require Scripting.Dictionary 参照設定
'''  */
' FIXME Out of Stack space
Dim dictTemp As Dictionary
Set dictTemp = New Dictionary
'
Dim dummy As String
dummy = "dummy-----------------------------"
Set dictTemp = Me.Filter(dict, dummy, False)
Set DeepCopy = dictTemp
'
End Function

Public Function FilterMatchItem(ByRef dict As Dictionary, ByVal str As String) As Dictionary
''' /********************************************************
'''  *@function FilterMatchItem
'''  *<br/> 検索文字列がアイテムに含まれるもので抽出して、
'''  *<br/> サブセットのDictionaryを返却する。
'''  *<br/>   注意：Itemは文字列が前提
'''  *<br/>
'''  * @param {Dictionary} dict
'''  * @param {String} str
'''  * @return {Dictionary} FilterMatchItem
'''  *
'''  * @require Scripting.Dictionary 参照設定
'''  */
    'ディクショナリのkey,itemを追加する。
    Dim dictRet As Dictionary
    Set dictRet = New Dictionary
    Dim aryKey As Variant
    Dim aryKeyTemp() As Variant
    Dim aryItem As Variant
    Dim i As Long
    Dim cnt As Long
    cnt = -1
    '
    aryKey = dict.keys
    aryItem = dict.Items
    '
    For i = LBound(aryItem) To UBound(aryItem)
        If aryItem(i) Like ("*" & str & "*") Then
                cnt = cnt + 1
                ReDim Preserve aryKeyTemp(cnt)
                aryKeyTemp(cnt) = aryKey(i)
        End If
    Next i
    '
    If cnt >= 0 Then
        For i = LBound(aryKeyTemp) To UBound(aryKeyTemp)
            dictRet.Add aryKeyTemp(i), dict(aryKeyTemp(i))
        Next i
        Set FilterMatchItem = dictRet
    Else
        Set FilterMatchItem = dictRet
    End If
End Function

Public Function FilterMatchPrefix(ByRef dict As Dictionary, _
                                  ByVal prefix As String) As Dictionary
''' /********************************************************
'''  *@function FilterMatchPrefix
'''  *<br/>
'''  * @param {Dictionary} dict
'''  * @param {String} prefix
'''  * @return {Dictionary}
'''  */
Dim dictRet As Dictionary
Set dictRet = New Dictionary
Dim ary As Variant
Dim aryKeyTemp As Variant
Dim i As Long
'
ary = dict.keys
aryKeyTemp = VBA.Filter(ary, prefix)
'
For i = LBound(aryKeyTemp) To UBound(aryKeyTemp)
    If aryKeyTemp(i) Like (prefix & "*") Then
        dictRet.Add aryKeyTemp(i), dict(aryKeyTemp(i))
    End If
Next i
'
Set FilterMatchPrefix = dictRet
'
End Function

Public Function FilterMatchWithRegexp(ByRef dict As Dictionary, _
                                  ByVal regStr As String) As Dictionary
''' /********************************************************
'''  *@function FilterMatchWithRegexp
'''  *<br/>
'''  * @param {Dictionary} dict
'''  * @param {String} regStr
'''  * @return {Dictionary}
'''  */
Dim C_String As C_String
Set C_String = New C_String
Dim dictRet As Dictionary
Set dictRet = New Dictionary
Dim ary As Variant
Dim i As Long
'
ary = dict.keys
'
For i = LBound(ary) To UBound(ary)
    If C_String.Test(ary(i), regStr) Then
        dictRet.Add ary(i), dict(ary(i))
    End If
Next i
'
Set FilterMatchWithRegexp = dictRet
'
End Function


Public Function Contains(ByRef dict As Dictionary, ByVal likestr As String) As Boolean
''' /********************************************************
'''  *@function  Contains
'''  *<br/>指定した文字列をLikeでキーの存在を確認する
'''  *<br/>
'''  * @return {Dictionary} dict  As Dictionary
'''  * @return {String} like As String
'''  *
'''  * @return {Boolean} Contains
'''  *
'''  */
    '指定した文字列がディクショナリのキーにLikeで含まれているかどうかを返却
    Dim bln As Boolean
    '
    Dim dictTemp As Dictionary
    Set dictTemp = Me.Filter(dict, likestr)
    '
    If dictTemp.Count = 0 Then
        bln = False
    Else
        bln = True
    End If
    Contains = bln
End Function



Public Function ToArray(ByRef dict As Dictionary, Optional ByRef aryHeader As Variant) As Variant
''' /********************************************************
'''  *@function ToArray
'''  *<br/>aryHeaderに対応した配列に変換する。
'''  *<br/>
'''  * @param {Dictionary} dict
'''  * @param {Variant<String>()} aryHeader
'''  * @return {Variant()} ToArray
'''  *
'''  */
    Dim aryRtn() As Variant
    If IsMissing(aryHeader) Then
        Let aryHeader = dict.keys
    End If
    ReDim aryRtn(LBound(aryHeader) To UBound(aryHeader))
    '
    Dim i As Long
    For i = LBound(aryHeader) To UBound(aryHeader)
        If dict.Exists(aryHeader(i)) Then
            If IsObject(dict(aryHeader(i))) Then
                Set aryRtn(i) = dict(aryHeader(i))
            Else
                Let aryRtn(i) = dict(aryHeader(i))
            End If
        Else
            aryRtn(i) = ""
        End If
    Next i
    ToArray = aryRtn
End Function

Public Function ToArrayArray(ByRef dict As Dictionary _
                           , Optional ByRef aryHeader As Variant _
                           , Optional arytype As Variant = "V") As Variant
''' /********************************************************
'''  *@function ToArray
'''  *<br/>aryHeaderに対応した配列に変換する。
'''  *<br/>
'''  * @param {Dictionary} dict
'''  * @param {Variant<String>()} aryHeader
'''  * @return {Variant()} ToArrayArray
'''  */
    Dim C_Array As C_Array
    Set C_Array = New C_Array
    Dim aryaryRtn As Variant
    If IsMissing(aryHeader) Then
        Let aryHeader = dict.keys
    End If
    Dim aryTemp As Variant
    aryTemp = Me.ToArray(dict, aryHeader)
    If arytype = "V" Then
        aryaryRtn = C_Array.ArrayToArrayArrayVertical(aryTemp)
    Else
        aryaryRtn = C_Array.ArrayToArrayArray(aryTemp)
    End If
    ToArrayArray = aryaryRtn
    Set C_Array = Nothing
End Function

Public Function Extract(ByRef dict As Dictionary, ByRef aryHeader As Variant) As Dictionary
''' /********************************************************
'''  *@function ToArray
'''  *<br/>aryHeaderのキーで抽出したDictionaryを返却
'''  *<br/>
'''  * @param {Dictionary} dict
'''  * @param {Variant<String>()} aryHeader
'''  * @return {Variant()} Extract
'''  *
'''  */
    Dim dictRtn As Dictionary
    Set dictRtn = New Dictionary
    '
    Dim i As Long
    For i = LBound(aryHeader) To UBound(aryHeader)
        If dict.Exists(aryHeader(i)) Then
            Call dictRtn.Add(aryHeader(i), dict(aryHeader(i)))
        Else
            Call dictRtn.Add(aryHeader(i), "")
        End If
    Next i
    Set Extract = dictRtn
    Set dictRtn = Nothing
End Function


'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Sub UnitTest()

    '//--------------------------------------
    '//変数宣言
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Dictionary.unitTest"
    Dim dict As Dictionary
    Dim dictTemp As Dictionary
    '
On Error GoTo EXCEPTION
    '//--------------------------------------
    '// 初期処理
    '
    Set dict = New Dictionary
    Set dictTemp = New Dictionary
    '//--------------------------------------
    '// 処理
'''primitive
    '
    
    
    
    GoSub FINALLY
    Exit Sub
    '
FINALLY: 'Label
    '//--------------------------------------
    Set dict = Nothing
    Set dictTemp = Nothing
    Return

EXCEPTION: 'Label
    '//--------------------------------------
    GoSub FINALLY
    Console.info ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    '
End Sub
