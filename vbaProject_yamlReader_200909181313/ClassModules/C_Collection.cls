VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Collection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Collection.cls
'''  *<pre>Collection Utility
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  *
'''  */
'
''' /********************************************************
'''  * @class C_Collection
'''  * @classdesc コレクション用ユーティリティ・クラス
'''  */
'
'//--------------------------------------------------------
'// Class Event
'//
Private Sub Class_Initialize()
''' /********************************************************
'''  * @event Class_Initialize
'''  *<br/>Initialize 関数
'''  *<br/>Class生成時に、実行される。
'''  */
End Sub

'//--------------------------------------------------------
'// Class Method
'//
Public Function TypeNameEx(ByRef clts As Variant) As String
''' /********************************************************
'''  *@function TypeNameEx
'''  *<br/> Collectionオブジェクトの型を確認する。（覚書として）
'''  *  @param {Variant} clts ByRef
'''  *  @return {String} TypeNameEx
'''  */
'
Dim strType As String
Dim strRtn As String
strType = TypeName(clts)
Select Case strType
        Case "Collection"
            strRtn = "Collection<Item>(idx_or_key)" & " <" & strType & ">"
        Case "Workbooks"
            strRtn = "Collection<Item>(idx_or_name).Name" & " <" & strType & ">"
        Case "Worksheets"
            strRtn = "Collection<Item>(idx_or_name).Name" & " <" & strType & ">"
        Case "Sheets"
            strRtn = "Collection<Item>(idx_or_name).Name" & " <" & strType & ">"
        Case "Shapes"
            strRtn = "Collection<Item>(idx_or_name).Name/TextFrame.Charactors.Text" & " <" & strType & ">"
        Case "ListObjects"
            strRtn = "Collection<Item>(idx_or_name).Name/Range" & " <" & strType & ">"
        Case "Names"
            strRtn = "Collection<Item>(idx_or_name).Name/Address/Value" & " <" & strType & ">"
        Case "Areas"
            strRtn = "Collection<Item>(idx).Range.Address" & " <" & strType & ">"
        Case "Files"
            strRtn = "Collection<Item>(idx).Name" & " <" & strType & ">"
        Case "Folders"
            strRtn = "Collection<Item>(idx).Name" & " <" & strType & ">"
        Case "SubFolders"
            strRtn = "Collection<Item>(idx).Name" & " <" & strType & ">"
        Case Else
            strRtn = "Unknown"
End Select
TypeNameEx = strRtn
''
End Function

Public Function VarTypeEx(ByRef clts As Variant) As String
''' /********************************************************
'''  *@function TypeNameEx
'''  *<br/> Collectionオブジェクトの型を確認する。（覚書として）
'''  *  @param {Variant} clts ByRef
'''  *  @return {String} TypeNameEx
'''  */
'
Dim strType As String
Dim strRtn As String
strType = TypeName(clts)
Select Case strType
        Case "Collection"
            strRtn = "clts(idx_or_key)"
        Case "Workbooks"
            strRtn = "clts(idx_or_name).CodeName"
        Case "Worksheets"
            strRtn = "clts(idx_or_name).Index"
        Case "Sheets"
            strRtn = "clts(idx_or_name).Index"
        Case "Shapes"
            strRtn = "clts(idx_or_name)"
        Case "ListObjects"
            strRtn = "clts(idx_or_name).Range"
        Case "Names"
            strRtn = "clts(idx_or_name).Index"
        Case "Areas"
            strRtn = "clts(idx).Address"
        Case "Files"
            strRtn = "clts(idx).Name"
        Case "Folders"
            strRtn = "clts(idx).Name"
        Case "SubFolders"
            strRtn = "clts(idx).Name"
        Case Else
            strRtn = "Unknown"
End Select
VarTypeEx = strRtn
''
End Function

Public Function toString(ByRef clts As Variant) As String
'''' /********************************************************
''''  *@function ToString
''''  *<pre> 可視できる文字列に変換する
''''  *</pre>
''''  *  @param {Variant} clts ByRef
''''  *  @return {String} ToString
''''  */
''
Dim strRtn As String
Dim strName As String
strName = TypeName(clts)
Select Case strName
        Case "Collection"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "Workbooks"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowWorkbooks(clts)
        Case "Worksheets"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "Sheets"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "Shapes"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "ListObjects"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "Names"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "Areas"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "Files"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "Folders"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case "SubFolders"
            strRtn = "<" & strName & "> " & clts.Count & ": " & C_Commons.ShowCollection(clts)
        Case Else
            strRtn = "<Unknown>" & C_Commons.ShowCollection(clts)
End Select
''
toString = strRtn
''
End Function

Public Function Exists(ByRef clts As Variant, ByVal key As Variant) As Boolean
'''' /********************************************************
''''  *@function Exists
''''  *<pre> コレクションに、指定したKeyが存在するかを判定する。
''''  *</pre>
''''  *  @param {Variant} clts ByRef
''''  *  @param {Variant} key index or key
''''  *  @return {Boolean}
''''  *
''''  */
''
Dim dammy As Variant
Dim blnRtn As Boolean
''
On Error Resume Next
    Err.clear
    ''
    dammy = clts(key)
    ''
    If Err.Number = 0 Then
        blnRtn = True
    Else
        Err.clear
        dammy = IsObject(clts.Item(key)) '''item(key)が取得できないと実行エラー発生
            If Err.Number = 0 Then
                blnRtn = True
            Else
                blnRtn = False
                If Err.Number = 438 Then
                    Err.Description = Err.Description & vbCrLf & "このコレクションにItemプロパティが存在しないため、判別できません"
                    Console.log Err.Description
                End If
            End If
    End If
    Exists = blnRtn
End Function

Public Sub PutValue(ByRef clts As Collection, ByVal Value As Variant, ByVal key As Variant)
''' /********************************************************
'''  *@function
'''  *<pre> コレクションに、指定したKey-Valueをputする。
'''  *</pre>
'''  *  @param {Variant} clts ByRef
'''  *  @param {Variant} value
'''  *  @param {Variant} key index or key
'''  *  @note Collection変数をNew で作成した場合に、当該functionの利用を想定
'''  */
'
If Me.Exists(clts, key) Then
        Call clts.remove(key)
        Call clts.Add(Value, key)
Else
        Call clts.Add(Value, key)
End If
End Sub

Public Function ExistsNameOfObject(ByRef clts As Variant, ByVal Name As String, Optional ByVal ignorecase As Variant = False) As Boolean
'''' /********************************************************
''''  *@function ExistsName
''''  *<br/> Objectのコレクションに、指定の名前が存在するかを判定する。
''''  *  @param {Variant} clts ByRef
''''  *  @param {String} Name
''''  *  @return {Boolean}
''''  *
''''  */
''
Dim C_String As C_String
Set C_String = New C_String
Dim itm As Object
Dim blnRtn As Boolean
blnRtn = False
For Each itm In clts
    If C_String.EqualIgnore(itm.Name, Name, ignorecase) Then
            blnRtn = True
            Exit For
        End If
Next itm
ExistsNameOfObject = blnRtn
End Function

Public Function FilterNameMatchOfObject(ByRef clts As Variant, ByVal str As String) As Collection
'''' /********************************************************
''''  *@function FilterNameMatch
''''  *<pre> コレクションの名前を、Likeで絞り込んで、絞り込んだコレクションを返却
''''  *</pre>
''''  *@param {Collection} clts ByRef
''''  *@param {String} str
''''  *@return {Variant}  FilterNameMatch
''''  *
''''  */
''
Dim cltRtns As Collection
Set cltRtns = New Collection
Dim cl As Variant
For Each cl In clts
        If LCase(cl.Name) Like "*" & LCase(str) & "*" Then
            cltRtns.Add cl
        End If
Next cl
''
Set FilterNameMatchOfObject = cltRtns
Set cltRtns = Nothing
''
End Function

Public Function FilterNameMatchPrefixOfObject(ByRef clts As Variant, ByVal prefix As String) As Collection
'''' /********************************************************
''''  *@function FilterNameMatchPrefix
''''  *<br/> コレクションの名前を、前方一致で絞り込んで、絞り込んだコレクションを返却
''''  *  @param {Collection} clts ByRef
''''  *  @param {String} name
''''  *  @return {Variant}  FilterNameMatchPrefix
''''  *
''''  */
''
Dim cltRtns As Collection
Set cltRtns = New Collection
Dim cl As Object
For Each cl In clts
        If cl.Name Like prefix & "*" Then
            cltRtns.Add cl
        End If
Next cl
Set FilterNameMatchPrefixOfObject = cltRtns
Set cltRtns = Nothing
''
End Function

Public Function ContainsNameOfObject(ByRef clts As Variant, ByVal str As String) As Boolean
'''' /********************************************************
''''  *@function ExistsName
''''  *<br/> コレクションに、指定の文字列がLikeで存在するかを判定する。
''''  *  @param {Variant} clts ByRef
''''  *  @param {String} str
''''  *  @return {Boolean}
''''  *
''''  */
''
Dim itm As Object
Dim blnRtn As Boolean
blnRtn = False
''
Dim cltsTemp As Collection
''
Set cltsTemp = Me.FilterNameMatchOfObject(clts, str)
''
If cltsTemp.Count > 0 Then
        blnRtn = True
Else
        blnRtn = False
End If
''
ContainsNameOfObject = blnRtn
Set cltsTemp = Nothing
''
End Function

Public Function ContainsNameMatchPrefixOfObject(ByRef clts As Variant, ByVal str As String) As Boolean
'''' /********************************************************
''''  *@function ContainsNameMatchPrefix
''''  *<br/> コレクションに、指定の文字列がLikeで存在するかを判定する。
''''  *  @param {Variant} clts ByRef
''''  *  @param {String} str
''''  *  @return {Boolean}
''''  *
''''  */
''
Dim itm As Object
Dim blnRtn As Boolean
blnRtn = False
''
Dim cltsTemp As Collection
''
Set cltsTemp = Me.FilterNameMatchPrefixOfObject(clts, str)
''
If cltsTemp.Count > 0 Then
        blnRtn = True
Else
        blnRtn = False
End If
''
ContainsNameMatchPrefixOfObject = blnRtn
Set cltsTemp = Nothing
''
End Function

Public Function ToArray(ByRef clts As Collection) As Variant
'''' /********************************************************
''''  *@function ToArray
''''  *<br/> コレクションのアイテムを配列にして返却（keyは無視）
''''  *  @param {Collection} clts ByRef
''''  *  @return {Variant} ToArray
''''  */
''
Dim aryRtn() As Variant
Dim cnt As Long
Dim i As Long
''
cnt = clts.Count
If cnt < 1 Then
        ToArray = Array()
        Exit Function
End If
ReDim aryRtn(1 To cnt)
''
For i = 1 To cnt
        If IsObject(clts.Item(i)) Then
        Set aryRtn(i) = clts.Item(i)
        Else
        aryRtn(i) = clts.Item(i)
        End If
Next i
''
ToArray = aryRtn
''
End Function
    
Public Function ToArrayList(ByRef clts As Collection) As Object
'''' /********************************************************
''''  *@function ToArray
''''  *<br/> コレクションのアイテムを配列にして返却（keyは無視）
''''  *  @param {Collection} clts ByRef
''''  *  @return {Variant} ToArray
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim aryList As Object
Dim ary() As Variant
Let ary = Me.ToArray(clts)
''
Set aryList = C_Array.ToArrayList(ary)
''
Set ToArrayList = aryList
''
End Function

Public Function ToArrayNameOfObject(ByRef clts As Variant) As Variant
'''' /********************************************************
''''  *@function ToArrayName
''''  *<pre> コレクションの名前を配列にして返却。
''''  *※コレクション内のObjectに、Nameプロパティがあるものを前提とする。
''''  *</pre>
''''  *  @param {Collection} clts ByRef
''''  *  @return {Variant}  ToArrayName
''''  *
''''  */
''
Dim aryRtn As Variant
Dim cnt As Long
Dim i As Long
Dim cl As Variant
''
cnt = clts.Count
If cnt > 0 Then
        ReDim aryRtn(cnt - 1)
        i = 0
        For Each cl In clts
            aryRtn(i) = cl.Name
            i = i + 1
        Next cl
    ToArrayNameOfObject = aryRtn
Else
    ToArrayNameOfObject = ""
End If
''
End Function

Public Function ToDictionary(ByVal clts As Variant, ByRef aryHeader As Variant) As Dictionary
'''' /********************************************************
''''  *@function ToDictionary
''''  *<pre> コレクションを、指定のヘッダ配列をもとに連想配列にして返却。
''''  *</pre>
''''  *  @param {Variant} ary ByRef
''''  *  @return {Dictionary}  ToDictionary
''''  *  @note Collection変数をNew で作成した場合に、当該functionの利用を想定
''''  */
''
If Not IsArray(aryHeader) Then Err.Raise 13
If IsObject(clts) Then
        If Not TypeName(clts) = "Collection" Then
            Err.Raise 13
        End If
End If
Dim dict As Dictionary
Set dict = New Dictionary
Dim i As Long
For i = LBound(aryHeader) To UBound(aryHeader)
        If Me.Exists(clts, aryHeader(i)) Then
            Call dict.Add(aryHeader(i), clts(aryHeader(i)))
        Else
            Call dict.Add(aryHeader(i), "")
        End If
Next i
Set ToDictionary = dict
'Set dict = Nothing
End Function

Public Function ArrayToCollection(ByRef ary As Variant) As Collection
'''' /********************************************************
''''  *@function ArrayToCollection
''''  *<pre> 配列をコレクションにして返却。
''''  *</pre>
''''  *  @param {Variant} ary ByRef
''''  *  @return {Collection}
''''  *
''''  */
''
If Not IsArray(ary) Then Err.Raise 13
Set ArrayToCollection = New Collection
Dim vnt As Variant
For Each vnt In ary
        ArrayToCollection.Add vnt
Next vnt
''
End Function

Public Function ArrayListToCollection(ByRef aryList As Object) As Collection
'''' /********************************************************
''''  *@function ArrayListToCollection
''''  *<pre> ArrayList配列をコレクションにして返却。
''''  *</pre>
''''  *  @param {Object} arylist ByRef
''''  *  @return {Collection}
''''  *
''''  */
''
If Not IsObject(aryList) Then Err.Raise 13
Set ArrayListToCollection = New Collection
Dim vnt As Variant
Dim i As Integer
For i = 0 To aryList.Count - 1
    ArrayListToCollection.Add aryList.Item(i)
Next i
''
End Function

Public Sub Push(ByRef clts As Collection, ByVal Value As Variant)
'''' /********************************************************
''''  *@function Push
''''  *<pre> 値をpush
''''  *</pre>
''''  *  @param {Collection} clts ByRef
''''  *  @param {Variant} value ByRef
''''  */
''
clts.Add Value
''
End Sub

Public Function Pop(ByVal clts As Collection) As Variant
'''' /********************************************************
''''  *@function Pop
''''  *<pre> 値をPop
''''  *</pre>
''''  *  @param {Collection} clts ByRef
''''  *  @return {Variant} Pop
''''  */
''
Dim i As Long
i = clts.Count
If IsObject(clts.Item(i)) Then
        Set Pop = clts.Item(i)
Else
        Let Pop = clts.Item(i)
End If
clts.remove i
''
End Function

Public Sub Unshift(ByVal clts As Collection, ByVal Value As Variant, Optional ByVal pos As Long = 1)
'''' /********************************************************
''''  *@function Shift
''''  *<pre> 値をShift
''''  *</pre>
''''  *  @param {Collection} clts ByRef
''''  *  @param {Variant} value ByRef
''''  */
''
If clts.Count < 1 Then
        clts.Add Value
Else
    clts.Add Value, , pos
End If
''
End Sub

Public Function shift(ByVal clts As Collection) As Variant
'''' /********************************************************
''''  *@function Unshift
''''  *<pre> 値をUnshift
''''  *</pre>
''''  *  @param {Collection} clts ByRef
''''  *  @return {Variant} Pop
''''  */
''
Dim i As Long
i = 1
If IsObject(clts.Item(i)) Then
    Set shift = clts.Item(i)
Else
    Let shift = clts.Item(i)
End If
clts.remove i
''
End Function

Public Function Rearrange(ByRef clt As Collection, ByVal aryHeader As Variant, Optional ByRef dict As Variant) As Collection
'''' /********************************************************
''''  *@function Rearrange
''''  *</br> コレクションを並び変える
''''  *  @param {Collection} clt ByRef
''''  *  @param {String()} aryHeader
''''  *  @param {Dictionary} dict Optional
''''  *  @return {Collection}
''''  */
''
Dim cltOut As Collection
Set cltOut = New Collection
Dim c As Long
If VBA.IsMissing(dict) Then
    Set dict = New Dictionary
    For c = LBound(aryHeader) To UBound(aryHeader)
        dict.Add aryHeader(c), aryHeader(c)
    Next c
End If
''
For c = LBound(aryHeader) To UBound(aryHeader)
    cltOut.Add clt.Item(aryHeader(c)), dict.Item(aryHeader(c))
Next c
''
Set Rearrange = cltOut
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
'//--------------------------------------
'//variables
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "C_Collection.unittest"
''
On Error GoTo EXCEPTION
    'Call debugC_Collection.unittest1
    'Call debugC_Collection.unittest2
    'Call debugC_Collection.unittest3
    ''
    GoSub FINALLY
    Exit Sub
    ''
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
