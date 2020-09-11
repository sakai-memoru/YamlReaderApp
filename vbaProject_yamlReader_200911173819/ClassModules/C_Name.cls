VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Name.cls
'''  *<pre>Name共通クラス
'''  *Name Objectのデータを、操作するユーティリティ
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  */
'
''' /********************************************************
'''  * @class C_Name
'''  * @classdesc Name Object共通クラス
'''  * @require C_String, C_ListObject
'''  */
'
Public Function GetNames(Optional ByRef wb As Variant) As Object
''' /********************************************************
'''  *@function getNames
'''  *<br/>対象ブックから、Nameの値を取得する。
'''  *<br/>　 対象ブックにあるNamesコレクションから取得
'''  * @param {Object<Workbook>} wb ワークブックオブジェクト
'''  * @return {Object<Dictionary>} getNames NamesコレクションのDictionary形式
'''  *
'''  * @require C_String
'''  */
'
    '------------------------------
    ' 変数定義
    '------------------------------
    Const CONS_MODULE_NAME As String = "C_Name.getNames"
    Dim nms As Variant
    Dim i As Long
    Dim C_String As New C_String
    
    ' 返却用
    Dim dict As Object
    Set dict = createObject("Scripting.Dictionary")
    
On Error GoTo EXCEPTION
    '------------------------------
    ' 初期処理
    '------------------------------
    ' 引数判定
    If IsMissing(wb) Then
        Set wb = ThisWorkbook
    End If
    ' ブックのNameコレクションを設定
    Set nms = wb.Names
    
    '------------------------------
    ' 主処理
    '------------------------------
    For i = 1 To nms.Count
        If C_String.ContainWord(nms(i).RefersTo, "#REF") Then
            '
            'Console.log CONS_MODULE_NAME & ":" & nms(i).name & " " & nms(i).RefersTo
        Else
            dict.Add nms(i).Name, wb.Names(nms(i).Name).RefersToRange.Value ''FIXME
            'Console.log "key = " & rngNames(i).Name & " , value = " & Range(rngNames(i).RefersTo).value
        End If
    Next
        
    '返却
    Set GetNames = dict
    
    GoSub FINALLY
    Exit Function
    
FINALLY:
    ' Object開放
    Set nms = Nothing
    Set dict = Nothing
    Return
    
EXCEPTION:
    GoSub FINALLY
    Dim desc As String
    desc = CONS_MODULE_NAME & vbCr & Err.Description
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Description = " & desc)
    Call Err.Raise(Err.Number, Err.source & "<-" & CONS_MODULE_NAME, desc)
    Exit Function
End Function

Public Function GetNameStatement( _
            ByVal shtName As String, _
            ByVal nmStr As String, _
            Optional ByRef wb As Variant _
) As String
'''' ********************************************************
''
Dim CONS_IN_SHEET As String
Let CONS_IN_SHEET = "WORKSHEET"
Dim CONS_IN_BOOK As String
Let CONS_IN_BOOK = "WORKBOOK"
Dim CONS_NO_EXISTS As String
Let CONS_NO_EXISTS = "NOTEXISTS"
Dim nameScope As String
If VBA.IsMissing(wb) Then
    Let nameScope = Me.GetNamedRangeScope(shtName, nmStr)
Else
    Let nameScope = Me.GetNamedRangeScope(shtName, nmStr, wb)
End If
''
Dim nmStatement
If nameScope = CONS_IN_BOOK Then
    Let nmStatement = nmStr
ElseIf nameScope = CONS_IN_SHEET Then
    Let nmStatement = Me.GetNameLocal(shtName, nmStr)
Else
    Console.info "Named range with " & nmStr & " is " & nameScope
    Let nmStatement = ""
End If
Let GetNameStatement = nmStatement
''
End Function

Public Function GetNamedRangeScope( _
            ByVal shtName As String, _
            ByVal paramNm As String, _
            Optional ByRef wb As Variant _
) As String
'''' ********************************************************
''
Dim strRtn As String
''
Dim CONS_IN_SHEET As String
Let CONS_IN_SHEET = "WORKSHEET"
Dim CONS_IN_BOOK As String
Let CONS_IN_BOOK = "WORKBOOK"
Dim CONS_NO_EXISTS As String
Let CONS_NO_EXISTS = "NOTEXISTS"
Dim C_Collection As C_Collection
Set C_Collection = New C_Collection
''
If VBA.IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
Dim sht As Worksheet
Set sht = wb.Worksheets(shtName)
''
Dim nmExp As String
Let nmExp = Me.GetNameLocal(shtName, paramNm)
''
Dim cltWbNames As Names
Set cltWbNames = wb.Names
Dim cltShtNames As Names
Set cltShtNames = sht.Names
''
If C_Collection.ExistsNameOfObject(cltShtNames, nmExp, True) Then
    Let strRtn = CONS_IN_SHEET
ElseIf C_Collection.ExistsNameOfObject(cltWbNames, paramNm, True) Then
    Let strRtn = CONS_IN_BOOK
Else
    Let strRtn = CONS_NO_EXISTS
End If
''
Let GetNamedRangeScope = strRtn
''
End Function

Public Function GetNameLocal( _
            ByVal shtName As String, _
            ByVal paramNm As String _
) As String
'''' ********************************************************
''
Dim C_String As C_String
Set C_String = New C_String
''
If C_String.Contain(paramNm, "!") Then
    Let GetNameLocal = paramNm
Else
    Let GetNameLocal = shtName & "!" & paramNm
End If
End Function


Public Function GetNamesWithSheet(ByVal strSheetName As String, _
            Optional ByVal excludingPrintArea As Variant = True, _
            Optional ByRef wb As Variant, _
            Optional ByVal excludingPrefix As Variant = True) As Object
''' /********************************************************
'''  *@function getNamesWithSheet
'''  *<br/>
'''  *<br/>
'''  * @param {String} strSheetName
'''  * @param {Object<Workbook>} wb Optional
'''  * @param {Boolean} exculedingPrintArea
'''  * @return {Object<Dictionary>}
'''  *
'''  * @require C_String
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Name.getNamesWithSheet"
    Dim nms As Variant
    Dim i As Long
    Dim C_String As New C_String
    Dim dict As Object
    Set dict = createObject("Scripting.Dictionary")
    '
On Error GoTo EXCEPTION
    If IsMissing(wb) Then
        Set wb = ThisWorkbook
    End If
    Dim sht As Worksheet
    Set sht = wb.Worksheets(strSheetName)
    Set nms = wb.Names
    Dim strName As String
    For i = 1 To nms.Count
        If C_String.ContainWord(nms(i).RefersTo, "#REF") Then
            '
            'Console.log CONS_MODULE_NAME & ":" & nms(i).name & " " & nms(i).RefersTo
        Else
            If C_String.StartsWith(nms(i).RefersTo, "=" & strSheetName) Then
                If excludingPrintArea Then
                    'Console.log "key = " & nms(i).Name
                    Let strName = nms(i).Name
                    If Not CheckPrintArea(strName) Then  ''FIXME
                        If excludingPrefix Then
                            If Not C_String.StartsWith(strName, "_") Then
                        dict.Add nms(i).Name, sht.Range(nms(i).RefersTo).Value
                    End If
                Else
                    dict.Add nms(i).Name, sht.Range(nms(i).RefersTo).Value
                End If
            End If
                Else
                    dict.Add nms(i).Name, sht.Range(nms(i).RefersTo).Value
                End If
            End If
        End If
    Next
    '
    Set GetNamesWithSheet = dict
    '
    GoSub FINALLY
    Exit Function
    '
FINALLY:
    ' release and return
    Set nms = Nothing
    Set dict = Nothing
    Set sht = Nothing
    Set C_String = Nothing
    Return
    '
EXCEPTION:
    GoSub FINALLY
    Dim desc As String
    desc = CONS_MODULE_NAME & vbCr & Err.Description
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Description = " & desc)
    Call Err.Raise(Err.Number, Err.source & "<-" & CONS_MODULE_NAME, desc)
    Exit Function
    '
End Function
Private Function CheckPrintArea(ByVal strName As String) As Boolean
''' ********************************************************
'
' const
Dim CONS_PRINT_AREA_NAME As String
CONS_PRINT_AREA_NAME = "Print_Area"
Dim CONS_PRINT_TITLE_NAME As String
CONS_PRINT_TITLE_NAME = "Print_Titles"
' variables
Dim C_String As C_String
Set C_String = New C_String
' release and return
CheckPrintArea = C_String.ContainWord(strName, CONS_PRINT_AREA_NAME) Or C_String.ContainWord(strName, CONS_PRINT_TITLE_NAME)
'Console.log "checkPrintArea : " & checkPrintArea
Set C_String = Nothing
'
End Function

Public Function GetValue(strName As Variant, _
        Optional strSheetName As Variant, _
        Optional wb As Variant _
        ) As Variant
''' /********************************************************
'''  *@function getValue
'''  *<br/>　対象の範囲名から値を取得する。
'''  *<br/>
'''  * @param {String} strName　対象範囲名
'''  * @param {String} Optional strSheetName　対象シート名(ブックで範囲名が一意でない場合に指定）
'''  * @param {Workbook} Optional wb　対象workbook
'''  * @return {Variant} getValue 範囲を表す値
'''  * 　セルが１つの場合は、Value。セルが２つ以上の場合は、Range。
'''  */
Dim C_String As C_String
Set C_String = New C_String
Dim vntRtn As Variant
Dim strRef As Variant
Dim rng As Range
Dim strRngName As String
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
If IsMissing(strSheetName) Then
        strRngName = strName
Else
    If C_String.IsAddress(strName) Then
        strRngName = strSheetName & "!" & strName
    Else
        strRngName = strName
    End If
End If
''
strRef = wb.Names(strRngName).RefersTo
''
If C_String.Contain(strRef, ":") Then
        Set vntRtn = ThisWorkbook.Names(strRngName).RefersToRange '''Cellが2つ以上の範囲の場合は、Rangeを返却
Else
        Set rng = ThisWorkbook.Names(strRngName).RefersToRange
        vntRtn = rng.Value
End If
GetValue = vntRtn
''
End Function

Public Function GetText(strName As Variant, _
        Optional strSheetName As Variant, _
        Optional wb As Variant _
        ) As Variant
''' /********************************************************
'''  *@function getValue
'''  *<br/>　対象の範囲名からテキスト値を取得する。
'''  *<br/>
'''  * @param {String} strName　対象範囲名
'''  * @param {String} strSheetName　対象シート名
'''  * @return {Variant} getText 範囲を表す値
'''  * 　セルが１つの場合は、Text。セルが２つ以上の場合は、Range。
'''  */
Dim C_String As C_String
Set C_String = New C_String
Dim vntRtn As Variant
Dim strRef As Variant
Dim rng As Range
Dim strRngName As String
''
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
If IsMissing(strSheetName) Then
        strRngName = strName
Else
    If C_String.IsAddress(strName) Then
        strRngName = strSheetName & "!" & strName
    Else
        strRngName = strName
    End If
End If
''
strRef = wb.Names(strRngName).RefersTo
''
If C_String.Contain(strRef, ":") Then
        Set vntRtn = ThisWorkbook.Names(strRngName).RefersToRange '''Cellが2つ以上の範囲の場合は、Rangeを返却
Else
        Set rng = ThisWorkbook.Names(strRngName).RefersToRange
        vntRtn = rng.text
End If
GetText = vntRtn
''
End Function


Public Function GetAddress(strName As Variant, _
        Optional strSheetName As Variant, _
        Optional wb As Variant _
        ) As Variant
''' /********************************************************
'''  *@function getValue
'''  *<br/>　対象の範囲名から参照アドレスを取得する。
'''  *<br/>
'''  * @param {String} strName　対象範囲名
'''  * @param {String} strSheetName　対象シート名
'''  * @return {Variant} getAddress 範囲を表すアドレス
'''  *
'''  */
Dim C_String As C_String
Set C_String = New C_String
Dim strRef As Variant
Dim strRngName As String
''
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
If IsMissing(strSheetName) Then
        strRngName = strName
Else
    If C_String.IsAddress(strName) Then
        strRngName = strSheetName & "!" & strName
    Else
        strRngName = strName
    End If
End If
''
strRef = wb.Names(strRngName).RefersTo
''
GetAddress = strRef
''
End Function

Public Function GetRange(strName As Variant, _
        Optional strSheetName As Variant, _
        Optional wb As Variant _
        ) As Variant
''' /********************************************************
'''  *@function getRange
'''  *<br/>　対象の範囲名から参照範囲オブジェクト（Range）を取得する。
'''  *<br/>
'''  * @param {String} strName　対象範囲名 対象範囲がAddressの場合はシート名が必要
'''  * @param {String} strSheetName　対象シート名
'''  * @return {Variant} getRange　Range Object
'''  */
Dim C_String As C_String
Set C_String = New C_String
Dim strRngName As String
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
If IsMissing(strSheetName) Then
        strRngName = strName
Else
    If C_String.IsAddress(strName) Then
        strRngName = strSheetName & "!" & strName
    Else
        strRngName = strName
    End If
End If
''
'Console.log "strRngName = " & strRngName
'Console.log "rngAddress = " & GetAddress(strRngName)
Set GetRange = wb.Names(strRngName).RefersToRange
End Function

Public Function GetComment(strName As Variant, _
        Optional strSheetName As Variant, _
        Optional wb As Variant _
        ) As Variant
''' /********************************************************
'''  *@function getComment
'''  *<br/>　対象の範囲名からコメントを取得する。
'''  *<br/>
'''  * @param {String} strName　対象範囲名
'''  * @param {String} strSheetName　対象シート名
'''  * @return {Variant} getComment　Commentに記述されている文字列
'''  */
Dim C_String As C_String
Set C_String = New C_String
Dim strRngName As String
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
If IsMissing(strSheetName) Then
        strRngName = strName
Else
    If C_String.IsAddress(strName) Then
        strRngName = strSheetName & "!" & strName
    Else
        strRngName = strName
    End If
End If
GetComment = wb.Names(strRngName).comment
End Function

Public Function GetCategory(strName As Variant, _
        Optional strSheetName As Variant, _
        Optional wb As Variant _
        ) As Variant
''' /********************************************************
'''  *@function getCategory
'''  *<br/>　対象の範囲名からカテゴリを取得する。
'''  *<br/>
'''  * @param {String} strName　対象範囲名
'''  * @param {String} strSheetName　対象シート名
'''  * @return {Variant} getCategory　範囲名のカテゴリ文字列（ユーザ定義関数やマクロシートのコマンドマクロ）
'''  */
Dim C_String As C_String
Set C_String = New C_String
Dim strRngName As String
If IsMissing(wb) Then
    Set wb = ThisWorkbook
End If
If IsMissing(strSheetName) Then
        strRngName = strName
Else
    If C_String.IsAddress(strName) Then
        strRngName = strSheetName & "!" & strName
    Else
        strRngName = strName
    End If
End If
If wb.Names(strRngName).MacroType <> xlNone Then
    GetCategory = wb.Names(strRngName).Category
Else
        GetCategory = ""
End If
End Function

Public Function SetNameByKeyword(ByVal strSheetName As Variant _
            , ByVal strKeyString As Variant _
            , strName As Variant _
            , Optional Row As Variant = 0, Optional column As Variant = 0) As Variant
''' /********************************************************
'''  * @function setNameByKeyword
'''  *<br/>　検索キーワードでヒットしたセルから
'''  *<br/>　Offsetで指定したセルにNameを設定する｡
'''  * @param {String} strSheetName　検索対象シート名
'''  * @param {String} strKeyword　検索対象文字
'''  * @param {Variant} row　   【Optional】Offset行位置　【Default】0
'''  * @param {Variant} column　【Optional】Offset列位置　【Default】0
'''  * @return {Variant} setNameByKeyword　NameセットができたらTrue
'''  *<br/>　検索がヒットする前提（最初にヒットしたセルを対象とする）
'''  */
Dim rngFindResult As Range
Dim rngUsed As Range
Dim nm As Name
Set rngUsed = ThisWorkbook.Worksheets(strSheetName).UsedRange
Set rngFindResult = rngUsed.find(what:=strKeyString, Lookat:=xlWhole)
    
If IsEmpty(Row) Or Row = "" Then
        Row = 0
End If
If IsEmpty(column) Or column = "" Then
        column = 0
End If
    
If rngFindResult Is Nothing Then
        SetNameByKeyword = False
Else
        For Each nm In ThisWorkbook.Names
            If UCase(nm.Name) = UCase(strName) Then
                rngFindResult.Offset(Row, column).Name.Delete
            End If
        Next
        rngFindResult.Offset(Row, column).Name = strSheetName & "!" & strName
        SetNameByKeyword = True
End If
    
End Function

Public Sub ChangeNamedRangeAddress(bk As Workbook, rangename As String, newRange As Range)
''' ********************************************************
''
bk.Names(rangename).RefersTo = newRange
''
End Sub

Sub ChangeLocalNameAndOrScope()
'Programmatically change a sheet-level range name and/or scope to a new name and/or scope
Dim nm As Name, Ans As Integer, newNm As String
For Each nm In ActiveWorkbook.Names
    If nm.Name Like "*!*" Then 'It is sheet level
        newNm = Replace(nm.Name, "*!", "")
        Range(nm.RefersTo).Name = newNm
        nm.Delete
    End If
Next nm
End Sub


'Public Sub GlobalNamesToLocalNames(Optional Void As Variant)
'   Dim list() As GlobalNamesToLocalNames_Type
'   Dim count As Long
'   Dim Name As Name
'   Dim Dat() As String
'   Dim X As Long
'
'   ' count the size
'   For Each Name In ActiveWorkbook.Names
'      count = count + 1
'   Next
'   ReDim list(count - 1)
'   count = 0
'
'   ' Collecect all name data
'   For Each Name In ActiveWorkbook.Names
'      With list(count)
'         ' Pick up only the name
'         If InStr(Name.Name, "!") > 0 Then
'            Dat = Split(Name.Name, "!")
'            .Name = Dat(1)
'         Else
'            .Name = Name.Name
'         End If
'         ' pick up the sheet and refer
'         Dat = Split(Name.RefersTo, "!")
'         .Sheet = Mid(Dat(0), 2)
'         .Ref = Dat(1)
'         ' make local sheet name
'         .Name = .Sheet & "!" & .Name
'      End With
'      count = count + 1
'   Next
'
'   ' Delete all names
'   For Each Name In ActiveWorkbook.Names
'      Name.Delete
'   Next
'
'   'rebuild all the names
'   For X = 0 To count - 1
'      With list(X)
'         If Left(.Ref, 1) <> "#" Then
'            ActiveWorkbook.Names.Add Name:=.Name, RefersToLocal:="=" & .Sheet & "!" & .Ref
'         End If
'      End With
'   Next
'End Sub

Public Sub ChangeGlobalNamedRangeAddress(bk As Workbook, rangename As String, newRange As Range)
''
'this sub only changes named range scoped to the workbook
'and ignores any ranges scoped to the worksheet with the same name.
Dim n As Name
    
For Each n In bk.Names
        If InStr(1, n.Name, rangename) > 0 And InStr(1, n.NameLocal, "!") = 0 Then
            n.RefersTo = newRange
            Exit For
        End If
Next
''
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RescopeNamedRangesToWorksheet
' Author    : Jesse Stratton
' Date      : 11/18/2013
' Purpose   : Rescopes each workbook scoped named range to the specific worksheet to
' which the range refers for each named range that refers to the active worksheet.
'---------------------------------------------------------------------------------------
Public Sub RescopeNamedRangesToWorksheet()
Dim wb As Workbook
Dim ws As Worksheet
Dim objName As Name
Dim sWsName As String
Dim sWbName As String
Dim sRefersTo As String
Dim sObjName As String

Set wb = ActiveWorkbook
Set ws = ActiveSheet
sWsName = ws.Name
sWbName = wb.Name

'Loop through names in worksheet.
For Each objName In wb.Names
    'Check name is visble.
    If objName.Visible = True Then
        'Check name refers to a range on the active sheet.
        If InStr(1, objName.RefersTo, sWsName, vbTextCompare) Then
            sRefersTo = objName.RefersTo
            sObjName = objName.Name
            'Check name is scoped to the workbook.
            If objName.Parent.Name = sWbName Then
            'Delete the current name scoped to workbook replacing with worksheet
            'scoped name.
                objName.Delete
                ws.Names.Add Name:=sObjName, RefersTo:=sRefersTo
            End If
        End If
    End If
Next objName
''
End Sub

'///////////////////////////////////////////////////////////
'/////  Debug          /////////////////////////////////////
'///////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////
'-----------------------------------------------------------
' debug entry
'-----------------------------------------------------------
Public Sub UnitTest()
' Collectionの確認
'    Console.dump ThisWorkbook.Names
'    Dim cltNames As Names
'    Set cltNames = ThisWorkbook.Names
'
'    Dim cltLists As ListObjects
'    Set cltLists = ThisWorkbook.Worksheets("__CONFIG__").ListObjects
    
    
' getValue
'    Console.dump Me.getValue("seqId", "__FORMAT__")
'    Console.dump Me.getValue("TEST_STRING")
'    Console.dump Me.getValue("TEST_RANGE")  '''FIXME□□Rangeの場合に何を返却するか？
    
' getAddress
'    Console.dump Me.getAddress("seqId", "__FORMAT__")
'    Console.dump Me.getAddress("TEST_STRING")
'    Console.dump Me.getAddress("TEST_RANGE")

' getRange
'    Console.dump Me.getRange("TEST_STRING")
'    Console.dump Me.getRange("TEST_RANGE")
'    Console.dump Me.getRange("seqId", "__FORMAT__")
'    Console.dump Me.getRange("SeqId", "__CONFIG__")
'    Console.dump Me.getRange("SeqId")
   
' getComment
'    Console.dump Me.getComment("seqId", "__FORMAT__")
'    Console.dump Me.getComment("TEST_STRING")
'    Console.dump Me.getComment("TEST_RANGE")
'    Console.info "---------------------"

' getCategory
    Console.Dump Me.GetCategory("seqId", "__FORMAT__")
    Console.Dump Me.GetCategory("TEST_STRING")
    Console.Dump Me.GetCategory("TEST_RANGE1")

   
   
'    Console.dump Me.getValue("テーブル1") '''ListObjectsの名前は、Namesには無い
    
    
    
'    Dim vnt As Variant
'    vnt = setName("画面項目定義", "No.", "seqId")
'    Console.dump vnt
End Sub
