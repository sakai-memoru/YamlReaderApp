VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Sheet.cls
'''  *<br/>Worksheet用ユーティリティ・クラス
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  *
'''  */
'
''' /********************************************************
'''  * @class C_Sheet
'''  * @classdesc Worksheet用ユーティリティ・クラス
'''  */
'
Public Sub ChangeVisible(ByVal sheetName As String)
''' /********************************************************
'''  *@function changeVisible
'''  *</br> シートの表示、非表示を変更する
'''  * @param {String} sheetname
'''  */
    If ThisWorkbook.Worksheets(sheetName).Visible Then
        Call ToInvisible(sheetName)
    Else
        Call ToVisible(sheetName)
    End If
End Sub

Private Sub ToVisible(ByVal sheetName As String)
    ThisWorkbook.Worksheets(sheetName).Visible = True
End Sub

Private Sub ToInvisible(ByVal sheetName As String)
    ThisWorkbook.Worksheets(sheetName).Visible = False
End Sub

Public Function IsSheetEmpty(ByVal bookName As String, ByVal sheetName As String) As Boolean
''' /********************************************************
'''  *@function isSheetEmpty
'''  *</br> シートが空かどうかを返却する
'''  * @param {String} bookname
'''  * @param {String} sheetname
'''  * @return {Boolean} isSheetEmpty
'''  */
        '
    If Application.WorksheetFunction.CountA(Workbooks(bookName).Worksheets(sheetName).UsedRange) = 0 Then
        IsSheetEmpty = True
    Else
        IsSheetEmpty = False
    End If
End Function

Public Function ExistsSheet(ByVal bookName As String, ByVal sheetName As String) As Boolean
''' /********************************************************
'''  *@function isSheetEmpty
'''  *</br> シートが存在しているか否かを返却する
'''  * @param {String} bookname
'''  * @param {String} sheetname
'''  * @return {Boolean} isSheetEmpty
'''  */
''
Dim sht As Object
ExistsSheet = False
For Each sht In Workbooks(bookName).Worksheets
        If sht.Name = sheetName Then
            ExistsSheet = True
            Exit For
        End If
Next sht
''
End Function

Public Function GetRenameSheetName(ByVal bookName As String, ByVal sheetName As String) As String
''' /********************************************************
'''  *@function getRenameSheetName
'''  *</br> "_nn"でReNameされたシート名を取得する。
'''  * @param {String} bookname
'''  * @param {String} sheetname
'''  * @return {Boolean} isSheetEmpty
'''  *
'''  */
'
Dim sht As Object
Dim str As String
Dim i As Long
i = 0
Dim bln As Boolean
bln = True
str = sheetName
''
Do While bln
        If Me.ExistsSheet(bookName, str) Then
            i = i + 1
            str = sheetName & "_" & format(i, "00")
        Else
            bln = False
        End If
Loop
''
GetRenameSheetName = str
''
End Function

Public Function CopyForce(ByVal bookName As String, ByRef sht As Worksheet)
''' /********************************************************
'''  *@function copyForce
'''  *</br> 指定されたブックに、指定されたシートオブジェクトを
'''  *</br> 強制的にコピーする
'''  * @param {String} bookname
'''  * @param {Workbook} sht
'''  *
'''  */
'
Dim i As Long
i = 0
Dim bk As Workbook
Set bk = Workbooks(bookName)
Application.DisplayAlerts = False
''
If Me.ExistsSheet(bookName, sht.Name) Then
        bk.Worksheets(sht.Name).Delete
        sht.Copy After:=bk.Worksheets(bk.Worksheets.Count)
Else
        sht.Copy After:=bk.Worksheets(bk.Worksheets.Count)
End If
''
Application.DisplayAlerts = True
''
End Function

Public Function AddSheetTemplate(ByVal bookName As String, ByRef sht As Worksheet, ByVal tempsheetname As String)
''' /********************************************************
'''  *@function AddSheetTemplate
'''  *</br> 指定されたブックに、テンプレートシートをコピーして
'''  *</br> 指定した名前に変更する。
'''  * @param {String} bookname
'''  * @param {Workbook} sht
'''  * @param {String} tempsheetname
'''  */
''
Dim i As Long
i = 0
Dim str
Dim bk As Workbook
Set bk = Workbooks(bookName)
Dim sh As Worksheet
str = Me.GetRenameSheetName(bookName, tempsheetname)
''
sht.Copy After:=bk.Worksheets(bk.Worksheets.Count)
ActiveSheet.Name = str
''
End Function

Public Sub Sort(ByVal bookName As String)
''' /********************************************************
'''  *@function Sort
'''  *</br> 対象ブック内のシートを並び替える
'''  * @param {String} bookname
'''  *
'''  */
'
Dim i As Long
    Dim j As Long
Dim aryList As Object
Set aryList = createObject("system.Collections.ArrayList")
Dim cnt As Long
cnt = Workbooks(bookName).Worksheets.Count

'ワークシート名をArrayListオブジェクトに追加
For i = 1 To cnt
    aryList.Add Workbooks(bookName).Worksheets(i).Name
Next i

'配列の要素をソートする
aryList.Sort

'ワークシートの位置を並べ替える
Worksheets(aryList(1)).Move Before:=Worksheets(1)
For i = 2 To cnt
    Worksheets(aryList(i)).Move After:=Worksheets(i - 1)
Next i
End Sub

Public Function GetBookSheetExpression(ByVal wbName As String, ByVal shtName As String) As String
'''' *****************************************************
''
GetBookSheetExpression = "'" & wbName & "'!" & shtName
''
End Function

Public Function SplitBookSheetExpression(ByVal expression As Variant) As Variant
'''' *****************************************************
''
Dim ary As Variant
ary = VBA.Split(expression, "!")
Let ary(LBound(ary)) = C_String.ChopDouble(ary(LBound(ary)), 1)
SplitBookSheetExpression = ary
''
End Function
