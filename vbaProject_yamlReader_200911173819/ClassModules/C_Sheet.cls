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
'''  *<br/>Worksheet�p���[�e�B���e�B�E�N���X
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  *
'''  */
'
''' /********************************************************
'''  * @class C_Sheet
'''  * @classdesc Worksheet�p���[�e�B���e�B�E�N���X
'''  */
'
Public Sub ChangeVisible(ByVal sheetName As String)
''' /********************************************************
'''  *@function changeVisible
'''  *</br> �V�[�g�̕\���A��\����ύX����
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
'''  *</br> �V�[�g���󂩂ǂ�����ԋp����
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
'''  *</br> �V�[�g�����݂��Ă��邩�ۂ���ԋp����
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
'''  *</br> "_nn"��ReName���ꂽ�V�[�g�����擾����B
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
'''  *</br> �w�肳�ꂽ�u�b�N�ɁA�w�肳�ꂽ�V�[�g�I�u�W�F�N�g��
'''  *</br> �����I�ɃR�s�[����
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
'''  *</br> �w�肳�ꂽ�u�b�N�ɁA�e���v���[�g�V�[�g���R�s�[����
'''  *</br> �w�肵�����O�ɕύX����B
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
'''  *</br> �Ώۃu�b�N���̃V�[�g����ёւ���
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

'���[�N�V�[�g����ArrayList�I�u�W�F�N�g�ɒǉ�
For i = 1 To cnt
    aryList.Add Workbooks(bookName).Worksheets(i).Name
Next i

'�z��̗v�f���\�[�g����
aryList.Sort

'���[�N�V�[�g�̈ʒu����בւ���
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