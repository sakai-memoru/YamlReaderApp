VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Book.cls
'''  *<pre>Workbook utility
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  *
'''  */
'
''' /********************************************************
'''  * @class C_Book
'''  * <pre> Workbook utility Class
'''  * </pre>
'''  *
'''  * @require C_File,C_Array,C_Sheet
'''  */
'

Public Function GetXlsxes(Optional ByVal strPath As Variant _
                 , Optional ByVal strExt As Variant = "*.xls?")
'''/****************************************
''' * @function getXlsxes
''' * <br/> get xlsx files in target folder
''' * <br/>
''' * @param {Variant}strPath Optional ByVal If isMissing, set ThisWorkbook.Path
''' * @param {Variant}strExt Optional ByVal wildcard Default:"*.xls?"
''' * @return {String[]} array of books' path
''' */
Dim C_File As C_File
Set C_File = New C_File
Dim ary As Variant
If IsMissing(strPath) Then
    strPath = ThisWorkbook.path
End If
GetXlsxes = C_File.GetFilesByWildcard(strPath, strExt)
''
End Function

Public Sub OpenBook(ByVal filePath As String, Optional updateLinks As Boolean = False, Optional readOnly As Boolean = True)
'''/****************************************
''' * @function openBook
''' * <br/> open a workbook file
''' * <br/>
''' * @param {Variant}filePath
''' */
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "C_Book.openBook"
Dim tempFilename As String
Dim bk As Workbook
''
On Error GoTo EXCEPTION
    ' if it exists
    tempFilename = Dir(filePath)
    If tempFilename = "" Then
        Call Err.Raise(12001, CONS_MODULE_NAME, CONS_MODULE_NAME & vbCr & "File Not Found", Err.HelpFile, Err.HelpContext)
    End If
    ' check already open the same book
    For Each bk In Workbooks
        If bk.Name = tempFilename Then
            Call Err.Raise(12001, CONS_MODULE_NAME, CONS_MODULE_NAME & vbCr & "The same file already has been opened ", Err.HelpFile, Err.HelpContext)
        End If
    Next bk
    Workbooks.Open filePath, updateLinks, readOnly
    '// release and return
    GoSub FINALLY
    Exit Sub
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end process
    Return

EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
''
End Sub

Public Function CloseBook(ByVal filePath As String, Optional saveChanges As Boolean = True)
'''/****************************************
''' * @function closeBook
''' * <br/> close a workbook
''' * <br/>
''' * @param {Variant}filePath
''' */
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "C_Book.closeBook"
Dim C_File As C_File
Set C_File = New C_File
Dim C_Collecton As C_Collection
Set C_Collection = New C_Collection
On Error GoTo EXCEPTION
    Dim clt As Collection
    Set clt = New Collection
    '' FIXME
    ''If C_Collection.ExistsNameOfObject(clt, C_File.GetFileName(filepath)) Then
    ''    Workbooks(GetBookName(filepath)).Close saveChanges:=saveChanges
    ''End If
    Workbooks(GetBookName(filePath)).Close saveChanges:=saveChanges
    '//release and return
    GoSub FINALLY
    Exit Function
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end process
    Return
    ''
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
''
End Function

Public Sub CreateBook(ByRef wb As Workbook)
'''/****************************************
''' * @function createBook
''' * <br/> create a new workbook
''' * <br/>
''' * @param {Workbook}
''' */
Set wb = Workbooks.Add
''
End Sub

Public Sub CreateBookByTemplate(ByRef wb As Workbook, ByVal templatePath As String)
'''/****************************************
''' * @function createBookByTemplate
''' * <br/> create a workbook by template
''' * <br/>
''' * @param {Variant}filePath
''' */
Set wb = Workbooks.Add(templatePath)
''
End Sub

Public Sub Save(ByRef wb As Workbook)
'''/****************************************
''' * @function createBook
''' * <br/> close a workbook
''' * <br/>
''' * @param {Variant}filePath
''' */
wb.Save
''
End Sub

Public Sub SaveAndClose(ByRef wb As Workbook)
'''/****************************************
''' * @function createBook
''' * <br/> close a workbook
''' * <br/>
''' * @param {Variant}filePath
''' */
wb.Save
wb.Close
''
End Sub

Public Sub SaveNotAndClose(ByRef wb As Workbook)
'''/****************************************
''' * @function createBook
''' * <br/> close a workbook
''' * <br/>
''' * @param {Variant}filePath
''' */
wb.Close saveChanges:=False
''
End Sub

Public Sub CloseAllWithoutThisWorkbook()
'''' ****************************************
''
Dim wb As Workbook
For Each wb In Workbooks
    If Not wb.Name = ThisWorkbook.Name Then
        Me.SaveNotAndClose wb
    End If
Next wb
''
End Sub

Public Sub SaveAllAndQuit()
'''/****************************************
''' * @function createBook
''' * <br/> close a workbook
''' * <br/>
''' * @param {Variant}filePath
''' */
Dim wb As Workbook
For Each wb In Application.Workbooks
    wb.Save
Next wb
Application.Quit
''
End Sub

Public Sub CreateAndSaveAs(ByRef wb As Workbook, ByVal savePath As String, Optional ByVal templatePath As Variant)
'''/****************************************
''' * @function createBook
''' * <br/> close a workbook
''' * <br/>
''' * @param {Variant}filePath
''' */
If IsMissing(templatePath) Then
    Set wb = Workbooks.Add
    wb.SaveAs fileName:=savePath
Else
    Set wb = Workbooks.Add(templatePath)
    wb.SaveAs fileName:=savePath
End If
''
End Sub

Public Function GetBookName(Optional ByVal bookpath As Variant) As String
'''/****************************************
''' * @function getBookname
''' * <br/>get target book's name
''' * <br/>
''' * @param {String}bookPath ByVal
''' * @return {String}
''' */
If IsMissing(bookpath) Then
    bookpath = ThisWorkbook.FullName
End If
Dim objFSO As Object
Set objFSO = createObject("Scripting.FileSystemObject")
GetBookName = objFSO.GetFileName(bookpath)
''
Set objFSO = Nothing
''
End Function

Public Function GetSheets(Optional ByVal bookpath As Variant) As Variant
'''/****************************************
''' * @function getSheets
''' * <br/>ブックをオープン後に、Workbooksコレクションの
''' * <br/>指定ブックパスからシート名の配列を取得する
''' * <br/>
''' * @param {String}bookpath　ByVal
''' */
If IsMissing(bookpath) Then
    bookpath = ThisWorkbook.FullName
End If
'
Dim wb  As Workbook
Dim sht As Worksheet
Dim ary() As Variant
Dim i As Long
i = 0
'
Set wb = Workbooks(GetBookName(bookpath))
'
For Each sht In wb.Worksheets
    If sht.Visible Then
    ReDim Preserve ary(i)
    ary(i) = sht.Name
    End If
    i = i + 1
Next sht
'
GetSheets = ary
'
Set wb = Nothing
''
End Function

Public Function GetSheetsWithSpec(ByVal keyword As String, Optional ByVal bookpath As Variant) As Variant
'''/****************************************
''' * @function getSheetsWithSpec
''' * <br/>ブックをオープン後に、Workbooksコレクションの
''' * <br/>指定ブックに含まれるシートを配列で取得する。
''' * <br/>　PREFIX"__"は除く。
''' * <br/>　どの色のシートを取り込むかは指定可能。
''' * <br/>　「色なし」は取り込み対象とする。
''' * <br/>　キーワード指定可能とする。
''' * <br/>　空のワークシートは取り込まない
''' * @param {String}bookpath ByVal
''' * @param {String}keyword ByVal
''' * @param {ParamArray}colorAry()
''' * @return {Variant} シート名の配列
''' */
Const CONS_COLOR_NONE As Long = xlColorIndexNone
Const CONS_PREFIX As String = "__"
'
Dim bk  As Workbook
Dim sht As Worksheet
Dim ary() As Variant
Dim strBookName As String
Dim i As Long
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_Sheet As C_Sheet
Set C_Sheet = New C_Sheet
'
i = 0
Dim aryTemp() As Variant
'
''対象の色を配列化（Defaultで「色無し」を追加）
ReDim Preserve aryTemp(0 To 0)
aryTemp(UBound(aryTemp)) = CONS_COLOR_NONE
''ブック名を取得し、オブジェクトを生成
strBookName = GetBookName(bookpath)
Set bk = Workbooks(strBookName)
'
For Each sht In bk.Worksheets
    If C_Sheet.IsSheetEmpty(strBookName, sht.Name) Then
        ''空のシートは取り込み対象としない
    Else
        If C_Array.Contains(aryTemp, sht.Tab.ColorIndex) Then
            If (left(sht.Name, 2) <> CONS_PREFIX) Then
                ReDim Preserve ary(i)
                ary(i) = sht.Name
                i = i + 1
            End If
        End If
    End If
Next sht
GetSheetsWithSpec = Filter(ary, keyword, True, vbTextCompare)
Set bk = Nothing
''
End Function


Public Function GetSheetsWithColor(ByVal bookpath As String, ParamArray colorAry() As Variant) As Variant
'''/****************************************
''' * @function getSheetsWithColor
''' * <br/>ブックをオープン後に、Workbooksコレクションの
''' * <br/>指定ブックに含まれるシートを配列で取得する。
''' * <br/>　どの色のシートを取り込むかは指定可能。
''' * <br/>　「色なし」は取り込み対象とする。
''' * @param {String}bookpath ByVal
''' * @param {ParamArray}colorAry()
''' * @return {Variant} シート名の配列
''' */
Const CONS_COLOR_NONE As Long = xlColorIndexNone
''
Dim bookName As String
Dim bk  As Workbook
Dim sht As Worksheet
Dim ary() As Variant
Dim i As Long
Dim C_Array As C_Array
Set C_Array = New C_Array
i = 0
Dim aryTemp As Variant
aryTemp = colorAry
ReDim Preserve aryTemp(UBound(aryTemp) + 1)
aryTemp(UBound(aryTemp)) = CONS_COLOR_NONE
''
Set bk = Workbooks(GetBookName(bookpath))
''
For Each sht In bk.Worksheets
    If C_Array.Contains(aryTemp, sht.Tab.ColorIndex) Then
        ReDim Preserve ary(i)
        ary(i) = sht.Name
        i = i + 1
    End If
Next sht
''
GetSheetsWithColor = ary
''
Set bk = Nothing
''
End Function

Public Function GetSheetsByPrefix(ByVal bookpath As String, Optional prefix As Variant = "__" _
                    , Optional StartWith As Variant = False) As Variant
'''/****************************************
''' * @function getSheetsByPrefix
''' * <br/>ブックをオープン後に、Workbooksコレクションの
''' * <br/>指定ブックに含まれるシートを配列で取得する。
''' * @param {String}bookpath ByVal
''' * @param {Variant}prefix 【Optional】 Default:"__"
''' * @param {Boolean}startWith 【Optional】 Default:False
''' * @return {Variant} シート名の配列
''' */
'//--------------------------------------
'//variables
Const CONS_MODULE_NAME As String = "C_Book.getSheetsByPrefix"
Dim ary As Variant
Dim retAry() As Variant
Dim i As Long
Dim j As Long
'
On Error GoTo EXCEPTION
    '//--------------------------------------
    '//処理
    j = 0
    ary = GetSheets(bookpath)
    '
    For i = LBound(ary) To UBound(ary)
        If StartWith Then
            If (left(ary(i), Len(prefix)) = prefix) Then
                ReDim Preserve retAry(j)
                retAry(j) = ary(i)
                j = j + 1
            End If
        Else
            If (left(ary(i), Len(prefix)) <> prefix) Then
                ReDim Preserve retAry(j)
                retAry(j) = ary(i)
                j = j + 1
            End If
        End If
    Next i
    '//release and return
    GoSub FINALLY
    GetSheetsByPrefix = retAry
    Exit Function
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
'''****************************************
End Function

Public Function CopySheets(ByVal sourceBookpath As String, ByVal destinationBookpath As String _
                    , sheetsArray As Variant)
'''/****************************************
''' * @function copySheets
''' * <br/>対象ブックをオープン後に、
''' * <br/>コピー元ブックに含まれるシートを配列を渡して
''' * <br/>コピー先ブックに対象シートをコピーする｡
''' * @param {String}bookpath ByVal
''' * @param {Variant}prefix 【Optional】 Default:"__"
''' * @param {Boolean}startWith 【Optional】 Default:False
''' * @return {Variant} シート名の配列
''' */
Dim bkSrc As Workbook
Dim bkDest As Workbook
Set bkSrc = Workbooks(GetBookName(sourceBookpath))
Set bkDest = Workbooks(GetBookName(destinationBookpath))
'
Dim i As Long
'
For i = LBound(sheetsArray) To UBound(sheetsArray)
    bkSrc.Worksheets(sheetsArray(i)).Copy After:=bkDest.Worksheets(bkDest.Worksheets.Count)
Next i
'
Set bkSrc = Nothing
Set bkDest = Nothing
'''****************************************
End Function

Public Function MergeBook(ByVal sourceBookpath As String, ByVal destinationBookpath As String)
'''/****************************************
''' * @function MergeBook
''' */
'' FIXME
End Function

Public Function DeleteSheet(ByVal bookName As String, ByVal sheetName)
'''/****************************************
''' * @function deleteSheet
''' * <br/>対象シートを、削除する。
''' * @param {String}bookname ByVal
''' * @param {Variant}sheetname ByVal
''' */
Application.DisplayAlerts = False
Workbooks(bookName).Worksheets(sheetName).Delete
'
Application.DisplayAlerts = True
'''****************************************
End Function

Public Function GetBookProperties(ByVal bookName As String) As Object
'''/****************************************
''' * @function getBookProperties
''' * <br/>対象ブックのプロパティ情報を取得する
''' * @param {String}bookname ByVal
''' */
Dim i As Long
Dim bk As Workbook
Set bk = Workbooks(bookName)
Dim dict As Object
Set dict = createObject("Scripting.Dictionary")
'
On Error Resume Next
For i = 1 To bk.BuiltinDocumentProperties.Count
    Call dict.Add(bk.BuiltinDocumentProperties(i).Name, bk.BuiltinDocumentProperties(i).Value)
Next i
'
Set GetBookProperties = dict
'
Set bk = Nothing
'''****************************************
End Function

Public Function SetBookProperties(ByVal bookName As String, dict As Dictionary)
'''/****************************************
''' * @function setBookProperties
''' * <br/>対象ブックのプロパティ情報を設定する。
''' * @param {String}bookname ByVal
''' */
' BuiltinDocumentPropertiesにあるキー値は、当該値を、BuiltinDocumentPropertiesに
' 対して、更新する。
' BuiltinDocumentPropertiesに無いキー値は、当該値を、CustomDocumentPropertiesに
' 対して、追加・更新する。
Dim i As Long
Dim bk As Workbook
Set bk = Workbooks(bookName)
'Dim dict As Object
'Set dict = CreateObject("Scripting.Dictionary")
'
On Error Resume Next
For i = 1 To bk.BuiltinDocumentProperties.Count
    Call dict.Add(bk.BuiltinDocumentProperties(i).Name, bk.BuiltinDocumentProperties(i).Value)
Next i
'
Set bk = Nothing
Set dict = Nothing
'''****************************************
End Function

Public Function GetThisWorkbookFolder()
'''' ****************************************
''
Dim C_File As C_File
Set C_File = New C_File
''
Dim book_path As String
Let book_path = ThisWorkbook.FullName
Dim base_folder As String
Let base_folder = C_File.GetParentFolder(C_File.GetLocalName(book_path))
Let GetThisWorkbookFolder = base_folder
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
''
'//--------------------------------------
'//variables
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "C_Book.unitTest"
Dim dict As Dictionary
'
On Error GoTo EXCEPTION
    '//--------------------------------------
    '// initialize
    '
    Set dict = Me.GetBookProperties(ThisWorkbook.Name)
    '//--------------------------------------
    '// process
    dict.Add "sakamoto", "misa"
    dict.Add "kida", "keisuke"
    
    Console.Dump dict
    
    Call Me.SetBookProperties(ThisWorkbook.Name, dict)
    
    Set dict = Me.GetBookProperties(ThisWorkbook.Name)
    
    Console.Dump dict
    
    GoSub FINALLY
    Exit Sub
    
FINALLY: 'Label
    '//--------------------------------------
    '// end process
    Set dict = Nothing
    Return

EXCEPTION: 'Label
    '//--------------------------------------
    ' error process
    GoSub FINALLY
    Console.info ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    
End Sub

