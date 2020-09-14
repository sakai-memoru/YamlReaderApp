VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_FileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_FileIO.cls
'''  *<br/>FileIO共通クラス
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  */
'
''' /********************************************************
'''  * @class FileIO
'''  * @classdesc FileIO Common Class
'''  * @dependency C_File, C_String, C_Array, C_Config
'''  */
'
'//-----------------------------------------
'// Instance Property　Constant
'    Const CONS_STREAM_CHARSET_UTF As String = "UTF-8"
''' /** @name CONS_STREAM_TYPE_TEXT @constant @type {long} @default @desc {1 : Binary, 2 : Text} */
Const CONS_STREAM_TYPE_TEXT As Long = 2 ' {1 : Binary, 2 : Text}
''' /** @name CONS_STREAM_MODE_TYPE @constant @type {long} @default @desc {1 : New File, 2 : Update} */
Const CONS_STREAM_MODE_TYPE As Long = 1 ' {1 : New File, 2 : Update}

'//-----------------------------------------
'// Instance Property　Variable
''' /** @name objStream @property @type {Object} @desc ADODB.Stream用Object */
Dim objStream As Object

Private Sub Class_Initialize()
''' /********************************************************
'''  * @event Class_Initialize
'''  * @fire Class 生成時
'''  *<br/>Initialize 関数
'''  *<br/>Class生成時に、実行される。
'''  */
Set objStream = createObject("ADODB.Stream")
objStream.Type = CONS_STREAM_TYPE_TEXT
objStream.Charset = "UTF-8"
End Sub


Private Sub Class_Terminate()
''' /********************************************************
'''  * @event Class_Terminate
'''  * @fire Class 廃棄時
'''  *<br/>Terminate 関数
'''  *<br/>Class生成時に、実行される。
'''  */
Set objStream = Nothing
End Sub


'//-----------------------------------------
'// Instance Method
Public Function WriteTextAllAsUTF8(Optional ByVal filePath As String, Optional ByVal src As String) As Long
''' /********************************************************
'''  *@function writeTextAllAsUTF8
'''  *</br>UTF-8 BOM付でテキストを出力する
'''  * @param {string} filepath　【Optional】出力するファイルパス名
'''  * @param {string} src 　【Optional】出力内容
'''  * @param {int} 実行結果コード
'''  */
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "C_FileIO.writeTextAllAsUTF8"
Dim adSaveCreateOverWrite As Long
Let adSaveCreateOverWrite = 2
''
On Error GoTo EXCEPTION
    Call objStream.Open
    Call objStream.WriteText(src)
    Call objStream.SaveToFile(filePath, adSaveCreateOverWrite)
    Call objStream.Close
    '返却
    Let WriteTextAllAsUTF8 = 0
    Exit Function
    ''
EXCEPTION:
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    ''
End Function

Public Function WriteTextAllAsUTF8NoneBOM(Optional ByVal filePath As String, Optional ByVal src As String) As Long
''' /********************************************************
'''  *@function WriteTextAllAsUTF8NoneBom
'''  *</br>UTF-8でテキストを出力する
'''  * @param {string} filepath　【Optional】出力するファイルパス名
'''  * @param {string} src 　【Optional】出力内容
'''  * @param {int} 実行結果コード
'''  * @reference https://unot13.hatenablog.com/entry/2013/12/04/194009
'''  */
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "C_FileIO.WriteTextAllAsUTF8NoneBOM"
Dim adTypeBinary As Long
Let adTypeBinary = 1
Dim adTypeText As Long
Let adTypeText = 2
Dim adSaveCreateOverWrite As Long
Let adSaveCreateOverWrite = 2
Dim byt As Variant
''
Dim objStreamBom As Object
Set objStreamBom = createObject("ADODB.Stream")
Let objStreamBom.Type = adTypeText
Let objStreamBom.Charset = "UTF-8"
Dim objStreamNoBom As Object
Set objStreamNoBom = createObject("ADODB.Stream")
Let objStreamNoBom.Type = adTypeBinary
''
On Error GoTo EXCEPTION
    Call objStreamBom.Open
    Call objStreamBom.WriteText(src)
    Let objStreamBom.Position = 0
    Let objStreamBom.Type = adTypeBinary
    Let objStreamBom.Position = 3
    Let byt = objStreamBom.Read()
    Call objStreamBom.Close
    ''
    Call objStreamNoBom.Open
    Call objStreamNoBom.write(byt)
    Console.log filePath
    Call objStreamNoBom.SaveToFile(filePath, adSaveCreateOverWrite)
    Call objStreamNoBom.Close
    '返却
    Let WriteTextAllAsUTF8NoneBOM = 0
    Set objStreamBom = Nothing
    Set objStreamNoBom = Nothing
    Exit Function
    ''
EXCEPTION:
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    ''
End Function

Public Function ReadTextAllAsUTF8(strFilePath As String) As String
''' /****************************************
'''  *@function readTextAllAsUTF8
'''  *</br>UTF-8でテキストを入力する
'''  * @param  {String} strFilePath ファイルパス名
'''  * @return {String} 読み取ったテキスト（全部）
'''  */
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "C_FileIO.readTextAllAsUTF8"
Dim str As String
''
On Error GoTo EXCEPTION
    Call objStream.Open
    Call objStream.LoadFromFile(strFilePath)
    Let str = objStream.ReadText()
    Call objStream.Close
    Let ReadTextAllAsUTF8 = str
    Exit Function
    ''
EXCEPTION:
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    ''
End Function

Public Function ReadTextAllAsUTF8ToArray(strFilePath As String) As Variant
''' /****************************************
'''  *@function readTextAllAsUTF8ToArray
'''  *</br>UTF-8でテキストを入力する
'''  * @param  {String} strFilePath ファイルパス名
'''  * @return {Variant} 読み取ったテキスト（全部）
'''  */
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "C_FileIO.readTextAllAsUTF8ToArray"
''
Dim str As String
Dim aryRtn As Variant
''
On Error GoTo EXCEPTION
    Call objStream.Open
    Call objStream.LoadFromFile(strFilePath)
    Let str = objStream.ReadText()
    Call objStream.Close
    ''
    Let aryRtn = Split(str, vbCrLf)
    Let ReadTextAllAsUTF8ToArray = aryRtn
    Exit Function
    ''
EXCEPTION:
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    ''
End Function

Public Function ReadJsonFile(ByVal file_path As String) As Variant
'''' *************************************************
''
Dim C_String As C_String
Set C_String = New C_String
''
Dim vntRtn As Variant
''
Dim json_str As String
Let json_str = C_String.TrimSpace(Me.ReadTextAllAsUTF8(file_path))
If C_String.IsJsonString(json_str) Then
    Set vntRtn = JsonConverter.ParseJson(json_str)
    Set ReadJsonFile = vntRtn
Else
    vntRtn = New Collection
End If
End Function

Sub SaveAsTSV(Optional ByVal aryary As Variant, Optional ByVal file_path As Variant)
'''' *************************************************
Dim CONS_MODULE_NAME As String
Let CONS_MODULE_NAME = "C_FileIO.SaveAsTSV"
''
Dim C_File As C_File
Set C_File = New C_File
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim dtmStr As String
Let dtmStr = format(Now(), "_yymmddhhnnss")
Dim outputPath As String
Dim outputFile As String
Dim outputFolder As String
Dim delimiter As String
''
If IsMissing(file_path) Then
    Let outputFolder = ThisWorkbook.path
    Let outputFile = C_File.GetBaseName(ThisWorkbook.FullName) & dtmStr & ".tsv"
    Let outputPath = C_File.BuildPath(outputFolder, outputFile)
Else
    Let outputPath = file_path
End If
Dim ext As String
Let ext = C_File.GetExtensionName(file_path)
If ext = "csv" Then
    Let delimiter = ","
Else
    Let delimiter = vbTab
End If
''
Dim output_str As String
Let output_str = C_Array.ToCsv(aryary, delimiter)
''
On Error GoTo EXCEPTION
    ''Call Me.WriteTextAllAsUTF8(outputPath, output_str)
    Call Me.WriteTextAllAsUTF8NoneBOM(outputPath, output_str)
    Exit Sub
''
EXCEPTION:
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
''
End Sub

Function LoadTsvToArrayArray(ByVal file_name As String, _
                Optional useHeader As Variant, _
                Optional folder_path As Variant) As Variant
'''' *************************************************
''
Dim C_String As C_String
Set C_String = New C_String
Dim C_File As C_File
Set C_File = New C_File
Dim C_Config As C_Config
Set C_Config = New C_Config
Dim C_Array As C_Array
Set C_Array = New C_Array
''
If IsMissing(folder_path) Then
    Let folder_path = ThisWorkbook.path & "\"
Else
    If Not C_String.EndsWith(folder_path, "\") Then
        Let folder_path = folder_path & "\"
    End If
End If
''
Dim templCnStatement As String
Dim templCnStatementCsvHead As String
Let templCnStatementCsvHead = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=${folder_path};Extended Properties=""text;HDR=Yes;FMT=Delimited"";"
Dim templCnStatementCsvNoHead As String
Let templCnStatementCsvNoHead = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=${folder_path};Extended Properties=""text;HDR=NO;FMT=Delimited"";"
Dim templCnStatementTsvHead As String
Let templCnStatementTsvHead = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=${folder_path};Extended Properties=""text;HDR=YES;FMT=Delimited;TAB;"";"
Dim templCnStatementTsvNoHead As String
Let templCnStatementTsvNoHead = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=${folder_path};Extended Properties=""text;HDR=NO;FMT=Delimited;TAB;"";"
Dim templSqlStatement As String
Let templSqlStatement = "SELECT * FROM ${file_name};"
''
If IsMissing(useHeader) Then
    Let useHeader = True
End If
''
Dim filePath As String
Dim ext As String
Dim basename As String
Let filePath = C_File.BuildPath(folder_path, file_name)
Let basename = C_File.GetBaseName(file_name)
Let ext = C_File.GetExtensionName(file_name)
''
Dim tempFile As String
Dim tempPath As String
Dim iniPath As String
Dim iniName As String
Let iniName = "schema.ini"
''
Dim dict As Dictionary
Set dict = New Dictionary
''
If ext = "csv" Then
    If useHeader Then
        Let templCnStatement = templCnStatementCsvHead
    Else
        Let templCnStatement = templCnStatementCsvNoHead
    End If
    Call dict.Add("folder_path", folder_path)
    Call dict.Add("file_name", file_name)
ElseIf ext = "tsv" Then
    If useHeader Then
        Let templCnStatement = templCnStatementTsvHead
    Else
        Let templCnStatement = templCnStatementTsvNoHead
    End If
    Let tempFile = basename & ".txt"
    Let tempPath = C_File.BuildPath(folder_path, tempFile)
    Let iniPath = C_File.BuildPath(folder_path, iniName)
    Call C_File.CopyFile(filePath, tempPath)
    ''
    Call dict.Add("folder_path", folder_path)
    Call dict.Add("file_name", tempFile)
ElseIf ext = "txt" Then
    If useHeader Then
        Let templCnStatement = templCnStatementTsvHead
    Else
        Let templCnStatement = templCnStatementTsvNoHead
    End If
    Call dict.Add("folder_path", folder_path)
    Call dict.Add("file_name", file_name)
    Let iniPath = C_File.BuildPath(folder_path, iniName)
End If
Dim flgNoScheme As String
If ext = "csv" Then
    Let flgNoScheme = "CsvDelimited"
Else
    Let flgNoScheme = C_Config.ReadIniFile(iniPath, dict.Item("file_name"), "Format")
End If
''
''
Dim cnStatement As String
Let cnStatement = C_String.RenderTemplate(templCnStatement, dict)
Dim sqlStatement As String
Let sqlStatement = C_String.RenderTemplate(templSqlStatement, dict)
''
Dim cn As Object
Set cn = createObject("ADODB.Connection")
Dim rs As Object
Set rs = createObject("ADODB.Recordset")
Call cn.Open(cnStatement)
Set rs = cn.Execute(sqlStatement)
''
Dim rsAryary() As Variant
Let rsAryary = WorksheetFunction.Transpose(rs.GetRows)
If C_String.IsEmpty(flgNoScheme) Then
    Dim r As Long
    Dim jagary() As Variant
    ReDim jagary(LBound(rsAryary) To UBound(rsAryary))
    If C_Array.TypeArrayPattern(rsAryary) = "Array" Then
        jagary(LBound(jagary)) = VBA.Split(rsAryary(LBound(rsAryary)), vbTab)
    Else
        For r = LBound(rsAryary) To UBound(rsAryary)
            jagary(r) = VBA.Split(rsAryary(r, 1), vbTab)
        Next r
    End If
    Let rsAryary = C_Array.JagArrayToArrayArray(jagary)
End If
''
Call rs.Close
Set cn = Nothing
If C_File.fileExists(tempPath) Then
    Call C_File.DeleteFile(tempPath)
End If
''
Let LoadTsvToArrayArray = rsAryary
'Console.info cnStatement
'Console.info sqlStatement
'Console.dump rsAryary
''
End Function

