VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Config.cls
'''  *<pre>Config utility
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/04/30
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  *
'''  */
'
''' /********************************************************
'''  * @class C_Config
'''  * <pre> Config utility Class
'''  * </pre>
'''  *
'''  * @require GlobalVariable, C_File
'''  */
'
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Sub GetConfig(Optional eachWb As Variant)
'''' ********************************************************
Dim C_File As C_File
Set C_File = New C_File
Dim C_FileIO As C_FileIO
Set C_FileIO = New C_FileIO
''
Dim CONS_CONFIG_FILE As String
Let CONS_CONFIG_FILE = "config.json"
Dim book_fullpath As String
Let book_fullpath = ThisWorkbook.FullName
Dim script_folder As String
Let script_folder = C_File.GetParentFolder(C_File.GetLocalName(book_fullpath))
Dim book_baseName As String
Let book_baseName = C_File.GetBaseName(C_File.GetLocalName(book_fullpath))
Dim config_path As String

If IsMissing(eachWb) Then
    Let config_path = C_File.BuildPath(script_folder, CONS_CONFIG_FILE)
Else
    Let config_path = C_File.BuildPath(script_folder, book_baseName & "." & CONS_CONFIG_FILE)
End If
''Console.info config_path
''
Dim json_str As String
Let json_str = C_FileIO.ReadTextAllAsUTF8(config_path)
Dim dict As Dictionary
Set dict = JsonConverter.ParseJson(json_str)
''
cdset.PutValue dict, "CONF"
''
Console.log ">> Have set config.json into common dataset (cdset) ..."
''
End Sub


Public Function ReadIniFile(ByVal iniFilePath As String, ByVal sction As String, ByVal ky As String) As String
'''' *************************************************
'''' https://stackoverflow.com/questions/660138/recommended-way-to-read-and-write-ini-files-in-vba/662225
Dim pos As Long
Dim strRtn As String * 128
Dim strSize As Long
''
Dim posOfCharInIni As Long
Let posOfCharInIni = 0
Dim strIni As String
Let strIni = ""
Dim strProfile
''
If sction = "" Or ky = "" Then
  Debug.Print "Section Or Key To Read Not Specified !!!", vbExclamation, "INI"
Else
  strProfile = ""
  strRtn = Space(128)
  strSize = Len(strRtn)
  pos = GetPrivateProfileString(sction, ky, "", strRtn, strSize, iniFilePath)
  If pos Then
    posOfCharInIni = pos
    strIni = left$(strRtn, pos)
  End If
End If
''
ReadIniFile = strIni
''
End Function

Public Function WriteIniFile(ByVal iniFilePath As String, ByVal sction As String, ByVal ky As String, ByVal Value As String) As String
'''' *************************************************
Dim pos As Long
''
Dim posOfCharInIni As Long
Let posOfCharInIni = 0
Dim strIni As String
Let strIni = ""
If sction = "" Or ky = "" Then
  Debug.Print "Section Or Key To Read Not Specified !!!", vbExclamation, "INI"
Else
  pos = WritePrivateProfileString(sction, ky, Value, iniFilePath)
  If pos Then
    posOfCharInIni = pos
    strIni = Value
  End If
  WriteIniFile = strIni
End If
''
End Function


