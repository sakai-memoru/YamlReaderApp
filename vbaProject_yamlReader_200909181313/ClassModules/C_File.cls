VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_File.cls
'''  * <br/> File Utillity Classs
'''  * @version 1.00
'''  * @since 2020/03/31
'''  *
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  */
'
''' /********************************************************
'''  * @class C_File
'''  * @require Be avaiable for "Microsoft Scripting Runtime"
'''  * @require C_String
'''  */
'
'//-----------------------------------------
'// Instance Property　Constant
''' /** @property @const @name CONS_DESKTOP */
Const CONS_DESKTOP As String = "Desktop"
''' /** @property @const @name CONS_MYDOC */
Const CONS_MYDOC As String = "MyDocuments"
'//-----------------------------------------
'// Instance Property　Variable
''' /** @property @Array @name aryCONS */
Dim aryCONS As Variant
''' /** @property @object {Scripting.FileSystemObject} @name objFSO */
Dim objFSO As Object

'''//--------------------------------------------------------
'''// Constructor
'''//
Private Sub Class_Initialize()
''' /********************************************************
'''  * @event Class_Initialize
'''  */
    aryCONS = Array(CONS_DESKTOP, CONS_MYDOC)
    Set objFSO = createObject("Scripting.FileSystemObject")
End Sub

Private Sub Class_Terminate()
''' /********************************************************
'''  * @event Class_Terminate
'''  */
'Destructor
    Set objFSO = Nothing
End Sub

'//-----------------------------------------
'// Instance Method
'
Public Function GetSpecialFolder(Optional ByVal strFoldername As Variant = "Desktop") As String
''' /********************************************************
'''  *@function getSpecialFolder
'''  *<br/> [WScript.Shell]::SpecialFolders property wrapper
'''  *<br/> Refer WScript.Shell#SpecialFolders
'''  *@param {String} FolderPath
'''  *@return {String} FilePath
'''  */
    '//--------------------------------------
    '// variables
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.getSpecialFolder"
    Dim objWShell As Object
    Set objWShell = createObject("WScript.Shell")
    Dim strRet As String
    '
On Error GoTo EXCEPTION
    '//--------------------------------------
    '// process
    strRet = objWShell.SpecialFolders(strFoldername)
    '
    '//release and return
    GoSub FINALLY
    GetSpecialFolder = strRet
    Exit Function
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Set objWShell = Nothing
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Function

Public Function GetLocalName(ByVal fullPath As String) As String
'''' ********************************************************
'''' https://stackoverflow.com/questions/33734706/excels-fullname-property-with-onedrive
''
Dim i As Long, j As Long
Dim oneDrivePath As String
Dim endFilePath As String
Dim iDocumentsPosition As Integer
''
''Check if it looks like a OneDrive location
If InStr(1, fullPath, "https://", vbTextCompare) > 0 Then
    ''
    'for commercial OneDrive file path :like "https://companyName-my.sharepoint.com/personal/userName_domain_com/Documents" & file.FullName)
    If InStr(1, fullPath, "my.sharepoint.com") <> 0 Then
        'find "/Documents" in string and replace everything before the end with OneDrive local path
        iDocumentsPosition = InStr(1, fullPath, "/Documents") + Len("/Documents")
        endFilePath = Mid(fullPath, iDocumentsPosition)
    Else
        'for personal onedrive it looks like "https://d.docs.live.net/d7bbaa#######1/" & file.FullName,
        endFilePath = Mid(fullPath, 9)
        For i = 1 To 2
            endFilePath = Mid(endFilePath, InStr(endFilePath, "/") + 1)
        Next
    End If
    ''
    endFilePath = Replace(endFilePath, "/", Application.PathSeparator)
    
    ' Loop to see if the tentative LocalWorkbookName is the name of a file that actually exists,
    ' if so return the name
    For j = 1 To 3
        oneDrivePath = Environ(Choose(j, "OneDrive", "OneDriveCommercial", "OneDriveConsumer"))
        If Len(oneDrivePath) > 0 Then
            GetLocalName = oneDrivePath & Application.PathSeparator & endFilePath
            If Dir(GetLocalName) <> "" Then
                Exit Function
            End If
        End If
    Next j
    GetLocalName = ""
Else
    GetLocalName = fullPath
End If
''
End Function

Public Sub ChangeDirectory(ByVal strFoldername As Variant)
''' /********************************************************
'''  *@function changeDirectory
'''  *<br/> [WScript.Shell]::CurrentDirectory property wrapper
'''  *<br/> Refer WScript.Shell#CurrentDirectory
'''  *　param {String} FolderPath
'''  *
'''  *@deprecated
'''  */
    '//--------------------------------------
    '//variables
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.changeCDirectory"
    '
    Dim objWShell As Object
    Set objWShell = createObject("WScript.Shell")
    '
On Error GoTo EXCEPTION
    '//--------------------------------------
    '// process
    objWShell.CurrentDirectory = strFoldername 'Notice:: Not be able to use with VBA
    '
    '//release and return
    GoSub FINALLY
    Exit Sub
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Set objWShell = Nothing
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Sub

Public Sub ChDirectory(ByVal strFoldername As Variant)
''' /********************************************************
'''  *@function chDirectory
'''  *<br/> [VBA]::ChDir Wrapper
'''  *<br/> Refer VBA.FileSystem#ChDir
'''  *<br/> Use this methon on ExcelVBA
'''  *@param {String} FolderPath
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.chDirectory"
    '//--------------------------------------
    '// process
On Error GoTo EXCEPTION
    ChDir strFoldername 'For VBA
    '
    '//release and return
    GoSub FINALLY
    Exit Sub
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Sub

Public Function GetCurrentDirectory(Optional ByVal driveString As String = "") As String
''' /********************************************************
'''  *@function getCurrentDirectory
'''  *<br/> [VBA]::CurDir wrapper
'''  *<br/> Refer to CurDir function
'''  *@param {String} driveString
'''  *@return {String} FolderPath
'''  */
GetCurrentDirectory = CurDir(driveString)
End Function

Public Function BuildPath(ByVal folderPath As String, ByVal fileName As String) As String
''' /********************************************************
'''  *@function buildPath
'''  *<br/> [FileSystemObject]::BuildPath wrapper
'''  *<br/> Refer Scripting.FileSytemObject
'''  *@param {String} folderPath
'''  *@param {String} fileName
'''  *@return {String} FilePath
'''  */
BuildPath = objFSO.BuildPath(folderPath, fileName)
End Function

Public Function GetAbsolutePathName(ByVal pathspec As String) As String
''' /********************************************************
'''  *@function GetAbsolutePathName
'''  *<br/> [FileSystemObject]::GetAbsolutePathName wrapper
'''  *<br/> Refer Scripting.FileSytemObject
'''  *@param {String} pathspec
'''  *@return {String} FilePath
'''  */
GetAbsolutePathName = objFSO.GetAbsolutePathName(pathspec)
End Function

Public Function GetTempName(Optional ByVal strFoldername As Variant) As String
''' /********************************************************
'''  *@function getTempName
'''  *<br/> [FileSystemObject]::getTempName wrapper
'''  *<br/> Refer Scripting.FileSytemObject
'''  *　param {String} strFolderName
'''  */
If IsMissing(strFoldername) Then
    strFoldername = GetSpecialFolder()
End If
GetTempName = Me.BuildPath(strFoldername, objFSO.GetTempName)
''
End Function

Public Function GetParentFolderName(ByVal strFilePath As String)
''' /********************************************************
'''  *@function getParentFolderName
'''  *<br/> [FileSystemObject]::GetParentFolderName wrapper
'''  *<br/> Refer Scripting.FileSytemObject
'''  *@param {String} strFilePath
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.getParentFolderName"
    '
On Error GoTo EXCEPTION
    '//--------------------------------------
    '// process and return
    GetParentFolderName = objFSO.GetParentFolderName(strFilePath)
    '//release
    GoSub FINALLY
    Exit Function
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Function

Public Function GetParentFolder(ByVal strFilePath As String)
''' /********************************************************
'''  *@function getParentFolder
'''  *<br/> [FileSystemObject]::GetFile().ParentFolder wrapper
'''  *<br/> Refer Scripting.FileSytemObject
'''  *@param {String} strFilePath
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.getParentFolder"
On Error GoTo EXCEPTION
    '//--------------------------------------
    '// process and return
    If Not Me.fileExists(strFilePath) Then
        GetParentFolder = GetParentFolderName(strFilePath)
    Else
        GetParentFolder = objFSO.GetFile(strFilePath).parentFolder
    End If
    '
    '//release
    GoSub FINALLY
    Exit Function
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Function

Public Function GetFileName(ByVal strFilePath As String)
''' /********************************************************
'''  *@function getParentFolder
'''  *<br/> [FileSystemObject]::GetFile().Name wrapper
'''  *<br/> Refer Scripting.FileSytemObject
'''  *@param {String} strFilePath
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.getFileName"
    Dim strFoldername As String
On Error GoTo EXCEPTION
    '//--------------------------------------
    '// process and return
    If Not Me.fileExists(strFilePath) Then
        GetFileName = objFSO.GetFileName(strFilePath)
    Else
        GetFileName = objFSO.GetFile(strFilePath).Name
    End If
    '
    '//release
    GoSub FINALLY
    Exit Function
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Function

Public Function GetBaseName(ByVal strFilePath As String)
''' /********************************************************
'''  *@function getParentFolder
'''  *<br/> [FileSystemObject]::getBaseName wrapper
'''  *<br/> Refer Scripting.FileSytemObject
'''  *@param {String} strFilePath
'''  */
'
GetBaseName = objFSO.GetBaseName(strFilePath)
End Function

Public Function GetExtensionName(ByVal strFilePath As Variant)
''' /********************************************************
'''  *@function getExtensionName
'''  *<br/> [FileSystemObject]::GetExtensionName wrapper
'''  *<br/> Refer Scripting.FileSytemObject
'''  *@param {String} strFilePath
GetExtensionName = objFSO.GetExtensionName(strFilePath)
End Function

Public Function GetFiles(ByVal strFolderPath As String) As Variant
''' /********************************************************
'''  *@function getFiles
'''  *<br/> Get array of files in parameter of folder path
'''  *<br/>
'''  *@param {String} strFolderPath
'''  *@return {String[]}
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.getFiles"
    '
On Error GoTo EXCEPTION
    Dim cltFiles As Variant
    Dim clt As Variant
    Dim ary() As Variant
    Dim i As Long
    '//--------------------------------------
    '// process and return
    Set cltFiles = objFSO.getFolder(strFolderPath & "\").Files
    '
    ReDim ary(cltFiles.Count - 1)
    '
    i = 0
    For Each clt In cltFiles
        ary(i) = strFolderPath & "\" & clt.Name
        i = i + 1
    Next
    '
    GetFiles = ary
    '
    '//release
    GoSub FINALLY
    Exit Function
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Function

Public Function GetFilesByKeyword(ByVal strFolderPath As String, ByVal keyword As String, Optional ignorecase As Boolean = False, Optional Include As Boolean = True) As Variant
''' /********************************************************
'''  *@function GetFilesByKeyword
'''  *<br/> Get array of files in parameter of folder path
'''  *<br/>with Fileter() condition
'''  *@param {String} strFolderPath
'''  *@param {String} keyword
'''  *@param {Boolean} ignoreCase: in keyword:true, excluding out of keyword :false
'''  *@param {Boolean} include : in keyword:true, excluding out of keyword :false
'''  *@return {String[]}
'''  *
'''  *@require C_String
'''  * FIXME□modify code for treatment of path
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.getFilesByKeyword"
    '
    Dim C_String As C_String
    Set C_String = New C_String
    Dim C_Array As C_Array
    Set C_Array = New C_Array
    Dim ary As Variant
    Dim aryBuff As Variant
    Dim aryTemp As Variant
    ReDim aryBuff(0)
    Dim i As Long
    '
    '//--------------------------------------
    '// process and return
    ary = GetFiles(strFolderPath)
    If Not ignorecase Then
        GetFilesByKeyword = Filter(ary, keyword, Include) 'Use Filter function
    Else
        For i = LBound(ary) To UBound(ary)
            If C_String.ContainWord(ary(i), keyword, ignorecase) Then
                aryBuff(UBound(aryBuff)) = ary(i)
                ReDim Preserve aryBuff(0 To UBound(aryBuff) + 1)
                'C_Array.add aryBuff, ary(i) FIXME
            End If
        Next i
        C_Array.DeleteElementAt aryBuff, UBound(aryBuff)
        GetFilesByKeyword = aryBuff
    End If
    '
End Function

Public Function GetFilesByWildcard(ByVal strFolderPath As String, ByVal wildcard As String) As Variant
''' /********************************************************
'''  *@function getFilesByWildcard
'''  *<br/> Get array of files in parameter of folder path
'''  *<br/> with wildcard condition
'''  *
'''  *　param {String} strFolderPath
'''  *　param {String} wildcard  : can use *, ? wild card
'''  *   The use of filter is provided from the windows file system
'''  *
'''  *@return {String[]}
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.getFilesByWildcard"
    '
    Dim cltFiles As Files
    Dim f As Object
    '
    Dim ary As Variant
    ary = Array()
    Dim i     As Long
    i = 0
    '//--------------------------------------
    '// process and return
    If strFolderPath = "" Then
        strFolderPath = GetCurrentDirectory()
    End If
    Set cltFiles = objFSO.getFolder(strFolderPath).Files
    '
    If cltFiles.Count <> 0 Then
        For Each f In cltFiles
            If (objFSO.GetFileName(f.Name) Like wildcard) Then
                ReDim Preserve ary(i)
                ary(i) = strFolderPath & "\" & f.Name
                i = i + 1
            End If
        Next f
    End If
    GetFilesByWildcard = ary
    '
End Function

Public Sub CopyFile(ByVal source As String, ByVal destination As String, Optional overWrite As Boolean = True)
''' /********************************************************
'''  *@function copyFile
'''  *<br/> copy a source file to a destination file
'''  *<br/> [FileSystemObject]::CopyFile Wrapper
'''  *<br/> Can set wildcard in a source param. If no match with wildcard,
'''  *<br/> it occurs an error. Can apply a wildcard to files' name.
'''  *<br/> Can set a directory (with last character "/") to destination,
'''  *<br/> it's transleted as a folder. If a folder does not exists,
'''  *<br/> it occurs an error.
'''  *
'''  *@param {String} source
'''  *@param {String} destination
'''  *@param {Boolean} overwrite
'''  *@Exception when occurred errors on operating files
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.copyFile"
    '
On Error GoTo EXCEPTION
    objFSO.CopyFile source, destination, overWrite
    '
    '//release
    GoSub FINALLY
    Exit Sub
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Sub

Public Sub CopyFolder(ByVal source As String, ByVal destination As String, Optional overWrite As Boolean = True)
''' /********************************************************
'''  *<br/> copy a source file to a destination file
'''  *<br/> [FileSystemObject]::copyFolder Wrapper
'''  *
'''  *@param {String} source : a source folder
'''  *@param {String} destination : a destination folder
'''  *@param {Boolean} overwrite
'''  *@Exception when occurred errors on operating files
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.copyFolder"
    '
On Error GoTo EXCEPTION
    objFSO.CopyFolder source, destination, overWrite
    '
    '//release
    GoSub FINALLY
    Exit Sub
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Sub

Public Sub MoveFile(ByVal source As String, ByVal destination As String)
''' /********************************************************
'''  *@function moveFile
'''  *<br/> move a source file to a destination file
'''  *<br/> [FileSystemObject]::moveFile Wrapper
'''  *
'''  *@param {String} source : file's wildcard format
'''  *@param {String} destination : a destination files
'''  *@Exception when occurred errors on operating files
'''  *  If the same name file exists in distination folder,
'''  *  it occurs error.
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.moveFile"
    '
On Error GoTo EXCEPTION
    objFSO.MoveFile source, destination
    '
    '//release
    GoSub FINALLY
    Exit Sub
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Sub

Public Sub MoveFolder(ByVal source As String, ByVal destination As String)
''' /********************************************************
'''  *@function moveFolder
'''  *<br/> move a source file to a destination file
'''  *<br/> [FileSystemObject]::moveFolder Wrapper
'''  *
'''  *@param {String} source : a source folder
'''  *@param {String} destination : a destination folder that will be created newly
'''  *@Exception when occurred errors on operating files
'''  *  If the same name folder of distination folder exists,
'''  *  it occurs error.
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.moveFolder"
    '
On Error GoTo EXCEPTION
    objFSO.MoveFolder source, destination
    '
    '//release
    GoSub FINALLY
    Exit Sub
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Sub

Public Sub DeleteFile(ByVal filespec As String, Optional ByVal force As Boolean = False)
''' /********************************************************
'''  *@function deleteFile
'''  *<br/> delete a spec file
'''  *<br/> [FileSystemObject]::deleteFile Wrapper
'''  *
'''  *@param {String} filespec : target file (wildcard format)
'''  *@Exception when occurred errors on operating files
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.deleteFile"
    '
On Error GoTo EXCEPTION
    If Me.fileExists(filespec) Then
        objFSO.DeleteFile filespec, force
    End If
    '
    '//release
    GoSub FINALLY
    Exit Sub
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Sub

Public Sub DeleteFolder(ByVal folderspec As String, Optional ByVal force As Boolean = False)
''' /********************************************************
'''  *@function deleteFolder
'''  *<br/> delete a spec folder
'''  *<br/> [FileSystemObject]::DeleteForlder Wrapper
'''  *
'''  *@param {String} folderspec : target folder (wildcard format)
'''  *@Exception when occurred errors on operating files
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.deleteFolder"
    '
On Error GoTo EXCEPTION
    If Me.FolderExists(folderspec) Then
        objFSO.DeleteFolder folderspec, force
    End If
    '
    '//release
    GoSub FINALLY
    Exit Sub
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Sub

Public Sub CreateFile(ByVal filespec As String)
''' /********************************************************
'''  *@function createFile
'''  *<br/> create a new folder with folderspec statement,
'''  *<br/> if exists no work.
'''  *<br/> [FileSystemObject]::createFolder Wrapper
'''  *
'''  *@param {String} folderspec
'''  *@Exception when error occur in file operation
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.createFile"
    '
On Error GoTo EXCEPTION
    Dim folderspec As String
    'Console.log (filespec)
    If Not Me.fileExists(filespec) Then
        objFSO.CreateTextFile filespec
    End If
    '
    '//release
    GoSub FINALLY
    Exit Sub
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Sub

Public Sub CreateFolder(ByVal folderspec As String)
''' /********************************************************
'''  *@function createFolder
'''  *<br/> create a new spec folder
'''  *<br/> [FileSystemObject]::createFolder Wrapper
'''  *
'''  *@param {String} folderspec : target folder spec
'''  *@Exception when occurred errors on operating files
'''  *
'''  *@param {String} folderspec
'''  *@Exception when occurred errors on operating files
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.createFolder"
    '
On Error GoTo EXCEPTION
    If Not Me.FolderExists(folderspec) Then
      objFSO.CreateFolder folderspec
    End If
    '
    '//release
    GoSub FINALLY
    Exit Sub
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Sub

Public Sub CreateFolders(ByVal folderspec As String)
''' /********************************************************
'''  *@function createFolders
'''  *<br/> create a new spec folder in the consideration
'''  *<br/> for path structure
'''  *<br/> [FileSystemObject]::createFolder Wrapper
'''  *
'''  *@param {String} folderspec : target folder spec
'''  *@Exception when occurred errors on operating files
'''  *
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.createFolders"
    Dim strParentFolderPath As String
    '
On Error GoTo EXCEPTION
    If Not Me.FolderExists(folderspec) Then
        strParentFolderPath = Me.GetParentFolderName(folderspec)
        If Not Me.FolderExists(strParentFolderPath) Then
            CreateFolders strParentFolderPath
        End If
        objFSO.CreateFolder folderspec
    End If
    '
    '//release
    GoSub FINALLY
    Exit Sub
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Sub

Public Function CreateNamespase(ByVal rootfolderspec As String, ByVal namespace As String, Optional ByVal delimiter As String = ".")
''' /********************************************************
'''  *@function createNamespase
'''  *<br/> folderspecに指定したファルダ階層を新規に作成する。
'''  *
'''  * @param {String} rootfolderspec
'''  * @param {String} namespace
'''  * @param {String} delimiter 【Optional】Default="."
'''  * @Exception フォルダ作成時にエラーが発生した場合
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.createNamespase"
    Dim strNamePath As String
    Dim target_path As String
    Dim ary As Variant
    '
On Error GoTo EXCEPTION
    If rootfolderspec = "" Then
        rootfolderspec = GetSpecialFolder()
    End If
    If namespace = "" Then
        namespace = format(Now(), "yyyymmddhhnnss")
    End If
    '
    ary = Split(namespace, delimiter)
    strNamePath = Join(ary, "\")
    '
    target_path = objFSO.BuildPath(rootfolderspec, strNamePath)
    Me.CreateFolders (target_path)
    '
    '//release
    CreateNamespase = target_path
    GoSub FINALLY
    
    Exit Function
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Function

Public Function fileExists(filespec As String) As Boolean
''' /********************************************************
'''  *@function FileExists
'''  *<br/> filespecに指定したファイルが存在するか否かを評価する
'''  *<br/> FileSystemObject#FileExistsメソッドのRapper
'''  *<br/>
'''  *
'''  *@param {String} filespec
'''  *@return {Boolean}ファイルが存在する場合:true
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.fileExists"
    '
On Error GoTo EXCEPTION
    If objFSO.fileExists(filespec) Then
      fileExists = True
    Else
      fileExists = False
    End If
    '
    '//release
    GoSub FINALLY
    Exit Function
    '
FINALLY: 'Label
    '//--------------------------------------
    '// end  process
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Function

Public Function FolderExists(folderspec As String) As Boolean
''' /********************************************************
'''  *@function folderExists
'''  *<br/> FileSystemObject#FolderExists wrapper
'''  *<br/>
'''  *
'''  *　param {String} filespec
'''  *　return {Boolean}
'''  */
'
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.folderExists"
    '
On Error GoTo EXCEPTION
    'Console.log ("folderExists param: " & folderspec)
    If objFSO.FolderExists(folderspec) Then
      FolderExists = True
    Else
      FolderExists = False
    End If
    '
    '// end process
    GoSub FINALLY
    Exit Function
    '
FINALLY: 'Label
    '//--------------------------------------
    '//
    Return
    '
EXCEPTION: 'Label
    GoSub FINALLY
    Console.log ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.log ("Err.Number = " & Err.Number)
    Console.log ("Err.Description = " & Err.Description)
    Call Err.Raise(Err.Number, Err.source, CONS_MODULE_NAME & vbCr & Err.Description, Err.HelpFile, Err.HelpContext)
    '
End Function

'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Sub UnitTest()

    '//--------------------------------------
    '// const
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_File.unitTest"
    
On Error GoTo EXCEPTION
    '//--------------------------------------
    '// variables
    '
    Dim C_File As C_File
    Set C_File = New C_File
    Dim C_Array As C_Array
    Set C_Array = New C_Array
    Dim str As String
    '//--------------------------------------
    '// process
    Console.log ("// -----------------// debug start")
'''getSpecialFolder
    'str = C_File.getSpecialFolder(CONS_MYDOC)
    'Console.Dump str
    'str = C_File.getSpecialFolder(CONS_DESKTOP)
    'Console.Dump str
'''chDirectory
'''getCurrentDirectory
    'str = C_File.getSpecialFolder(CONS_DESKTOP)
    ''Me.chDirectory (str)
    'C_File.chDirectory (str)
    'Dim currentPath As String
    'currentPath = C_File.getCurrentDirectory
    'Console.log (currentPath)
'''buildPath
'''createFolder
    'Dim childPath As String
    'childPath = "temp"
    'dim folder_path As String
    'folder_path = C_File.buildPath(currentPath, childPath)
    'Console.log (folder_path)
'''getTempName
    'Dim str1 As String
    'str1 = C_File.getTempName
    'Console.log (str1)
    '
    'Dim currentPath As String
    'currentPath = C_File.getCurrentDirectory
    'Dim childPath As String
    'childPath = "temp"
    'Dim folder_path As String
    'folder_path = C_File.buildPath(currentPath, childPath)
    'str = C_File.getTempName(folder_path)
    'Console.log (folder_path)
    'Console.log (str)

'''folderExists
    'Dim currentPath1 As String
    'currentPath1 = C_File.getCurrentDirectory
    'Dim childPath1 As String
    'childPath1 = "temp"
    'Dim folder_path1 As String
    'folder_path1 = C_File.buildPath(currentPath1, childPath1)
    'Dim bln1 As Boolean
    'bln1 = C_File.folderExists(folder_path1)
    'Console.log (bln1)
    ''
    'Dim childPath2 As String
    'childPath2 = "temp2"
    'Dim folder_path2 As String
    'folder_path2 = C_File.buildPath(currentPath1, childPath2)
    'Dim bln2 As Boolean
    'bln2 = C_File.folderExists(folder_path2)
    'Console.log (bln2)

'''fileExists
    'Dim currentPath1 As String
    'currentPath1 = C_File.getCurrentDirectory
    'Dim childPath1 As String
    'childPath1 = "temp"
    'Dim folder_path1 As String
    'folder_path1 = C_File.buildPath(currentPath1, childPath1)
    'Dim file1 As String
    'file1 = "test_notExist.txt"
    'Dim file_path1 As String
    'file_path1 = C_File.buildPath(folder_path1, file1)
    'Dim bln3 As Boolean
    'bln3 = C_File.fileExists(file_path1)
    'Console.log (bln3)
    ''
    'Dim file As String
    'file = "test_Exist.txt"
    'Dim file_path As String
    'file_path = C_File.buildPath(folder_path1, file)
    'Dim bln4 As Boolean
    'bln4 = C_File.fileExists(file_path)
    'Console.log (bln4)

'''getParentFolder
'''createFolder
'''createFile
    'Dim currentPath As String
    'currentPath = C_File.getCurrentDirectory
    'Dim childPath As String
    'childPath = "temp"
    'Dim folder_path As String
    'folder_path = C_File.buildPath(currentPath, childPath)
    'Dim file_path As String
    'file_path = C_File.getTempName(folder_path)
    'Console.log (file_path)
    'C_File.createFolder (folder_path)
    'C_File.createFile (file_path)
    ''
    'Dim parentFolder As String
    'parentFolder = C_File.getParentFolder(file_path)
    'Console.log ("file_path : " & file_path)
    'Console.log ("folder_path : " & folder_path)
    'Console.log ("parent_path : " & parentFolder)
'''getParentFolderName
    'Dim currentPath3 As String
    'currentPath3 = C_File.getCurrentDirectory
    'Dim childPath3 As String
    'childPath3 = "temp3"
    'Dim folder_path3 As String
    'folder_path3 = C_File.buildPath(currentPath3, childPath3)
    'Dim file_path3 As String
    'file_path3 = C_File.getTempName(folder_path3)
    'Dim parentFolder3 As String
    'parentFolder3 = C_File.getParentFolderName(file_path)
    'Console.log ("-----------------------getParentFolderName")
    'Console.log (parentFolder3)
'''getParentFolder
'''getFileName
'''getBaseName
'''getExtensionName
    'Dim currentPath2 As String
    'currentPath2 = C_File.getCurrentDirectory
    'Dim childPath2 As String
    'childPath2 = "temp2"
    'Dim folder_path2 As String
    'folder_path2 = C_File.buildPath(currentPath2, childPath2)
    'Dim file_path2 As String
    'file_path2 = C_File.getTempName(folder_path2)
    'Dim parentFolder2 As String
    'parentFolder2 = C_File.getParentFolder(file_path2)
    'Dim file_name2 As String
    'file_name2 = C_File.getFileName(file_path2)
    'Dim file_basename2 As String
    'file_basename2 = C_File.getBaseName(file_path2)
    'Dim file_ext2 As String
    'file_ext2 = C_File.getExtensionName(file_path2)
    ''
    'Console.log ("-----------------------etc")
    'Console.log ("file_path2 : " & file_path2)
    'Console.log ("folder_path2 : " & folder_path2)
    'Console.log ("parent_path2 : " & parentFolder2)
    'Console.log ("file_name2 : " & file_name2)
    'Console.log ("file_basename2 : " & file_basename2)
    'Console.log ("file_ext2 : " & file_ext2)
'''getFiles
'''getFilesByKeyword
'    Dim basePath As String
'    basePath = C_File.getSpecialFolder
'    Dim childPath As String
'    childPath = "temp"
'    Dim folder_path As String
'    folder_path = C_File.buildPath(basePath, childPath)
'    Dim aryFiles As Variant
'    '
'    'Console.log ("-----------------------getFiles")
'    'Console.log ("folder_path : " & folder_path)
'    'aryFiles = C_File.getFiles(folder_path)
'    'Console.Dump (aryFiles)
'    Console.log ("-----------------------getFilesByKeyword")
'    aryFiles = C_File.getFilesByKeyword(folder_path, "rad")
'    Console.Dump (aryFiles)
'    'C_Array.initialize (aryFiles)
'    aryFiles = C_File.getFilesByKeyword(folder_path, "exist")
'    Console.Dump (aryFiles)
'    aryFiles = C_File.getFilesByKeyword(folder_path, "exist", True)
'    Console.Dump (aryFiles)
'    '
'''getFilesByWildcard
'    Dim basePath As String
'    basePath = C_File.getSpecialFolder
'    Dim childPath As String
'    childPath = "temp"
'    'Dim folder_path As String
'    'folder_path = C_File.buildPath(basePath, childPath)
'    'Dim aryFiles As Variant
'    '
'    Console.log ("-----------------------getFilesByWildcard")
'    'aryFiles = C_File.getFilesByWildcard(folder_path, "rad*.tmp")
'    'Console.Dump (aryFiles)
'    aryFiles = C_File.getFilesByWildcard("", "rad*.tmp")
'    Console.Dump (aryFiles)
'    aryFiles = C_File.getFilesByWildcard(folder_path, "radf*.tmp")
'    Console.Dump (aryFiles)
'    aryFiles = C_File.getFilesByWildcard(folder_path, "radF*.tmp")
'    Console.Dump (aryFiles)
'''copyFile
'''createFolder
'    Dim basePath As String
'    basePath = C_File.getSpecialFolder
'    Dim childPath As String
'    childPath = "temp"
'    Dim childPath2 As String
'    childPath2 = "temptemp"
'    Dim folder_path As String
'    folder_path = C_File.buildPath(basePath, childPath)
'    Dim folder_path2 As String
'    folder_path2 = C_File.buildPath(basePath, childPath2)
'    Dim aryFiles As Variant
'    aryFiles = C_File.getFilesByWildcard(folder_path, "*.txt")
'    C_File.createFolder folder_path2
'    Console.Dump aryFiles
'    '
'    Dim src_path As String
'    src_path = aryFiles(UBound(aryFiles))
'    Dim src_filename As String
'    src_filename = C_File.getFileName(src_path)
'    Dim dist_path As String
'    dist_path = C_File.buildPath(folder_path2, src_filename)
'    C_File.copyFile src_path, dist_path
'    Dim aryFiles2 As Variant
'    aryFiles2 = C_File.getFilesByWildcard(folder_path2, "*.txt")
'    Console.Dump (aryFiles2)
'''copyFile2
'    Dim basePath As String
'    basePath = C_File.getSpecialFolder
'    Dim childPath As String
'    childPath = "temp"
'    Dim childPath2 As String
'    childPath2 = "temptemp"
'    Dim folder_path As String
'    folder_path = C_File.buildPath(basePath, childPath)
'    Dim folder_path2 As String
'    folder_path2 = C_File.buildPath(basePath, childPath2)
'    '
'    Dim src As String
'    Dim dist As String
'    src = folder_path & "/" & "radF*.tmp"
'    dist = folder_path2 & "/"
'    C_File.copyFile src, dist
    
'''moveFile
'    Dim basePath As String
'    basePath = C_File.getSpecialFolder
'    Dim childPath As String
'    childPath = "temptemptemp"
'    Dim childPath2 As String
'    childPath2 = "temptemptemptemp"
'    Dim folder_path As String
'    folder_path = C_File.buildPath(basePath, childPath)
'    Dim folder_path2 As String
'    folder_path2 = C_File.buildPath(basePath, childPath2)
'    '
'    Dim src As String
'    Dim dist As String
'    src = folder_path & "/" & "radF*.tmp"
'    dist = folder_path2 & "/"
'    C_File.moveFile src, dist
    
'''moveFolder
'    Dim basePath As String
'    basePath = C_File.getSpecialFolder
'    Dim childPath As String
'    childPath = "temptemptemp"
'    Dim childPath2 As String
'    childPath2 = "temptemptemptemp"
'    Dim folder_path As String
'    folder_path = C_File.buildPath(basePath, childPath)
'    Dim folder_path2 As String
'    folder_path2 = "G:\temp2"
'    '
'    Dim src As String
'    Dim dist As String
'    src = folder_path
'    dist = folder_path2
'    C_File.moveFolder src, dist
    
'''deleteFile
'    Dim basePath As String
'    basePath = C_File.getSpecialFolder
'    Dim childPath As String
'    childPath = "temptemp"
'    Dim folder_path As String
'    folder_path = C_File.buildPath(basePath, childPath)
'    '
'    Dim src As String
'    Dim dist As String
'    src = folder_path & "\" & "*.txt"
'    'Console.log src
'    C_File.deleteFile src
'
'    '//--------------------------------------
'    '// end process
'    GoSub FINALLY
'    Exit Sub
    
'''deleteFile
'    Dim basePath As String
'    basePath = C_File.getSpecialFolder
'    Dim childPath As String
'    childPath = "temptemptemptemp"
'    Dim folder_path As String
'    folder_path = C_File.buildPath(basePath, childPath)
'    '
'    Dim src As String
'    Dim dist As String
'    src = folder_path
'    'Console.log src
'    C_File.deleteFolder src
'''createFolders
'    Dim basePath As String
'    basePath = C_File.getSpecialFolder
'    Dim childPath As String
'    childPath = "temptemp"
'    Dim folder_path As String
'    folder_path = C_File.buildPath(basePath, childPath)
'    Dim childPathNext As String
'    childPathNext = "temp"
'    Dim folder_pathNext As String
'    folder_pathNext = C_File.buildPath(folder_path, childPathNext)
'    Console.log folder_pathNext
'    '
'    C_File.createFolders folder_pathNext
'''createNamespace
    Dim namespace As String
    namespace = "jp.sse.vba"
    '
    Call C_File.CreateNamespase("", namespace)

    '//--------------------------------------
    '// end process
    GoSub FINALLY
    Exit Sub
    
FINALLY: 'Label
    '//--------------------------------------
    '// end
    Console.log ("// -----------------// debug end.....")
    Console.log ("")
    Set C_File = Nothing
    Return

EXCEPTION: 'Label
    '//--------------------------------------
    ' exception
    GoSub FINALLY
    Console.log ("// ================== error detail..")
    Console.info ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    Console.log ("")
    
End Sub

