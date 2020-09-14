VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_Console.cls
'''  *<br/>Console Class (Common)
'''  *<br/>Debug.Print on Immediate Window Wrapper
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  */
'
''' /********************************************************
'''  * @class C_Console
'''  * @classdesc Console Classe
'''  *
'''  * @require C_Commons (in GlobalVariable.bas)
'''  * @require If output log file, require O_StringBuilder ,C_File
'''  */
'
'//--------------------------------------------------------
'// Class Property
'//
'
'//--------------------------------------------------------
'// Class Event
'//
Private Sub Class_Initialize()
''' /********************************************************
'''  * @event Class_Initialize
'''  *<br/>Initialize function when creating class
'''  */
On Error Resume Next
#If LANG_EN = 1 Then
    Application.VBE.Windows("immediate").Visible = True
#Else
    Application.VBE.Windows("イミディエイト").Visible = True ''For ja
#End If
End Sub

Private Sub Class_Terminate()
''' /********************************************************
'''  *Class_Terminate
'''  *<br/>Terminate function
'''  */
End Sub


'//--------------------------------------------------------
'// Class Method (Function)
'//
Public Sub log(vnt As Variant)
''' /********************************************************
'''  *@function log
'''  *</br>Console.log
'''  *</br>if Compile Option'DEBUG_MODE' is True, log out
'''  *@param {variant} vnt
'''  */
    '#Const DEBUG_MODE = 0
    #If DEBUG_MODE = 1 Then
        Debug.Print "[DEBUG] " & vnt
    #End If
End Sub

Public Sub info(vnt As Variant)
''' /********************************************************
'''  * @function info
'''  *</br>Console.info
'''  *@param {variant} vnt
'''  */
    Debug.Print vnt
End Sub

Public Sub error(vnt As Variant)
''' /********************************************************
'''  * @function error
'''  *</br>Console.error
'''  *@param {variant} vnt
'''  */
    Debug.Print "<err> " & vnt
End Sub

Public Sub Dump(ByRef vnt As Variant)
''' /********************************************************
'''  * @function dump
'''  *</br>Console.dump
'''  *
'''  *@param {variant} vnt
'''  */
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Console.dump"
On Error GoTo EXCEPTION
    #If DEBUG_MODE = 1 Then
    Debug.Print C_Commons.toString(vnt)
    #End If
    Exit Sub
EXCEPTION:
    Err.Raise 21001, CONS_MODULE_NAME & vbCr & Err.Description, "????????"
End Sub

Public Sub Dir(ByRef vnt As Variant)
''' /********************************************************
'''  * @function dump
'''  *</br>Console.dump
'''  *
'''  *@param {variant} vnt
'''  */
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "C_Console.dir"
On Error GoTo EXCEPTION
    Debug.Print C_Commons.toString(vnt)
    Exit Sub
EXCEPTION:
    Err.Raise 21001, CONS_MODULE_NAME & vbCr & Err.Description, "????????"
End Sub

'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Sub UnitTest()
' /************************************
'  * debug
'  */
    Dim CONS_MODULE_NAME As String
    CONS_MODULE_NAME = "Console"
    '//--------------------------------------
    '// variables
    Dim lng
    lng = 1234
    Dim str
    str = "sak.mit"
    Dim dtm
    dtm = #1/5/2015#
    Dim ary
    ary = Array("sak", "miss", "sakot", "mit")
    Dim dict2
    Set dict2 = createObject("Scripting.Dictionary")
    dict2.Add "sak", "mits"
    dict2.Add "sakot", "miss"
    dict2.Add "ao", "saki"
    Dim dict  As Object
'    Set dict = Dictionary(Array("sak", "mit", "sakot", "miss", "ayo", "saki"))
'    Set dict = Dictionary(Array("sak", "mit", "sakot", "miss", "ayo", "saki", "aka", "gira"))
    
On Error GoTo EXCEPTION
    '//--------------------------------------
    '// initialize
    '//--------------------------------------
    '// do Process
    Console.log ("DEBUG MODE log1")
    Console.log ("DEBUG MODE log2")
    Console.info ("DEBUG MODE info1")
    Console.log ("DEBUG MODE log3")
    'Console.Dump (ObjLogBuff.ToString)
    'Console.info "console.info"
'    Console.Dump (lng)
'    Console.Dump (str)
'    Console.Dump (dtm)
'    Console.Dump (ary)
'
    'Console.Dump dict2
'    'Console.Dump dict
    
    Dim rng As Range
''    Set rng = ThisWorkbook.Names("TEST_RANGE_2").RefersToRange
'    Set rng = ThisWorkbook.ActiveSheet.Range("A1:B2")
'    Console.Dump rng
'
    Dim lstObject As ListObject
'    If ActiveSheet.ListObjects.Count() > 0 Then
'        Set lstObject = ActiveSheet.ListObjects(1)
'        Console.Dump lstObject
'    End If
    Dim i As Integer
    i = 12 / 0

    GoSub FINALLY
    Exit Sub
    
FINALLY: 'Label
    '//--------------------------------------
    '// end process
    Set lstObject = Nothing
    Set rng = Nothing
    Set dict = Nothing
    Set dict2 = Nothing
    Return

EXCEPTION: 'Label
    '//--------------------------------------
    ' catch exception
    GoSub FINALLY
    Console.error ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.error ("Err.Number = " & Err.Number)
    Console.error ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    
End Sub
