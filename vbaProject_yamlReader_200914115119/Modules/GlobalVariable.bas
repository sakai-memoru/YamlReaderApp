Attribute VB_Name = "GlobalVariable"
Option Explicit

'''/**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
''' * @file GlobalVariable_xx.bas
''' * @version 1.00
''' * @since 2020/03/31
''' *
''' * @require Refer Scripting.Dictionary
''' * @require Set vbaproject property "LANG_EN = 1: DEBUG_MODE = 1"
''' * @require C_Commons, C_Console, C_String (not use C_Commons.DumpEx)
''' */
'
'''/** @global @public @name C_Commons */
Public C_Commons As New C_Commons
'''/** @global @public @name Console */
Public Console As New C_Console
'''/** @global @public @name objDSet @description Common Data Set */
Public cdset As New O_DataSet
'''/** @global @public @name Console @description Buffer of Logging  */
Public objLogBuff As New O_StringBuilder

'///////////////////////////////////////////////////////////
'/////  Debug          /////////////////////////////////////
'///////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////
'-----------------------------------------------------------
' debug entry
'-----------------------------------------------------------
Sub helloWorld()
'''' ************************************************
Console.info "//-info--------------------"
Console.info format(Now(), "yyyy/mm/dd hh:mm:ss")
Console.info "Hello, Console.info"
Console.info "//--------------------------"

Console.log "//-log----------------------"
Console.log format(Now(), "yyyy/mm/dd hh:mm:ss")
Console.log "Hello, Console.log"
Console.log "//--------------------------"

''Console’P‘Ì
Call Console.UnitTest

''C_Commons’P‘Ì
Call C_Commons.UnitTest

End Sub

