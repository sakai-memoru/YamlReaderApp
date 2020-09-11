VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_String"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_String.cls
'''  *<br/>String Utility Class
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  */
'
''' /********************************************************
'''  * @class C_String
'''  * @classdesc String Utility
'''  * @require C_Commons, C_Array
'''  */
'
'//--------------------------------------------------------
'// Class Property (Instance Variable)
'//
''' /** @property @public @name objRegExp_ @type {Object:"VBScript.RegExp"} @desc RegExp Object */
Dim objRegExp_ As Object

Private Enum RE_PTN
    DIGIT = 0
    num = 1
    ALPHA = 2
    ASCII = 3
    KANA = 4
    MAIL = 5
    url = 6
    POSTNO = 7
    TEL = 8
    ASCII2 = 9
End Enum

Private aryPATTERN As Variant


'//--------------------------------------------------------
'// Class Event
'//
Private Sub Class_Initialize()
''' /********************************************************
'''  * @event Class_Initialize
'''  *<br/>Initialize 関数
'''  *<br/>Class生成時に、実行される。
'''  */
Set objRegExp_ = createObject("VBScript.RegExp")
Dim RE_PTN_DIGIT As String
RE_PTN_DIGIT = "^[0-9]+$"
Dim RE_PTN_NUMBER As String
RE_PTN_NUMBER = "^[-]?[0-9]+(\.[0-9]+)?$"
Dim RE_PTN_ALPHA As String
RE_PTN_ALPHA = "^[a-zA-Z]+$"
Dim RE_PTN_ASCII As String
RE_PTN_ASCII = "^[a-zA-Z0-9]+$"
Dim RE_PTN_KANA As String
RE_PTN_KANA = "^[ァ-ン]+$"
Dim RE_PTN_MAIL As String
RE_PTN_MAIL = "^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$"
Dim RE_PTN_URL As String
RE_PTN_URL = "^(https?|ftp)(:\/\/[-_.!~*\'()a-zA-Z0-9;\/?:\@&=+\$,%#]+)$"
Dim RE_PTN_POSTNO As String
RE_PTN_POSTNO = "^\d{3}\-\d{4}$"
Dim RE_PTN_TEL As String
RE_PTN_TEL = "^[0-9-]{6,9}$|^[0-9-]{12}$"
'Dim RE_PTN_ASCII2 As String
'RE_PTN_ASCII2= "^[a-zA-Z0-9]+$"
aryPATTERN = Array(RE_PTN_DIGIT _
                , RE_PTN_NUMBER _
                , RE_PTN_ALPHA _
                , RE_PTN_ASCII _
                , RE_PTN_KANA _
                , RE_PTN_MAIL _
                , RE_PTN_URL _
                , RE_PTN_POSTNO _
                , RE_PTN_TEL _
                )
End Sub

Private Sub Class_Terminate()
''' /********************************************************
'''  * @event Class_Terminate
'''  *<br/>Terminate 関数
'''  *<br/>Class廃棄時(Set Nothing)に、実行される。
'''  */
Set objRegExp_ = Nothing
End Sub


'//--------------------------------------------------------
'// Instance Method (Function)
'//
Public Function Contain(ByVal str As String, ByVal strSet As String, Optional ignorecase As Variant = False) As Boolean
''' /********************************************************
'''  * @function contain
'''  *</br>判定したい文字列に、文字パターンが含まれるかを
'''  *</br>判定する
'''  *　@param {string} str 判定したい文字列 複数の文字列にも対応する
'''  *　@param {string} strSet 文字集合
'''  *　@return {boolean} 含むか否か {True :含む,False:含まない}
'''  */
Dim blnRet As Boolean
blnRet = False
Dim aryPTN() As String
ReDim aryPTN(Len(str) - 1)
Dim i As Integer
For i = 1 To Len(str)
        aryPTN(i - 1) = Mid$(str, i, 1)
Next i
'    Console.dump (aryPTN)
objRegExp_.Pattern = strSet
objRegExp_.ignorecase = ignorecase
objRegExp_.Global = True
For i = 1 To Len(str)
        If objRegExp_.Test(aryPTN(i - 1)) Then
            blnRet = True
        End If
Next i
    
Contain = blnRet
End Function

Public Function ContainWord(ByVal strSentence As String, ByVal strWord As String, Optional ignorecase As Boolean = False) As Boolean
''' /********************************************************
'''  * @function containWord
'''  *</br>判定したい文字列(sentence)に、指定文字列(Word)が含まれるかを
'''  *</br>判定する
'''  *　@param {string} strSentence 判定したい文字列
'''  *　@param {string} strWord 文字列
'''  *　@return {boolean} 含むか否か {True :含む,False:含まない}
'''  */
Dim blnRet As Boolean
blnRet = False
objRegExp_.Pattern = strWord
objRegExp_.ignorecase = ignorecase
objRegExp_.Global = True
ContainWord = objRegExp_.Test(strSentence)
End Function

Public Function ContainRegExp(ByVal strSentence As String _
            , ByVal Pattern As String _
            , Optional ignorecase As Variant = False) As Boolean
''' /********************************************************
'''  * @function containRegExp
'''  *</br>判定したい文字列(sentence)に、指定文字列(Word)が含まれるかを
'''  *</br>判定する
'''  *　@param {string} strSentence 判定したい文字列
'''  *　@param {string} pattern RegPattern
'''  *　@return {boolean} 含むか否か {True :含む,False:含まない}
'''  */
Dim blnRet As Boolean
blnRet = False
objRegExp_.Pattern = Pattern
objRegExp_.ignorecase = ignorecase
objRegExp_.Global = True
ContainRegExp = objRegExp_.Test(strSentence)
End Function

Public Function Test(ByVal strSentence As String _
            , ByVal Pattern As String _
            , Optional ignorecase As Variant = False) As Boolean
''' ********************************************************
Test = ContainRegExp(strSentence, Pattern, ignorecase)
End Function

Public Function EqualIgnore(ByVal str1 As String, ByVal str2 As String, Optional ignorecase As Variant = True) As Boolean
'''' ********************************************************
If ignorecase Then
    If VBA.UCase(str1) = VBA.UCase(str2) Then
        EqualIgnore = True
    Else
        EqualIgnore = False
    End If
Else
    If str1 = str2 Then
        EqualIgnore = True
    Else
    EqualIgnore = False
    End If
End If
''
End Function

Public Function IsEmpty(ByVal str As String) As Boolean
''' /********************************************************
'''  * @function isEmpty
'''  *</br>対象が空白か否かを判定する
'''  *</br>
'''  *　@param {string} str 判定したい文字変数
'''  *　@return {boolean} 空欄か否か {True :空欄,False:空欄でない}
'''  */
If str = "" Then
        IsEmpty = True
Else
        IsEmpty = False
End If
End Function

Public Function IsAlphanumeric(ByVal str As String) As Boolean
''' /********************************************************
'''  * @function isAlphanumeric
'''  *</br>対象がアルファベットもしくは数字で構成されているか
'''  *</br>否かを判定する。
'''  *　@param {string} str 判定したい文字列 複数の文字列にも対応する
'''  *　@return {boolean}{ True : OK, False : NG }
'''  */
Dim blnRet As Boolean
blnRet = True
Dim aryPTN() As String
ReDim aryPTN(Len(str) - 1)
Dim i As Integer
For i = 1 To Len(str)
        aryPTN(i - 1) = Mid$(str, i, 1)
Next i
'    Console.dump (aryPTN)
objRegExp_.Pattern = "^[a-zA-Z0-9]+$"
objRegExp_.ignorecase = False
objRegExp_.Global = True
For i = 1 To Len(str)
        If objRegExp_.Test(aryPTN(i - 1)) Then
            '
        Else
            blnRet = False
        End If
Next i
'
IsAlphanumeric = blnRet
''
End Function

Public Function IsAlphanumericSpace(ByVal str As String) As Boolean
''' /********************************************************
'''  * @function isAlphanumericSpace
'''  *</br>対象がアルファベットもしくは数字で構成されているか
'''  *</br>否かを判定する。
'''  *　@param {string} str 判定したい文字列 複数の文字列にも対応する
'''  *　@return {boolean}{ True : OK, False : NG }
'''  */
Dim blnRet As Boolean
blnRet = True
Dim aryPTN() As String
ReDim aryPTN(Len(str) - 1)
Dim i As Integer
For i = 1 To Len(str)
        aryPTN(i - 1) = Mid$(str, i, 1)
Next i
'    Console.dump (aryPTN)
objRegExp_.Pattern = "^[a-zA-Z0-9 ,\.]+$"
objRegExp_.ignorecase = False
objRegExp_.Global = True
For i = 1 To Len(str)
        If objRegExp_.Test(aryPTN(i - 1)) Then
            '
        Else
            blnRet = False
        End If
Next i
IsAlphanumericSpace = blnRet
''
End Function

Public Function IsAlphaSpace(ByVal str As String) As Boolean
''' /********************************************************
'''  * @function isAlphaSpace
'''  *</br>対象がアルファベットもしくは数字で構成されているか
'''  *</br>否かを判定する。
'''  *　@param {string} str 判定したい文字列 複数の文字列にも対応する
'''  *　@return {boolean}{ True : OK, False : NG }
'''  */
Dim blnRet As Boolean
blnRet = True
Dim aryPTN() As String
ReDim aryPTN(Len(str) - 1)
Dim i As Integer
For i = 1 To Len(str)
        aryPTN(i - 1) = Mid$(str, i, 1)
Next i
'    Console.dump (aryPTN)
objRegExp_.Pattern = "^[a-zA-Z ,\.]+$"
objRegExp_.ignorecase = False
objRegExp_.Global = True
For i = 1 To Len(str)
        If objRegExp_.Test(aryPTN(i - 1)) Then
            '
        Else
            blnRet = False
        End If
Next i
IsAlphaSpace = blnRet
''
End Function

Public Function IsEmail(ByVal vnt As Variant) As Boolean
''' /********************************************************
'''  *@function IsNumber
'''  * FIXME
'''  */
''
Let IsEmail = Me.Test(vnt, aryPATTERN(RE_PTN.MAIL))
End Function

Public Function IsUrl(ByVal vnt As Variant) As Boolean
''' /********************************************************
'''  *@function IsNumber
'''  * FIXME
'''  */
''
Let IsUrl = Me.Test(vnt, aryPATTERN(RE_PTN.url))
End Function

Public Function IsTel(ByVal vnt As Variant) As Boolean
''' /********************************************************
'''  *@function IsNumber
'''  * FIXME
'''  */
''
Let IsTel = Me.Test(vnt, aryPATTERN(RE_PTN.TEL))
End Function

Public Function IsPostCode(ByVal vnt As Variant) As Boolean
''' /********************************************************
'''  *@function IsNumber
'''  * FIXME
'''  */
''
Let IsPostCode = Me.Test(vnt, aryPATTERN(RE_PTN.POSTNO))
End Function

Public Function IsKana(ByVal vnt As Variant) As Boolean
''' /********************************************************
'''  *@function IsNumber
'''  * FIXME
'''  */
''
Let IsKana = Me.Test(vnt, aryPATTERN(RE_PTN.KANA))
End Function

Public Function IsAddress(ByVal vnt As Variant) As Boolean
''' /********************************************************
'''  *@function IsNumber
'''  * FIXME
'''  */
''

Dim ptn As String
Let ptn = "^[a-zA-Z]{1,3}[0-9]{1,7}$"
Dim bln As Boolean
Let bln = Me.Test(vnt, ptn)
On Error Resume Next
Dim rng As Range
If bln Then
    Set rng = Range(vnt)
End If
If Err.Number <> 0 Then
    Let bln = False
End If
Let IsAddress = bln
End Function


Public Function IsZenkaku(chr As String) As Boolean
''' /********************************************************
'''  *@function Is2Byte
'''  *<pre>２バイト文字を含むか否かを判定
'''  *</pre>
'''  *　@param {String} chr 対象文字
'''  *　@return {Boolean} Is2Byte
'''  */
IsZenkaku = Not (Len(chr) = Me.Length(chr))
End Function

Public Function IsNumber(ByVal vnt As Variant) As Boolean
''' /********************************************************
'''  *@function IsNumber
'''  * FIXME
'''  */
''
Dim blnRtn As Boolean
If (VBA.VarType(vnt) = VBA.vbInteger) Then
    Let blnRtn = True
ElseIf (VBA.VarType(vnt) = VBA.vbSingle) Then
    Let blnRtn = True
ElseIf (VBA.VarType(vnt) = VBA.vbLong) Then
    Let blnRtn = True
ElseIf (VBA.VarType(vnt) = VBA.vbDouble) Then
    Let blnRtn = True
ElseIf (VBA.VarType(vnt) = VBA.vbCurrency) Then
    Let blnRtn = True
ElseIf (VBA.VarType(vnt) = VBA.vbVariant) Then
    If (Me.IsNumeric2(vnt)) Then
        Let blnRtn = Me.Test(vnt, aryPATTERN(RE_PTN.num))
    Else
        Let blnRtn = False
    End If
ElseIf (VBA.VarType(vnt) = VBA.vbString) Then
    If (Me.IsNumeric2(vnt)) Then
        Let blnRtn = Me.Test(vnt, aryPATTERN(RE_PTN.num))
    Else
        Let blnRtn = False
    End If
Else
    Let blnRtn = False
End If
Let IsNumber = blnRtn
End Function


Public Function IsNumeric(ByVal str As String) As Boolean
''' /********************************************************
'''  * @function isNumeric
'''  *</br>対象が数字で構成されているか
'''  *</br>否かを判定する。
'''  *　@param {string} str 判定したい文字列 複数の文字列にも対応する
'''  *　@return {boolean}{ True : OK, False : NG }
'''  */
Dim blnRet As Boolean
blnRet = True
Dim aryPTN() As String
ReDim aryPTN(Len(str) - 1)
Dim i As Integer
For i = 1 To Len(str)
        aryPTN(i - 1) = Mid$(str, i, 1)
Next i
'    Console.dump (aryPTN)
objRegExp_.Pattern = "^[0123456789]+$"
objRegExp_.ignorecase = False
objRegExp_.Global = True
For i = 1 To Len(str)
        If objRegExp_.Test(aryPTN(i - 1)) Then
            '
        Else
            blnRet = False
        End If
Next i
IsNumeric = blnRet
''
End Function

Public Function IsNumeric2(ByVal str As String) As Boolean
''' /********************************************************
'''  * @function isNumeric
'''  *</br>対象がアルファベットもしくは数字で構成されているか
'''  *</br>否かを判定する。
'''  *　@param {string} str 判定したい文字列 複数の文字列にも対応する
'''  *　@return {boolean}{ True : OK, False : NG }
'''  */
Dim blnRet As Boolean
blnRet = True
Dim aryPTN() As String
ReDim aryPTN(Len(str) - 1)
Dim i As Integer
For i = 1 To Len(str)
    aryPTN(i - 1) = Mid$(str, i, 1)
Next i
'    Console.dump (aryPTN)
objRegExp_.Pattern = "^[-0123456789.]+$"
objRegExp_.ignorecase = False
objRegExp_.Global = True
For i = 1 To Len(str)
    If objRegExp_.Test(aryPTN(i - 1)) Then
        '
    Else
        blnRet = False
    End If
Next i
IsNumeric2 = blnRet
''
End Function

Public Function IsNumericSpace(ByVal str As String) As Boolean
''' /********************************************************
'''  * @function isNumericSpace
'''  *</br>対象が数字もしくは空白で構成されているか
'''  *</br>否かを判定する。
'''  *　@param {string} str 判定したい文字列 複数の文字列にも対応する
'''  *　@return {boolean}{ True : OK, False : NG }
'''  */
Dim blnRet As Boolean
blnRet = True
Dim aryPTN() As String
ReDim aryPTN(Len(str) - 1)
Dim i As Integer
For i = 1 To Len(str)
        aryPTN(i - 1) = Mid$(str, i, 1)
Next i
'    Console.dump (aryPTN)
objRegExp_.Pattern = "^[0123456789 ]+$"
objRegExp_.ignorecase = False
objRegExp_.Global = True
For i = 1 To Len(str)
        If objRegExp_.Test(aryPTN(i - 1)) Then
            '
        Else
            blnRet = False
        End If
Next i
IsNumericSpace = blnRet
''
End Function

Public Function IsAlpha(ByVal str As String) As Boolean
''' /********************************************************
'''  * @function isAlpha
'''  *</br>対象がアルファベットもしくは数字で構成されているか
'''  *</br>否かを判定する。
'''  *　@param {string} str 判定したい文字列 複数の文字列にも対応する
'''  *　@return {boolean}{ True : OK, False : NG }
'''  */
Dim blnRet As Boolean
blnRet = True
Dim aryPTN() As String
ReDim aryPTN(Len(str) - 1)
Dim i As Integer
For i = 1 To Len(str)
        aryPTN(i - 1) = Mid$(str, i, 1)
Next i
'    Console.dump (aryPTN)
objRegExp_.Pattern = "^[a-zA-Z]+$"
objRegExp_.ignorecase = False
objRegExp_.Global = True
For i = 1 To Len(str)
        If objRegExp_.Test(aryPTN(i - 1)) Then
            '
        Else
            blnRet = False
        End If
Next i
IsAlpha = blnRet
End Function

Public Function IsString(ByVal vnt As Variant) As Boolean
''' /********************************************************
'''  *@function IsString
'''  */
Dim blnRtn As Boolean
If (VBA.VarType(vnt) = VBA.vbString) Then
    Let blnRtn = True
ElseIf (VBA.VarType(vnt) = VBA.vbVariant) Then
    Let blnRtn = Not Me.IsNumeric(vnt)
Else
    Let blnRtn = False
End If
Let IsString = blnRtn
End Function

Public Function IsJsonString(ByVal str As String) As Boolean
''' /********************************************************
'''  *@function isJsonObject
'''  *<pre>文字列が、JSONのObject形式かを判定する。
'''  *※開始と終了が、{} or []で括られている場合、Object形式と判定する。
'''  *</pre>
'''  *　@param {string} str 判定したい文字列 複数の文字列にも対応する
'''  *　@return {boolean}{ True : OK, False : NG }
'''  */
Dim blnRet As Boolean
If Me.StartWith(Me.TrimSpace(str)) = "{" And Me.EndWith(Me.TrimSpace(str)) = "}" Then
    blnRet = True
ElseIf Me.StartWith(Me.TrimSpace(str)) = "[" And Me.EndWith(Me.TrimSpace(str)) = "]" Then
        blnRet = True
Else
        blnRet = False
End If
IsJsonString = blnRet
End Function

Public Function SqueezeSpace(ByVal str As String) As String
''' /********************************************************
'''  * @function squeezeSpace
'''  *</br>文字列内の連続するSpace文字を、1つの半角Spaceにする。
'''  *</br>
'''  *　@param {string} str 変換したい文字列
'''  *　@return {String} 変換後の文字列
'''  */
Dim strTemp As String
objRegExp_.Pattern = "[　\t]"
objRegExp_.ignorecase = False
objRegExp_.Global = True
strTemp = objRegExp_.Replace(str, " ")

objRegExp_.Pattern = "\s{2,}"
objRegExp_.ignorecase = False
objRegExp_.Global = True
SqueezeSpace = objRegExp_.Replace(strTemp, " ")
End Function

Public Function DeleteSpace(ByVal str As String) As String
''' /********************************************************
'''  * @function deleteWhiteSpace
'''  *</br>文字列内のSpace文字を、削除する。
'''  *</br> 全角スペース、半角スペース
'''  *</br>
'''  *　@param {string} str 変換したい文字列
'''  *　@return {String} 変換後の文字列
'''  */
objRegExp_.Pattern = "[ 　]"
objRegExp_.ignorecase = False
objRegExp_.Global = True
DeleteSpace = objRegExp_.Replace(str, "")
End Function

Public Function DeleteWhiteSpace(ByVal str As String) As String
''' /********************************************************
'''  * @function deleteWhiteSpace
'''  *</br>文字列内のSpace文字を、削除する。
'''  *</br> 半角スペース、\t、\n、\r、\f
'''  *</br>
'''  *　@param {string} str 変換したい文字列
'''  *　@return {String} 変換後の文字列
'''  */
Dim temp As String
objRegExp_.Pattern = "\s+"
objRegExp_.ignorecase = False
objRegExp_.Global = True
temp = objRegExp_.Replace(str, "")
objRegExp_.Pattern = "　"
DeleteWhiteSpace = objRegExp_.Replace(temp, "")
End Function

Public Function Strip(ByVal str As String, Optional strDel As String = vbLf) As String
''' /********************************************************
'''  * @function strip
'''  *</br>文字列内の指定文字を、削除する。
'''  *</br>
'''  *　@param {string} str 変換したい文字列
'''  *　@param {string} strDel 削除したい文字　Default=vbLf
'''  *　@return {String} 変換後の文字列
'''  */
Dim strBuff As String
strBuff = strDel & "+"
objRegExp_.Pattern = strBuff
objRegExp_.ignorecase = False
objRegExp_.Global = True
Strip = objRegExp_.Replace(str, "")
End Function

Public Function StripEx(ByVal str As String, regpattern As String) As String
''' /********************************************************
'''  * @function stripEx
'''  *</br>文字列内の指定文字のセットを、削除する。
'''  *</br>
'''  *　@param {string} str 変換したい文字列
'''  *　@param {string} strSet 削除したい正規表現の文字セット（例：[ABC]）
'''  *　@return {String} 変換後の文字列
'''  */
Dim strBuff As String
objRegExp_.Pattern = regpattern
objRegExp_.ignorecase = False
objRegExp_.Global = True
StripEx = objRegExp_.Replace(str, "")
End Function

Public Function StripCrLf(ByVal str As String) As String
''' /********************************************************
'''  * @function stripCrLf
'''  *</br>文字列内の改行文字を、削除する。
'''  *</br>
'''  *　@param {string} str 変換したい文字列
'''  *　@param {string} strSet 削除したい正規表現の文字セット（例：[ABC]）
'''  *　@return {String} 変換後の文字列
'''  */
StripCrLf = Me.StripEx(str, "[\n\r\f]")
End Function

Public Function Clean(ByVal str As String) As String
''' /********************************************************
'''  * @function Clean
'''  * </br> Removes all nonprintable characters from text.
'''  */
Clean = WorksheetFunction.Clean(str)
End Function

Public Function RandamString(Optional intLength As Integer = 8, Optional blnIncUpper As Boolean = True, Optional blnIncSign As Boolean = True)
''' /********************************************************
'''  * @function randamString
'''  *</br>randamな文字列を生成する
'''  *</br>
'''  *　@param {integer} intLength 【Optional】生成する文字数 規定値:8
'''  *　@param {boolean} blnIncUpper 【Optional】英大文字を含むか否か　規定値:True
'''  *　@param {boolean} blnIncSign 【Optional】記号文字を含むか否か　規定値:True
'''  *　@return {String} 変換後の文字列
'''  */
Dim strRet As String
Dim strLower As String
Let strLower = "abcdefghijklmnopqrstuvwxyz"
Dim strUpper As String
Let strUpper = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Dim strNumber As String
Let strNumber = "0123456789"
Dim strSign As String
Let strSign = "-^@[;],!#$%&'()=~`{+}<>_"   'ファイル名に使えない記号 /\?*:|".
Dim strBuff As String
Dim intMax As Integer
Dim i As Integer
    
strBuff = strLower & strNumber
If blnIncUpper = True Then
        strBuff = strBuff & strUpper
End If
If blnIncSign = True Then
        strBuff = strBuff & strSign
End If
''
intMax = Len(strBuff)
For i = 1 To intLength
        strRet = strRet & Mid(strBuff, Int((intMax - 1 + 1) * Rnd + 1), 1)
Next i
RandamString = strRet
End Function

Public Function MatchWith(ByVal str As String, ByVal match_str As String) As Boolean
'''' ********************************************************
''
Dim start_str As String
Let start_str = Me.StartWith(str, Len(match_str))
Let MatchWith = Me.EqualIgnore(start_str, match_str, ignorecase:=False)
''
End Function


Public Function StartWith(ByVal str As String, Optional lngNum As Long = 1) As String
''' /********************************************************
'''  * @function startWith
'''  *</br>文字列内の最初の文字を返却する。
'''  *</br>
'''  *　@param {string} str 文字列
'''  *　@param {Long} lngNum 【Optional】文字数 Default = 1
'''  *　@return {String} 最初の文字
'''  */
StartWith = left(str, lngNum)
End Function

Public Function EndWith(ByVal str As String, Optional lngNum As Long = 1) As String
''' /********************************************************
'''  * @function endWith
'''  *</br>文字列内の最後の文字を返却する。
'''  *</br>
'''  *　@param {string} str 文字列
'''  *　@param {Long} lngNum 【Optional】文字数 Default = 1
'''  *　@return {String} 最後の文字
'''  */
EndWith = Right(str, lngNum)
End Function

Public Function StartsWith(ByVal str As String, ByVal substr As String) As Boolean
''' /********************************************************
'''  * @function startsWith
'''  *</br>
'''  *</br>
'''  * @param {String} str
'''  * @param {String} substr
'''  * @return {Boolean}
'''  */
'
''' *********************************************************
Dim buff As String
buff = VBA.left(str, VBA.Len(substr))
If buff = substr Then
    StartsWith = True
Else
    StartsWith = False
End If
''' *********************************************************
End Function

Public Function EndsWith(ByVal str As String, ByVal substr As String) As Boolean
''' /********************************************************
'''  * @function endsWith
'''  *</br>
'''  *</br>
'''  * @param {String} str
'''  * @param {String} substr
'''  * @return {Boolean}
'''  */
'
''' *********************************************************
Dim buff As String
buff = VBA.Right(str, VBA.Len(substr))
If buff = substr Then
    EndsWith = True
Else
    EndsWith = False
End If
''' *********************************************************
End Function

Public Function TrimSpace(ByVal str As String) As String
'''' /********************************************************
''''  * @function trimSpace
''''  *</br>文字列の前後の半角・全角空白文字をカットする
''''  *</br> Trim関数のRapper
''''  *　@param {string} str 文字列
''''  *　@return {String} 変換後の文字列
''''  */
Dim strRtn As String
Let strRtn = str
Let strRtn = Me.TrimLeftSpace(strRtn)
Let strRtn = Me.TrimRightSpace(strRtn)
Let TrimSpace = strRtn
''
End Function

Public Function TrimRightSpace(ByVal str As String) As String
'''' /********************************************************
''''  * @function trimLeftSpace
''''  *<pre>文字列の後ろの半角・全角空白文字をカットする
''''  *</pre>
''''  *　@param {string} str 文字列
''''  *　@return {String} 変換後の文字列
''''  */
Dim temp As String
objRegExp_.Pattern = "[\s　]+$"
objRegExp_.ignorecase = False
objRegExp_.Global = True
TrimRightSpace = objRegExp_.Replace(str, "")
''
End Function

Public Function TrimLeftSpace(ByVal str As String) As String
'''' /********************************************************
''''  * @function trimLeftSpace
''''  *<pre>文字列の前の半角・全角空白文字をカットする
''''  *</pre>
''''  *　@param {string} str 文字列
''''  *　@return {String} 変換後の文字列
''''  */
Dim temp As String
objRegExp_.Pattern = "^[\s　]+"
objRegExp_.ignorecase = False
objRegExp_.Global = True
TrimLeftSpace = objRegExp_.Replace(str, "")
''
End Function

Public Function GetLeftIndent(ByVal str As String) As String
'''' /********************************************************
''''  * @function trimLeftSpace
''''  *<pre>文字列の後ろの半角・全角空白文字をカットする
''''  *</pre>
''''  *　@param {string} str 文字列
''''  *　@return {String} 変換後の文字列
''''  */
objRegExp_.Pattern = "^([\s]+)"
objRegExp_.ignorecase = False
objRegExp_.Global = False
Dim matches As MatchCollection
Set matches = objRegExp_.Execute(str)
If matches.Count = 0 Then
    Let GetLeftIndent = ""
Else
    Let GetLeftIndent = matches.Item(0).Value
End If
''
End Function

Public Function GetIndentAndElement(ByVal str As String, Optional ByVal num As Variant) As Variant
'''' ********************************************************
'''' * @function GetIndentAndElement
'''' *
If VBA.IsMissing(num) Then
    Let num = 4
End If
Dim ary As Variant
Dim indent As String
Dim level As Long
Dim elm As String
Let indent = Me.GetLeftIndent(str)
Let level = VBA.Len(indent) / num
Let elm = Me.TrimLeftSpace(str)
Let ary = Array(level, indent, elm)
Let GetIndentAndElement = ary
''
End Function

Public Function ToUpper(ByVal str As String) As String
''' /********************************************************
'''  * @function toUpper
'''  *</br>文字列を英大文字に変換する
'''  *</br> UCase関数のRapper
'''  *　@param {string} str 文字列
'''  *　@return {String} 変換後の文字列
'''  */
ToUpper = UCase(str)
End Function

Public Function ToLower(ByVal str As String) As String
''' /********************************************************
'''  * @function toLower
'''  *</br>文字列を英小文字に変換する
'''  *</br> LCase関数のRapper
'''  *　@param {string} str 文字列
'''  *　@return {String} 変換後の文字列
'''  */
ToLower = LCase(str)
End Function

Public Function HankakuKanaToZen(ByVal str As String) As String
'''' /********************************************************
''''  * @function HankakuKanaToZen
''''  * FIXME: Not be able to use WorksheetFunction.JIS in English Language mode
''''  */
Dim strRtn As String
Dim strTemp As String
Dim i As Long
Dim s As Variant
For i = 1 To Len(str)
    Let s = Mid(str, i, 1)
    '' from 161 to 223
    If asc(s) >= 161 And asc(s) <= 223 Then
        Let strTemp = strTemp & s
    Else
        If strTemp <> "" Then
            strRtn = strRtn & StrConv(strTemp, vbWide)
        End If
    End If
    strRtn = strRtn & s
Next i
If strTemp <> "" Then
    strRtn = strRtn & StrConv(strTemp, vbWide)
End If
HankakuKanaToZen = strRtn
End Function

Public Function Length(ByVal str As String) As Long
''' /********************************************************
'''  * @function length
'''  *</br>２バイト文字列は2として、1バイト文字は１で、長さを返却する
'''  *</br> Len関数のRapper
'''  *　@param {string} str 文字列
'''  *　@return {Long}文字列の長さ
'''  *　FIXME□□単体未
'''  */
    Length = LenB(StrConv(str, vbFromUnicode))
End Function

Public Function Chop(ByVal str As String, Optional lngNum As Long = 2) As String
''' /********************************************************
'''  * @function chop
'''  *</br>文字列の後ろを切り取る
'''  *</br> Mid関数のRapper
'''  *　@param {string} str 文字列
'''  *　@param {Long} lngNum 【Optional】文字数 Default = 1
'''  *　@return {String}変換後の文字列
'''  */
Chop = Mid(str, 1, Len(str) - lngNum)
End Function

Public Function Chomp(ByVal str As String, Optional lngNum As Long = 2) As String
''' /********************************************************
'''  * @function chomp
'''  *</br>文字列の前を切り取る
'''  *</br> Mid関数のRapper
'''  *　@param {string} str 文字列
'''  *　@param {Long} lngNum 【Optional】文字数 Default = 1
'''  *　@return {String}変換後の文字列
'''  */
Chomp = Mid(str, lngNum + 1, Len(str) - lngNum + 1)
End Function

Public Function ChopDouble(ByVal str As String, Optional lngNum As Long = 2) As String
''' /********************************************************
'''  * @function chopDouble
'''  *</br>文字列の前後１文字を切り取る
'''  *</br> Mid関数のRapper
'''  *　@param {string} str 文字列
'''  *　@param {Long} lngNum 【Optional】文字数 Default = 1
'''  *　@return {String}変換後の文字列
'''  */
Dim strBuff As String
strBuff = Me.Chomp(str, lngNum)
ChopDouble = Me.Chop(strBuff, lngNum)
End Function

Public Function CutPrefix(ByVal str As String, ByVal prefix As String) As String
''' ********************************************************
''
If Me.StartsWith(str, prefix) Then
    Let CutPrefix = Me.Chomp(str, Len(prefix))
Else
    Let CutPrefix = str
End If
''
End Function

Public Function Abbreviate(ByVal str As String, Optional lngNum As Long = 15) As String
''' /********************************************************
'''  * @function abbreviate
'''  *</br>文字列の後ろを切り取り、後ろに、"..."を付ける
'''  *</br>
'''  *　@param {string} str 文字列
'''  *　@param {Long} lngNum 【Optional】文字数 Default = 15
'''  *　@return {String}変換後の文字列
'''  */
Abbreviate = Mid(str, 1, Len(str) - lngNum - 3) & "..."
''
End Function

Public Function Quote(ByVal str As String, Optional strQuote As Variant = vbTab) As String
''' /********************************************************
'''  * @function quote
'''  *</br>引用符をつける
'''  *</br>
'''  *　@param {string} str 文字列
'''  *　@param {Variant<string>} strQuate 引用符文字列
'''  *　@return {String}変換後の文字列
'''  */
objRegExp_.Pattern = "^"
Quote = objRegExp_.Replace(str, strQuote)
End Function

Public Function Quotes(ByVal str As String _
                    , Optional strQuote As Variant = vbTab _
                    , Optional delimiter As Variant = vbCrLf) As String
''' /********************************************************
'''  *@function quotes
'''  *</br>行に対して引用符をつける
'''  *</br>
'''  *　@param {string} str 文字列
'''  *　@param {string} strQuate 引用符文字列
'''  *　@return {String}変換後の文字列
'''  */
Dim aryPTNTemp As Variant
aryPTNTemp = Split(str, vbCrLf)
objRegExp_.Pattern = "^"
Dim r As Long
For r = LBound(aryPTNTemp) To UBound(aryPTNTemp)
        aryPTNTemp(r) = objRegExp_.Replace(aryPTNTemp(r), strQuote)
Next r
Quotes = Join(aryPTNTemp, delimiter)
End Function

Public Function IsQuoted(ByVal str As String _
                    , Optional strQuote As String = """") As Boolean
''' /********************************************************
'''  *@function isQuoted
'''  *</br>行に対して引用符がついているか判定
'''  *</br>
'''  *　@param {string} str 文字列
'''  *　@param {string} strQuate 引用符文字列
'''  *　@return {Boolean} isQuoted
'''  */
Dim blnRtn As Boolean
Dim cnt As Long
cnt = Len(strQuote)
If Me.StartWith(str, cnt) = strQuote Then
        blnRtn = True
Else
        blnRtn = False
End If
IsQuoted = blnRtn
End Function

Public Function indent(ByVal str As String _
                    , Optional strIndent As Variant = vbTab _
                    , Optional lngNum As Variant = 1) As String
''' /********************************************************
'''  *@function indent
'''  *</br>インデントをつける
'''  *</br>
'''  *　@param {string} str 文字列
'''  *　@param {Variant<String>} strIndent インデント文字列
'''  *　@param {Variant<Long>} lngNum 繰り返し数
'''  *　@return {String}変換後の文字列
'''  */
Dim i As Long
Dim strTemp As String
For i = 1 To lngNum
        strTemp = strTemp & strIndent
Next i
objRegExp_.Pattern = "^"
indent = objRegExp_.Replace(str, strTemp)
End Function

Public Function Indents(ByVal str As String _
                    , Optional strIndent As Variant = vbTab _
                    , Optional lngNum As Variant = 1 _
                    , Optional delimiter As Variant = vbCrLf) As String
''' /********************************************************
'''  *@function indent
'''  *<pre>行に対してインデントをつける
'''  *</pre>
'''  *　@param {string} str 文字列
'''  *　@param {Variant<String>} strIndent インデント文字列
'''  *　@param {Variant<Long>} lngNum 繰り返し数
'''  *　@param {Variant<String>} delimiter 行区切り文字
'''  *　@return {String}変換後の文字列
'''  */
Dim i As Long
Dim strTemp As String
For i = 1 To lngNum
        strTemp = strTemp & strIndent
Next i
Dim aryPTNTemp As Variant
aryPTNTemp = Split(str, vbCrLf)

objRegExp_.Pattern = "^"
Dim r As Long
For r = LBound(aryPTNTemp) To UBound(aryPTNTemp)
        aryPTNTemp(r) = objRegExp_.Replace(aryPTNTemp(r), strTemp)
Next r
Indents = Join(aryPTNTemp, delimiter)
''
End Function

Public Function Quotation(ByVal str As String, Optional strQuotation As Variant = """") As String
''' /********************************************************
'''  * @function quotation
'''  *</br>引用符で囲む
'''  *</br>
'''  *　@param {string} str 文字列
'''  *　@param {Variant<String>} strQuotation 引用符文字列
'''  *　@return {String}変換後の文字列
'''  */
Quotation = strQuotation & str & strQuotation
End Function

Public Function IsQuotation(ByVal str As String, Optional strQuotation As Variant = """") As Boolean
''' /********************************************************
'''  * @function isQuotation
'''  *</br>引用符で囲まれているか判定
'''  *</br>
'''  *　@param {string} str 文字列
'''  *　@param {Variant<String>} strQuotation 引用符文字列
'''  *　@return {Boolean} Boolean
'''  */
Dim blnRtn As Boolean
Dim cnt As Long
cnt = Len(strQuotation)
If (Me.StartWith(Me.TrimSpace(str), cnt) = strQuotation) And (Me.EndWith(Me.TrimSpace(str), cnt) = strQuotation) Then
        blnRtn = True
Else
        blnRtn = False
End If
IsQuotation = blnRtn
End Function

Public Function FormatString(ByVal str As String, ByVal strFormat As String) As String
''' /********************************************************
'''  * @function formatString
'''  *</br>文字列をフォーマットする
'''  *</br>　Format関数のRapper
'''  *</br>文字列書式記号：
'''  *</br>@ : 任意の一文字。存在しない場合はSpace
'''  *</br>& : 任意の一文字。存在しない場合はSpaceを返さない。
'''  *</br>! : 左から詰める
'''  *　@param {string} str 文字列
'''  *　@param {string} strFormat フォーマット定義文字列
'''  *　@return {String}変換後の文字列
'''  */
FormatString = format(str, strFormat)
End Function

Public Function ToCharArray(ByVal str As String) As Variant
'''' ********************************************************
''
Dim ary() As Variant
ReDim ary(1 To VBA.Len(str))
Dim i As Long
For i = LBound(ary) To UBound(ary)
    Let ary(i) = VBA.Mid(str, i, 1)
Next i
Let ToCharArray = ary
''
End Function

Public Function Repeat(ByVal str As String, lngNum As Long) As String
''' /********************************************************
'''  *@function repeat
'''  *</br>連続する文字列を生成する
'''  *</br>
'''  *　@param {string} str 文字列
'''  *　@param {Long} lngNum 文字数
'''  *　@return {String}変換後の文字列
'''  */
Dim aryPTN() As String
ReDim aryPTN(lngNum - 1)
Dim i As Integer
    
For i = 0 To lngNum - 1
        aryPTN(i) = str
Next
Repeat = Join(aryPTN(), "")
End Function

Public Function Rept(ByVal str As String, lngNum As Long) As String
''' /********************************************************
'''  *@function Rept
'''  */
Rept = WorksheetFunction.Rept(str, lngNum)
End Function

Public Function Spaces(ByVal lngNum As Long) As String
''' /********************************************************
'''  * @function spaces
'''  *</br>連続するSpace文字列を生成する
'''  *</br>
'''  *　@param {Long} lngNum 文字数
'''  *　@return {String}文字列
'''  */
Spaces = Me.Repeat(" ", lngNum)
End Function

Public Function RightPad(str As String, lngNum As Long, Optional strPad As String = " ") As String
''' /********************************************************
'''  * @function rightPad
'''  *</br>指定文字数で、文字列を右寄せする
'''  *</br>
'''  *　@param {string} str 文字列
'''  *　@param {Long} lngNum 文字数
'''  *　@param {String} strPad 【Optional】パッドする文字 default=半角Space
'''  *　@return {String}変換後の文字列
'''  */
RightPad = Me.Repeat(strPad, lngNum - Len(str)) & str
End Function

Public Function LeftPad(str As String, lngNum As Long, Optional strPad As String = " ") As String
''' /********************************************************
'''  * @function leftPad
'''  *</br>指定文字数で、文字列を左寄せする
'''  *</br>
'''  *　@param {string} str 文字列
'''  *　@param {Long} lngNum 文字数
'''  *　@param {String} 【Optional】strPad パッドする文字 default=半角Space
'''  *　@return {String}変換後の文字列
'''  */
LeftPad = str & Me.Repeat(strPad, lngNum - Len(str))
End Function

Public Function CapitalizeFully(str As String, Optional ByVal strSplit As String = " ") As String
''' /********************************************************
'''  * @function capitalizeFully
'''  *</br>文字列の文で、空白の後は、大文字とする
'''  *</br>
'''  *　@param {string} str 文字列
'''  *　@param {string} strSplit 返却時の結合区切り文字
'''  *　@return {String}変換後の文字列
'''  */
Dim aryPTN As Variant 'variant型で受ける
aryPTN = Split(Me.SqueezeSpace(Trim(str)), " ")
'    Console.dump (aryPTN)
Dim aryPTNRet() As String
ReDim aryPTNRet(UBound(aryPTN))
Dim i As Integer
For i = 0 To UBound(aryPTN)
    aryPTNRet(i) = UCase(aryPTN(i))
Next i
CapitalizeFully = Join(aryPTNRet, strSplit)
End Function

Public Function Capitalize(str As String, Optional strSplit As String = " ") As String
''' /********************************************************
'''  * @function capitalize
'''  *</br>文字列の文で、空白の後は、大文字とする
'''  *</br>（大文字は小文字に変換）
'''  *　@param {string} str 文字列
'''  *　@param {string} strSplit 返却時の結合区切り文字
'''  *　@return {String}変換後の文字列
'''  */
Dim aryPTN As Variant 'variant型で受ける
aryPTN = Split(Me.SqueezeSpace(Trim(LCase(str))), " ")
'    Console.dump (aryPTN)
Dim aryPTNRet() As String
ReDim aryPTNRet(UBound(aryPTN))
Dim i As Integer
For i = 0 To UBound(aryPTN)
    aryPTNRet(i) = UCase(Me.StartWith(aryPTN(i))) & Me.Chomp(aryPTN(i), 1)
Next i
Capitalize = Join(aryPTNRet, strSplit)
End Function

Public Function Camelize(str As String, Optional strSplit As Variant = " ", _
                        Optional blnJoinStr As Variant) As String
''' /********************************************************
'''  * @function camelize
'''  *</br>文字数の文で、空白の後は、大文字とする
'''  *</br>
'''  *　@param {string} str 文字列
'''  *　@param {string} strSplit 返却時の結合区切り文字
'''  *　@return {String}変換後の文字列
'''  */
Dim aryPTN As Variant 'variant型で受ける
aryPTN = Split(Me.SqueezeSpace(Trim(LCase(str))), strSplit)
'    Console.dump (aryPTN)
Dim aryPTNRet() As String
ReDim aryPTNRet(UBound(aryPTN))
Dim i As Integer
For i = 0 To UBound(aryPTN)
        If i = 0 Then
        aryPTNRet(i) = LCase(Me.StartWith(aryPTN(i))) & Me.Chomp(aryPTN(i), 1)
        Else
        aryPTNRet(i) = UCase(Me.StartWith(aryPTN(i))) & Me.Chomp(aryPTN(i), 1)
        End If
Next i
If IsMissing(blnJoinStr) Then
    Let strSplit = ""
End If
Camelize = Join(aryPTNRet, strSplit)
End Function

Public Function Decamelize(str As String, Optional ByVal sep As Variant) As String
''' /********************************************************
'''  * @function decamelize
'''  */
If IsMissing(sep) Then
    Let sep = " "
End If
Dim converted As String
''
objRegExp_.Pattern = "([a-z\\d])([A-Z])"
objRegExp_.Global = True
Let converted = objRegExp_.Replace(str, "$1" & sep & "$2")
objRegExp_.Pattern = "([A-Z]+)([A-Z][a-z\\d]+)"
objRegExp_.Global = True
Let converted = objRegExp_.Replace(converted, "$1" & sep & "$2")
''
Let Decamelize = VBA.LCase(converted)
End Function

Public Function SnakeUpperCase(str As String, Optional ByVal strSplit As Variant) As String
''' /********************************************************
'''  * @function SnakeUpperCase
'''  */
If IsMissing(strSplit) Then
    Let strSplit = " "
End If
Dim converted As String
Let converted = Me.CapitalizeFully(str, strSplit)
Let SnakeUpperCase = Me.ReplaceString(converted, strSplit, "_")
End Function

Public Function SnakeLowerCase(str As String, Optional ByVal strSplit As Variant) As String
''' /********************************************************
'''  * @function SnakeUpperCase
'''  */
If IsMissing(strSplit) Then
    Let strSplit = " "
End If
Dim converted As String
Let converted = Me.SnakeUpperCase(str, strSplit)
Let SnakeLowerCase = Me.ToLower(converted)
End Function

Public Function Substitute(ByVal str As String, _
          ByVal orgStr As String, ByVal destStr As String, _
          Optional num As Variant) As String
''' /********************************************************
'''  * @function Substitute
'''  */
If IsMissing(num) Then
    Substitute = WorksheetFunction.Substitute(str, orgStr, destStr)
Else
    Substitute = WorksheetFunction.Substitute(str, orgStr, destStr, num)
End If
End Function

Public Function EncodeUrl(ByVal str As String) As String
''' /********************************************************
'''  * @function encodeUrl
'''  */
EncodeUrl = WorksheetFunction.EncodeUrl(str)
End Function

Public Function ReplaceString(ByVal str As String, _
          ByVal orgStr As String, ByVal destStr As String, _
          Optional regParam As Variant = "g") As String
''' /********************************************************
'''  * @function replaceString
'''  *</br>文字列を、ある文字を、指定の文字に置換する
'''  *</br>VBScript.RegExp#Replaceのラッパー関数
'''  *　@param {string} str 文字列
'''  *　@param {string} orgStr 置換対象文字列
'''  *　@param {string} destStr 置換文字列
'''  *　@param {string} regParam RegExpパラメータ
'''  *　@return {String}変換後の文字列
'''  */
    '"$"文字を置換文字列の接頭辞とする前提
    If Me.Contain(orgStr, "$") Then
        orgStr = Replace(orgStr, "$", "\$")
    End If
    If Me.Contain(LCase(regParam), "g") Then
        objRegExp_.Global = True '全体を対象とする
    Else
        objRegExp_.Global = False '全体を対象しない
    End If
    If Me.Contain(LCase(regParam), "i") Then
        objRegExp_.ignorecase = False '大小区別する場合
    Else
        objRegExp_.ignorecase = True  '大小区別しない
    End If
    objRegExp_.Pattern = orgStr
    ReplaceString = objRegExp_.Replace(str, destStr)
End Function

'Public Function getArray(str As String, Optional delimiter As Variant = ",") As String()
'''' /********************************************************
''''  * @function getArray　＝＞Splitそのものを利用する
''''  *</br>文字列を、セパレータ(delimiter)で区切り、配列に格納する。
''''  *</br>Split関数のラッパー
''''  *　@param {string} str 文字列
''''  *　@param {string} delimiter 区切り文字
''''  *　@return {String()}変換後の配列
''''  */
'    getArray = Split(str, delimiter)
'End Function

Public Function ToJsonForParamArray(ParamArray paryPTN()) As String
''' /********************************************************
'''  * @function ToJsonForParamArray
'''  *　@param {String()} paryPTN 文字列（パラメータ配列型）
'''  *　@return {String}Json文字列
'''  */
    Dim aryPTN() As Variant
    ReDim aryPTN(UBound(paryPTN))
    Dim i As Long
    For i = LBound(paryPTN) To UBound(paryPTN)
        aryPTN(i) = C_Commons.ToJson(paryPTN(i))
    Next i
    
    ToJsonForParamArray = "[" & Join(aryPTN, ",") & "]"
End Function

Public Function Concatenate(delimiter As String, ParamArray aryPTN()) As String
''' /********************************************************
'''  * @function concatenate
'''  *</br>文字列を、セパレータ(delimiter)で区切り、文字列を生成する。
'''  *</br>Split関数のラッパー
'''  *　@param {String} delimiter 文字列
'''  *　@param {String()} aryPTN 文字列（パラメータ配列型）
'''  *　@return {String}結合した文字列
'''  */
Concatenate = Join(aryPTN, delimiter)
End Function

Public Function Charactors(str) As Variant
''' /********************************************************
'''  * @function concatenate
'''  *</br>文字列を、セパレータ(delimiter)で区切り、配列に格納する。
'''  *</br>Split関数のラッパー
'''  *　@param {String()} aryPTN 文字列（パラメータ配列型）
'''  *　@return {String}結合した文字列
'''  */
Dim aryPTN() As Variant
ReDim aryPTN(Len(str) - 1)
Dim i As Long
For i = 1 To Len(str)
        aryPTN(i - 1) = Mid(str, i, 1)
Next i
Charactors = aryPTN
End Function

Public Function IndexOf(str As String _
                    , strKey As String _
                    , Optional lngStartPos As Variant) As Long
''' /********************************************************
'''  * @function indexOf
'''  *</br>InStr関数のRapper
'''  *</br>
'''  *　@param {String} str 文字列
'''  *　@param {String} strKey 文字列
'''  *　@param {Variant<Long>} lngStartPos
'''  *　@param {Variant<Long>} compareMode vbBinaryPTNCompare | vbTextCompare
'''  *　@return {Long} indexOf 先頭からの文字位置
'''  */
If IsMissing(lngStartPos) Then
        IndexOf = InStr(str, strKey)
Else
        IndexOf = InStr(lngStartPos, str, strKey)
End If
End Function

Public Function SubStringBefore(str As String _
                    , strKey As String _
                    , Optional lngStartPos As Variant) As String
''' /********************************************************
'''  * @function subStringBefore
'''  *<pre>指定の位置から、指定の文字までの文字列を返却
'''  *</pre>
'''  *　@param {String} str 文字列
'''  *　@param {String} strKey 文字列
'''  *　@param {Variant<Long>} lngStartPos
'''  *　@param {Variant<Long>} compareMode vbBinaryPTNCompare | vbTextCompare
'''  *　@return {String} subStringBefore 先頭からの指定の文字までの文字列
'''  */
Dim pos As Long
pos = IndexOf(str, strKey, lngStartPos)
If IsMissing(lngStartPos) Then
        SubStringBefore = left(str, pos)
Else
        If pos = 0 Then
            SubStringBefore = ""
        Else
            SubStringBefore = Mid(str, lngStartPos, pos - lngStartPos + 1)
        End If
End If
End Function

Public Function SubStringAfter(str As String _
                    , strKey As String _
                    , Optional lngEndPos As Variant) As String
''' /********************************************************
'''  * @function subStringAfter
'''  *<pre>指定の位置から、指定の文字までを、後ろから検索して、
'''  *抽出文字列を返却
'''  *</pre>
'''  *　@param {String} str 文字列
'''  *　@param {String} strKey 文字列
'''  *　@param {Variant<Long>} lngStartPos
'''  *　@param {Variant<Long>} compareMode vbBinaryPTNCompare | vbTextCompare
'''  *　@return {String} subStringAfter 最後尾からの指定の文字までの文字列
'''  */
Dim pos As Long
If IsMissing(lngEndPos) Then
        pos = InStrRev(str, strKey, Len(str))
Else
        pos = InStrRev(str, strKey, lngEndPos)
End If
If IsMissing(lngEndPos) Then
        SubStringAfter = Right(str, Len(str) - pos)
Else
        If pos = 0 Then
            SubStringAfter = ""
        Else
            If pos > lngEndPos Then
                SubStringAfter = ""
            Else
                SubStringAfter = Mid(str, pos, lngEndPos - pos + 1)
            End If
        End If
End If
End Function

Public Function ConvertAster(strPattern As String) As String
''' /********************************************************
'''  * @function convertAster
'''  *<pre>アスタリスクを、任意の一文字と*に変換して返却する。
'''  *（注意）
'''  * LikeでのWildカードのような指定を、RegExpで実現する際に
'''  * 利用を想定
'''  *</pre>
'''  *　@param {String} strPattern 文字列
'''  *　@return {String}変換後文字
'''  *　FIXME□□RegExpの利用も含め、どう使うかを整理。
'''  */
Dim s As String
objRegExp_.Pattern = "\."
s = objRegExp_.Replace(strPattern, "\.")
objRegExp_.Pattern = "\*"
ConvertAster = objRegExp_.Replace(s, ".*")
End Function

Public Function DefaultString(vnt As Variant, Optional defaultStr As Variant = "N/A") As String
''' /********************************************************
'''  *@function defaultString
'''  *<pre>渡された値が、空の文字列、Empty等である場合にDefault値を設定する
'''  *</pre>
'''  *　@param {Variant<String>} vnt 文字列
'''  *　@param {Variant<String>} defaultStr Default文字列
'''  *　@return {String} defaultString
'''  */
Dim str As Variant
str = vnt
Select Case vnt
        Case ""
            str = defaultStr
        Case vbNullString
            str = defaultStr
        Case Else
            If IsNull(vnt) Then
                str = defaultStr
            ElseIf IsEmpty(vnt) Then
                str = defaultStr
            End If
End Select
DefaultString = str
End Function

Public Function InnerString(ByVal str As String _
                    , leftString As String _
                    , rightString As String) As String
''' /********************************************************
'''  *@function innerString
'''  *<pre>指定の文字に囲われた文字を取得する。
'''  *</pre>
'''  *　@param {String} str 対象文字列
'''  *　@param {String} leftString 左側文字列
'''  *　@param {String} rightString 右側文字列
'''  *　@return {String} innerString
'''  */
Dim lngLeftPos As Long
Dim lngRightPos As Long
Dim lng As Long
'右側文字列の最初のヒット位置を取得
lngLeftPos = InStr(str, leftString)
If lngLeftPos = 0 Then
        InnerString = ""
        Exit Function
End If
    '左側文字列の最後のヒット位置を取得
lngRightPos = InStrRev(str, rightString)
If lngLeftPos = 0 Then
        InnerString = ""
        Exit Function
End If
''
If lngLeftPos > lngRightPos Then
        InnerString = ""
        Exit Function
End If
''
'間にある文字列を取得
InnerString = Mid$(str, lngLeftPos + Len(leftString), lngRightPos - lngLeftPos - Len(leftString))
''
End Function

Public Function IsVariableInTemplate(ByVal str As String) As Boolean
''' /********************************************************
'''  *@function isVariableInTemplate
'''  *<pre>テンプレート文字列の中に、${xxx}のVariable設定が
'''  *あるかを判定する。
'''  *</pre>
'''  *　@param {String} str 対象文字列
'''  *　@return {Boolean} isVariableInTemplate
'''  */
Dim aryPTNTemp As Variant
aryPTNTemp = Split(str, "$")
Dim aryPTN() As Variant
ReDim aryPTN(LBound(aryPTNTemp) To UBound(aryPTNTemp))
Dim i As Long
For i = LBound(aryPTNTemp) To UBound(aryPTNTemp)
        aryPTN(i) = Me.InnerString(aryPTNTemp(i), "{", "}")
Next i
''
For i = LBound(aryPTN) To UBound(aryPTN)
        If Not aryPTN(i) = "" Then
            IsVariableInTemplate = True
            Exit Function
        End If
Next i
IsVariableInTemplate = False
''
End Function

Public Function IsVariableInTemplateOfMustache(ByVal str As String) As Boolean
''' /********************************************************
'''  *@function IsVariableInTemplateOfMustache
'''  *<pre>テンプレート文字列の中に、{{xxx}}のVariable設定が
'''  *あるかを判定する。
'''  *</pre>
'''  *　@param {String} str 対象文字列
'''  *　@return {Boolean} IsVariableInTemplateOfMustache
'''  */
Dim aryPTNTemp As Variant
aryPTNTemp = Split(str, "{{")
Dim aryPTN() As Variant
ReDim aryPTN(LBound(aryPTNTemp) To UBound(aryPTNTemp))
Dim i As Long
For i = LBound(aryPTNTemp) To UBound(aryPTNTemp)
    If Not aryPTNTemp(i) = "" Then
        Let aryPTN(i) = VBA.Mid(aryPTNTemp(i), 1, VBA.InStr(aryPTNTemp(i), "}}") - 1)
    End If
Next i
''
If Not Join(aryPTN, "") = "" Then
    Let IsVariableInTemplateOfMustache = True
Else
    Let IsVariableInTemplateOfMustache = False
End If
''
End Function

Public Function GetVariablesInTemplate(ByVal str As String) As Variant
''' /********************************************************
'''  *@function isVariableInTemplate
'''  *<pre>テンプレート文字列の中に、${xxx}のVariable設定が
'''  *あるかを判定する。
'''  *</pre>
'''  *　@param {String} str 対象文字列
'''  *　@return {Boolean} isVariableInTemplate
'''  */
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim aryPTNTemp As Variant
aryPTNTemp = Split(str, "$")
Dim aryPTN() As Variant
ReDim aryPTN(LBound(aryPTNTemp) To UBound(aryPTNTemp))
Dim i As Long
For i = LBound(aryPTNTemp) To UBound(aryPTNTemp)
    aryPTN(i) = Me.InnerString(aryPTNTemp(i), "{", "}")
Next i
''
Call C_Array.DeleteEmptyElement(aryPTN)
GetVariablesInTemplate = aryPTN
''
End Function

Public Function GetVariablesInTemplateOfMustache(ByVal str As String) As Variant
''' /********************************************************
'''  *@function isVariableInTemplate
'''  *<pre>テンプレート文字列の中に、${xxx}のVariable設定が
'''  *あるかを判定する。
'''  *</pre>
'''  *　@param {String} str 対象文字列
'''  *　@return {Boolean} isVariableInTemplate
'''  */
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim aryPTNTemp As Variant
aryPTNTemp = Split(str, "{{")
Dim aryPTN() As Variant
ReDim aryPTN(LBound(aryPTNTemp) To UBound(aryPTNTemp))
Dim i As Long
For i = LBound(aryPTNTemp) To UBound(aryPTNTemp)
    If Not aryPTNTemp(i) = "" Then
        Let aryPTN(i) = VBA.Mid(aryPTNTemp(i), 1, VBA.InStr(aryPTNTemp(i), "}}") - 1)
    End If
Next i
''
Call C_Array.DeleteEmptyElement(aryPTN)
GetVariablesInTemplateOfMustache = aryPTN
''
End Function

Public Function Substitute4Template(ByVal str As String, _
          ByVal orgStr As String, ByVal destStr As String, _
          Optional num As Variant) As String
''' /********************************************************
'''  * @function Substitute
'''  */
Dim orgStrVar As String
Let orgStrVar = "${" + orgStr + "}"
If IsMissing(num) Then
    Substitute4Template = WorksheetFunction.Substitute(str, orgStrVar, destStr)
Else
    Substitute4Template = WorksheetFunction.Substitute(str, orgStrVar, destStr, num)
End If
End Function

Public Function Substitute4Mustache(ByVal str As String, _
          ByVal orgStr As String, ByVal destStr As String, _
          Optional num As Variant) As String
''' /********************************************************
'''  * @function Substitute
'''  */
Dim orgStrVar As String
Let orgStrVar = "{{" + orgStr + "}}"
If IsMissing(num) Then
    Substitute4Mustache = WorksheetFunction.Substitute(str, orgStrVar, destStr)
Else
    Substitute4Mustache = WorksheetFunction.Substitute(str, orgStrVar, destStr, num)
End If
End Function

Public Function RenderTemplate(ByVal templ As String, ByRef dict As Dictionary, Optional ByVal style As Variant) As String
''' /********************************************************
'''  *@function RenderTemplate
'''  */
Dim aryVariables() As Variant
Dim rtn As String
Let rtn = templ
''
If VBA.IsMissing(style) Then
    Let style = "${}"
End If
If style = "${}" Then
    If Not Me.IsVariableInTemplate(rtn) Then
        RenderTemplate = rtn
        Exit Function
    End If
ElseIf style = "{{}}" Then
    If Not Me.IsVariableInTemplateOfMustache(rtn) Then
    RenderTemplate = rtn
    Exit Function
    End If
End If
''
If style = "${}" Then
    Let aryVariables = Me.GetVariablesInTemplate(templ)
ElseIf style = "{{}}" Then
    Let aryVariables = Me.GetVariablesInTemplateOfMustache(templ)
End If
''
Dim elm As Variant
Dim i As Long
For i = LBound(aryVariables) To UBound(aryVariables)
    Let elm = aryVariables(i)
    If dict.Exists(elm) Then
        If style = "${}" Then
        Let rtn = Me.Substitute4Template(rtn, elm, dict.Item(elm))
        ElseIf style = "{{}}" Then
            ''Console.log elm
            Let rtn = Me.Substitute4Mustache(rtn, elm, dict.Item(elm))
        End If
    End If
Next i
Let RenderTemplate = rtn
''
End Function

Public Function GetTemporaryStr() As String
''' /********************************************************
'''  *@function RenderTemplate
'''  */
''
Dim rtn As String
Dim objFSO As Object
Set objFSO = createObject("Scripting.FileSystemObject")
Let rtn = objFSO.GetBaseName(objFSO.GetTempName)
Let GetTemporaryStr = rtn
Set objFSO = Nothing
''
End Function

Public Function DeleteEmptyRow(ByVal str As String) As String
'''' ********************************************************
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim ary As Variant
Let ary = Split(str, vbLf)
Call C_Array.DeleteEmptyElement(ary)
Let DeleteEmptyRow = Join(ary, vbCrLf)
''
End Function

'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Sub UnitTest()
    Dim CONS_MODULE_NAME As String
    Let CONS_MODULE_NAME = "C_String.unitTest"
    '//--------------------------------------
    '//変数宣言
    Dim blnRet As Boolean
    Dim str As String
    str = "  Sakai   Mitsuru  "
    Dim strPass As String
    
    Dim C_String As C_String
    Set C_String = New C_String
    '//--------------------------------------
    '// 初期処理
On Error GoTo EXCEPTION
    '//--------------------------------------
    '// 処理
    Console.info ("//----------------------------------// Start")
''stripCrLf
    str = "sakai" & vbCrLf & "mitsuru" & vbLf _
        & "sakamoto" & vbCr & "misa"
    Console.info str
    Console.info Me.StripCrLf(str)


    GoSub FINALLY
    Exit Sub
    
FINALLY: 'Label
    '//--------------------------------------
    '// 終了処理
    Console.info ("//-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=// End")
    Return

EXCEPTION: 'Label
    '//--------------------------------------
    ' 例外処理
    GoSub FINALLY
    Console.info ("Err.Source = " & Err.source & "." & CONS_MODULE_NAME)
    Console.info ("Err.Number = " & Err.Number)
    Console.info ("Err.Description = " & vbCr & CONS_MODULE_NAME & vbCr & Err.Description)
    Console.info ("//-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-// Error End")
End Sub


Public Function passwordDesp(str As String)
''' /********************************************************
'''  * passwordDesp
'''  *</br>パスワードの読み方を表す
'''  *</br>
'''  *　@param {String} str 読み方を表示するパスワード
'''  *　@return {String} 読み方を表す文字列
'''  */
    Dim strRet As String
    Dim objDict As Object
    Dim aryPTN() As String
    ReDim aryPTN(Len(str) - 1)
    Dim aryPTNBuff() As String
    ReDim aryPTNBuff(Len(str) - 1)
    Dim i As Integer
    For i = 1 To Len(str)
        aryPTN(i - 1) = Mid$(str, i, 1)
    Next i

    Set objDict = createObject("Scripting.DictionaryPTN")
    objDict.Add "a", "エイ"
    objDict.Add "b", "ビィ"
    objDict.Add "c", "シィ"
    objDict.Add "d", "ディ"
    objDict.Add "e", "イー"
    objDict.Add "f", "エフ"
    objDict.Add "g", "ジィ"
    objDict.Add "h", "エッチ"
    objDict.Add "i", "アイ"
    objDict.Add "j", "ジェイ"
    objDict.Add "k", "ケイ"
    objDict.Add "l", "エル"
    objDict.Add "m", "エム"
    objDict.Add "n", "エヌ"
    objDict.Add "o", "オー"
    objDict.Add "p", "ピィ"
    objDict.Add "q", "キュー"
    objDict.Add "r", "アール"
    objDict.Add "s", "エス"
    objDict.Add "t", "ティ"
    objDict.Add "u", "ユー"
    objDict.Add "v", "ブィ"
    objDict.Add "w", "ダブリュー"
    objDict.Add "x", "エックス"
    objDict.Add "y", "ワイ"
    objDict.Add "z", "ゼット"
    objDict.Add "A", "エイ（大文字）"
    objDict.Add "B", "ビィ（大文字）"
    objDict.Add "C", "シィ（大文字）"
    objDict.Add "D", "ディ（大文字）"
    objDict.Add "E", "イー（大文字）"
    objDict.Add "F", "エフ（大文字）"
    objDict.Add "G", "ジィ（大文字）"
    objDict.Add "H", "エッチ（大文字）"
    objDict.Add "I", "アイ（大文字）"
    objDict.Add "J", "ジェイ（大文字）"
    objDict.Add "K", "ケイ（大文字）"
    objDict.Add "L", "エル（大文字）"
    objDict.Add "M", "エム（大文字）"
    objDict.Add "N", "エヌ（大文字）"
    objDict.Add "O", "オー（大文字）"
    objDict.Add "P", "ピィ（大文字）"
    objDict.Add "Q", "キュー（大文字）"
    objDict.Add "R", "アール（大文字）"
    objDict.Add "S", "エス（大文字）"
    objDict.Add "T", "ティ（大文字）"
    objDict.Add "U", "ユー（大文字）"
    objDict.Add "V", "ブィ（大文字）"
    objDict.Add "W", "ダブリュー（大文字）"
    objDict.Add "X", "エックス（大文字）"
    objDict.Add "Y", "ワイ（大文字）"
    objDict.Add "Z", "ゼット（大文字）"
    objDict.Add "0", "０"
    objDict.Add "1", "１"
    objDict.Add "2", "２"
    objDict.Add "3", "３"
    objDict.Add "4", "４"
    objDict.Add "5", "５"
    objDict.Add "6", "６"
    objDict.Add "7", "７"
    objDict.Add "8", "８"
    objDict.Add "9", "９"
    objDict.Add "-", "ハイフン"
    objDict.Add "^", "ハット"
    objDict.Add "@", "アト"
    objDict.Add "[", "大括弧左"
    objDict.Add ";", "セミコロン"
    objDict.Add "]", "大括弧右"
    objDict.Add ",", "コロン"
    objDict.Add "!", "アテンション"
    objDict.Add "#", "シャープ"
    objDict.Add "$", "ドル"
    objDict.Add "%", "パーセント"
    objDict.Add "&", "アンド"
    objDict.Add "'", "シングルクォート"
    objDict.Add "(", "丸括弧左"
    objDict.Add ")", "丸括弧右"
    objDict.Add "=", "イコール"
    objDict.Add "~", "チルド"
    objDict.Add "`", "グラーブアクセント"
    objDict.Add "{", "波括弧左"
    objDict.Add "+", "プラス"
    objDict.Add "}", "波括弧右"
    objDict.Add "<", "小なり"
    objDict.Add ">", "大なり"
    objDict.Add "_", "アンダーバー"
    
    For i = 0 To Len(str) - 1
        aryPTNBuff(i) = objDict(aryPTN(i))
    Next i
    
    strRet = Join(aryPTNBuff(), "・")
    
    passwordDesp = strRet
End Function


