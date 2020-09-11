VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "O_StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file O_StringBuilder.cls
'''  *<br/>.NET StringBuilder Wrapper
'''  * @version 1.00
'''  * @since 2020/03/31
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  *
'''  * @require C_String,C_File,C_FileIO
'''  *
'''  */
'
''' /********************************************************
'''  * @class O_StringBuilder
'''  * @classdesc String Buffer Utility Class
'''  */
'
'//--------------------------------------------------------
'// Class Property
'//
''' /** @property @private @name builder @type {Object:System.Text.StringBuilder} @desc Object  */
'vba_stringbuilder class
Private builder_ As Object

'//--------------------------------------------------------
'// Class Event
'//
Private Sub Class_Initialize()
''' /********************************************************
'''  * @event Class_Initialize
'''  */
'Constructor
    Set builder_ = createObject("System.Text.StringBuilder")
End Sub

Private Sub Class_Terminated()
''' /********************************************************
'''  *Class_Terminate
'''  */
'Deconstructor
    Set builder_ = Nothing
End Sub

'//--------------------------------------------------------
'// setter/getter
'//
Property Get Chars(ByVal Index As Integer) As String
    Chars = ChrW(builder_.Chars(Index))
End Property

Property Let Chars(ByVal Index As Integer, ByVal Value As String)
    builder_.Chars(Index) = AscW(Value)
'System.Text.StringBuilder Chars(int index)
End Property

Property Get Capacity() As Integer
    Capacity = builder_.Capacity
End Property

Property Let Capacity(ByVal Value As Integer)
    builder_.Capacity = Value
'Definition : System.Int32 Capacity {get;set;}
End Property

Property Get Length() As Integer
    Length = builder_.Length
End Property

Property Get LengthW() As Integer
    Dim C_String As C_String
    Set C_String = New C_String
    LengthW = C_String.Length(builder_.toString)
    Set C_String = Nothing
End Property

Property Let Length(ByVal Value As Integer)
    builder_.Length = Value
'Definition : System.Int32 Length {get;set;}
End Property

Property Get MaxCapacity()
    MaxCapacity = builder_.MaxCapacity
'Definition : System.Int32 MaxCapacity {get;}
End Property

'//--------------------------------------------------------
'// Class Method (Function)
'//
Public Function append(ByVal val_or_char As Variant _
         , Optional ByVal repeat_or_startIndex As Integer = 0 _
         , Optional ByVal Count As Integer = 0)
''' /********************************************************
'''  * @function Append
'''  *</br>[System.Text.StringBuilder]::Append wrapper
'''  *</br>
'''  *　@param {variant} val_or_char
'''  *　@param {Integer} repeat_or_startIndex : Optional start position, Default=0
'''  *　@param {Integer} count : Optional string count Default=0
'''  */
'Append overload
    If repeat_or_startIndex > 0 Then
        If Count > 0 Then
            Call builder_.Append_4(val_or_char, repeat_or_startIndex, Count)
        Else
            Call builder_.append(val_or_char, repeat_or_startIndex)
        End If
    Else
        ' if paramter's count is one
        Select Case TypeName(val_or_char)
            Case "Byte": builder_.Append_7 (val_or_char)
            Case "Integer": builder_.Append_10 (val_or_char)
            Case "Long": builder_.Append_11 (val_or_char)
            Case "Single": builder_.Append_12 (val_or_char)
            Case "Double": builder_.Append_13 (val_or_char)
            Case "Currency": builder_.Append_14 (val_or_char)
            Case "String": builder_.Append_3 (val_or_char)
            Case "Boolean": builder_.Append_5 (val_or_char)
            Case "Date", "Object"
                builder_.Append_18 (val_or_char)
            Case "Error", "Empty", "Null", "Unknown", "Nothing"
            Case Else
                If IsArray(val_or_char) Then
                    builder_.Append_3 (C_Commons.toString(val_or_char))
                Else
                    builder_.Append_18 (val_or_char)
                End If
        End Select
    End If
    Set append = Me
'System.Text.StringBuilder Append(char value, int repeatCount),
'System.Text.StringBuilder Append_2(char[] value, int startIndex, int charCount),
'System.Text.StringBuilder Append_3(string value),
'System.Text.StringBuilder Append_4(string value, int startIndex, int count),
'System.Text.StringBuilder Append_5(bool value),
'System.Text.StringBuilder Append_6(System.SByte value),
'System.Text.StringBuilder Append_7(byte value),
'System.Text.StringBuilder Append_8(char value),
'System.Text.StringBuilder Append_9(System.Int16 value),
'System.Text.StringBuilder Append_10(int value),
'System.Text.StringBuilder Append_11(long value),
'System.Text.StringBuilder Append_12(float value),
'System.Text.StringBuilder Append_13(double value),
'System.Text.StringBuilder Append_14(decimal value),
'System.Text.StringBuilder Append_15(System.UInt16 value),
'System.Text.StringBuilder Append_16(System.UInt32 value),
'System.Text.StringBuilder Append_17(System.UInt64 value),
'System.Text.StringBuilder Append_18(System.Object value),
'System.Text.StringBuilder Append_19(char[] value)
End Function

Public Function Appends(ParamArray pAry())
''' /********************************************************
'''  *@function Appends
'''  *</br>Append with ParamArra
'''  *</br>
'''  *@param {ParamArray} ary
'''  */
    Dim i As Long
    For i = LBound(pAry) To UBound(pAry)
        Call Me.append(pAry(i))
    Next i
    Set Appends = Me
End Function

Public Function AppendFormat(ByVal format As Variant, ParamArray arg() As Variant)
''' /********************************************************
'''  * @function AppendFormat
'''  *</br>[System.Text.StringBuilder]::AppendFormat wrapper
'''  *</br>  Refer to AppendFormat
'''  *　@param {variant} format
'''  *　@param {ParamArray} arg()
'''  */
'AppendFormat overload
    Dim obj() As Variant, i As Integer
    
    If TypeName(format) = "String" Then
        Select Case UBound(arg)
        Case 0
            If IsArray(arg) Then
                ReDim obj(UBound(arg(0)))
                For i = 0 To UBound(arg(0))
                    obj(i) = arg(0)(i)
                Next
                Call builder_.AppendFormat_4(format, obj)
            Else
                Call builder_.AppendFormat(format, arg(0))
            End If
        Case 1: Call builder_.AppendFormat_2(format, arg(0), arg(1))
        Case 2: Call builder_.AppendFormat_3(format, arg(0), arg(1), arg(2))
        Case Else
            ReDim obj(UBound(arg))
            For i = 0 To UBound(arg)
                obj(i) = arg(i)
            Next
            Call builder_.AppendFormat_4(format, obj)
        End Select
    Else
        ReDim obj(UBound(arg) - 1)
        For i = 0 To UBound(arg) - 1
            obj(i) = arg(i + 1)
        Next i
        Call builder_.AppendFormat_5(format, arg(0), obj)
    End If
    Set AppendFormat = Me
    
'System.Text.StringBuilder AppendFormat(string format, System.Object arg0),
'System.Text.StringBuilder AppendFormat_2(string format, System.Object arg0, System.Object arg1),
'System.Text.StringBuilder AppendFormat_3(string format, System.Object arg0, System.Object arg1, System.Object arg2),
'System.Text.StringBuilder AppendFormat_4(string format, Params System.Object[] args),
'System.Text.StringBuilder AppendFormat_5(System.IFormatProvider provider, string format, Params System.Object[] args)
End Function

Public Function AppendLine(Optional ByVal str As String = "")
''' /********************************************************
'''  * @function AppendLine
'''  *</br>[System.Text.StringBuilder]::AppendLine wrapper
'''  *</br>
'''  *　@param {String} str
'''  */
    If Len(str) > 0 Then
        Me.append (str)
    End If
    Me.append (vbCrLf)
    Set AppendLine = Me
End Function

Public Function AppendsLine(ParamArray ary() As Variant)
''' /********************************************************
'''  * @function AppendsLine
'''  *</br>append a ParamArray element of string as a line
'''  *</br>
'''  *@param {String} ary
'''  */
    Dim i As Long
    For i = LBound(ary) To UBound(ary)
        Me.append (ary(i))
    Next i
    Me.append (vbCrLf)
    Set AppendsLine = Me
End Function

Public Function AppendLines(ParamArray ary() As Variant)
''' /********************************************************
'''  * @function AppendLines
'''  *</br>append a ParamArray element inculde object as a line
'''  *</br>
'''  *@param {String[]} ary
'''  */
    Dim i As Long
    For i = LBound(ary) To UBound(ary)
        If TypeName(ary(i)) = "String" Then
            Me.AppendLine (ary(i))
        Else
            Me.AppendLine (C_Commons.toString(ary(i)))
        End If
    Next i
    append (vbCrLf)
    Set AppendLines = Me
End Function

Public Function AppendFormatLine(ByVal format As Variant, ParamArray arg() As Variant)
''' /********************************************************
'''  * @function AppendFormatLine
'''  *</br>Append a formatted line
'''  *</br>
'''  *@param {Variant} format
'''  *@param {Variant[]} arg
'''  */
    Call AppendFormat(format, arg()).AppendLine
    Set AppendFormatLine = Me
End Function
        
Public Function EnsureCapacity(ByVal Capacity As Integer) As Integer
''' /********************************************************
'''  * @function EnsureCapacity
'''  *</br>[System.Text.StringBuilder]::EnsureCapacity Property
'''  *</br>
'''  *@param {Integer} Capacity
'''  */
    EnsureCapacity = builder_.EnsureCapacity(Capacity)
End Function

Public Function Equals(ByRef obj As O_StringBuilder) As Boolean
''' /********************************************************
'''  * @function Equals
'''  *</br>[System.Text.StringBuilder]::Equals wrapper
'''  *</br>
'''  * @param {Objct} obj : builder object for comparing
'''  */
       'equals = builder_.equals(obj.inner_stringbuilder) ''FIXME
End Function

Public Function GetHashCode() As Long
''' /********************************************************
'''  * @function GetHashCode
'''  *</br>[System.Text.StringBuilder]::GetHashCode wrapper
'''  *</br>
'''  *　@return {Integer} This object's hashcode
'''  */
    GetHashCode = builder_.GetHashCode
End Function

Public Function insert(ByVal Index As Integer, ByVal arg As Variant _
         , Optional ByVal count1 As Integer = 0 _
         , Optional ByVal count2 As Integer = 0)
''' /********************************************************
'''  * @function Insert
'''  *</br>[System.Text.StringBuilder]::Insert wrapper
'''  *</br>
'''  *　@param {Integer} index target buffer index
'''  *　@param {Variant} arg strings
'''  *　@param {Integer} count1 sart position of arg
'''  *　@param {Integer} count2 substing length
'''  */
    If count1 > 0 Then
        If count2 > 0 Then
            Dim str As String
            str = Mid(arg, count1 + 1, count2)
            builder_.Insert_2 Index, str
        Else
            builder_.insert Index, arg, count1
        End If
    Else
        Select Case TypeName(arg)
            Case "Byte": builder_.Insert_5 Index, arg
            Case "Integer": builder_.Insert_6 Index, arg
            Case "Long": builder_.Insert_10 Index, arg
            Case "Single": builder_.Insert_12 Index, arg
            Case "Double": builder_.Insert_13 Index, arg
            Case "Currency": builder_.Insert_14 Index, arg
            Case "String": builder_.Insert_2 Index, arg
            Case "Boolean": builder_.Insert_3 Index, arg
            Case "Date", "Object"
                builder_.Insert_18 Index, arg
            Case "Error", "Empty", "Null", "Unknown", "Nothing"
            Case Else
                builder_.Insert_18 Index, arg
        End Select
    End If
'System.Text.StringBuilder Insert(int index, string value, int count)
'System.Text.StringBuilder Insert_2(int index, string value)
'System.Text.StringBuilder Insert_3(int index, bool value)
'System.Text.StringBuilder Insert_4(int index, System.SByte value)
'System.Text.StringBuilder Insert_5(int index, byte value)
'System.Text.StringBuilder Insert_6(int index, System.Int16 value)
'System.Text.StringBuilder Insert_7(int index, char value)
'System.Text.StringBuilder Insert_8(int index, char[] value)
'System.Text.StringBuilder Insert_9(int index, char[] value, int startIndex, int charCount)
'System.Text.StringBuilder Insert_10(int index, int value)
'System.Text.StringBuilder Insert_11(int index, long value)
'System.Text.StringBuilder Insert_12(int index, float value)
'System.Text.StringBuilder Insert_13(int index, double value)
'System.Text.StringBuilder Insert_14(int index, decimal value)
'System.Text.StringBuilder Insert_15(int index, System.UInt16 value)
'System.Text.StringBuilder Insert_16(int index, System.UInt32 value)
'System.Text.StringBuilder Insert_17(int index, System.UInt64 value)
'System.Text.StringBuilder Insert_18(int index, System.Object value)
End Function

Public Function remove(ByVal startIndex As Integer, ByVal Length As Integer)
''' /********************************************************
'''  * @function Remove
'''  *</br>[System.Text.StringBuilder]::Remove wrapper
'''  *</br>
'''  *　@param {Integer} startIndex :start position
'''  *　@param {Integer} Length :string length to delete
'''  */
    builder_.remove startIndex, Length
'System.Text.StringBuilder Remove(int startIndex, int length)
End Function

Public Function Replace(ByVal oldstr As String _
          , ByVal newstr As String _
          , Optional ByVal startIndex As Integer = -1 _
          , Optional ByVal Count As Integer = 0)
''' /********************************************************
'''  * @function Replace
'''  *</br>[System.Text.StringBuilder]::Replace wrapper
'''  *</br>
'''  *　@param {String} oldstr :strings for replacement
'''  *　@param {String} newstr :replace oldstr to newstr
'''  *　@param {Integer} startIndex :start position
'''  *　@param {Integer} count :substring length
'''  */
    If startIndex > -1 And Count > 0 Then
        builder_.Replace_2 oldstr, newstr, startIndex, Count
    Else
        builder_.Replace oldstr, newstr
    End If
'System.Text.StringBuilder Replace(string oldValue, string newValue)
'System.Text.StringBuilder Replace_2(string oldValue, string newValue, int startIndex, int count)
'System.Text.StringBuilder Replace_3(char oldChar, char newChar)
'System.Text.StringBuilder Replace_4(char oldChar, char newChar, int startIndex, int count)
End Function

Public Function clear()
''' /********************************************************
'''  * @function Clear
'''  *</br> clear stringbuilder object
'''  */
    Set builder_ = Nothing
    Set builder_ = createObject("System.Text.StringBuilder")
End Function

Public Function Flush(Optional ByVal fileName As Variant _
                , Optional ByVal folderPath As Variant)
''' /********************************************************
'''  * @function flush
'''  *</br> flush buffer into a file
'''  *@param {Variant} filename Optional : default = workbook name and yymmdd.log
'''  *@param {Variant} folderpath
'''  *@require C_File,C_FileIO
'''  */
    Dim C_File As C_File
    Set C_File = New C_File
    Dim C_FileIO As C_FileIO
    Set C_FileIO = New C_FileIO
    
    Dim strFilename As String
    Dim strFilePath As String
    Dim strBookPath As String
    Dim strDtm As String
    strDtm = format(Now(), "yymmdd")
    strBookPath = ThisWorkbook.FullName
    
    
    If IsMissing(folderPath) Then
        folderPath = C_File.GetSpecialFolder()
    End If
    If IsMissing(fileName) Then
        strFilename = C_File.GetBaseName(strBookPath) & "_" & strDtm & ".log"
    Else
        strFilename = fileName
    End If
    
    strFilePath = C_File.BuildPath(folderPath, strFilename)
        
    Call C_FileIO.WriteTextAllAsUTF8(strFilePath, Me.toString())
        
    Set C_File = Nothing
    Set C_FileIO = Nothing

End Function

Public Function toString(Optional ByVal startIndex As Integer = -1 _
                , Optional ByVal Length As Integer = 0) As String
''' /********************************************************
'''  * @function ToString
'''  *</br>[System.Text.StringBuilder]::ToString wrapper
'''  *</br>
'''  *　@param {Integer} startIndex Optional: start position
'''  *　@param {Integer} Length : Optional substring length
'''  *　@return {String}
'''  */
    If startIndex > -1 And Length > 0 Then
        toString = builder_.ToString_2(startIndex, Length)
    Else
        toString = builder_.toString
    End If
'System.Text.StringBuilder ToString()
'System.Text.StringBuilder ToString_2(int startIndex, int length)
End Function

Public Function ToSource() As String
'''/********************************************************
''' * @function toSource
''' *<br/> stringify object
''' * @return {String}
''' */
    ToSource = "Object<O_StringBuilder> " & """" & toString() & """"
    
End Function

Public Function GetTypeName() As String
''' /********************************************************
'''  *@function GetTypeName
'''  *<br/> get data type name
'''  */
    GetTypeName = TypeName(Me)
End Function

Public Function GetVerType() As Long
''' /********************************************************
'''  *@function GetVerType
'''  *<br/> get VarType
'''  */
    GetVerType = VarType(builder_)
End Function

Public Function getObject() As Variant
''' /********************************************************
'''  *@function GetObject
'''  *<br/> get builder object
'''  */
    Set getObject = builder_
End Function

'/////////////////////////////////////////////////////////////
'//  Debug Area  /////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
'-------------------------------------------------------------
' debug entry
'-------------------------------------------------------------
Public Sub UnitTest()
    Dim obj As Object
''AppendLine
    Me.AppendLine ("sak.mit")
    Me.AppendLine ("sakot.miss")
    Console.Dump Me
''AppendLines
    Call Me.AppendLines("saki.gira", "sk.ao", "sai.iyi", "ski.osu")
''Capacity
    Console.info Me.Capacity
''Length
    Console.info Me.Length
''MaxCapacity
    Console.info Me.MaxCapacity
''ToSource
    Console.info Me.ToSource
''ToString
    Console.info Me.toString
''AppendFormat
    Call Me.AppendFormat("{0} is {1}", "sak", "mit")
    Call Me.AppendFormat("{0} is {1}", "sak", "mit")
''AppendsLine
    Call Me.AppendsLine("")
    Console.Dump Me
''AppendFormatLine
    Call Me.AppendFormatLine("{0} is {1}", "sakot", "miss")
''flush
    Call Me.Flush
''Object
    Console.info Me.AppendLine("sak").GetTypeName
    Console.info Me.AppendLine("mit").toString
    Console.info Me.GetVerType
''Clear
    Call Me.clear
    Console.Dump Me
    
End Sub



