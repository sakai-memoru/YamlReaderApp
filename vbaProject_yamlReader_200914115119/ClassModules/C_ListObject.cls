VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_ListObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''' /**-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
'''  * @file C_ListObject.cls
'''  *<pre>Table(ListObject)用ユーティリティ・クラス
'''  *</pre>
'''  * @version 1.00
'''  * @since 2020/04/30
'''  * @author sakai-memoru
'''  * @license MIT (http://www.opensource.org/licenses/mit-license.php)
'''  *
'''  */
'
'//--------------------------------------------------------
'// Class Method
'//
Public Function Exists(ByVal sheetName As String, ByVal tblName As String) As Boolean
'''' /********************************************************
''''  *@function Exists
''''  *<pre>
''''  * If a table of name Exists or not in a sheet
''''  *</pre>
''''  *@param {String} sheetname
''''  *@param {String} tblName
''''  *
''''  */
''
    Dim clts As ListObjects
    Dim lstObj As ListObject
    Dim blnRtn As Boolean
On Error Resume Next
    blnRtn = False
    Set clts = ActiveWorkbook.Worksheets(sheetName).ListObjects
    If Err.Number <> 0 Then
        Err.clear
        GoTo FINALLY
    End If
    For Each lstObj In clts
        If lstObj.Name = tblName Then
            blnRtn = True
        End If
    Next lstObj
    GoSub FINALLY
    Exit Function
FINALLY:
    Exists = blnRtn
    Set clts = Nothing
    Return
End Function

Public Sub CreateNewTable(ByRef rng As Range, ByVal tblName As String)
'''' /********************************************************
'' FIXME どのタイミングで使うのかで機能を整理する
''
Dim tbl As ListObject
Dim sht As Worksheet
Set sht = ThisWorkbook.Worksheets(rng.Parent.Name)
Set tbl = sht.ListObjects.Add(xlSrcRange, rng, XlListObjectHasHeaders:=xlYes)
tbl.Name = tblName
If tbl.InsertRowRange Is Nothing Then
    ResetTable tbl
End If
''
End Sub

Public Sub ResetTable(ByRef tblObj As ListObject)
'''' /********************************************************
''
tblObj.DataBodyRange.Delete
End Sub

Public Sub ClearTable(ByRef tblObj As ListObject)
'''' /********************************************************
''
Dim rng As Range
Set rng = tblObj.Range
tblObj.Unlist
rng.ClearFormats
''
End Sub

Public Function GetHeaderArray(ByRef tblObj As ListObject) As Variant
'''' /********************************************************
''''  *@function GetHeaderArray
''''  *@param {ListObject} tblObject
''''  *@return {Variant()}
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim C_String As C_String
Set C_String = New C_String
Dim aryRtn() As Variant
Dim c As Long
Let aryRtn = C_Array.ChangeIndexOfBegin(C_Array.ArrayArrayToArray(tblObj.HeaderRowRange.Value))
''
For c = LBound(aryRtn) To UBound(aryRtn)
    aryRtn(c) = C_String.ReplaceString(aryRtn(c), " ", "_")
Next c
''
GetHeaderArray = aryRtn
''
End Function

Public Function GetRowIndexByPrimaryKey(ByRef tblObj As ListObject, _
                    ByVal colName As String, ByVal key As Variant) As Long
'''' /********************************************************
''''  *@function GetNextNumber
''''  *@param {ListObject} tblObject
''''  *@param {Long} colIdx
''''  *@param {Long} inc : increase by step
''''  *@return {Long}
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim idx As Long
Dim headerAry() As Variant
Let headerAry = Me.GetHeaderArray(tblObj)
Let idx = C_Array.IndexOf(headerAry, colName)
Dim keyAry() As Variant
Let keyAry = C_Array.ArrayArrayToArray(tblObj.ListColumns(idx).DataBodyRange.Value)
Let keyAry = C_Array.ChangeIndexOfBegin(keyAry)
GetRowIndexByPrimaryKey = C_Array.IndexOf(keyAry, key)
''
End Function

Public Function GetNextNumber(ByRef tblObj As ListObject, ByVal colIdx As Long, Optional inc As Long = 1) As Long
'''' /********************************************************
''''  *@function GetNextNumber
''''  *@param {ListObject} tblObject
''''  *@param {Long} colIdx
''''  *@param {Long} inc : increase by step
''''  *@return {Long}
''''  *
''''  */
Dim C_Array As C_Array
Set C_Array = New C_Array
''
Dim colAry As Variant
Let colAry = C_Array.ArrayArrayToArray(tblObj.ListColumns(colIdx).DataBodyRange.Value)
'Console.Dump colAry
C_Array.Sort colAry
''
GetNextNumber = colAry(UBound(colAry)) + inc
''
End Function

Public Function GetNewRowArray(ByRef tblObj As ListObject) As Variant
'''' /********************************************************
''''  *@function GetNewRowArray
''''  *<pre>
''''  * If a table of name Exists or not in a sheet
''''  *</pre>
''''  *@param {String} sheetname
''''  *@param {String} tblName
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim c As Long
Let c = tblObj.ListColumns.Count
Dim ary() As Variant
ReDim ary(1 To c)
''
C_Array.FillEmptySpaceForArray ary
GetNewRowArray = ary
''
End Function

Public Function GetNewRowArrayArray(ByRef tblObj As ListObject) As Variant
'''' /********************************************************
''''  *@function GetNewRowArray
''''  *<pre>
''''  * If a table of name Exists or not in a sheet
''''  *</pre>
''''  *@param {String} sheetname
''''  *@param {String} tblName
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim c As Long
Let c = tblObj.ListColumns.Count
Dim aryary() As Variant
ReDim aryary(1, 1 To c)
''
C_Array.FillEmptySpaceForArrayArray aryary
GetNewRowArrayArray = aryary
''
End Function

Public Function GetRowArrayArray(ByRef tblObj As ListObject, _
                    ByVal rowIdx As Long) As Variant
'''' /********************************************************
''''  *@function GetRow
''''  *@param {ListObject} tblObj
''''  *@param {Long} rowIdx
''''  *
''''  */
''
GetRowArrayArray = tblObj.ListRows(rowIdx).Range.Value
''
End Function

Public Function GetRowArray(ByRef tblObj As ListObject, _
                    ByVal rowIdx As Long) As Variant
'''' /********************************************************
''''  *@function GetRow
''''  *@param {ListObject} tblObj
''''  *@param {Long} rowIdx
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
GetRowArray = C_Array.ArrayArrayToArray(tblObj.ListRows(rowIdx).Range.Value)
''
End Function

Public Sub AddRow(ByRef tblObj As ListObject, ByVal rowAryary As Variant)
'''' /********************************************************
''''  *@function AddRow
''''  *@param {ListObject} tblObj
''''  *@param {Variant()} rowAryary
''''  *
''''  */
''
Dim blnAutoExpandBuf As Boolean
Let blnAutoExpandBuf = Application.AutoCorrect.AutoExpandListRange
Dim rngInsert As Range
''
Application.AutoCorrect.AutoExpandListRange = True
If tblObj.InsertRowRange Is Nothing Then
    Set rngInsert = tblObj.ListRows.Add.Range.Resize(UBound(rowAryary))
Else
    Set rngInsert = tblObj.InsertRowRange
End If
''
rngInsert.Value = rowAryary
Application.AutoCorrect.AutoExpandListRange = blnAutoExpandBuf
''
End Sub

Public Sub DeleteRow(ByRef tblObj As ListObject, ByVal rowIdx As Long)
'''' /********************************************************
''''  *@function DeleteRow
''''  *@param {ListObject} tblObj
''''  *@param {Long} rowIdx
''''  *
''''  */
''
tblObj.ListRows(rowIdx).Range.Delete
''
End Sub

Public Sub UpdateRow(ByRef tblObj As ListObject, _
                    ByVal rowIdx As Long, _
                    ByVal rowAryary As Variant)
'''' /********************************************************
''''  *@function UpdateRow
''''  *@param {ListObject} tblObj
''''  *@param {Long} rowIdx
''''  *@param {Variant()} rowAryary
''''  *
''''  */
''
tblObj.ListRows(rowIdx).Range.Value = rowAryary
''
End Sub

Public Sub InsertRow(ByRef tblObj As ListObject, _
                    ByVal rowIdx As Long, _
                    ByVal rowAryary As Variant)
'''' /********************************************************
''''  *@function UpdateRow
''''  *@param {ListObject} tblObj
''''  *@param {Long} rowIdx '' 下に挿入
''''  *@param {Variant()} rowAryary
''''  *
''''  */
''
tblObj.ListRows(rowIdx).Range.insert
tblObj.ListRows(rowIdx).Range.Value = rowAryary
''
End Sub

Public Function GetNewColArray(ByRef tblObj As ListObject) As Variant
'''' /********************************************************
''''  *@function GetNewColArray
''''  *@param {String} tblName
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim r As Long
Let r = tblObj.ListRows.Count
Dim ary() As Variant
ReDim ary(1 To r)
''
C_Array.FillEmptySpaceForArray ary
GetNewColArray = ary
''
End Function

Public Function GetColArray(ByRef tblObj As ListObject, _
                    ByVal colIdx As Long) As Variant
'''' /********************************************************
''''  *@function GetColArray
''''  *@param {ListObject} tblObj
''''  *@param {Long} colIdx
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim tempAryary() As Variant
Let tempAryary = tblObj.ListColumns(colIdx).DataBodyRange.Value
Let tempAryary = C_Array.ArrayArrayToArray(tempAryary)
GetColArray = tempAryary
''
End Function

Public Sub AddColumn(ByRef tblObj As ListObject, ByVal colAry As Variant, _
                    ByVal colName As String)
'''' /********************************************************
''''  *@function AddColumn
''''  *@param {ListObject} tblObj
''''  *@param {Variant()} colAry
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim ary() As Variant
ReDim ary(0)
Let ary(0) = colName
C_Array.Concat ary, colAry
Dim colAryary() As Variant
Let colAryary = C_Array.ArrayToArrayArrayVertical(ary)
Dim blnAutoExpandBuf As Boolean
Let blnAutoExpandBuf = Application.AutoCorrect.AutoExpandListRange
''
Application.AutoCorrect.AutoExpandListRange = True
tblObj.ListColumns.Add.Range.Value = colAryary
Application.AutoCorrect.AutoExpandListRange = blnAutoExpandBuf
''
End Sub

Public Sub DeleteColumn(ByRef tblObj As ListObject, ByVal colIdx As Long)
'''' /********************************************************
''''  *@function DeleteRow
''''  *@param {ListObject} tblObj
''''  *@param {Long} rowIdx
''''  *
''''  */
''
tblObj.ListColumns(colIdx).Range.Delete
''
End Sub

Public Sub UpdateColumn(ByRef tblObj As ListObject, _
                    ByVal colIdx As Long, _
                    ByVal colAry As Variant, _
                    ByVal colName As String)
'''' /********************************************************
''''  *@function UpdateRow
''''  *@param {ListObject} tblObj
''''  *@param {Long} colIdx
''''  *@param {Variant()} colAry
''''  *@param {String} colName
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim ary() As Variant
ReDim ary(0)
Let ary(0) = colName
C_Array.Concat ary, colAry
Dim colAryary() As Variant
Let colAryary = C_Array.ArrayToArrayArrayVertical(ary)
tblObj.ListColumns(colIdx).Range.Value = colAryary
''
End Sub

Public Sub InsertColumn(ByRef tblObj As ListObject, _
                    ByVal colIdx As Long, _
                    ByVal colAry As Variant, _
                    ByVal colName As String)
'''' /********************************************************
''''  *@function InsertRow
''''  *@param {ListObject} tblObj
''''  *@param {Long} colIdx '' insert After colIndex
''''  *@param {Variant()} colAry
''''  *@param {String} colName
''''  *
''''  */
''
Dim C_Array As C_Array
Set C_Array = New C_Array
Dim ary() As Variant
ReDim ary(0)
Let ary(0) = colName
C_Array.Concat ary, colAry
Dim colAryary() As Variant
Let colAryary = C_Array.ArrayToArrayArrayVertical(ary)
tblObj.ListColumns(colIdx).Range.insert
tblObj.ListColumns(colIdx).Range.Value = colAryary
''
End Sub


Public Sub Sort(ByRef tblObj As ListObject, _
        ByVal sortIdx1 As Long, _
        Optional sortIdx2 As Variant, _
        Optional orderBy As String = "Asc")
'''' /********************************************************
''''  *@function Sort
''''  *@param {ListObject} tblObj
''''  *@param {Long} colIdx '' insert After colIndex
''''  *
''''  */
''
Dim order As Long
If VBA.LCase(orderBy) = "asc" Then
    order = xlAscending
ElseIf VBA.LCase(orderBy) = "desc" Then
    order = xlDescending
Else
    Exit Sub
End If
''
Dim rngSortColumn1 As Range
Dim rngSortColumn2 As Range
''
If IsMissing(sortIdx2) Then
    Set rngSortColumn1 = tblObj.ListColumns(sortIdx1).DataBodyRange
    With tblObj.Sort
        .SortFields.clear
        .SortFields.Add key:=rngSortColumn1, SortOn:=xlSortOnValues, order:=order
        .Header = xlYes
        .Apply
    End With
Else
    Set rngSortColumn1 = tblObj.ListColumns(sortIdx1).DataBodyRange
    Set rngSortColumn2 = tblObj.ListColumns(sortIdx2).DataBodyRange
    With tblObj.Sort
        .SortFields.clear
        .SortFields.Add key:=rngSortColumn1, SortOn:=xlSortOnValues, order:=order
        .SortFields.Add key:=rngSortColumn2, SortOn:=xlSortOnValues, order:=order
        .Header = xlYes
        .Apply
    End With
End If
''
End Sub

