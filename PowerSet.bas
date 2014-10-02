Attribute VB_Name = "Module1"
Option Explicit
 
' PGC Oct 2007
' Calculates a Power Set
' Set in A1, down. Result in C1, down and accross. Clears C:Z.
Sub PowerSet()
Dim vElements As Variant, vresult As Variant, vNumbers As Variant, vMedias As Variant
Dim lRow As Long, nRow As Long, i As Long
Dim ws As Worksheet
Dim match As Range, off As Range, off2 As Range
Dim match2 As Range
Dim findMe As String


Set ws = ThisWorkbook.Sheets("Dados")

findMe = "Rating"
Set match = ws.Cells.Find(findMe, , , xlWhole)
Set off = match.Offset(1, 0)

findMe = "Field"
Set match2 = ws.Cells.Find(findMe)
Set off2 = match2.Offset(1, 0)

vElements = Application.Transpose(Range(off2, off2.End(xlDown)))
vNumbers = Application.Transpose(Range(off, off.End(xlDown)))


Application.DisplayAlerts = False
On Error Resume Next
Sheets("Combinacoes").Delete
Sheets.Add.Name = "Combinacoes"
Sheets("Combinacoes").Select
Cells(1, 1).Value = "ID"
Cells(1, 2).Value = "N Combinação"
Cells(1, 3).Value = "Fields"
Cells(1, 4).Value = "Média"
 
lRow = 1
nRow = 1
For i = 1 To UBound(vElements)
    ReDim vresult(1 To i)
    ReDim vMedias(1 To i)
    Call CombinationsNP(vElements, i, vresult, lRow, 1, 1)
    Call MediasNP(vNumbers, i, vMedias, nRow, 1, 1)
Next i
End Sub
 
Sub CombinationsNP(vElements As Variant, p As Long, vresult As Variant, lRow As Long, iElement As Integer, iIndex As Integer)
Dim i As Long
Dim j As Long
Dim a As String
 
For i = iElement To UBound(vElements)
    vresult(iIndex) = vElements(i)
    If iIndex = p Then
        lRow = lRow + 1
        a = ""
        For j = 1 To p
        a = a & vresult(j) & ","
        Next j
        Range("A" & lRow) = lRow - 1
        Range("B" & lRow) = p
        Range("C" & lRow) = a
    Else
        Call CombinationsNP(vElements, p, vresult, lRow, i + 1, iIndex + 1)
    End If
Next i
End Sub

Sub MediasNP(vNumbers As Variant, p As Long, vMedias As Variant, lRow As Long, iElement As Integer, iIndex As Integer)
Dim i As Long
Dim j As Long
Dim a As Double
 
For i = iElement To UBound(vNumbers)
    vMedias(iIndex) = vNumbers(i)
    If iIndex = p Then
        lRow = lRow + 1
        a = 0
        For j = 1 To p
        a = a + vMedias(j)
        Next j
        a = a / (j - 1)
        Range("D" & lRow) = a
    Else
        Call MediasNP(vNumbers, p, vMedias, lRow, i + 1, iIndex + 1)
    End If
Next i
End Sub
