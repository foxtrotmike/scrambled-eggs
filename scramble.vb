Sub ScrambledEggs()
' GPL v 2 Licensed Software.
' If you use it or redistribute it after changes, please be sure to include and acknowledge the
' original creator of this macro. Thanks!
' Scrambled Eggs Macro by Dr. Fayyaz ul Amir Afsar Minhas
' (afsar at pieas dot edu dot pk), Senior Scientist, PIEAS, Islamabad, Pakistan  April 8, 2015.
' Objective: Shuffle MCQ Option tables without affecting formatting
' Assumptions:
'   Each MCQ has a table of options
'   All rows in a table have the same number of columns and all columns have the same number of rows
'   All 1x1 tables are ignored
'   All options are numbered using Word's list styles (A. B.,...)
'   The valid answer for all questions in the original file is 'A'
' 	When the macro is running, the user doesn't change the active window
' To generate multiple scrambled sets, run this macro on a copy of the original file every time to get a valid key.
' The key is generated in the Immediate Window (Pressing Ctrl+G or from the menu "View > Immediate Window" in the VBA Editor.
' It works by initially doubling the number of rows in each table and then copying cells randomly to the newly
' created rows and then deleting the original option rows.

    Dim i As Integer, J As Integer, tt As Integer, ttidx As Integer
    Dim tbl As Table
    Dim maxRow As Long
    Dim maxCol As Long
    If ActiveDocument.Tables.Count = 0 Then
        MsgBox "No Eggs to scramble."
        Exit Sub
    End If
    MsgBox ("Number of Tables Detected: " & ActiveDocument.Tables.Count)
    ttidx = 0
    Randomize Now
    For tt = 1 To ActiveDocument.Tables.Count
        ' For each table in the currently active document
        
        Set tbl = ActiveDocument.Tables(tt)
        maxRow = tbl.Rows.Count
        maxCol = tbl.Columns.Count
        If maxRow * maxCol > 1 Then
            ttidx = ttidx + 1
            Dim Indices() As Integer
            Dim Indices2D() As Integer
            ReDim Indices(maxRow * maxCol - 1) 'holds linear indices
            'printArray InArray:=Indices
            ReDim Indices2D(maxRow * maxCol - 1, 2) ' holds cell indices
            For i = 0 To maxRow - 1
                For J = 0 To maxCol - 1
                    Indices2D(J + (i) * maxCol, 0) = i + 1
                    Indices2D(J + (i) * maxCol, 1) = J + 1
                    Indices(J + (i) * maxCol) = J + (i) * maxCol
                    'Debug.Print i, J, J + (i) * maxCol
                Next J
            Next i
            
            'printArray InArray:=Indices
            PermuteArray InArray:=Indices  ' Shuffle the linear indices
            'printArray InArray:=Indices
                            
            For i = 1 To maxRow:
                Set rowNew = tbl.Rows.Add() 'Add dummy rows
            Next i
            
            For i = LBound(Indices) To UBound(Indices) 'copy the scrambled rows
                'Debug.Print k
                'Debug.Print Indices2D(k, 0), Indices2D(k, 1), Indices2D(Indices(k), 0), Indices2D(Indices(k), 1)
                tbl.Cell(Indices2D(i, 0), Indices2D(i, 1)).Range.Copy
                tbl.Cell(maxRow + Indices2D(Indices(i), 0), Indices2D(Indices(i), 1)).Range.PasteAndFormat (wdFormatOriginalFormatting)
            Next i
            
            For i = 1 To maxRow: 'delete original rows
                tbl.Rows(1).Delete
            Next i
            'printing the key
            Debug.Print ttidx, tbl.Cell(Indices2D(Indices(0), 0), Indices2D(Indices(0), 1)).Range.ListFormat.ListString
        End If
        
    Next tt
    MsgBox ("Scrambling Complete! Number of tables scrambled: " & ttidx)
End Sub
Sub PermuteArray(ByRef InArray() As Integer)
    ''''''''''''''''''''''''''''''''''''
    ' Permutes the given Array in-place
    ' Part of the ScrambledEggs
    ''''''''''''''''''''''''''''''''''''
    Dim N As Long
    Dim Temp As Variant
    Dim J As Long
    For N = LBound(InArray) To UBound(InArray)
        J = CLng(((UBound(InArray) - N) * Rnd) + N)
        If N <> J Then
            Temp = InArray(N)
            InArray(N) = InArray(J)
            InArray(J) = Temp
        End If
    Next N
End Sub
Sub printArray(ByRef InArray() As Integer)
    ''''''''''''''''''''''''''''''''''''
    ' Display 1D array
    ' Part of the ScrambledEggs
    ''''''''''''''''''''''''''''''''''''
    Debug.Print "--------------"
    Dim i As Integer
    For i = LBound(InArray) To UBound(InArray)
        Debug.Print InArray(i)
    Next i
    Debug.Print "--------------"
End Sub
