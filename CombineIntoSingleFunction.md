# CombineIntoSingleFunction

> Made by me asking a bunch of questions to my friend Copilot.

This function takes a cell that contains a formula and returns the same formula but expands the inner referenced cells that contains another formulas leaving only the references of cells that contains values.

This is useful to consctruct a complex formula from multiple
simple formulas. Making it possible to automatically combine
everything in a single one.

```VBScript
Function CombineIntoSingleFunction(cell As Range) As String
    Dim formula As String
    Dim matches As Object
    Dim regex As Object
    Dim result As String
    Dim i As Integer
    Dim refCell As Range
    Dim refFormula As String
   
    ' Get the formula from the cell
    formula = cell.formula
   
    ' Create a regex object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "(\$?[A-Za-z]+\$?\d+)"
   
    ' Find all matches
    Set matches = regex.Execute(formula)
   
    ' Replace cell references with their formulas if they contain formulas
    For i = 0 To matches.Count - 1
        Set refCell = Range(matches(i).Value)
        If refCell.HasFormula Then
            refFormula = refCell.formula
            ' Recursively expand the nested formula
            refFormula = CombineIntoSingleFunction(refCell)
            ' Remove the '=' sign from the nested formula
            If Left(refFormula, 1) = "=" Then
                refFormula = Mid(refFormula, 2)
            End If
            formula = Replace(formula, matches(i).Value, "(" & refFormula & ")")
        End If
    Next i
   
    ' Replace commas with semicolons (may be needed depending on the configured functions' arguments separator)
    'formula = Replace(formula, ",", ";")
   
    ' Return the expanded formula
    CombineIntoSingleFunction = formula
End Function
```
