For g = 1 To 200
    If CallByName(CodeCollection(i), "row" & g, VbGet) <> "" Then
        RowItems = Split(CallByName(CodeCollection(i), "row" & g, VbGet), "*")
        numNewRows = UBound(RowItems)
        
        ' Check if the first item starts with LI# or LIA#
        Dim isValidRow As Boolean
        Dim regex As Object
        Set regex = CreateObject("VBScript.RegExp")
        
        isValidRow = False
        
        If numNewRows >= 0 Then
            For yy = 0 To numNewRows
                splitArray = Split(RowItems(yy), "|")
                If UBound(splitArray) >= 0 Then
                    ' Set up regex pattern for LI# or LIA#
                    regex.Global = False
                    regex.IgnoreCase = True
                    
                    ' Check for LI# pattern
                    regex.Pattern = "^LI[0-9]+\|"
                    If regex.Test(splitArray(0)) Then
                        isValidRow = True
                        Exit For
                    End If
                    
                    ' Check for LIA# pattern
                    regex.Pattern = "^LIA[0-9]+\|"
                    If regex.Test(splitArray(0)) Then
                        isValidRow = True
                        Exit For
                    End If
                End If
            Next yy
        End If
        
        ' Only proceed if it's a valid row
        If isValidRow Then
            ' Insert new rows below the current checkrow + g - 1
            If numNewRows > 0 Then
                Rows(checkrow + g & ":" & checkrow + g + numNewRows - 1).Insert Shift:=xlDown
                
                ' Convert cell references to absolute row references using regex
                Dim cell As Range
                Dim formula As String
                Dim regex As Object
                Set regex = CreateObject("VBScript.RegExp")
                
                regex.Global = True
                regex.IgnoreCase = True
                ' Modified pattern to match cell references but capture the column letter separately
                regex.Pattern = "([A-Z]+)([0-9]+)"
                
                For Each cell In Rows(checkrow + g - 1).Cells
                    If Left(cell.formula, 1) = "=" Then
                        formula = cell.formula
                        ' Use a custom replacement function that checks the column letter
                        cell.formula = RegExReplace(formula, regex)
                    End If
                Next cell
                
                Rows(checkrow + g - 1).Copy
                Rows(checkrow + g & ":" & checkrow + g + numNewRows - 1).PasteSpecial Paste:=xlPasteFormats
                Rows(checkrow + g & ":" & checkrow + g + numNewRows - 1).PasteSpecial Paste:=xlPasteFormulas
                Application.CutCopyMode = False
            End If

            
            ' Populate each row
            For yy = 0 To numNewRows
                splitArray = Split(RowItems(yy), "|")
                
                For x = LBound(splitArray) To UBound(splitArray)
                    If splitArray(x) <> "" And splitArray(x) <> "F" Then
                        Cells(checkrow + g - 1 + yy, columnSequence(x)).value = splitArray(x)
                    End If
                Next x
            Next yy
            
            ' Adjust checkrow to account for added rows
            checkrow = checkrow + numNewRows
        End If
    End If
Next g