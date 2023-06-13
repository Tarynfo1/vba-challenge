# vba-challenge
Module 2 assignment

This is the first time I have ever used VBA so I relied on many resources, including how to format acknowledgements on github. I was unable to complete the Bonus question on this assignment due to my beginner level of VBA.

## Acknowledgements

The code implementation for sending output to separate worksheet was assisted by AskBCS LA, Richie Garafola. The provided code snippet helped in sending output to a separate worksheet.

### Code Snippet

``
' ...

'  Set outputWs = ThisWorkbook.Worksheets.Add
'  Set the parameters of the For Loop
'  Dim outputRow As Long    
' Variable for output row counter
' outputRow = 2
' and
'' Increment the output row counter
'outputRow = outputRow + 1

' ...


The insights and code helped enhance the functionality and efficiency of the assignment. 

The code implementation for populating the column with the Price_Change_Percent and Yearly_Change was assisted by AskBCS LA, Shreha (unknown lastname). The provided code snippet helped in fixing an error where only 0 was populating the column.

### Code Snippet

``
' ...

' Calculate the Yearly Change and Percentage Change
            'Yearly_Change = close_price - open_price
            'If open_price <> 0 Then
                'Price_Change_Percent = (Yearly_Change / open_price) * 100
            'Else
                'Price_Change_Percent = 0
            'End If

' ...

The insights and code helped enhance the functionality and efficiency of the assignment. 


The code implementation for conditional formatting was assisted by https://www.wallstreetmojo.com/vba-conditional-formatting/#h-example-1. The provided code snippet helped in using this as a template for my own formatting.

### Code Snippet

``
' ...

' Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "=80")
 'Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "=50")
   'Defining and setting the format to be applied for each condition
  ' With condition1
   ' .Font.Color = vbBlue
   ' .Font.Bold = True
  ' End With

  ' With condition2
     '.Font.Color = vbRed
      '.Font.Bold = True
   'End With

' ...

'The insights and code helped enhance the functionality and efficiency of the assignment. 

The code implementation for colour conditions was assisted by TA Imaad Fakier. The provided code snippet helped in applying conditional formatting.

### Code Snippet

``
' ...
 'Apply colour conditions to variables
            Select Case Yearly_Change
                Case Is > 0
                    outputWs.Range("J" & 2 + i).Interior.ColorIndex = 4
                Case Is < 0
                    outputWs.Range("J" & 2 + i).Interior.ColorIndex = 3
                Case Else
                    outputWs.Range("J" & 2 + i).Interior.ColorIndex = 0
            End Select

' ...

'The insights and code helped enhance the functionality and efficiency of the assignment. 
