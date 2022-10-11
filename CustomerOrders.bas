Attribute VB_Name = "CustomerOrders"
Option Explicit

'==============================================+
' Description: Asks user for a total amount spent, then searches the Data
'   worksheet for users who spent over the total, and finally reporting these
'   overspent amounts in a new worksheet called Report.
'==============================================+

Public Sub Main()
    'Declaring variables
    Dim userInput As Variant
    Dim isValid As Boolean
    Dim ws As Worksheet
    Dim wsReport As Worksheet
    Dim dataCounter As Integer
    Dim reportCounter As Integer
    Dim cell As Range
    Dim customerID As String
    Dim total As Integer
    
    'Initializing variables
    isValid = False
    
    'Check if a worksheet named Report already exists in workbook -- if yes, delete it
    For Each ws In ThisWorkbook.Worksheets 'Runs through the worbook
        If ws.Name = "Report" Then
                Application.DisplayAlerts = False 'Turns off display alert
                ws.Delete
                Application.DisplayAlerts = True 'Turns back on display alert
        End If
    Next
    
'Checks if input is valid -- program continues to ask user for input until user input is valid
    Do
        userInput = InputBox("Enter a total amount spent:", "Total amount spent") 'Prompts user for input
        If IsNumeric(userInput) = True And Val(userInput) > 0 Then 'If input is valid
            isValid = True
        Else
            MsgBox "Please input a numeric value greater than $0.", vbOKOnly, "Error" 'If input is invalid
        End If
    Loop Until isValid = True 'Program stops asking user for total amount when input is valid

    'Creates new worksheet called Report
    Set wsReport = Worksheets.Add(after:=Worksheets(Worksheets.Count)) 'After the Data worksheet
    wsReport.Name = "Report" 'Gives it a name that users can see
    
    'Formatting the Report worksheet
    With wsReport.Range("A1")
        .Value = "Customers who spent more than $" & userInput
        .Font.Bold = True
    End With
    wsReport.Range("A3").Value = "Customer ID"
    wsReport.Range("B3").Value = "Total amount spent"
    wsReport.Range("A3", "B3").Columns.AutoFit
    
    'Sort entries in Data worksheet by Customer ID
    With wsData
        .Range("A3", .Range("A3").End(xlDown).End(xlToRight)).Sort Key1:=.Range("B3"), Order1:=xlAscending, Header:=xlYes
    End With
    
    'Run through amount purchased column of Data worksheet
    'Initialize tracking variables:
    dataCounter = 4 'Keeps track of row in Data
    reportCounter = 4 'Keeps track of row in Report
    customerID = wsData.Range("B" & dataCounter).Value
    total = wsData.Range("C" & dataCounter).Value
    dataCounter = 5 'Update row in Data to 5
    With wsData
        '.Activate 'NOTE: Using the For i To cell count loop is more robust than activating worksheet!!
        For Each cell In .Range("C5", .Range("C5").End(xlDown)) 'Run through amount column
            If .Range("B" & dataCounter).Value = customerID Then 'Customer ID is same as previous -- sum the amounts
                total = total + cell.Value
            ElseIf Val(total) > Val(userInput) Then 'Customer ID is different -- when previous ID's total is greater than user input, add info to Report
                wsReport.Range("A" & reportCounter).Value = customerID
                wsReport.Range("B" & reportCounter).Value = total
                reportCounter = reportCounter + 1 'Move onto next row in Report
                total = .Range("C" & dataCounter).Value 'Update total with current cell amount
            Else
                total = .Range("C" & dataCounter).Value 'Update total with current cell amount
            End If
            customerID = .Range("B" & dataCounter).Value 'Update ID associated with current cell
            dataCounter = dataCounter + 1 'Move onto next row in Report
        Next
    End With
    
    'Restore Data worksheet to orders being sorted by date
    With wsData
        .Range("A3", .Range("A3").End(xlDown).End(xlToRight)).Sort Key1:=.Range("A3"), Order1:=xlAscending, Header:=xlYes
    End With
    
'Sort entries in Report worksheet in descending order by total amount spent
    With wsReport
        .Range("A3", .Range("A3").End(xlDown).End(xlToRight)).Sort Key1:=.Range("B3"), Order1:=xlDescending, Header:=xlYes
        .Range("B4", .Range("B4").End(xlDown)).NumberFormat = "$#,##" 'Formatting dollar display in amount purchased cell
    End With
    
End Sub


Public Sub Reset()
'Clears cells that were edited by task
    Worksheets("Report").UsedRange.Clear
End Sub




