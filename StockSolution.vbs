Sub Homework2()

'Declaration and Initialization of Variables--------------------------------------------
Dim Ticker As String
Dim Change, DayChange, DiffCalc, DiffCalcDaily, Volume, VolTot As Double
Dim i, j, Days As Integer
Dim MaxVol, GInc, GDec, GDaily As Double
Dim MaxTic, GIncTic, GDecTic, GDailyTic As String

i = 2: j = 2: Days = 1
MaxVol = 0: GInc = 0: GDec = 0: GDaily = 0

'Disable screen updating to save processing power----------------------------------------
Application.ScreenUpdating = False

'Create and format new sheet-------------------------------------------------------------
Sheets.Add
ActiveSheet.Name = "Data Dump"

Sheets("Data Dump").Cells(1, "A") = "Ticker"
Sheets("Data Dump").Cells(1, "B") = "Total Change"
Sheets("Data Dump").Cells(1, "C") = "% Change"
Sheets("Data Dump").Cells(1, "D") = "Average Daily Change"
Sheets("Data Dump").Cells(1, "E") = "Total Volume"
Sheets("Data Dump").Cells(1, "I") = "Ticker"
Sheets("Data Dump").Cells(2, "G") = "Total Shares"
Sheets("Data Dump").Cells(5, "G") = "Greatest % Increase"
Sheets("Data Dump").Cells(8, "G") = "Greatest % Decrease"
Sheets("Data Dump").Cells(10, "G") = "Greatest Daily Avg."


Sheets("Data Dump").Rows(1).EntireRow.Font.Bold = True
Sheets("Data Dump").Columns("G").EntireColumn.Font.Bold = True
Sheets("Data Dump").Rows(1).EntireRow.HorizontalAlignment = xlCenter


'BEGIN OFFICIAL CALUCLATIONS-------------------------------------------------------------

'Loop until blank cell in Column A of "Stock_data_2016"
Do While Not IsEmpty(Sheets("Stock_data_2016").Cells(i, "A"))

    'Takes value of Ticker out of i-th cell in Column A
    Ticker = Sheets("Stock_data_2016").Cells(i, "A")
    
    'Temp variable stores single instance difference, and then combines it to the total amount.
    Change = Sheets("Stock_data_2016").Cells(i, "F") - Sheets("Stock_data_2016").Cells(i, "C")
    DiffCalc = DiffCalc + Change
    
    DayChange = Sheets("Stock_data_2016").Cells(i, "D") - Sheets("Stock_data_2016").Cells(i, "E")
    DiffCalcDaily = DiffCalcDaily + DayChange
    
    Volume = Sheets("Stock_data_2016").Cells(i, "G")
    VolTot = VolTot + Volume
    
    'compares Ticker to value below it
    If Ticker <> Sheets("Stock_data_2016").Cells(i + 1, "A") Then
    
        
        'The Four If Statements check and store the highest value for each column
        If VolTot > MaxVol Then
            MaxVol = VolTot
            MaxTic = Ticker
        End If
        
    
        If Round((DiffCalc / Days) * 100, 2) > GInc Then
            GInc = Round((DiffCalc / Days) * 100, 2)
            GIncTic = Ticker
        End If
        
        
        If Round((DiffCalc / Days) * 100, 2) < GDec Then
            GDec = Round((DiffCalc / Days) * 100, 2)
            GDecTic = Ticker
        End If
        
        
        If GDaily < (DiffCalcDaily / Days) Then
            GDaily = DiffCalcDaily / Days
            GDailyTic = Ticker
        End If
    
        'Sets data and then reinitializes it
        Sheets("Data Dump").Cells(j, "A") = Ticker
        
        Sheets("Data Dump").Cells(j, "B") = DiffCalc
        
        'Color Total Change column based on its value
        If DiffCalc > 0 Then
            Sheets("Data Dump").Cells(j, "B").Interior.ColorIndex = 4
        ElseIf DiffCalc < 0 Then
            Sheets("Data Dump").Cells(j, "B").Interior.ColorIndex = 3
        Else
            Sheets("Data Dump").Cells(j, "B").Interior.ColorIndex = 16
        End If
        
        
        Sheets("Data Dump").Cells(j, "C") = Round((DiffCalc / Days) * 100, 2) & "%"
        DiffCalc = 0
        
        Sheets("Data Dump").Cells(j, "D") = DiffCalcDaily / Days
        DiffCalcDaily = 0
        
        Sheets("Data Dump").Cells(j, "E") = VolTot
        VolTot = 0
        
        Days = 1
        
        'Increments variables
        j = j + 1

    End If
    'Increments variables
    Days = Days + 1
    i = i + 1
Loop

Sheets("Data Dump").Cells(2, "H") = MaxVol
Sheets("Data Dump").Cells(2, "I") = MaxTic

Sheets("Data Dump").Cells(5, "H") = GInc
Sheets("Data Dump").Cells(5, "I") = GIncTic

Sheets("Data Dump").Cells(8, "H") = GDec
Sheets("Data Dump").Cells(8, "I") = GDecTic

Sheets("Data Dump").Cells(10, "H") = GDaily
Sheets("Data Dump").Cells(10, "I") = GDailyTic



Sheets("Data Dump").Columns("A:H").EntireColumn.AutoFit

'Resumes Screen Updating
Application.ScreenUpdating = True

'Notifies user of application completion
MsgBox ("Application has run successfully. " & j - 2 & " Different Tickers.")
End Sub