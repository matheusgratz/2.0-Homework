Attribute VB_Name = "Macros"
Option Explicit

'--------
'Homework 2
'Student: Matheus Gratz
'Objectives:

'Create a script that will loop through all the stocks for one year
'and output the following information.
'
'   The ticker symbol.
'
'   Yearly change from opening price at the beginning of a given year
'   to the closing price at the end of that year.
'
'   The percent change from opening price at the beginning of a given
'   year to the closing price at the end of that year.
'
'   The total stock volume of the stock.
'
'You should also have conditional formatting that will highlight positive
'change in green and negative change in red.
'--------


Sub RunOverSheets()

'define and set local parameters:  worksheet
Dim ws As Worksheet

Application.ScreenUpdating = False

'loop through all worksheets
For Each ws In Worksheets
    ws.Activate
    
'call macro
    Call Macros.RunOverStockList
    MsgBox ("Sheet " + ws.Name + " Done! Press OK to run over the remaining Sheets")
Next

Application.ScreenUpdating = True

MsgBox ("Everything is Done!")

End Sub


Sub RunOverStockList()

'define and set local parameters: workbook and worksheet

    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    Set ws = ActiveSheet

'define variables
    Dim ticker As String                '-- ticker name
    Dim oPrice, cPrice As Double        '-- Open Price, Close Price
    Dim sVolume As Variant              '-- Total Negotiated Volume
    Dim yChange As Variant              '-- Yearly Change
    Dim pChange As Variant              '-- Percent Change
    
    Dim gIncV, gDecV, gVolV As Variant  '-- Greatest values
    Dim gIncT, gDecT, gVolT As String   '-- Greatest tickers
            
    Dim i, j As Variant                 '-- Loop variables
    Dim aStock As String                '-- Stock name evaluated
    Dim iDate, eDate As Long            '-- Initial and End Date
    Dim slastRow, llastRow As Variant   '-- Last rows large and small lists
    Dim sPoint As Long                  '-- Where I stopped after a ticker has changed
    
Application.ScreenUpdating = False
    
'set model total rows
    slastRow = 0
    slastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

'clear everything, just in case
    Columns("J:Q").Select
    Selection.Clear
    Range("J1").Select

'generate headers
    With ws.Range("J1")
        .Value = "Ticker"
        .EntireColumn.AutoFit
    End With
    
    With ws.Range("K1")
        .Value = "Yearly Change"
        .EntireColumn.AutoFit
    End With
    
    With ws.Range("L1")
        .Value = "Percent Change"
        .EntireColumn.AutoFit
    End With
    
    With ws.Range("M1")
        .Value = "Total Stock Volume"
        .EntireColumn.AutoFit
    End With

'CHALLENGE = generate headers
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    
    With ws.Range("O4")
        .Value = "Greatest Total Volume"
        .EntireColumn.AutoFit
    End With

    With ws.Range("P1")
        .Value = "Ticker"
        .EntireColumn.AutoFit
    End With

    With ws.Range("Q1")
        .Value = "Value"
        .EntireColumn.AutoFit
    End With


'generate list of stocks and remove duplicates

ws.Range("A2:A" & slastRow).Select
Selection.Copy
ws.Range("J2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
ws.Range("J2:J" & slastRow).RemoveDuplicates Columns:=1, Header:=xlNo

'set how many unique tickers we will evaluate
llastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

'set initial starting point
sPoint = 2

'CHALLENGE - set zeros
gIncV = 0
gDecV = 0
gVolV = 0

'Loop through small list
For i = 2 To llastRow

    aStock = Range("J" & i).Value
    iDate = 99999999
    eDate = 0
    sVolume = 0
    
    'Loop through large list
        For j = sPoint To slastRow
            If Cells(j, 1).Value = aStock Then
                
                'establish the initial date
                If Cells(j, 2).Value < iDate Then
                iDate = Cells(j, 2).Value
                oPrice = Cells(j, 3).Value
                    'establish the end date
                    ElseIf Cells(j, 2).Value > eDate Then
                    eDate = Cells(j, 2).Value
                    cPrice = Cells(j, 6).Value
                End If
            
            sVolume = sVolume + Cells(j, 7) 'acumulate volume
            sPoint = j + 1                  'define point
            Else: Exit For
            
            End If
        Next j
        
    'calculate yearly change
    yChange = cPrice - oPrice
    Range("K" & i).Value = yChange
    
    'calculate percent change
    If oPrice <> 0 Then
    pChange = yChange / oPrice
    Else: pChange = 0
    End If
    
    'format percent change
    Range("L" & i).Value = pChange
    With Range("L" & i)
            .Style = "Percent"
            .NumberFormat = "0.00%"
            
            If pChange < 0 Then
                .Interior.Color = 255
                Else
                .Interior.Color = 5287936
            End If
            .Font.ThemeColor = xlThemeColorDark1
            .Font.TintAndShade = 0
            .Font.Bold = True
    End With
    
    'store total volume
    Range("M" & i).Value = sVolume
    
    '-----------------
    '   CHALLENGE
    '-----------------
    If pChange > gIncV Then
        gIncV = pChange
        gIncT = aStock
    End If
    
    If pChange < gDecV Then
        gDecV = pChange
        gDecT = aStock
    End If
    
    If sVolume > gVolV Then
        gVolV = sVolume
        gVolT = aStock
    End If
    
Next i

    '-----------------
    '   CHALLENGE
    '-----------------
    Range("P2").Value = gIncT
    Range("Q2").Value = gIncV
    
    Range("P3").Value = gDecT
    Range("Q3").Value = gDecV
    
    Range("P4").Value = gVolT
    Range("Q4").Value = gVolV
    
    With Range("Q2:Q3")
        .Style = "Percent"
        .NumberFormat = "0.00%"
        .EntireColumn.AutoFit
    End With

Application.ScreenUpdating = True

'Msgbox("Done!)

End Sub



