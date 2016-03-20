# StockDataCompiler
Excel Workbooks that get stock info using various API's and URL queries in VBA
Open New Workbook with one worksheet named "Stocks"
#Copy and paste 
[the code]:(StockDataCompiler/README.md)
[copy me]       (http://raw.githubusercontent.com/HotBreakfast/StockDataCompiler/master/GetAllStockInfoCode "This Code") into a new module
##In the workbook module paste the following:

Private Sub Workbook_Open()
Dim MyCode As String
MyCode = MsgBox("Run Code? ", vbYesNo)
If MyCode = vbYes Then
        IDotheCounting
If MyCode = vbNo Then
        Sheets("Stocks").Select
End If
End If
End Sub

''close and save workbook as .xlms, open the workbook and it will ask you if you would like to run the code _
<P/>''***Upon opening the workbook you will know if you have done it correctly if you are asked a question to run the code
###run the code and wait 25 minutes to get all stock data***</P>
