Sub StockAnalysis():

	'---------------------------------
	'Declare variables
	'------------------------

    Dim r, c, counter, i As Integer
    Dim lastRow As Long
    Dim header As Variant
    Dim openAmount, totalStockVol As Double
    
	'---------------------------------
	'Loop through all worksheets
	'--------------------
  
	For i = 1 To Worksheets.Count

        Dim ws As Worksheet
        Set ws = Worksheets(i)
  
		'------------------------------------------------------------
		'Sort the data in ascending order by ticker and then date.
		'----------------------------------------------------
	
	    ws.Range("A:G").Sort _
	        Key1:=ws.Range("A1"), _
	        Order1:=xlAscending, _
	        Key2:=ws.Range("B1"), _
	        Order2:=xlAscending, _
	        header:=xlYes
	
		'----------------------------------------------
		'Count number of non-blank rows in column A.
		'----------------------------------
	
	    lastRow = Application.WorksheetFunction.CountA(ws.Range("A:A"))
	
		'---------------------------------
		'Create headers for new table.
		'------------------------
	
	    header = Array("Ticker", "Year Change", "Year Change (%)", "Total Stock Vol.")
	    ws.Range("J1:M1").Value = header
	    ws.Range("J1:M1").Font.Bold = True
	
		'------------------------
		'set counter equal to 2
		'---------------
	    
	    counter = 2
	
		'------------------------------------
		'Loop through intial data set
		'-------------------
	
	    For r = 2 To lastRow
	
			'-------------------------------------------------------------------------------------------------------------------
			'If previous <ticker> value (col A) is not equal to current <ticker> value, then save the <open> value (col C).
			'-------------------------------------------------------------------------------------------
	
	        totalStockVol = totalStockVol + ws.Cells(r, 7).Value
	        
	        If ws.Cells(r - 1, 1).Value <> ws.Cells(r, 1).Value Then
	            openAmount = ws.Cells(r, 3)
	            
		 	'-------------------------------------------------------------------------------------------------
		 	'If  following <ticker> value (col A) is not equal to current <tikcer> value, then
		 	'Add the <ticker> and the difference of the close and open amounts to the new table (col J-K)
		 	'------------------------------------------------------------------------
	            
	        ElseIf ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
	    
	            ws.Cells(counter, 10).Value = ws.Cells(r, 1).Value
	            ws.Cells(counter, 11).Value = ws.Cells(r, 6).Value - openAmount
	         
				'---------------------------------------------------------------------------------------------------------------
				'Calculate percent differnce of close and open amount to the new table (col L). Skip if open value is zero.
				'-------------------------------------------------------------------------------------
	            
	            If openAmount <> 0 Then
	                ws.Cells(counter, 12).Value = (ws.Cells(r, 6).Value - openAmount) / openAmount
	            End If
	            
				'----------------------
				'Sum stock volume
				'--------------
	
	            ws.Cells(counter, 13).Value = totalStockVol
	            
				'-----------------------------------------------------------------------
				'Highlight row red if price decreased and green if price increased
				'-----------------------------------------------
	            
	            If ws.Cells(counter, 11).Value < 0 Then
	                
	                ws.Range("J" + CStr(counter) + ":M" + CStr(counter)).Interior.ColorIndex = 22    '(Red highlight)
	                    
	            ElseIf ws.Cells(counter, 11).Value > 0 Then
	                
	                ws.Range("J" + CStr(counter) + ":M" + CStr(counter)).Interior.ColorIndex = 35    '(Green highlight)
	                
	            End If
	                            
				'-------------------------------------------------------------
				'Increase counter by one and reset stock volume total
				'--------------------------------------
	                            
	            counter = counter + 1
	            totalStockVol = 0
	        
	        End If
	            
	    Next r  'Next row
	        
		'-----------------------------------
		'format values in new table
		'---------------------
	        
	    ws.Range("K:K").Style = "Currency"
	    ws.Range("L:L").Style = "Percent"
	    ws.Range("M:M").Style = "Comma"
	       
		'--------------------------------------
		'Create headers for third table
		'--------------------------
	    
	    ws.Range("Q1:R1") = Array("Ticker", "Value")
	    ws.Range("Q1:R1").Font.Bold = True
	    
		'--------------------------------------------------------------------------
		'Sort second table by price descending and select greatest increase
		'------------------------------------------------
	    
	    ws.Range("J:M").Sort _
	        Key1:=ws.Range("L1"), _
	        Order1:=xlDescending, _
	        header:=xlYes
	
	    ws.Cells(2, 16).Value = "Greatest Increase (%)"
	    ws.Cells(2, 16).Font.Bold = True
	    ws.Cells(2, 17).Value = ws.Cells(2, 10).Value
	    ws.Cells(2, 18).Value = ws.Cells(2, 12).Value
	    ws.Cells(2, 18).Style = "Percent"
	   
		'--------------------------------------------------------------------------
		'Sort second table by price ascedning and select greatest decrease
		'-------------------------------------------------
	   
	    ws.Range("J:M").Sort _
	        Key1:=ws.Range("L1"), _
	        Order1:=xlAscending, _
	        header:=xlYes
	    
	    ws.Cells(3, 16).Value = "Greatest Decrease (%)"
	    ws.Cells(3, 16).Font.Bold = True
	    ws.Cells(3, 17).Value = ws.Cells(2, 10).Value
	    ws.Cells(3, 18).Value = ws.Cells(2, 12).Value
	    ws.Cells(3, 18).Style = "Percent"
	    
	    ws.Range("J:M").Sort _
	        Key1:=ws.Range("M1"), _
	        Order1:=xlDescending, _
	        header:=xlYes
	   
		'--------------------------------------------------------------------------
		'Sort second table by total stock volume descending and select largest
		'------------------------------------------------
	   
	    ws.Cells(4, 16).Value = "Greatest Total Vol."
	    ws.Cells(4, 16).Font.Bold = True
	    ws.Cells(4, 17).Value = ws.Cells(2, 10).Value
	    ws.Cells(4, 18).Value = ws.Cells(2, 13).Value
	    ws.Cells(4, 18).Style = "Comma"
	    
	    ws.Range("J:M").Sort _
	        Key1:=ws.Range("J1"), _
	        Order1:=xlAscending, _
	        header:=xlYes
	
        '-------------------------
        'Autofit columns width
        '----------------
	
	    ws.Range("A:R").Columns.AutoFit


	Next i  'Next worksheet

End Sub