Option Explicit

Private	ResponseDiameters(10) as Double
Private	ResponsePersonCounts(10) as Double
Private	Doc as Object
Private	MainSheet as Object

Sub Main

End Sub


Sub CalcCakeDiameters

	Dim FormType as String
	Dim AskedPersonCount as Integer
	
	Doc = ThisComponent
	MainSheet = Doc.Sheets.getByName("CALC")
	
	FormType = MainSheet.getCellRangeByName("VORM").String
	AskedPersonCount = MainSheet.getCellRangeByName("PERSONEN").Value
	
	Rem clear previous result
	MainSheet.getCellRangeByName("B4").String = ""
	MainSheet.getCellRangeByName("B7:C16").clearContents(com.sun.star.sheet.CellFlags.STRING+com.sun.star.sheet.CellFlags.VALUE)
	
	Dim Found as Boolean
	Found = DoCalcCakeDiameters(FormType, 10,  AskedPersonCount)
	
	If (Found = False) Then
		Found = DoCalcCakeDiameters(FormType, 12,  AskedPersonCount)
	End If
	

End Sub


Rem Fill in the global RespondeDiameters variable
Rem Returns True if a solution is found
Function DoCalcCakeDiameters(FormType as String, Height as Integer,  AskedPersonCount as Integer)  as Boolean

	Dim RefSheet as Object
	Dim RefName as String 
	
	Dim ResponseDiametersIdx as Integer
	
	Dim CurrentDiam as Double
	Dim CurrentPersonCount as Double
	Dim TotalPersonCount as Double
	
	Dim I as Integer
	Dim PersCell as Object
	Dim DiameterCell as Object
	Dim continue as Boolean

	RefName = FormType + Height
	RefSheet = Doc.Sheets.GetByName(RefName)
	
	ResponseDiametersIdx = 0

	Rem Find Base cake
	Rem Try first with a base equals to tthe AskedPersonCount (so just one cake in total)
	I = 2
	continue = true

	Do While continue
		PersCell = RefSheet.GetCellRangeByName("B" + I)
		If PersCell.Type = com.sun.star.table.CellContentType.VALUE Then
			If PersCell.Value >= AskedPersonCount  Then
				DiameterCell = RefSheet.GetCellRangeByName("D" + I)
				continue = false
			Else
				I = I + 1
			End If
		Else
			I = -1
			continue = false
		End If
	Loop
	
	Rem no diameter great enough for the AskedPersonCount, so try with a base for half the AskedPersonCount
	if i < 0 Then
		i= 2
		continue = true
	
		Do While continue
			PersCell = RefSheet.GetCellRangeByName("B" + I)
			If PersCell.Type = com.sun.star.table.CellContentType.VALUE Then
				If PersCell.Value >= AskedPersonCount / 2 Then
					DiameterCell = RefSheet.GetCellRangeByName("D" + I)
					continue = false
				Else
					I = I + 1
				End If
			Else
				I = -1
				continue = false
			End If
		Loop
	End If
	
	Rem no diameter great enough for the AskedPersonCount/2, so try with a base for a third of the AskedPersonCount
	if i < 0 Then
		I = 2
		continue = true	

		Do While continue
			PersCell = RefSheet.GetCellRangeByName("B" + I)
			If PersCell.Type = com.sun.star.table.CellContentType.VALUE Then
				If PersCell.Value >= AskedPersonCount / 3 Then
					DiameterCell = RefSheet.GetCellRangeByName("D" + I)
					continue = false
				Else
					I = I + 1
				End If
			Else
				I = -1
				continue = false
			End If
		Loop	
	
	End If
	
	Rem If a basis is found, then go on for the next cake levels
	Rem Each level must be 5 cm shorter is diameter than the previous one, for estheatic purpose
	If I > 0 Then
		ResponseDiameters(ResponseDiametersIdx) = DiameterCell.Value 
		ResponsePersonCounts(ResponseDiametersIdx) = PersCell.Value 
		ResponseDiametersIdx = ResponseDiametersIdx + 1
		TotalPersonCount = PersCell.Value
		if (TotalPersonCount < AskedPersonCount) Then 
					
			Dim NextDiameter as Double
			NextDiameter = DiameterCell.Value - 5
			I = I - 1
			continue = true
		
			Do While continue And I > 0
				DiameterCell = RefSheet.GetCellRangeByName("D" + I)
				If DiameterCell.Type = com.sun.star.table.CellContentType.VALUE Then
					If DiameterCell.Value <= NextDiameter Then
						PersCell = RefSheet.GetCellRangeByName("B" + I)	
						ResponseDiameters(ResponseDiametersIdx) = DiameterCell.Value 
						ResponsePersonCounts(ResponseDiametersIdx) = PersCell.Value 
						ResponseDiametersIdx = ResponseDiametersIdx + 1		
						TotalPersonCount = TotalPersonCount  + PersCell.Value
						if (TotalPersonCount >= AskedPersonCount) Then 
							continue = false
						Else
							NextDiameter = DiameterCell.Value - 5
							I = I - 1
						End If
					Else
						I = I - 1
					End If
				Else 
					I = -1
				End If
			Loop	
		End If		
	End If
	
	DoCalcCakeDiameters = (I > 0)
		
	If I > 0 Then
			MainSheet.getCellRangeByName("B4").String = "OK met H= " + Height + " cm"
	Else
			Rem No basis found or not enough in multi level to fit for the AskedPersonCount
			MainSheet.getCellRangeByName("B4").String = "Niet OK met H= " + Height + " cm"
	End If
	
	For I = 0 To ResponseDiametersIdx 
		MainSheet.getCellRangeByName("B" + (7 + I)).Value= ResponseDiameters(I)
		MainSheet.getCellRangeByName("C" + (7 + I)).Value = ResponsePersonCounts(I)
	Next
	
	MainSheet.getCellRangeByName("B" + (8 + ResponseDiametersIdx)).String = "Totaal:"
	MainSheet.getCellRangeByName("C" + (8 + ResponseDiametersIdx)).Value = 	TotalPersonCount

End Function


