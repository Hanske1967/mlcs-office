Option Explicit

Private	Doc as Object
Private	MainSheet as Object

Private	ResponseDiameters(6) as Double
Private	ResponsePersonCounts(6) as Double
Private ResponseDiametersIdx as Integer


Sub Main

End Sub

Rem  Main entry point
Sub CalcCakeDiameters

	Dim FormType as String
	Dim AskedPersonCount as Integer
	Dim RecipeName as String
	Dim I as Integer
	
	Doc = ThisComponent
	MainSheet = Doc.Sheets.getByName("CALC")
	FormType = MainSheet.getCellRangeByName("VORM").String
	AskedPersonCount = MainSheet.getCellRangeByName("PERSONEN").Value
	RecipeName = MainSheet.getCellRangeByName("RECIPE").String

	Rem clear previous result
	MainSheet.getCellRangeByName("RESULT").String = ""
	MainSheet.getCellRangeByName("C13:D22").clearContents(com.sun.star.sheet.CellFlags.STRING+com.sun.star.sheet.CellFlags.VALUE)
	
	Dim Height as Integer
	Dim Found as Boolean
	
	Height = 10
	Found = DoCalcCakeDiameters(FormType, Height,  AskedPersonCount)
	
	If (Found = False) Then
		Height = 12
		Found = DoCalcCakeDiameters(FormType, Height,  AskedPersonCount)
	End If
	
	If (Found = True) Then
		For I = 0 To ResponseDiametersIdx
			Dim Volume as Double
			If FormType = "ROND" Then
				Volume = PI() * ResponseDiameters(I) * ResponseDiameters(I) * Height / 4
			Else 
				Volume = ResponseDiameters(I) * ResponseDiameters(I) * Height
			End If
			
			DoCalcCakeComposition(I, RecipeName, Volume)
		Next
	End If
	
		
End Sub

Rem Search composition of each cake and fill in a table below result table
Sub DoCalcCakeComposition(CakeIdx as Integer, RecipeName as String, Volume as Double) 

	Dim RecipeRangeName as String
	Dim RecipeSheet as Object
	Dim Element as Object
	Dim StartIdx as Integer
	Dim RecipeRange as Object
	Dim RecipeRangeArray as Variant
	Dim DestRange as Object
	Dim DestRangeArray as Variant

	
	Rem remove all spaces to get range name
	Dim I as Integer
	Dim str as String
	For I = 1 to Len(RecipeName)
		str = Mid(RecipeName, I, 1) 
		if (str <> " ") Then
			RecipeRangeName = RecipeRangeName + str
		End If
	Next	

	RecipeSheet =  Doc.Sheets.getByName(RecipeName)
	RecipeRange = RecipeSheet.GetCellRangeByName(RecipeRangeName)
	RecipeRangeArray = RecipeRange.DataArray

	DestRange = MainSheet.getCellRangeByPosition(1, 42 + CakeIdx * 10, 1 + RecipeRange.Columns.Count - 1, 42 + CakeIdx * 10 + RecipeRange.Rows.Count)	
	DestRangeArray = DestRange.DataArray

	DestRangeArray(0)(0) = "Product"
	DestRangeArray(0)(1) = "Hoeveelheid"
	DestRangeArray(0)(2) = "Eenheid"
	DestRangeArray(0)(3) = "Prijs"

	Dim ProductCount as Integer
	ProductCount = UBound(RecipeRangeArray)
	For I = 1 to ProductCount
		DestRangeArray(I)(0) = RecipeRangeArray(I-1)(0)
		DestRangeArray(I)(1) = RecipeRangeArray(I-1)(1) * Volume / RecipeRangeArray(I-1)(3)
		DestRangeArray(I)(2) = RecipeRangeArray(I-1)(2)
		DestRangeArray(I)(3) = 0
	Next

	DestRange.DataArray = DestRangeArray


End Sub

Rem Fill in the global RespondeDiameters variable
Rem Returns True if a solution is found
Function DoCalcCakeDiameters(FormType as String, Height as Integer,  AskedPersonCount as Integer)  as Boolean

	Dim RefSheet as Object
	Dim RefName as String 
	
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
						ResponseDiametersIdx = ResponseDiametersIdx + 1
						ResponseDiameters(ResponseDiametersIdx) = DiameterCell.Value 
						ResponsePersonCounts(ResponseDiametersIdx) = PersCell.Value 
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
	
	Dim ResultTable as Object
	ResultTable = MainSheet.getCellRangeByName("RESULTTABLE")
	
	If I > 0 Then
			MainSheet.getCellRangeByName("RESULT").String = "OK met H= " + Height + " cm"

			Dim ResultTableValues As Variant
			ResultTableValues = ResultTable.DataArray
		
			For I = 0 To ResponseDiametersIdx 
				ResultTableValues(I)(0) = ResponseDiameters(I)
				ResultTableValues(I)(1) = ResponsePersonCounts(I)
			Next
			
			ResultTableValues(ResponseDiametersIdx+1)(0) = "Totaal pers.:"
			ResultTableValues(ResponseDiametersIdx+1)(1) = TotalPersonCount
		
			ResultTable.DataArray = ResultTableValues
	Else
			Rem No basis found or not enough in multi level to fit for the AskedPersonCount
			MainSheet.getCellRangeByName("RESULT").String = "Niet OK met H= " + Height + " cm"
			ResultTable.clearContents(com.sun.star.sheet.CellFlags.STRING+com.sun.star.sheet.CellFlags.VALUE)
	End If


End Function

