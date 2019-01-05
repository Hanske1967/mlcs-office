Option Explicit

Private	Doc as Object
Private	MainSheet as Object


Private	ResponseDiameters(6) as Double
Private	ResponsePersonCounts(6) as Double
Private ResponseDiametersIdx as Integer

Private LocalSettings As New com.sun.star.lang.Locale
Private NumberFormats As Object
Private Key as long 


Sub Init
	
	NumberFormats = Doc.numberFormats
	LocalSettings.language = "nl"
	LocalSettings.country = "be"
	
	Key = NumberFormats.queryKey("0,#", LocalSettings , true)
	If Key = -1 then 
    	Key = NumberFormats.addNew("0,#", LocalSettings)
	End If
	
End Sub

Rem  Main entry point
Sub CalcCakeDiameters

	Dim ResultTable as Object
	Dim FormType as String
	Dim AskedPersonCount as Integer
	Dim RecipeName as String
	Dim I as Integer
	Dim TotalPrice as Double
	
	Doc = ThisComponent
	Init()

	MainSheet = Doc.Sheets.getByName("CALC")
	FormType = MainSheet.getCellRangeByName("VORM").String
	AskedPersonCount = MainSheet.getCellRangeByName("PERSONEN").Value
	RecipeName = MainSheet.getCellRangeByName("RECIPE").String

	Rem clear previous result
	MainSheet.getCellRangeByName("RESULT").String = ""
	MainSheet.getCellRangeByName("C13:D22").clearContents(com.sun.star.sheet.CellFlags.STRING+com.sun.star.sheet.CellFlags.VALUE)
	MainSheet.getCellRangeByPosition(1, 45, 6, 100).clearContents(com.sun.star.sheet.CellFlags.STRING+com.sun.star.sheet.CellFlags.VALUE+com.sun.star.sheet.CellFlags.HARDATTR)
		
	Dim Found as Boolean
	Dim Continue as Boolean
	Dim Heights() as  Variant 
	Dim Coefs() as  Variant
	Dim HeightIdx, CoefIdx as Integer

 	Coefs = Array(1, 0.67, 0.5, 0.33)
	Heights = Array(10, 12)

	HeightIdx = 0
	CoefIdx = 0
	Continue = True
	Found = False
	
 	Do While (Continue And Not(Found))
		Found = DoCalcCakeDiameters(FormType, Heights(HeightIdx),  AskedPersonCount, Coefs(CoefIdx))
		
		If (Found = False) Then
			CoefIdx = CoefIdx +1
			Dim maxIdx as Integer
			maxIdx = UBound(Coefs)
			If CoefIdx > maxIdx	 Then
				CoefIdx = 0
				HeightIdx = HeightIdx + 1
			End If
			
			maxIdx = UBound(Heights)
			Continue = HeightIdx  <= maxIdx
		End If
	Loop
	
	TotalPrice = 0
	
	If (Found = True) Then
		For I = 0 To ResponseDiametersIdx
			Dim Volume as Double
			If FormType = "ROND" Then
				Volume = PI() * ResponseDiameters(I) * ResponseDiameters(I) *  Heights(HeightIdx) / 4
			Else 
				Volume = ResponseDiameters(I) * ResponseDiameters(I) *  Heights(HeightIdx)
			End If
			
			TotalPrice = TotalPrice + DoCalcCakeComposition(I, RecipeName, Volume)
		Next
	End If
	
	ResultTable = MainSheet.getCellRangeByName("RESULTTABLE")	
	MainSheet.getCellRangeByName("C" + (ResultTable.RangeAddress.StartRow + ResponseDiametersIdx + 4)).String = "Prijs grondstoffen: "
	MainSheet.getCellRangeByName("D" + (ResultTable.RangeAddress.StartRow + ResponseDiametersIdx + 4)).Value = TotalPrice
	MainSheet.getCellRangeByName("D" + (ResultTable.RangeAddress.StartRow + ResponseDiametersIdx + 4)).NumberFormat = Key
	
		
End Sub

Rem Search composition of each cake and fill in a table below result table
Rem Returns the price of the cake
Function DoCalcCakeComposition(CakeIdx as Integer, RecipeName as String, Volume as Double) as Double

	Dim RecipeRangeName as String
	Dim RecipeSheet as Object
	Dim Element as Object
	Dim StartIdx as Integer
	Dim RecipeRange as Object
	Dim RecipeRangeArray as Variant
	Dim DestRange as Object
	Dim DestRangeArray as Variant
	Dim ProductCount as Integer
	
	Dim TotalPrice as Double
	
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

	ProductCount = UBound(RecipeRangeArray)

	MainSheet.getCellRangeByName("A43").String = "Samenstelling"
	MainSheet.getCellRangeByName("A43").charWeight = com.sun.star.awt.FontWeight.BOLD
	
	DestRange = MainSheet.getCellRangeByPosition(1, 45 + CakeIdx * (ProductCount + 5), 1 + RecipeRange.Columns.Count - 1, 45 + CakeIdx * (ProductCount + 5) + RecipeRange.Rows.Count + 1)	
	DestRangeArray = DestRange.DataArray

	DestRangeArray(0)(0) = RecipeName + "  D = " + ResponseDiameters(CakeIdx) + " cm"
	DestRangeArray(0)(1) = ""
	DestRangeArray(0)(2) = ""
	DestRangeArray(0)(3) = ""

	DestRangeArray(1)(0) = "Product"
	DestRangeArray(1)(1) = "Hoeveelheid"
	DestRangeArray(1)(2) = "Eenheid"
	DestRangeArray(1)(3) = "Prijs"

	TotalPrice = 0
	
	For I = 0 to ProductCount
		DestRangeArray(I+2)(0) = RecipeRangeArray(I)(0)
		DestRangeArray(I+2)(1) = RecipeRangeArray(I)(1) * Volume / RecipeRangeArray(I)(3)
		DestRangeArray(I+2)(2) = RecipeRangeArray(I)(2)
		DestRangeArray(I+2)(3) = RecipeRangeArray(I)(4) * Volume / RecipeRangeArray(I)(3)
		TotalPrice = TotalPrice + DestRangeArray(I+2)(3)
	Next

	DestRange.DataArray = DestRangeArray
	DestRange.Rows(0).charWeight = com.sun.star.awt.FontWeight.BOLD
	DestRange.Rows(1).charWeight = com.sun.star.awt.FontWeight.BOLD
	
	DestRange.Columns(1).NumberFormat = Key
	DestRange.Columns(3).NumberFormat = Key

	DoCalcCakeComposition = TotalPrice
	
End Function

Rem Fill in the global RespondeDiameters variable
Rem Returns True if a solution is found
Function DoCalcCakeDiameters(FormType as String, Height as Integer,  AskedPersonCount as Integer, Coef as Double)  as Boolean

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
	I = -1
	
Rem 	If AskedPersonCount < 15 Then
		Rem For 15 or more persons, always choose multi level cake
		I = 2
		continue = true
		
		Do While continue
			PersCell = RefSheet.GetCellRangeByName("B" + I)
			If PersCell.Type = com.sun.star.table.CellContentType.VALUE Then
				If PersCell.Value >= AskedPersonCount * Coef  Then
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
Rem 	End If
	
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
			
			ResultTableValues(ResponseDiametersIdx+2)(0) = "Totaal pers.:"
			ResultTableValues(ResponseDiametersIdx+2)(1) = TotalPersonCount
		
			ResultTable.DataArray = ResultTableValues
	Else
			Rem No basis found or not enough in multi level to fit for the AskedPersonCount
			MainSheet.getCellRangeByName("RESULT").String = "Niet OK met H= " + Height + " cm"
			ResultTable.clearContents(com.sun.star.sheet.CellFlags.STRING+com.sun.star.sheet.CellFlags.VALUE)
	End If


End Function


