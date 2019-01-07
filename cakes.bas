Option Explicit

Private	Doc as Object
Private	MainSheet as Object

Private LocalSettings As New com.sun.star.lang.Locale
Private NumberFormats As Object
Private Key as long 

Public Type Component

	Label As String
	Quantity As Double
	Unit As String
	Price As Double
	
End Type


Public Type Cake

	FormType As String
	Diameter As Double
	Height as Double
	PersonCount As Integer
	Components()  'Array of Component
	
End Type


Public Type ShoppingBasket 

	Components() 
	
End Type

Public Type CakeSimulation
	
	ID as Integer
	Height as Integer
	FormType as String
	PersonCount as Integer
	Coef as Double
	Recipe as String
	Cakes()  'Array of Component

End Type



Sub Init
	
	NumberFormats = Doc.numberFormats
	LocalSettings.language = "nl"
	LocalSettings.country = "be"
	
	Key = NumberFormats.queryKey("0,#", LocalSettings , true)
	If Key = -1 then 
    	Key = NumberFormats.addNew("0,#", LocalSettings)
	End If
	
End Sub

Public Function GetSimulationPersonCount(oSimulation as CakeSimulation)

	Dim TotalPers as Integer
	TotalPers = 0
	
	Dim I as Integer 
	Dim CakeCount as Integer
	Dim Cake as Cake
		
	CakeCount =  UBound(oSimulation.Cakes)
	I =0
	
	For I = 0 to CakeCount
		TotalPers = TotalPers + oSimulation.Cakes(I).PersonCount
	Next
	
	GetSimulationPersonCount = TotalPers

End Function


Public Function GetSimulationPrice(oSimulation as CakeSimulation)

	Dim TotalPrice as Double
	TotalPrice = 0
	
	Dim I as Integer 
	Dim CakeCount as Integer
	Dim Cake as Cake
		
	CakeCount =  UBound(oSimulation.Cakes)
	I =0
	
	For I = 0 to CakeCount
		TotalPrice = TotalPrice + GetCakePrice((oSimulation.Cakes(I))
	Next
	
	GetSimulationPrice = TotalPrice

End Function

Public Function GetCakePrice(oCake as Cake) as Double 

	Dim TotalPrice as Double
	TotalPrice = 0
	
	Dim I as Integer
	Dim CompCount as Integer
	CompCount = UBound(oCake.Components)
	
	For I = 0 To CompCount
	Dim Component as Component
		Component = oCake.Components(I)
		TotalPrice = TotalPrice + Component.Price
	Next
	
	GetCakePrice = TotalPrice 

End Function


Public Function GetCakeVolume(oCake as Cake) as Double 

	Dim Volume as Double
	If oCake.FormType = "ROND" Then
		Volume = PI() * oCake.Diameter * oCake.Diameter * oCake.Height / 4
	Else 
		Volume = oCake.Diameter * oCake.Diameter *  oCake.Height
	End If
	
	GetCakeVolume = Volume

End Function

Sub ShowCakesSimulation

	Doc = ThisComponent
	init()
	
	Dim CakeSimulations() ' result array, passed as Variant pointer ipo Array value element
	Dim ResultTable as Object
	Dim ResultTableRange 
	Dim FormType as String 
	Dim AskedPersonCount  as Integer 
	Dim RecipeName  as String 

	MainSheet = Doc.Sheets.getByName("CALC")
	FormType = MainSheet.getCellRangeByName("VORM").String
	AskedPersonCount = MainSheet.getCellRangeByName("PERSONEN").Value
	RecipeName = MainSheet.getCellRangeByName("RECIPE").String

	ResultTable = MainSheet.getCellRangeByName("RESULT")
	ResultTable.clearContents(com.sun.star.sheet.CellFlags.STRING+com.sun.star.sheet.CellFlags.VALUE)
	ResultTableRange = ResultTable.DataArray
		
	DoCalcSimulations(CakeSimulations, FormType, AskedPersonCount, RecipeName)
	
	Dim SimulationIdx, SimulationCount As Integer
	SimulationCount = UBound(CakeSimulations)
	SimulationIdx = 0
	

	For SimulationIdx = 0 to SimulationCount
		ResultTableRange(0)(SimulationIdx) = "" & CakeSimulations(SimulationIdx).ID 
		ResultTableRange(1)(SimulationIdx) = "" & CakeSimulations(SimulationIdx).Height & " cm"
		ResultTableRange(2)(SimulationIdx) = "# " & GetSimulationPersonCount(CakeSimulations(SimulationIdx))
		ResultTableRange(3)(SimulationIdx) = GetSimulationPrice(CakeSimulations(SimulationIdx))
				
		Dim CakeCount, CakeIdx as Integer
		CakeCount = UBound(CakeSimulations(SimulationIdx).Cakes)
		For CakeIdx = 0 to CakeCount
			ResultTableRange(CakeIdx+5)(SimulationIdx) = CakeSimulations(SimulationIdx).Cakes(CakeIdx).Diameter
		Next
				
	Next

	ResultTable.DataArray = ResultTableRange

End Sub


Sub ShowCakeComposition()

	Doc = ThisComponent
	init()
	
	Dim CakeSimulations() 
	Dim CakeSimulation
	Dim ResultTable as Object
	Dim ResultTableRange 
	Dim FormType as String 
	Dim AskedPersonCount  as Integer 
	Dim CakeSimulationID as Integer
	Dim RecipeName  as String 

	MainSheet = Doc.Sheets.getByName("CALC")
	FormType = MainSheet.getCellRangeByName("VORM").String
	AskedPersonCount = MainSheet.getCellRangeByName("PERSONEN").Value
	RecipeName = MainSheet.getCellRangeByName("RECIPE").String
	CakeSimulationID = MainSheet.getCellRangeByName("CAKEID").Value
	
	DoCalcSimulations(CakeSimulations, FormType, AskedPersonCount, RecipeName)

	CakeSimulation = CakeSimulations(CakeSimulationID)	

	ResultTable = MainSheet.getCellRangeByName("SAMENSTELLING")
	ResultTable.clearContents(com.sun.star.sheet.CellFlags.STRING+com.sun.star.sheet.CellFlags.VALUE)
	ResultTableRange = ResultTable.DataArray
	
	Dim CakeIdx, CakeCount As Integer
	Dim ProductIdx, ProductCount As Integer
	
	CakeCount = UBound(CakeSimulation.Cakes)
	
	For CakeIdx = 0 To CakeCount
		With CakeSimulation.Cakes(CakeIdx) 
    		ProductCount = UBound(.Components)
			Dim x as Integer : x = CakeIdx * (ProductCount + 6)
			
    		ResultTableRange(x)(0) = "Cake D: " & .Diameter & " / H: " & .Height 
    		ResultTable.Rows(x).charWeight = com.sun.star.awt.FontWeight.BOLD

    		ResultTableRange(x+2)(0) = "Product"
    		ResultTableRange(x+2)(2) = "Hoeveelheid"
    		ResultTableRange(x+2)(3) = "Eenheid"
    		
    		ResultTable.Rows(x+2).charWeight = com.sun.star.awt.FontWeight.BOLD
		    		    		
    		For ProductIdx = 0 to ProductCount
	    		With CakeSimulation.Cakes(CakeIdx).Components(ProductIdx)
		    		ResultTableRange(x + ProductIdx + 3)(0) = .Label
		    		ResultTableRange(x + ProductIdx + 3)(2) = .Quantity
		    		ResultTableRange(x + ProductIdx + 3)(3) = .Unit
	    		End With
    		Next
    	End With
	Next
	
    ResultTable.Columns(2).NumberFormat = Key
	ResultTable.DataArray = ResultTableRange

End Sub


Sub DoCalcSimulations(CakeSimulations, FormType as String, AskedPersonCount as Integer, RecipeName as String)

	Dim arCakeSimulations(8) '  working array of CakeSimulation.  Some last elements may be null

	Dim Heights() : Heights = Array(10, 12)
	Dim Coefs() : 	Coefs = Array(1, 0.67, 0.5, 0.33)
		
	Dim HeightIdx as Integer: HeightIdx = 0
	Dim CoefIdx as Integer : CoefIdx = 0
	Dim SimulationIdx as Integer : 	SimulationIdx = -1
	Dim Continue as Boolean :  Continue = True
	Dim Simulation 
	
	
 	Do While Continue 

 		Simulation = New CakeSimulation
 		Simulation.Recipe = RecipeName
		Simulation.Height = Heights(HeightIdx)
		Simulation.FormType = FormType 
		Simulation.PersonCount = AskedPersonCount
		Simulation.Coef = Coefs(CoefIdx) 
			 	
		DoCalcSimulation(Simulation)
	
		Dim CakesCount as Integer
		CakesCount = UBound(Simulation.Cakes)
		If  CakesCount > -1 Then 
			SimulationIdx = SimulationIdx + 1
			Simulation.ID = SimulationIdx
			arCakeSimulations(SimulationIdx) = Simulation
		End If
		
		CoefIdx = CoefIdx +1
		Dim maxIdx as Integer
		maxIdx = UBound(Coefs)
		If CoefIdx > maxIdx	 Then
			CoefIdx = 0
			HeightIdx = HeightIdx + 1
		End If
		
		maxIdx = UBound(Heights)
		Continue = HeightIdx  <= maxIdx
	Loop

	If SimulationIdx > -1 Then
	    Redim CakeSimulations(SimulationIdx) 
	    Dim I as Integer : I = 0
		For I = 0 To SimulationIdx
			CakeSimulations(I) =  arCakeSimulations(I)
	
			Dim TotalPrice as Double
			TotalPrice = getSimulationPrice(Simulation)
		Next
	End If
		
End Sub


Sub DoCalcSimulation(CakeSimulation)

 	Dim Cakes()
 	
	DoCalcCakeDiameters(Cakes, CakeSimulation.FormType, CakeSimulation.Height,  CakeSimulation.PersonCount, CakeSimulation.Coef)
	
	Dim CakesCount as Integer
	CakesCount = UBound(Cakes)
	If  CakesCount > -1 Then 
		CakeSimulation.Cakes = Cakes

		Dim CakeIdx as Integer
		For CakeIdx = 0 to CakesCount 
			DoCalcCakeComposition(Cakes(CakeIdx), CakeSimulation.Recipe)
		Next
	End If
			
End Sub

Rem Search composition of each cake and fill in a table below result table
Rem Update Cake array
Sub DoCalcCakeComposition(Cake, RecipeName as String)  

	Dim RecipeRangeName as String
	Dim RecipeSheet as Object
	Dim Element as Object
	Dim StartIdx as Integer
	Dim RecipeRange as Object
	Dim RecipeRangeArray
	Dim DestRange as Object
	Dim ProductCount as Integer
	
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
	
	Dim Volume as Double
	Volume = GetCakeVolume(Cake)
	
	Dim oComponents(ProductCount)

	For I = 0 to ProductCount
	    Dim Component as new Component
		Component.Label = RecipeRangeArray(I)(0)
		Component.Quantity = RecipeRangeArray(I)(1) * Volume / RecipeRangeArray(I)(3)
		Component.Unit = RecipeRangeArray(I)(2)
		Component.Price = RecipeRangeArray(I)(4) * Volume / RecipeRangeArray(I)(3)
		oComponents(i) = Component
	Next

	Cake.Components = oComponents
	
End Sub

Rem Fill in the global RespondeDiameters variable
Rem Returns True if a solution is found
Sub DoCalcCakeDiameters(Cakes(), FormType as String, Height as Integer,  AskedPersonCount as Integer, Coef as Double) 

	Dim arCakes(6) as New Cake
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
	
	I = 2
	continue = True
	
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

	
	Rem If a basis is found, then go on for the next cake levels
	Rem Each level must be 5 cm shorter is diameter than the previous one, for estheatic purpose
	If I > 0 Then
	    Dim cakeIdx as Integer
	    CakeIdx = 0    
	    
	    Dim aCake as Cake

		aCake = New Cake
	    aCake.Diameter = DiameterCell.Value 
	    aCake.PersonCount = PersCell.Value
	    aCake.FormType = FormType
	    aCake.Height = Height
		arCakes(CakeIdx) = aCake
		 	    
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
						CakeIdx = CakeIdx + 1						
						aCake = New Cake
						aCake.Diameter = DiameterCell.Value 
						aCake.PersonCount = PersCell.Value
					    aCake.FormType = FormType
					    aCake.Height = Height
						arCakes(CakeIdx) = aCake
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
	
	If (TotalPersonCount >= AskedPersonCount) Then 
		ReDim Cakes(CakeIdx) 
		For I = 0 to CakeIdx
			Cakes(I) = arCakes(I)
		Next
	End If	
	
End Sub


