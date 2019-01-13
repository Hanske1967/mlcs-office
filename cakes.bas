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
	Density as Double
	
End Type


Public Type Cake

	FormType As String
	Diameter As Double
	Height as Double
	PersonCount As Integer
	Components()  'Array of component 
	
End Type


Public Type ShoppingList 

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

Sub UpdateShoppingBasket(Simulation as CakeSimulation, ShoppingList as ShoppingList)

	Dim ComponentCount as Integer
	Dim ComponentIdx as Integer
	Dim CakeIdx as Integer
	Dim CakeCount as Integer

	CakeCount = UBound(Simulation.Cakes)
	ComponentCount = UBound(Simulation.Cakes(0).Components)

	Dim Components(ComponentCount) 
	Dim Component as Component
	
	For ComponentIdx = 0 to ComponentCount
		With Simulation.Cakes(0).Components(ComponentIdx) 
			Component = new Component
			Component.	Label = .Label
			Component.	Quantity = .Quantity 
			Component.	Unit = .Unit 
			Component.Price = .Price 
			Components(ComponentIdx) = Component
		End With
	Next
		
	If CakeCount > 0	 Then
		For CakeIdx = 1 to CakeCount
			For ComponentIdx = 0 to ComponentCount
				Components(ComponentIdx).	Quantity = Components(ComponentIdx).Quantity + Simulation.Cakes(CakeIdx).Components(ComponentIdx).Quantity 
				Components(ComponentIdx).Price = Components(ComponentIdx).Price + Simulation.Cakes(CakeIdx).Components(ComponentIdx).Price 
			Next
		Next
	End If
	

	ShoppingList.Components = Components

End Sub

Public Function GetSimulationPersonCount(oSimulation as CakeSimulation) as Integer

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


Public Function GetSimulationPrice(oSimulation as CakeSimulation) as Double

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
		TotalPrice = TotalPrice + oCake.Components(I).Price
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


Public Function GetCakeTopSurface(oCake as Cake) as Double 

	Dim Surface as Double
	If oCake.FormType = "ROND" Then
		Surface = PI() * oCake.Diameter * oCake.Diameter / 4 
	Else 
		Surface = oCake.Diameter * oCake.Diameter 
	End If
	
	GetCakeTopSurface = Surface

End Function


Public Function GetCakeOveralSurface(oCake as Cake) as Double 

	Dim Surface as Double
	If oCake.FormType = "ROND" Then
		Surface = PI() * oCake.Diameter * oCake.Diameter / 4 + PI() * oCake.Diameter * oCake.Height
	Else 
		Surface = oCake.Diameter * oCake.Diameter + 4 *  oCake.Diameter * oCake.Height
	End If
	
	GetCakeOveralSurface = Surface

End Function


Sub ShowCakesSimulation

	Doc = ThisComponent
	init()
	
	Dim CakeSimulations() 
	
	Dim ResultTable as Object
	Dim ResultTableRange 
	Dim FormType as String 
	Dim AskedPersonCount  as Integer 
	Dim RecipeName  as String 

	MainSheet = Doc.Sheets.getByName("CALC")
	FormType = MainSheet.getCellRangeByName("VORM").String
	AskedPersonCount = MainSheet.getCellRangeByName("PERSONEN").Value
	RecipeName = MainSheet.getCellRangeByName("RECIPE").String

	ResultTable = MainSheet.getCellRangeByName("SIMULATIONS")
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

Function GetFillings()

	Dim Components()
	Dim Fillins(100) as Component
	Dim Idx, Count , Start as Integer
	
	Components = DoGetFillings("Vulling")
	Count = UBound(Components)
	Start = 0
	
	For Idx = 0 to Count
		Fillins(Start + Idx) = Components(Idx)
	Next	
	Start = Start + Count + 1
	
	Components = DoGetFillings("Afsmeren")
	Count = UBound(Components)
	
	For Idx = 0 to Count
		Fillins(Start + Idx) = Components(Idx)
	Next	
	Start = Start + Count + 1

	Components = DoGetFillings("Bekleding")
	Count = UBound(Components)
	
	For Idx = 0 to Count
		Fillins(Start + Idx) = Components(Idx)
	Next	
	Start = Start + Count
	
	Dim Result(Start)
	For Idx = 0 to Start 
		Result(Idx) = Fillins(Idx)
	Next

	GetFillings = Result

End Function

Rem Returns an array of all fillingfs as Component
Function DoGetFillings(Name as String) 

	Doc = ThisComponent
	Dim FillingSheet as Object
	Dim FillingSheetData
	FillingSheet = Doc.Sheets.getByName(Name)
	FillingSheetData = FillingSheet.GetCellRangeByName("A2:D20").DataArray
			
	Dim Components(20) as Component
	Dim 	Component as Component
	Dim RowIdx as Integer : RowIdx = 0
	Dim continue as Boolean : continue = True
	
	Do While Continue
		Component = New Component
		Component.Label = FillingSheetData(RowIdx)(0)
		Component.Quantity =FillingSheetData(RowIdx)(1)
		Component.Price = FillingSheetData(RowIdx)(2)
		Component.Density = FillingSheetData(RowIdx)(3)
		Component.	Unit = "g"
		
		Components(RowIdx) = Component
		RowIdx = RowIdx + 1
		
		Continue =  (FillingSheetData(RowIdx)(0) <> "")
	Loop
	RowIdx = RowIdx - 1
	
	Dim Result(RowIdx)
	Dim I as Integer : I = 0
	
	For I = 0 To RowIdx
		Result(I) = Components(I)
	Next
		
	DoGetFillings = Result

End Function

Function GetFilling(Components(), Name as String, Surface As Double) 

	Dim Component as Component
	Dim CompCount as Integer
	Dim CompIdx as Integer
	
	CompCount = UBound(Components)
	For CompIdx = 0 To CompCount
		If (Components(CompIdx).Label = Name) Then
			Component = Components(CompIdx)
		End If
	Next
	
	Dim Result
	
	Result = new Component
	Result.Label = Component.Label
	Result.Quantity = Component.Quantity * Surface * Component.Density
	Result.Unit = Component.Unit
	Result.Price = Component.Price * Result.Quantity
	
	GetFilling = Result

End Function

Sub FillCompositionTable()

	Dim CakeSimulations() 
	Dim CakeSimulation
	Dim ResultTable as Object
	Dim ResultTableRange 
	
	Dim FormType as String 
	Dim AskedPersonCount  as Integer 
	Dim RecipeName  as String 
	Dim CakeSimulationID as Integer	

	Doc = ThisComponent

	MainSheet = Doc.Sheets.getByName("CALC")
	FormType = MainSheet.getCellRangeByName("VORM").String
	AskedPersonCount = MainSheet.getCellRangeByName("PERSONEN").Value
	RecipeName = MainSheet.getCellRangeByName("RECIPE").String
	CakeSimulationID = MainSheet.getCellRangeByName("CAKEID").Value
		
	ResultTable = MainSheet.getCellRangeByName("SAMENSTELLINGIDS")
	ResultTable.clearContents(com.sun.star.sheet.CellFlags.STRING+com.sun.star.sheet.CellFlags.VALUE)
	ResultTableRange = ResultTable.DataArray
	
	DoCalcSimulations(CakeSimulations, FormType, AskedPersonCount, RecipeName)
	CakeSimulation = CakeSimulations(CakeSimulationID)
	
	Dim CakeIdx, CakeCount As Integer
	CakeCount = UBound(CakeSimulation.Cakes)
	CakeIdx = 0

	For CakeIdx = 0 to CakeCount
		ResultTableRange(CakeIdx)(0) = "D" & CakeSimulation.Cakes(CakeIdx).Diameter 
	Next
	
	ResultTable.DataArray = ResultTableRange

End Sub


Sub ShowShoppingList()

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


	' Samenvatting

	ResultTable = MainSheet.getCellRangeByName("PERCAKE")
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
    		ResultTableRange(x+2)(4) = "Prijs"
    		
    		ResultTable.Rows(x+2).charWeight = com.sun.star.awt.FontWeight.BOLD
		    		    		
    		For ProductIdx = 0 to ProductCount
	    		With CakeSimulation.Cakes(CakeIdx).Components(ProductIdx)
		    		ResultTableRange(x + ProductIdx + 3)(0) = .Label
		    		ResultTableRange(x + ProductIdx + 3)(2) = .Quantity
		    		ResultTableRange(x + ProductIdx + 3)(3) = .Unit
		    		ResultTableRange(x + ProductIdx + 3)(4) = .Price
	    		End With
    		Next

    	End With
	Next
	
    ResultTable.Columns(2).NumberFormat = Key
	ResultTable.DataArray = ResultTableRange


	' Shopping List

	Dim ShoppingList as New ShoppingList
	UpdateShoppingBasket(CakeSimulation, ShoppingList)

	ResultTable = MainSheet.getCellRangeByName("SHOPPINGLIST")
	ResultTable.clearContents(com.sun.star.sheet.CellFlags.STRING+com.sun.star.sheet.CellFlags.VALUE)
	ResultTableRange = ResultTable.DataArray
	
	Dim ComponentIdx, ComponentCount As Integer
	ComponentCount = UBound(ShoppingList.Components)
	
	ResultTableRange(0)(0) = "Product"
	ResultTableRange(0)(2) = "Hoeveelheid"
	ResultTableRange(0)(3) = "Eenheid"
	ResultTableRange(0)(4) = "Prijs"
	
	ResultTable.Rows(0).charWeight = com.sun.star.awt.FontWeight.BOLD

	For ComponentIdx = 0 To ComponentCount
		With ShoppingList.Components(ComponentIdx) 
    		ResultTableRange(ComponentIdx + 2)(0) = .Label
    		ResultTableRange(ComponentIdx + 2)(2) = .Quantity
    		ResultTableRange(ComponentIdx + 2)(3) = .Unit
    		ResultTableRange(ComponentIdx + 2)(4) = .Price
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
Rem Then take in account all filling and finition layers
Rem Returns Cake components updated
Sub DoCalcCakeComposition(Cake, RecipeName as String)  

	Dim RecipeRangeName as String
	Dim RecipeSheet as Object
	Dim Element as Object
	Dim StartIdx as Integer
	Dim RecipeRange as Object
	Dim RecipeRangeArray
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
	
	Dim oComponents(ProductCount+5)

	For I = 0 to ProductCount
	    Dim Component as new Component
		Component.Label = RecipeRangeArray(I)(0)
		Component.Quantity = RecipeRangeArray(I)(1) * Volume / RecipeRangeArray(I)(3)
		Component.Unit = RecipeRangeArray(I)(2)
		Component.Price = RecipeRangeArray(I)(4) * Volume / RecipeRangeArray(I)(3)
		oComponents(i) = Component
	Next

	' Fill in all values found in SAMENSTELLING
	Dim FillingRange as Object
	Dim FillingRangeRangeArray

	MainSheet = Doc.Sheets.getByName("CALC")
	FillingRange = MainSheet.GetCellRangeByName("SAMENSTELLING")
	FillingRangeRangeArray = FillingRange.DataArray	
	
	' Find cake row
	Dim RowIdx as Integer : RowIdx = -1
	Dim found as boolean : found = False
	Dim continue as boolean : continue = True
	str = "D" & Cake.Diameter
	
	Do While (NOT(found) AND Continue) 
		RowIdx = RowIdx + 1
		Continue = (FillingRangeRangeArray(RowIdx)(0) <> "")
		found = (FillingRangeRangeArray(RowIdx)(0) = str)
	Loop
	
	If found Then
		Dim Components
		Components = GetFillings()
		Dim Surface as Double

		Surface = GetCakeTopSurface(Cake)
		' Filling 
		For I = 1 To 3 
			If (FillingRangeRangeArray(RowIdx)(I) <> "") Then
				ProductCount = ProductCount + 1
				oComponents(ProductCount) =  GetFilling(Components, FillingRangeRangeArray(RowIdx)(I), Surface)
			End If
		Next
			
		Surface =  GetCakeOveralSurface(Cake)
		' Surface
		For I = 4 To 5 
			If (FillingRangeRangeArray(RowIdx)(I) <> "") Then
				ProductCount = ProductCount + 1
				oComponents(ProductCount) =  GetFilling(Components, FillingRangeRangeArray(RowIdx)(I), Surface)
			End If
		Next
	End If

	Dim CakeComponents(ProductCount) 
	For I = 0 to ProductCount
		CakeComponents(I) = oComponents(I)
	Next
	Cake.Components = CakeComponents
	
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


