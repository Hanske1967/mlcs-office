Public Type Component

	Label As String
	Quantity As Double
	Unit As String
	Price As Double
	
End Type


Public Type Cake

	Diameter As Double
	Components(25) As Component
	
End Type


Public Type ShoppingBasket 

	Components(25) As Component
	
End Type

Public Type CakeSimulation

	Height as Integer
	FormType as String
	Rem cakes is an array of cakes 
	Cakes as Variant

End Type


Public Sub GetCakePrice(Cake as Cake) as Double 

	Dim TotalPrice as Double
	TotalPrice = 0
	
	Dim I as Integer
	Dim CompCount as Integer
	CompCount = UBound(Cale.Components)
	
	For I = 0 To CompCount
		TotalPrice = TotalPrice + Components(I).Price
	Next
	
	GetCakePrice = TotalPrice 

End Sub
