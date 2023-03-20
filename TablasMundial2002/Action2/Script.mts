'Declaracion variables
Dim infCell
Set tableWiki1= Browser("Copa Mundial de Fútbol").Page("Copa Mundial de Fútbol").WebTable("xpath:=//*[@id=""mw-content-text""]/div[1]/table[8]")
Set tableWiki2= Browser("Copa Mundial de Fútbol_2").Page("Copa Mundial de Fútbol").WebTable("xpath:=//*[@id=""mw-content-text""]/div[1]/table[15]")
Set tableWiki3= Browser("Copa Mundial de Fútbol_2").Page("Copa Mundial de Fútbol").WebTable("xpath:=//*[@id=""mw-content-text""]/div[1]/table[22]")
Set tableWiki4= Browser("Copa Mundial de Fútbol_2").Page("Copa Mundial de Fútbol").WebTable("xpath:=//*[@id=""mw-content-text""]/div[1]/table[29]")
Set tableWiki5=Browser("Copa Mundial de Fútbol").Page("Copa Mundial de Fútbol").WebTable("xpath:=//*[@id=""mw-content-text""]/div[1]/table[36]")
Set tableWiki6=Browser("Copa Mundial de Fútbol_4").Page("Copa Mundial de Fútbol").WebTable("xpath:=//*[@id=""mw-content-text""]/div[1]/table[43]")
Set tableWiki7=Browser("Copa Mundial de Fútbol_4").Page("Copa Mundial de Fútbol").WebTable("xpath:=//*[@id=""mw-content-text""]/div[1]/table[50]")
Set tableWiki8= Browser("Copa Mundial de Fútbol_2").Page("Copa Mundial de Fútbol").WebTable("xpath:=//*[@id=""mw-content-text""]/div[1]/table[57]")


fila=3
Set hojaEditar=escribirExcel("\OneDrive\Documentos\FasesMundial.xlsx","Fases")   
recorrerTabla

Sub recorrerTabla()
	         wait 1	         
		  For contador = 1 To 8 Step 1
			  	Select Case (contador)
			  	     Case 1
			  	     	      escribirTabla tableWiki1,"A"
			  	     Case 2
				  	      escribirTabla tableWiki2,"B"
			  	     Case 3
			  	            escribirTabla tableWiki3,"C"
			  	     Case 4
			  	            escribirTabla tableWiki4,"D"
			  	     Case 5
			  	             escribirTabla tableWiki5,"E"
			  	     Case 6
			  	              escribirTabla tableWiki6,"F"
			  	     Case 7
			  	             escribirTabla tableWiki7,"G"
			  	     Case 8
			  	             escribirTabla tableWiki8,"H"
			  	End Select
	          Next
  End  Sub	
  
  
  Sub escribirTabla(tabla,letra)
    If objExistente(tabla) =1 Then
       hojaEditar.Cells(fila,(column+3)).Interior.Color=RGB(229, 252, 251)
       hojaEditar.Cells(fila,(column+3)).Value="Grupo "  &  letra 
       fila=fila+1
  	For row = 1 To  5 Step 1	   
	  	For column = 1 To 9 Step 1
	  		   If row=1 Then
	  		   	infCell=tabla.GetCellData(row,column)
	  		   	print(infCell)
	  		   	hojaEditar.Cells(fila,(column+2)).Interior.Color=RGB(255, 230, 252 )
	  		   	hojaEditar.Cells(fila,(column+2)).Font.Color=RGB(197, 20, 244 )
	  		       hojaEditar.Cells(fila,(column+2)).Value=infCell
	  		       
	  		    else
	  		       If column=1 Then
	  		                url=tabla.ChildItem(row,column,"Image",0).GetRoProperty("src") 
	  		                hojaEditar.Shapes.AddPicture  url,False,True, _
	  		                hojaEditar.Cells(fila,(column+1)).Left, hojaEditar.Cells(fila,(column+1)).Top, _
	  		                hojaEditar.Cells(fila,(column+1)).Width, hojaEditar.Cells(fila,(column+1)).Height
	  		                nombrePais = tabla.ChildItem(row,column, "Link",0).GetRoProperty("innertext")
	  		       	  hojaEditar.Cells(fila,(column+2)).Value=	nombrePais  
	  		       else
	  		               infCell=tabla.GetCellData(row,column)
	  		               hojaEditar.Cells(fila,(column+2)).Value=infCell
	  		       End If
	  		     End If 
	  		Next
	  		 fila=fila+1
	  	Next
	  	 fila=fila+1
   End  If	  	
  End Sub
