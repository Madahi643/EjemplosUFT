'declaracion variables
Set txtWiki= Browser("Wikipedia, la enciclopedia").Page("Wikipedia, la enciclopedia").WebEdit("WebEdit")
Set buttonSearch= Browser("Wikipedia, la enciclopedia_3").Page("Wikipedia, la enciclopedia").WebButton("WebButton")
Set  linkMunWiki= Browser("Wikipedia, la enciclopedia").Page("Mundial 2002 - Wikipedia,").Link("Copa Mundial de Fútbol")
Set  linkFaseWiki = Browser("Wikipedia, la enciclopedia").Page("Copa Mundial de Fútbol").Link("5 Fase de grupos")

'llamada de metodos
inicio

Sub inicio
	    abrirNavegador Parameter("Navegador"), "https://es.wikipedia.org/wiki/Wikipedia:Portada"
	    wait 2	   
	    if objExistente(txtWiki) =1Then
	       	txtWiki.Set("Mundial 2002")
	       	wait 1
	       	buttonSearch.Click     	
	     End If
	     Wait 1
	  
	      If objExistente(linkMunWiki)=1 Then
	       	linkMunWiki.Click
	      End If
	      Wait 1
	  
	      If objExistente(linkFaseWiki)=1 Then
	  	      linkFaseWiki.Click
	       End If
	      Wait 1
	  
  End  Sub
  
  
  
 
