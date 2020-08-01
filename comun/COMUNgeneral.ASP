<body >

<script ID="clientEventHandlersJS" LANGUAGE="javascript">
var pagesize = '<%=pagesize%>';

function GRABAR_onclick() 
{ var ok = VALIDA();
 //  alert(VALIDA())
  
       //alert(chk)
	if (ok == true )
	{   cad = completa()

        cad += '&cod=' + trim(thisForm.COD.value)  ;
		cad += '&des=' + Left(toAlpha(thisForm.DES.value),largo2)  ;
        cad += '&chk='  
	    cad += '&LARGO=' + largo   
        cad += '&url=' + url 
    //alert (funcionalidad + cad)
 /* document.all.ACTIV.style.visibility='visible'
    document.all.ACTIV.height="150"
    document.all.ACTIV.width="100%"
 */
    document.all.ACTIV.src=(funcionalidad + cad); }
}

function elimina()
{	var si
	{si = confirm("¿ Confirma la Eliminación de este Registro ? ");} 
	if (si == false)
	{ return false;}
	var cad =	'COD='	 + thisForm.COD.value  ;
	cad = cad + '&chk=1' 	
	cad = cad + '&url=' + url ;
	cad = cad + '&tbl=' + TBL
	cad = cad + '&PK=' + PK
	cad = cad + '&DS=' + DS		
    //alert (funcionalidad + cad)
 /*   document.all.ACTIV.style.visibility='visible'
    document.all.ACTIV.height="150"
    document.all.ACTIV.width="100%"
*/
	document.all.ACTIV.src=(funcionalidad + cad);
	return true;
}

function dd(ff, op) 
{	var pos = parseInt(ff,10)
	var t = document.all.TABLA;
	if ((oldrow%2) ==0)
		eval("document.all.fila" + oldrow + ".style.backgroundColor='<%=(Application("color1"))%>'");
	else
		eval("document.all.fila" + oldrow + ".style.backgroundColor='<%=(Application("color2"))%>'");
	// PINTA LA LINEA DEL COLOR OSCURO (PREDETERMINADO )
	eval("document.all.fila"+ff+".style.backgroundColor='<%=(Application("BARRA"))%>'");
    // Rellena los valores en los textbox
	    thisForm.COD.value = ltrim(t.rows(pos).cells(0).innerText );
	    thisForm.DES.value = ltrim(t.rows(pos).cells(1).innerText) ;
	    thisForm.COD.readOnly= true
        llena(pos);
	oldrow=pos
	return true;
}
function retrocede() 
{	var t = document.all.TABLA;
	var ff = 1;
	if (oldrow>1)
		dd(oldrow-1);
}
function avanza() 
{	var t = document.all.TABLA;
	var ff = 1;
	
  var max 
  ff =oldrow  + 1
  max =  t.rows.length -1;
  if (ff <= max)
   {dd(ff); }
}
function pagina() 
{	var t = document.all.TABLA;
	var i = t.rows.length -1;
	// captura el valor del último código mostrado en el GRID
	var cad = t.rows(i).cells(0).innerText;
    pag = '<%=urlbase%>' + '?pos='
   // alert(pag + cad)
	window.location.replace(pag + cad);
}
function primera(pag) 
{	window.location.replace('<%=urlbase%>' + '?pos=0');
}
function atras() 
{	var t = document.all.TABLA;
	var cod = t.rows(1).cells(0).innerText; // primer valor del browse
	var cad = 'comun/BACK0.asp?pos='+ trim(cod) + '&pag='+ pagesize ;
	cad = cad + '&alias='+ alias + '&pk='+PK + '&url='+ url+'?';
	//alert(cad)
	//return true;
	window.location.replace(cad)	;
}
function ultima() 
{var cad = 'comun/last0.asp?pag='+ pagesize  ;
	cad = cad + '&ALIAS='+ alias   ;
	cad = cad + '&PK='+ PK+ '&url='+url
	//alert(cad)
	window.location.replace( cad );
}
function imprime()
{	cad = 'REPORTES/prnTBL.asp?pk=' + PK + '&ds=' + DS+ '&tbl=' +alias + '&tit=' + '<%=titulo%>'
	window.open(cad)
}
	
function BUSCA(url, alias) 
{	if (trim(window.thisForm.kod.value) == '' && trim(window.thisForm.ds.value) == '')
	{	alert("Debe informar al menos el Código o la descripción");
		window.thisForm.kod.focus();
	}
    cad = url+'?pos=' + trim(window.thisForm.kod.value)
    cad += '&des=' + ltrim(window.thisForm.ds.value)
	//alert(cad)
    window.location.replace( cad );
    
	return true;
}
</SCRIPT>
