function redondea(num, DECIMAL)
{	cc = num.toString()
	punto = cc.indexOf('.')	
	num = parseFloat(num)
	nn = Math.floor(num)
	if (isNaN(num))
		return ''
	ceros = '0000000000000000000000000'	
	if(punto == -1)
		return Math.floor(num)+ '.'+ ceros.substr(0,DECIMAL)
	else
	{	residuo = (num-nn)*10000
		residuo = residuo.toString()+ ceros
		return Math.floor(num)+ '.'+ residuo.substr(0,DECIMAL)
	}
	
}
function decidoce(num)
{	num = parseFloat(num)
	if (num==0)
		return ("")
	nume = parseInt(num,10)
	VALOR =	num - nume
	VALOR = VALOR * 12
	VALOR = Math.round(VALOR)
//	alert(nume.toString() + " " + VALOR.toString())
	return (nume.toString() + " " + VALOR.toString())

}
function JULIANO(fecha)
{	a = trim(fecha.replace(' ',''));
	dia = parseInt(a.substring(0, 2),10)	// dia
	mes = parseInt(a.substring(3, 5),10)	// mes
	ano = parseInt(a.substring(6, 10),10)	// año
	JUL = ( 1461 * ( ano + 4800 + ( mes - 14) / 12 ) ) / 4 
	JUL += ( 367 * ( mes - 2 - 12 * ( ( mes - 14 ) /12 ) ) ) / 12 
	JUL -= ( 3 * ( ( ano + 4900 + ( mes - 14 ) / 12 ) / 100 ) ) / 4 
	JUL += dia - 32075
	//alert(JUL)
	return JUL
}
function addDays(myDate,days) 
{   dd =  new Date(myDate.getTime() + days*24*60*60*1000);
	DIA = dd.getDate()
	mes = dd.getMonth()
	mes+=1
	ano = dd.getFullYear()	
	fecha = fStrFecha(DIA+'-'+mes+'-'+ano)
	return fecha;
}

function reformat(field,len) 
{  var padding = len - field.length;
   for (i = 0; i < padding; i++) {
       field = " " + field;
    }
    return field;
}
function fStrFecha(fecha)
{
// previo a la validación de fechas
var	cFecha = "", ch, a;
// elimina blancos
	a=trim(fecha);
// reemplaza todo lo que no es dígito por "/"
	for (i = 0; i <  a.length; i++) {
		ch = a.substring(i, i+1);
		if (ch == "-") {
			ch="/"
		  }
		if ("0123456789/".indexOf(ch) < 0) {
			ch=""
		  }
		cFecha += ch;
		}
// si la primera ocurrencia es 1, agrega un cero
	if (cFecha.indexOf("/") == 1) {
		cFecha="0"+cFecha
		}
// pasa los 3 primeros caracteres a "a", cFecha queda con el resto		
	//alert('a:'+a+', cFecha:'+cFecha);
	a=cFecha.substring(0,3)		
	cFecha=cFecha.substr(3)		
	//alert('a:'+a+', cFecha:'+cFecha);
// si la segunda ocurrencia es 1, agrega un cero
	if (cFecha.indexOf("/") == 1) {
		cFecha="0"+cFecha
		}

// pasa los 3 siguientes caracteres a "a", cFecha queda con el resto		
	a+=cFecha.substring(0,3)		
	cFecha=cFecha.substr(3)		
	
// si el resto es una posicion, agrega "200"		
	if (cFecha.length == 1) {
		cFecha="200"+cFecha
		}

// si el resto son dos posiciones, agrega "20"		
	if (cFecha.length == 2) {
		cFecha="20"+cFecha
		}

// si el resto son tres posiciones, agrega "2"		
	if (cFecha.length == 3) {
		cFecha="2"+cFecha
		}

// pasa el resto a "a"
	a+=cFecha

return a;
}
function cuenta(obj, idd, numero) 
{   
	if (obj.length>parseInt(numero,10))
	{	alert("Los comentarios no deben exceder los " + numero + " caracteres")
		eval("thisForm."+idd+".focus()")
		return false;
	}
}


/*********************************************************************
funciones de caracteres, numeros, validacion de fechas, etc
*********************************************************************/
function deci(numero, deci)
{	// formatea un numero a "n" decimales
	/*var newNum = '';
	var CadNum = numero.toString(10)
	deci = parseInt(deci,10)
	punto = CadNum.indexOf('.')
	newNum  = parseInt(CadNum,10)
	xx =   numero - newNum
	newNum += '.';
	if (xx == 0)
		for (j = 0 ; j<	 deci; j++) 
			newNum += '0';
	else
		newNum += CadNum.substr(punto+1,deci)
	vv = newNum.indexOf('.')
	dd = newNum.substr(vv+1)
	if (dd.length < deci )
		for (g=dd.length; g<deci; g++ )
			newNum += '0';
	if (deci == 0)
		newNum = newNum.substr(0,newNum.length-1)		
return(newNum)*/
	var num= Number(numero)
	return num.toFixed(deci)
}
function barra(chk)
{	// devuelve una cadena con un separador antes del contenido
	if (trim(chk) != '')
		chk = ' / ' + trim(chk)
return chk;
}
function seleindice(subcadena, elemento)
{	// busca el elemento activo en un pop down dependiendo del descriptor
	var a = '';
	for (rt=0; rt<subcadena.length; rt++ )
	{	if (subcadena.substring(rt, rt+1)!= ' ' )
		{	a = a + subcadena.substring(rt, rt+1);
		}
	}
	
	for (rt=0; rt<elemento.length; rt++ )
	{	if(elemento.options(rt).value == a)
		{ return(rt);	
		}
		if(trim(elemento.options(rt).value) == trim(a))
		{ return(rt);	
		}
	} 
	
}
function trim(checkString)
{ //                    QUITA TODOS LOS BLANCOS DE LA CADENA
var	newString = "";
    // loop through string character by character
  for (i = 0; i < checkString.length; i++) 
  {	ch = checkString.substring(i, i+1);
    // quita blancos
    if (ch != ' ' ) 
    {  newString += ch;      }
  }
  return newString;
}
function ltrim(checkString)
{ //   ELIMINA LOS BLANCOS      A N T E S  DE LA CADENA
var	newString = "";
  // loop through string character by character
  for (i = 0; i < checkString.length; i++) 
  {	// quita blancos
	cont = checkString.length-i;
	ch = checkString.substring(i, i+1);
    if (ch != ' ' ) 
    {	newString = checkString.substring(i, cont);     
		break;
	}
  }
  return newString;
}
function toAlpha(checkString) 
{	// devuelve una cadena con elementos válidos para SQL
	newString = "";
	count = 0;
  // loop through string character by character
  for (i = 0; i < checkString.length; i++) 
  {	ch = checkString.substring(i, i+1);
    // concatena si es alfabético, sinó pone un espacio en blanco
    var cad = '(ch >= "a" && ch <= "z") || (ch >= "A" && ch <= "Z" ) ||';
    cad = cad + '(ch >= "0" &&  ch <= "9")|| (ch == "á")||' ;
    cad = cad + '(ch == "á")|| (ch == "é")|| (ch == "í")||';
    cad = cad + '(ch == "ó")|| (ch == "ú")|| (ch == "Á")||';
    cad = cad + '(ch == "É")||(ch == "Í") || (ch == "Ó")||' ;
    cad = cad + '(ch == " ")||' ;
    cad = cad + '(ch == "Ü")||';
    cad = cad + '(ch == "/")||' ;
    cad = cad + '(ch == "º")||' ;
    cad = cad + '(ch == "-")||' ;
    cad = cad + '(ch == "ñ")||' ;
    cad = cad + '(ch == "Ñ")||' ;
    cad = cad + '(ch == "ª")||' ;
    cad = cad + '(ch == "Ú")||(ch == ".") || (ch == "Ú")||(ch == "@")';
    if (eval(cad)== true )
    {  newString += ch;      }
    else
    {  newString += " ";}
  }
  return newString;
}
function FormatNumber(num,decimalNum,bolLeadingZero,bolParens,bolCommas, bolpunto)
/**********************************************************************
	IN:
		NUM - the number to format
		decimalNum - the number of decimal places to format the number to
		bolLeadingZero - true / false - display a leading zero for
										numbers between -1 and 1
		bolParens - true / false - use parenthesis around negative numbers
		bolCommas - put commas as number separators.
 
	RETVAL:
		The formatted number!
 **********************************************************************/
{ 
        if (isNaN(parseInt(num))) return "NaN";

	var tmpNum = num;
	var iSign = num < 0 ? -1 : 1;		// Get sign of number
	
	// Adjust number so only the specified number of numbers after
	// the decimal point are shown.
	tmpNum *= Math.pow(10,decimalNum);
	tmpNum = Math.round(Math.abs(tmpNum))
	tmpNum /= Math.pow(10,decimalNum);
	tmpNum *= iSign;					// Readjust for sign
	
	// Create a string object to do our formatting on
	var tmpNumStr = new String(tmpNum);

	// See if we need to strip out the leading zero or not.
	if (!bolLeadingZero && num < 1 && num > -1 && num != 0)
		if (num > 0)
			tmpNumStr = tmpNumStr.substring(1,tmpNumStr.length);
		else
			tmpNumStr = "-" + tmpNumStr.substring(2,tmpNumStr.length);
		
	// See if we need to put in the commas
	if (bolCommas && (num >= 1000 || num <= -1000)) {
		var iStart = tmpNumStr.indexOf(".");
		if (iStart < 0)
			iStart = tmpNumStr.length;

		iStart -= 3;
		while (iStart >= 1) {
			tmpNumStr = tmpNumStr.substring(0,iStart) + "," + tmpNumStr.substring(iStart,tmpNumStr.length)
			iStart -= 3;
		}		
	}

	// See if we need to use parenthesis
	if (bolParens && num < 0)
		tmpNumStr = "(" + tmpNumStr.substring(1,tmpNumStr.length) + ")";
		
	punto =	tmpNumStr.indexOf(".");
	ceros =""
	for (mm=0; mm<decimalNum; mm++)
		ceros += "0"
	if (punto == -1 && bolpunto== true)
		tmpNumStr += "." + ceros
	else
		tmpNumStr += ceros.substring(1,decimalNum-1)		
	return tmpNumStr;		// return our formatted string!
}
function strzero(num, largo)
{	// llena de ceros a la izquierda del numero
	var newNum = '';
	var CadNum = num.toString()
	CadNum= trim(CadNum)
	largo = parseInt(largo,10)
	ll = CadNum.length
	ceros = largo - ll
	for (i=0 ; i<ceros; i++)
	{	newNum += '0'
	}
	newNum += CadNum
return(newNum)
}
function decimal(numero, deci)
{	// formatea un numero a "n" decimales
	var newNum = '';
	var CadNum = numero.toString(10)
	deci = parseInt(deci,10)+1
	for (i = 0; i < CadNum.indexOf('.'); i++) 
		newNum += CadNum.substring(i, i+1); 
	for (j = CadNum.indexOf('.'); j<= deci; j++) 
		newNum += CadNum.substring(j, j+1);
return(newNum)
}
function toInt(checkString) 
{ // QUITA LOS CARACTERES NO NUMERICOS DE UNA CADENA DADA
	newString = "";
	count = 0;
    // loop through string character by character
  for (i = 0; i < checkString.length; i++) 
  {	ch = checkString.substring(i, i+1);
    // concatena si es numerico,
    if ((ch == "0" || ch == "1" || ch == "2" || ch == "3" || ch == "4" || ch == "5" || ch == "6" || ch == "7" || ch == "8" || ch == "9" || ch==".")) 
    {  newString += ch;      }
  }
  return newString;
}
function isInt(checkString) 
{ if (checkString.indexOf('.')>=0)
	{	alert("Este campo DEBE ser un número entero")
		return false}
 return true
}
function toTelf(checkString) 
{ // validación para teléfonos
	newString = "";
	count = 0;
  // loop through string character by character
  for (i = 0; i < checkString.length; i++) 
  {	ch = checkString.substring(i, i+1);
    // concatena si es numerico,
    if ((ch == "0" || ch == "1" || ch == "2" || ch == "3" || ch == "4" || ch == "5" || ch == "6" || ch == "7" || ch == "8" || ch == "9" || ch == "-")) 
    {  newString += ch;      }
    else {newString += " " ;}
  }
 return newString;
}
function fecha(fecha)
{	// validación de fechas
	var err=0
	var psj=0;
	a=trim(fecha.replace(' ',''));
	//formato de fecha 	 inválido
	if (a.length != 10) 
		err=1
	else	
	{	b = a.substring(3, 5)// mes
		c = a.substring(2, 3)// '/'
		d = a.substring(0, 2)// dia
		e = a.substring(5, 6)// '/'
		f = a.substring(6, 10)// año
		//errores básicos
		if (b<1 || b>12) err = 3
		if (c != '/') err = 4
		if (d<1 || d>31) err = 5
		if (e != '/') err = 4
		if (f<=2003 ) err = 2
	    if (f>2100) err= 6
		// meses de 30 dias
		if (b==4 || b==6 || b==9 || b==11){
			if (d==31) err=5
		}
		if (b==2)
		{	// febrero
			var g=parseInt(f/4)
			if (isNaN(g)) {err=1}
			if (d>29) err=5
			if (d==29 && ((f/4)!=parseInt(f/4))) err=5
		}
	}	
		
	switch (err)
	{	case 1 :
		alert('Formato de Fecha inválido!, Favor ingresar DD/MM/AAAA');
		return false;
		case 2 :
		alert("¿ Año 2003 ?")
		return false;
		case 3 :
		alert("Sólo hay 12 meses en un Año ")
		return false;
		case 4 :
		alert("El separador de fechas es una diagonal [/]")
		return false;
		case 5 :
		alert("¿Corresponde el número de días al MES ingresado ?")
		return false;
		case 6 :
		alert("¿Año mayor al 2100 ?")
		return false;
		case 0:
//		JULIANO(d,b,f)
		return true;	
		default:
			return true;
	}	
}

function isDate(fecha)
{	// validación de fechas
	var err=0
	var psj=0;
	a=trim(fecha.replace(' ',''));
	//formato de fecha 	 inválido
	if (a.length != 10) 
		err=1
	else	
	{	b = a.substring(3, 5)// mes
		c = a.substring(2, 3)// '/'
		d = a.substring(0, 2)// dia
		e = a.substring(5, 6)// '/'
		f = a.substring(6, 10)// año
		//errores básicos
		if (b<1 || b>12) err = 3
		if (c != '/') err = 4
		if (d<1 || d>31) err = 5
		if (e != '/') err = 4
		if (f<=2003 ) err = 2
	    if (f>2100) err= 6
		// meses de 30 dias
		if (b==4 || b==6 || b==9 || b==11){
			if (d==31) err=5
		}
		if (b==2)
		{	// febrero
			var g=parseInt(f/4)
			if (isNaN(g)) {err=1}
			if (d>29) err=5
			if (d==29 && ((f/4)!=parseInt(f/4))) err=5
		}
	}	
	if (err==0)
		return true
	else
		return false	
}
function isEmail(ele, op)
{  // valida campos tipo e-mail
	ele = trim(ele)
   // return false if e-mail field is blank.
   // Si op == 1, el mail puede ir en blanco
   if (ele == "" && op != 1 ) 
   {
      alert("\n El campo E-MAIL está en blanco\n\nIngrese la dirección e-mail.")
      return false; 
   }
   if (ele != '')
   // return false if e-mail field does not contain a '@' and '.' .
   if (ele.indexOf ('@',0) == -1 || 
       ele.indexOf ('.',0) == -1)
   {    alert("El E-MAIL  require  una  \"@\"y un \".\"necesariamente.\n\nRegistre adecuadamente el e-mail.")
      return false;
   }
return true;      
}
function PUNTO(obj)
{	var cad;
	cad = trim(obj);
	if(cad == '.')	
		return '' ;
	else 
		return obj;

}



/********************************/
 function FComson(nTotal)
/********************************/
{	var cStr = ""
	var nInt
// millones
nTotal = parseFloat(nTotal)
nInt = parseInt((nTotal/1000000),10)
if (nInt > 0)
{   if (nInt == 1)
      cStr = 'UN MILLON '
   else
      cStr = cStr+FCommil(nInt)+' MILLONES '
   
   nTotal  = nTotal-nInt*1000000
}
nInt=parseInt(nTotal,10)
if (nInt > 0)
{   cStr+=FCommil(nInt)+' Y '
    nTotal = nTotal-nInt
}
nTotal=nTotal*100
if (nTotal== 0)
	cStr+= '00/100 NUEVOS SOLES'
else
	cStr += parseInt(nTotal,10)+ '/100 NUEVOS SOLES'
return cStr
}
/*********************************/
 function FCommil(nTotal)
/*********************************/
{	var nInt = 0
	var cStr = ""
	nTotal = parseFloat(nTotal)
	nInt = parseInt((nTotal/1000),10)
if (nInt > 0)
{	if (nInt == 1)
		cStr+='UN MIL '
	else if (nInt == 21)
		cStr+='VENTIUN MIL '
	else
		cStr =cStr + FComcien(nInt) + ' MIL '
nTotal = nTotal - nInt * 1000
}   

nInt = nTotal

if (nInt>0)
{   if (nInt == 1)
      cStr+= 'UN '
    else
      cStr+= FComcien(nInt)
}   

return cStr
}
/***********************************/
  function FComcien(nTotal)
/***********************************/
{	//alert(nTotal)
	var cStr=""
	var nInt =0
	nTotal = parseFloat(nTotal)
	nInt = parseInt((nTotal /100	),10)
	if (nInt > 0)
	{	if (nInt == 1)
	    {	if (nTotal == 100)
				cStr = 'CIEN '
			else
				cStr = 'CIENTO '
	    }
		else if (nInt == 5)
			cStr = 'QUINIENTOS '
		else if (nInt == 7)
			cStr = 'SETECIENTOS '
		else if (nInt == 9)
			cStr = 'NOVECIENTOS '
		else
			cStr = cStr + FComunidad(nInt)  + 'CIENTOS '
	nTotal = nTotal - nInt * 100
	}
	nInt = nTotal
	if (nInt<=20)
	{  if (nInt == 0)
	      return cStr
	   else if (nInt < 10)
	      return cStr = FComunidad(nInt) // OJO ACA IBAN SOLO cStr + xxxx
	   else if (nInt == 10)
	      return cStr + 'DIEZ'
	   else if (nInt == 20)
	      return cStr + 'VEINTE'
	   else if (nInt == 11)
	      return cStr + 'ONCE'
	   else if (nInt == 12)
	      return cStr + 'DOCE'
	   else if (nInt == 13)
	      return cStr = 'TRECE'
	   else if (nInt == 14)
	      return cStr + 'CATORCE'
	   else if (nInt == 15)
	      return cStr + 'QUINCE'
	   else
	    	return cStr + 'DIECI' + FComUNIDAD(nInt-10)
	}
	nInt = parseInt((nTotal/10),10)
	if (nInt == 2)
	   cStr+='VEINT'
	else if (nInt == 3)
	   cStr+='TREINT'
	else if (nInt == 4)
	   cStr+='CUARENT'
	else if (nInt == 5)
	   cStr+='CINCUENT'
	else if (nInt == 6)
	   cStr+='SESENT'
	else if (nInt == 7)
	   cStr+='SETENT'
	else if (nInt == 8)
	   cStr+='OCHENT'
	else if (nInt == 9)
	   cStr+='NOVENT'
	   
nTotal= nTotal - nInt * 10
if (nTotal == 0)
   return cStr+'A'
	   
return cStr+'I'+FComunidad(nTotal)
}
/******************************************/
  function FComunidad(nTotal)
/******************************************/
{	
	nTotal = parseFloat(nTotal)
	if (nTotal == 2)
	   return 'DOS'
	else if (nTotal == 3)
	   return 'TRES'
	else if (nTotal == 4)
	   return 'CUATRO'
	else if (nTotal == 5)
	   return 'CINCO'
	else if (nTotal == 6)
	   return 'SEIS'
	else if (nTotal == 7)
	   return 'SIETE'
	else if (nTotal == 8)
	   return 'OCHO'
	else if (nTotal == 9)
	   return 'NUEVE'
	return 'UNO'
}
///////////////////
function FormatoHora(Caja)
{

	if(window.event.keyCode!=8)//verifico que no hallan presionado backspace
		if(Caja.value.length == 2)
			Caja.value+=":";
}
function EsNro(obj)
{
		if(obj.toString().match(/^\d+$/g))
		return true
		
	return false
}

/*Manejo de Cookies*/
function getCookie(name){
  var cname = name + '='; 
  var dc = document.cookie;             
  if (dc.length > 0) {              
    begin = dc.indexOf(cname);       
    if (begin != -1) {           
      begin += cname.length;       
      end = dc.indexOf(';', begin);
      if (end == -1) end = dc.length;
        return unescape(dc.substring(begin, end));
    } 
  }
  return null;
}
/*funcion setCookie mas simple funciona con varios browser's*/
function setCookie(name,value) {
  document.cookie = name+'='+escape(value);
}
/*funcion delCookie mas simple funciona con varios browser's*/
function delCookie (name) {
    document.cookie = name + '=' +'; expires=Thu, 01-Jan-00 00:00:01 GMT';
  
}
function EsFloat(val)
{
		if(val.toString().match(/^[0-9]+(\.[0-9]+)?$/g))
		return true
		
	return false
}