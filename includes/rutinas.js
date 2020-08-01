function FormatoHora(Caja)
{

	if(window.event.keyCode!=8)//verifico que no hallan presionado backspace
		if(Caja.value.length == 2)
			Caja.value+=":";
}
function EsNro(val)
{
		if(val.toString().match(/^\d+$/g))
		return true
		
	return false
}
function EsFloat(val)
{
		if(val.toString().match(/^[0-9]+(\.[0-9]+)?$/g))
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
function MensajeForm(txt,obj){
	alert(txt)
	obj.blur()
	obj.focus()
}
function Formatotxt(obj,esp,conv)
{	
	/*Elimina espacios*/
	if(esp==true)
	obj.value = obj.value.replace(/ /g,'');
	/*Convierte a Mayusculas*/
	
	if(conv==1)
		obj.value=obj.value.toUpperCase();
	else if(conv==2)
		obj.value=obj.value.toLowerCase();
}
function FormatCeroIsq(obj,nceros){
	var ceros,val,lon,ini
	
	for (a=1;a < nceros;a++) 
		ceros += '0'
			
	val = ceros + obj.value
	lon = val.length
	ini = Number(lon)-nceros	
	obj.value=val.substr(ini,lon)
}
function pausecomp(millis) 
{
date = new Date();
var curDate = null;

do { var curDate = new Date(); } 
while(curDate-date < millis);
} 
function objetus() {
	try {
		objetus = new ActiveXObject("Msxml2.XMLHTTP");        
	} 
	catch (e) {
		try {                        
			objetus= new ActiveXObject("Microsoft.XMLHTTP");                
		} 
		catch (E){
			objetus= false;                
		}        
	}        
	if (!objetus && typeof XMLHttpRequest!='undefined') {
		objetus = new XMLHttpRequest();
	}
	return objetus
}

function deci(numero, deci)
{	
	var num= Number(numero)
	return num.toFixed(deci)
}
//Number.toFixed(4) le da formato en decimales al numero
