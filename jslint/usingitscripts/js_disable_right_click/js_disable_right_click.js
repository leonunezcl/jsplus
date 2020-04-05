/*---------------------------------------------------------------------------

Disable Right Click Script for Web Designers Toolkit
(C)1999-2003 USINGIT.COM, All Rights Reserved.
To get more free and professional scripts, visit:
http://www.usingit.com/
http://www.usingit.com/products/webtoolkit
email: support@usingit.com

---------------------------------------------------------------------------*/

function onMouseDownIE4(){
	if(event.button==2){
		alert(warningMessage);
		return false;
	}
};

function onMouseDownNS4(e){
	if(document.layers||document.getElementById&&!document.all){
		if(e.which==2||e.which==3){
			alert(warningMessage);
			return false;
		}
	}
};

if(document.layers){
	document.captureEvents(Event.MOUSEDOWN);
	document.onmousedown=onMouseDownNS4;
}
else if(document.all&&!document.getElementById){
	document.onmousedown=onMouseDownIE4;
};

document.oncontextmenu=new Function("alert(warningMessage);return false");
