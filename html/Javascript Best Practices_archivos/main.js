// Turn on background image caching in IE
// --------------------------------------
/*@cc_on
if (document && document.execCommand) {
	try { document.execCommand("BackgroundImageCache",false,true); }
	catch (e) { }
}
@*/ 

// Returns true if a tag is assigned a given class
// ===============================================
function hasClass(obj,name) {
	if (obj.className) {
		if ((" "+obj.className+" ").indexOf(name)>-1) { 
			return true;
		}
	}
	return false;
}

// Get the closest DIV to the object with class="example"
// ======================================================
function getExampleDiv(obj) {
	while (obj.nextSibling && (obj = obj.nextSibling)) {
		if (hasClass(obj,"example")) { 
			return obj;
		}
	}
	return null;
}

var uniqueDivId = 1;
// Show an object's associated example
// ===================================
function showExample(obj) {
	var exampleDiv = getExampleDiv(obj);
	if (!exampleDiv.id) {
		var newId = "div"+(uniqueDivId++);
		while (document.getElementById(newId)!=null) {
			var newId = "div"+(uniqueDivId++);
		}
		exampleDiv.id = newId;
	}
	if (exampleDiv!=null && exampleDiv.id) {
		toggleDiv(exampleDiv.id);
	}
}