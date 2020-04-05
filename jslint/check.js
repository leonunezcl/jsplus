function go(){
var o = {};
jslint(document.forms.jslint.input.value, o);
document.getElementById('output').innerHTML = jslint.report();
}

