// Rev. 09/07/2003

function Toggle(item) {
   obj=document.getElementById(item);
   visible=(obj.style.display!="none")
   key=document.getElementById("x" + item);
   if (visible) {
     obj.style.display="none";
     key.innerHTML="<img src='folder.gif' width='16' height='16' hspace='0' vspace='0' border='0'>";
   } else {
      obj.style.display="block";
      key.innerHTML="<img src='textfolder.gif' width='16' height='16' hspace='0' vspace='0' border='0'>";
   }
}

function Expand() {
   divs=document.getElementsByTagName("DIV");
   for (i=0;i<divs.length;i++) {
     divs[i].style.display="block";
     key=document.getElementById("x" + divs[i].id);
     key.innerHTML="<img src='textfolder.gif' width='16' height='16' hspace='0' vspace='0' border='0'>";
   }
}

function Collapse() {
   divs=document.getElementsByTagName("DIV");
   for (i=0;i<divs.length;i++) {
     divs[i].style.display="none";
     key=document.getElementById("x" + divs[i].id);
     key.innerHTML="<img src='folder.gif' width='16' height='16' hspace='0' vspace='0' border='0'>";
   }
}

