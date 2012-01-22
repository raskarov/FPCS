<!--
	// code from http://www.hiermenuscentral.com/bulletins/6/
   HM_DOM = document.getElementById ? true : false;
   HM_IE  = document.all ? true : false;
   HM_NS4 = document.layers ? true : false;

   function HM_f_ToggleElementList(show,elList,toggleBy) {
      if(!(HM_DOM||HM_IE||HM_NS4)) return true;

      if(HM_NS4&&(toggleBy=="tag")) return true;

      for(var i=0; i<elList.length; i++) {
         var ElementsToToggle = [];
         switch(toggleBy) {
            case "tag":
               ElementsToToggle = 
     (HM_DOM) ? document.getElementsByTagName(elList[i]) :
     document.all.tags(elList[i]);
               break;
            case "id":
               ElementsToToggle[0] = 
     (HM_DOM) ? document.getElementById(elList[i]) :
     (HM_IE) ? document.all(elList[i]) : 
     document.layers[elList[i]];
               break;
         }
         for(var j=0; j<ElementsToToggle.length; j++) {
            var theElement = ElementsToToggle[j];
            if(!theElement) continue;
            if(HM_DOM||HM_IE) {
               theElement.style.visibility = 
                  show ? "inherit" : "hidden";
            } else {
               theElement.visibility = 
                  show ? "inherit" : "hide";
            }
         }
      }
      return true;
   }
// -->