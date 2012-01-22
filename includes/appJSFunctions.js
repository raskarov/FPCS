////////////////////////////////////////////////////////////////////////////////
// These javascript functions are used through out the application
////////////////////////////////////////////////////////////////////////////////

// Cookie code from http://techpatterns.com/downloads/javascript_cookies.php
// this function gets the cookie, if it exists

function Get_Cookie( name ) {
	var start = document.cookie.indexOf( name + "=" );
	var len = start + name.length + 1;
	if ( ( !start ) &&
	( name != document.cookie.substring( 0, name.length ) ) )
	{
	return null;
	}
	if ( start == -1 ) return null;
	var end = document.cookie.indexOf( ";", len );
	if ( end == -1 ) end = document.cookie.length;
	return unescape( document.cookie.substring( len, end ) );
}

function Set_Cookie( name, value, expires, path, domain, secure ) 
{
	// set time, it's in milliseconds
	var today = new Date();
	today.setTime( today.getTime() );

	/*
	if the expires variable is set, make the correct 
	expires time, the current script below will set 
	it for x number of days, to make it for hours, 
	delete * 24, for minutes, delete * 60 * 24
	*/
	if ( expires )
	{
	expires = expires * 1000 * 60 * 60 * 24;
	}
	
	var expires_date = new Date( today.getTime() + (expires) );

	document.cookie = name + "=" +escape( value ) +
	( ( expires ) ? ";expires=" + expires_date.toGMTString() : "" ) + 
	( ( path ) ? ";path=" + path : "" ) + 
	( ( domain ) ? ";domain=" + domain : "" ) +
	( ( secure ) ? ";secure" : "" );
}


// this deletes the cookie when called
function Delete_Cookie( name, path, domain ) {
	if ( Get_Cookie( name ) ) document.cookie = name + "=" +
	( ( path ) ? ";path=" + path : "") +
	( ( domain ) ? ";domain=" + domain : "" ) +
	";expires=Thu, 01-Jan-1970 00:00:01 GMT";
}

function SetPageScroll(pName){ 
	// Used to remember scroll location of page so we 
	// can later reset scroll loaction when user returns to page
	var scrOfX = 0, scrOfY = 0;
	if( typeof( window.pageYOffset ) == 'number' ) {
		//Netscape compliant
		scrOfY = window.pageYOffset;
		scrOfX = window.pageXOffset;
	} else if( document.body && ( document.body.scrollLeft || document.body.scrollTop ) ) {
		//DOM compliant
		scrOfY = document.body.scrollTop;
		scrOfX = document.body.scrollLeft;
	} else if( document.documentElement &&
		( document.documentElement.scrollLeft || document.documentElement.scrollTop ) ) {
		//IE6 standards compliant mode
		scrOfY = document.documentElement.scrollTop;
		scrOfX = document.documentElement.scrollLeft;
	}
	Set_Cookie( pName, scrOfY, '', '/', '', '' );
}
	
function RestoreScroll(pName){
	var yVal;
	yVal = Get_Cookie(pName);
	
	if (yVal != null) {
		window.scroll(0,yVal);
	}
	Set_Cookie( pName, 0, '', '/', '', '' );
}

function jfChanged()	{
	// This function tracks if a change has been made to a form element.
	// If so it sets a form variable 'changed' to yes which you can use to
	// alert a script whether it needs to make updates to a database or not.
	document.main.changed.value = "yes";
}

function jfUpper(objField) {
	objField.value = objField.value.toUpperCase();
}

function jfToggle(toggleThis){
	//This toggles the display of a document element with the id of 
	//the parameter toggleThis
	if (document.all.item(toggleThis).style.display == 'block')	{
		document.all.item(toggleThis).style.display = 'none'
	}
	else{
		document.all.item(toggleThis).style.display = 'block'
	}
}

function jfMultiply(item1,item2,result){
	//Multipies 2 numbers togather and returns the product.
	//all three parameters are passed in form elements.
	//item1 and item2 are the form elements from which we get the values we are to multiply
	//result is the form element in which the product will be displayed
	var a = parseFloat(item1.value);
	var b = parseFloat(item2.value);
	result.value = a * b;
}	

function jfMaxSize(intMaxSize,formElement){
	if (formElement.value.length >= intMaxSize) {
		alert("You have gone over maximum number of characters you can enter into this text box. The max is " + intMaxSize + ".");
		formElement.value = formElement.value.substring(0,intMaxSize-1);
	}
}

function jfSelectItemFromTo(selectFrom, selectTo) {
	//based on ideas from excite.com's weather selection - heavily modified
	selectFrom = document.all.item(selectFrom);
	selectTo = document.all.item(selectTo);
	var blnSelected = false;
	var selected = selectFrom.selectedIndex;
	if (selected != -1){
		for (j=0; j<selectFrom.length; j++) {
			if (selectFrom.options[j].selected){
				var selectedText = selectFrom.options[j].text;
				var selectedValue = selectFrom.options[j].value;
				if (selectedValue != "") {
					var toLength = selectTo.length;
					var i;
					// If item is already added, give it focus
					for (i=0; i<toLength; i++) {
						if (selectTo.options[i].value == selectedValue) {
							blnSelected = true;
						}
					}
					if (!blnSelected){
						// Add new option 
						selectTo.options[selectTo.length] = new Option(selectedText, selectedValue);
					}
				}
			}
			blnSelected = false;
		}
	}
}	

function jfRemoveItems(pobjSelect){
	//remove items from multiple select list
	//Since setting an option to NULL changes the index
	//value of the item beneath it, we have to make a couple
	//of passes at the object.  We first grab the quantity of
	//items selected, then we use that as a counter to remove
	//the selected items
	pobjSelect = document.all.item(pobjSelect)
	var iCnt = 0;
	for (i=0; i<pobjSelect.length; i++) {
		if (pobjSelect.options[i].selected){
			iCnt ++;
		}
	}
	for (j=0; j<iCnt; j++){
		for (i=0; i<pobjSelect.length; i++) {
			if (pobjSelect.options[i].selected){
				pobjSelect.options[i] = null;
			}
		}
	}
}


function SwapElement(pselObj, nDirection){
	// from http://groups.google.com/groups?q=move+items+option+up+down+javascript&hl=en&lr=&ie=UTF-8&scoring=r&selm=%23GOaRzPnAHA.2164%40tkmsftngp05&rnum=7 
	// Take the currently selected element in SelectColumns
	// and swap it with the element that is nDirection from
	// the current element
	with(pselObj)  {
		// If there is more than one item selected, alert the user
		var nCount = 0;
		for(var x = 0; x < length; x++) {
			if(options[x].selected) {nCount++;}
		}

		if(nCount > 1) {
			alert("Please select a single column to move up or down");
			return;
		}

		var nIndex = selectedIndex;
		if(nIndex == -1){
			alert("Please select a column to move up or down");
			return;
		}

		// Make sure we are not the top element
		// or bottom element trying to move too far
		var nSwapIndex = nIndex + nDirection;
		if(nSwapIndex < 0 || nSwapIndex >= length){return;}

		var nValue = options[nIndex].value;
		var strText = options[nIndex].text;

		var nSwapValue = options[nSwapIndex].value;
		var strSwapText = options[nSwapIndex].text;

		options[nIndex] = new Option(strSwapText, nSwapValue);
		options[nSwapIndex] = new Option(strText, nValue);

		selectedIndex = nSwapIndex;
	}
}

