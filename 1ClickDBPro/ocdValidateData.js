<SCRIPT LANGUAGE=JavaScript>
<!-- Begin
//1 Click DB Client Side JavaScript Validation Routines
//copyright 1997-2001 AccessHelp.net Custom Solutions
//
//This program and its source code is fully protected by
//international copyright laws and may not be used, altered 
//or redistributed in any way contrary to its licensing.
//
//Redistribution and use in source and binary forms, with or without 
//modification, are permitted provided that the following conditions are met:
//
//Redistributions of source code must retain the above copyright notice, 
//this list of conditions and the following disclaimer. 
//
//Redistributions in binary form must reproduce the above copyright notice, 
//this list of conditions and the following disclaimer in the documentation 
//and/or other materials provided with the distribution. 
//
//Neither the name of 1 Click DB or AccessHelp.net nor the names of its 
//contributors may be used to endorse or promote products derived from 
//this software without specific prior written permission. 
//
//THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
//AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE 
//IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE 
//ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDERS OR CONTRIBUTORS BE 
//LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR 
//CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF 
//SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS 
//INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN 
//CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) 
//ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE 
//POSSIBILITY OF SUCH DAMAGE.
//
//More Info via email to info@1ClickDB.com

var acwValidationIsExecuting = false;
function acwValidateNumericInput(objInput,strOKChars,numMinValue, numMaxValue)
	{
	// Assume Success
	var inputOK = true;
	// List of valid input characters
	strOKChars = "0123456789" + strOKChars
	var strJustNumber = "";
	var intI;
	var intJ;
	var strCheck;
	for (intI = 0;  intI < objInput.value.length;  intI++)
		{
		strCheck = objInput.value.charAt(intI);
		for (intJ = 0;  intJ < strOKChars.length;  intJ++)
		if (strCheck == strOKChars.charAt(intJ))
			{
			break;
			}
		if (intJ == strOKChars.length)
			{
			inputOK = false;
			break;
			}
		if (strCheck != ",")
			{
			strJustNumber += strCheck;
			}
		}
	if (!inputOK)
		{
		alert("Please enter a numeric value.");
		}
		
		if (strJustNumber != "")
			{
			var numJustNumber = parseInt(strJustNumber);
			if (numMinValue != null && numMaxValue != null)
				{
				if (!(numJustNumber >= numMinValue && numJustNumber <= numMaxValue))
					{
					alert("Please enter a value between " + numMinValue + " and " + numMaxValue + ".");
					inputOK = false;
					}
				}
			}	
		if (!inputOK)
			{
			objInput.select();
			objInput.focus();
			return false;
			}
		else
			{
			return true;
			}
	}

	//end function
	
	function acwValidateTextInput(objInput,strOKChars, strBadChars, numMinChars, numMaxChars)
	{
	if (!acwValidationIsExecuting)
	{
	acwValidationIsExecuting = true;
	// Assume Success
	var inputOK = true;
	var intI;
	var intJ;
	var strCheck;

	if (strOKChars != null)
		{
		for (intI = 0;  intI < objInput.value.length;  intI++)
			{
			strCheck = objInput.value.charAt(intI);
			for (intJ = 0;  intJ < strOKChars.length;  intJ++)
			if (strCheck == strOKChars.charAt(intJ))
				{
				break;
				}
			if (intJ == strOKChars.length)
				{
				inputOK = false;
				break;
				}
			}
		}
	if (!inputOK)
		{
		alert("Only the following characters are allowed " + strOKChars + ".");
		}
	if (inputOK)
		{
		if (strBadChars != null)
			{
			intI = 0;
			ingJ = 0
			for (intI = 0;  intI < objInput.value.length;  intI++)
				{
				strCheck = objInput.value.charAt(intI);
				for (intJ = 0;  intJ < strBadChars.length;  intJ++)
				if (strCheck == strBadChars.charAt(intJ))
					{
					inputOK = false;
					break;
					}
				}
			}
		if (!inputOK)
			{
			alert("The following characters are not allowed " + strBadChars + ".");
			}
		}
	//Min Characters Check
	if (inputOK)
		{	
		if (numMinChars != null)
			{
			if (objInput.value.length < numMinChars)
				{
				inputOK = false;
				} 
			}	
		if (!inputOK)
			{
			alert("Please enter at least " + numMinChars + " characters");
			}	
	}
	//Max Characters Check
	if (inputOK)
		{
		if (numMaxChars != null)
			{
			if (objInput.value.length > numMaxChars)
				{
				inputOK = false;
				} 
			}
			if (!inputOK)
				{
				alert("Please enter no more than " + numMaxChars + " characters");
				}	
		}
	if (!inputOK)
		{

		objInput.select();
		objInput.focus();
		acwValidationIsExecuting = false;
		return false;
		}
	else
		{
		acwValidationIsExecuting = false;
		return true;
		}
	}
	}

//  End -->
</script>
