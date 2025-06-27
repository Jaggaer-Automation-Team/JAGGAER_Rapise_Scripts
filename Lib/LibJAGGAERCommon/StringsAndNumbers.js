// package resources.common // Original Groovy package

SeSGlobalObject("StringsAndNumbers");

//Describe the library purpose
/**
 * The library contains common String and number methods
 * The methods can return a derived value or a boolean value.
 */

//Add Imports
// import org.apache.commons.lang3.StringUtils // Conceptual import

//Add Shared Resources
// import resources.common.GlobalVariables as GVars // Conceptual import
// import resources.common.TestCaseExecParams as TSExecParams // Conceptual import


// --- MOCK IMPLEMENTATIONS FOR EXTERNAL DEPENDENCIES ---
// IMPORTANT: These are highly simplified mocks for structural translation.
// Replace with actual implementations or library calls.

/*
const GVars = {
	GblNull: (value) => null, // Or a specific string like "##NULL##" if that's the convention
	GblUndefined: (value) => undefined, // Or a specific string
	GBLSpace: (value) => " " // Mock for a potential GVars method (not explicitly defined but used in JComm_ReturnValueFromString)
};

const TSExecParams = {
	getBoolDoDebug: () => false // Default mock value
};
*/

// Mock for org.apache.commons.lang3.StringUtils (if it were used)
/*
const StringUtils = {
	// Example: isNumeric: (str) => !isNaN(parseFloat(str)) && isFinite(str)
};
*/


// -------------------------------------
// --- FUNCTIONS (from StringsAndNumbers) ---
// -------------------------------------

/** ---------------------------- String Functions ------------------------------*/
/**
*-------------------- JComm_HandleNoData -----------------------------
	 * Trim the string and determine if the value is null or empty
	 * @param strInput The value to test
	 * @return result The output value
	 * @author pkanaris
	 * @author Created: 04/23/2021
	 * @author Last Edited:
	 * @author Last Edited By:
	 * @author Edit Comments: (Include email, date and details)
*/
// public static String JComm_HandleNoData (String strInput){
function StringsAndNumbers_JComm_HandleNoData (strInput){
	let result;
	if (strInput != null){ // Checks for null and undefined in JS
		// Groovy's stripLeading() removes leading whitespace. JS trimStart() or trimLeft()
		// Groovy's trim() removes leading and trailing. JS trim()
		// The sequence strInput.stripLeading().trim() seems redundant if trim() is comprehensive.
		// Let's assume strInput.trim() is sufficient for removing leading/trailing.
		// And .stripLeading() might be a custom Groovy or Java extension not standard.
		// For JS:
		result = String(strInput).trimStart().trim(); // Equivalent to strInput.stripLeading().trim()
		// In Groovy, strInput.isEmpty() checks if length is 0.
		// result == "" checks if after stripping/trimming it's empty.
		if (String(strInput).trim() === "" || result === ""){ // Check original trimmed and final result
			result = GblNull;
		}
	}
	else {
		result = GblNull; //Unable to assign the gblVariables
	}
	return result;
}

/**
 * ------------------ JComm_GetLeftTextInString -----------------------
	 * Gets the text left of the assigned in string value.
	 * @param strInputText The value to parse
	 * @param textToFind The value to find and parse to
	 *
	 * @return strLeftText The text value to be returned.
	 *
	 * @author pkanaris
	 * @author Created 04/23/2021
	 * @author Last Edited:
	 * @author Last Edited By:
	 * @author Edit Comments: (Include email, date and details)
 */
 // public static String JComm_GetLeftTextInString (String strInputText, String textToFind) {
function StringsAndNumbers_JComm_GetLeftTextInString (strInputText, textToFind) {
	//Define variables
	let intTextPos;
	let strLeftText = null; // Match Groovy initialization

	if (strInputText == null) return null; // Handle null input gracefully

	// intTextPos = strInputText.indexOf(textToFind, 0) //Makes sure we start at 0
	intTextPos = String(strInputText).indexOf(String(textToFind)); // JS indexOf starts at 0 by default

	if (intTextPos >= 0){
		//Return the left text
		// strLeftText = this.JComm_HandleNoData(strInputText.substring(0, intTextPos).trim())//removed intTextPos - 1 cutoff last character of text PGK 07252019
		strLeftText = JComm_HandleNoData(String(strInputText).substring(0, intTextPos).trim()); // `this.` removed
	}
	else {
		strLeftText = strInputText; // Return original if not found
	}
	return strLeftText;
}

/**
* ------------------ JComm_GetRightTextInString -----------------------
 * Gets the text right of the assigned in string value.
 * @param strInputText The value to parse
 * @param strToFind The value to find and parse to
 *
 * @return strRightText
 * @author pkanaris
 * @author Created 4/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static String JComm_GetRightTextInString (String strInputText, String strToFind) {
function StringsAndNumbers_JComm_GetRightTextInString (strInputText, strToFind) {
	//Define variables
	let intTextPos;
	let intStrLen, intLeftLen; // intLeftLen in Groovy was length of strToFind
	let strRightText = ''; // Initialize to empty string like Groovy

	if (strInputText == null) return ''; // Handle null input gracefully, returning empty like Groovy's default

	const sInputText = String(strInputText);
	const sToFind = String(strToFind);

	// intTextPos = strInputText.lastIndexOf(strToFind)
	intTextPos = sInputText.lastIndexOf(sToFind);
	// intLeftLen = strToFind.length()
	intLeftLen = sToFind.length;
	// intStrLen = strInputText.length()
	intStrLen = sInputText.length;

	if (intTextPos >= 0){
		//Return the left text // Comment says left, but it's right
		// strRightText = this.JComm_HandleNoData(strInputText.substring(intTextPos + intLeftLen, intStrLen).trim())
		strRightText = JComm_HandleNoData(sInputText.substring(intTextPos + intLeftLen, intStrLen).trim()); // `this.` removed
	}
	else {
		strRightText = sInputText; // Return original if not found
	}
	return strRightText;
}

// public static String JComm_ReplaceValue (String strInput, String strOldValue, String strNewValue) {
function StringsAndNumbers_JComm_ReplaceValue (strInput, strOldValue, strNewValue) {
	if (strInput == null) return null;
	// Groovy's replace() behaves like JS replaceAll if strOldValue is a string.
	// For precise behavior match (if strOldValue could be regex in Groovy use case):
	// return String(strInput).replace(new RegExp(strOldValue.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), strNewValue);
	// Assuming simple string replacement:
	return String(strInput).split(String(strOldValue)).join(String(strNewValue)); // More robust for all occurrences
	// Or, if only first occurrence was intended (though Groovy's String.replace is usually all for char/string):
	// return String(strInput).replace(String(strOldValue), String(strNewValue));
}

// public static String JComm_ReplaceLineFeedWithGblLineFeed(String strInput) {
function StringsAndNumbers_JComm_ReplaceLineFeedWithGblLineFeed(strInput) {
	//May need to check the OS to determine the value to replace and potentially try other values based on the object
	if (strInput == null) return null;
	// String strTempVal = strInput.replaceAll("\r\n", "%n").replace("\n", "%n")
	// The "%n" is usually for printf-style formatting, not literal replacement.
	// Assuming the goal is to normalize to a GVars.GblLineFeed.
	// Let's assume GVars.GblLineFeed gives the desired consistent line feed (e.g., "\n").
	let strTempVal = String(strInput).replace(/\r\n/g, GVars.GblLineFeed('Value')); // Replace CR LF
	strTempVal = strTempVal.replace(/\n/g, GVars.GblLineFeed('Value'));   // Replace LF
	strTempVal = strTempVal.replace(/\r/g, GVars.GblLineFeed('Value'));   // Also replace CR just in case
	return strTempVal;
}

/**
* ------------------ JComm_VerifyTextPresent -----------------------
 * Verifies if the specified text is displayed within the string with the corresponding rule.
 * @param strDispText The value to parse
 * @param strExpText The value to find and parse to
 * @param strTextRule The rule on which expression to determine if the text is displayed.
 * @return boolTextIsPresent Is the text present in the correct location?
 * @author pkanaris
 * @author Created 4/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static boolean JComm_VerifyTextPresent (String strDispText, String strExpText, String strTextRule) {
function StringsAndNumbers_JComm_VerifyTextPresent (strDispText, strExpText, strTextRule) {
	let boolTextIsPresent = false;
	// Integer intTextLocation, intExpTextLength, intDispTextLength // Used later

	if (strDispText == null || strExpText == null) return false; // Cannot verify if either is null

	const sDispText = String(strDispText);
	const sExpText = String(strExpText);

	//check if text is present
	// intTextLocation = this.JComm_TextLocationInString(strDispText, strExpText)
	const intTextLocation = JComm_TextLocationInString(sDispText, sExpText); // `this.` removed

	if (intTextLocation >= 0) { // Text is found (contains)
		switch (strTextRule) {
			case '%==%': //Matches
				if (sDispText == sExpText) {
					boolTextIsPresent = true;
				}
				break;
			case '%Con%': //Contains
				boolTextIsPresent = true; // Already true if intTextLocation >= 0
				break;
			case '%St%': //Starts with
				// if (intTextLocation == 1) { // Groovy String.indexOf is 0-based. So starts with means index 0.
				if (intTextLocation == 0) {
					boolTextIsPresent = true;
				}
				break;
			case '%End%': //Ends With
				const intExpTextLength = sExpText.length;
				const intDispTextLength = sDispText.length;
				// if (intTextLocation == intDispTextLength - intExpTextLength) {
				// This condition is a bit tricky. String.endsWith is simpler in JS.
				// sDispText.endsWith(sExpText)
				if (sDispText.endsWith(sExpText)) {
					boolTextIsPresent = true;
				}
				break;
			default: // Unknown rule
				boolTextIsPresent = false;
				break;
		}
	} else {
		boolTextIsPresent = false; // Text not found at all
	}
	return boolTextIsPresent;
}

/**
* ------------------ JComm_VerifyTextPresent ----------------------- <!-- Title duplicated -->
 * Verifies if the specified text is displayed within the string with the corresponding rule. <!-- Description Duplicated -->
 * @param strDispText The value to parse
 * @param strExpText The value to find and parse to
 * @param strTextRule The rule on which expression to determine if the text is displayed.
 * @return boolTextIsPresent Is the text present in the correct location? <!-- Return type in Groovy is Map -->
 * @author pkanaris
 * @author Created 4/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static Map JComm_ReturnValueFromString (String strDispText, String strExpText, String strTextRule) {
function StringsAndNumbers_JComm_ReturnValueFromString (strDispText, strExpText, strTextRule) {
	let boolTextIsPresent = true; // Initial assumption, might be set to false
	let strValue = GblNull;
	let strExpTextTemp;
	let strMethodDetails = ""; // Initialize
	const mapReturn = {}; // Plain JS object
	// Integer intTextLocation, intExpTextLength, intDispTextLength // Not directly used here

	if (strDispText == null) { // Handle null input
		boolTextIsPresent = false;
		strMethodDetails = "FAILED!!! The displayed value is null!!!";
		mapReturn['boolTextIsPresent'] = boolTextIsPresent.toString();
		mapReturn['strMethodDetails'] = strMethodDetails;
		mapReturn['strValue'] = strValue;
		return mapReturn;
	}

	if (strExpText == GVars.GBLSpace('Value')) { // GBLSpace needs to be defined in GVars mock
		strExpTextTemp = " ";
	}
	else {
		strExpTextTemp = strExpText;
	}

	const sDispText = String(strDispText);
	const sExpTextTemp = String(strExpTextTemp);

	//check if text is present
	// intTextLocation = this.JComm_TextLocationInString(strDispText, strExpTextTemp)
	const intTextLocation = JComm_TextLocationInString(sDispText, sExpTextTemp);

	if (intTextLocation >= 0) {
		switch (strTextRule) {
			case '%==%': //Matches
				if (sDispText == sExpTextTemp) {
					// boolTextIsPresent = true; // Already true
					strMethodDetails = "The displayed text '" + sDispText + "' matches the expected text.";
				} else {
					boolTextIsPresent = false;
					strMethodDetails = "The displayed text '" + sDispText + "' DOES NOT EXACTLY MATCH the expected text '" + sExpTextTemp + "'.";
				}
				break;
			case '%Left%':
				// strValue = this.JComm_HandleNoData(this.JComm_GetLeftTextInString(strDispText, strExpTextTemp))
				strValue = JComm_HandleNoData(JComm_GetLeftTextInString(sDispText, sExpTextTemp));
				strMethodDetails = `Returned text to the left of '${sExpTextTemp}': '${strValue}'.`;
				break;
			case '%Right%': //Starts with <!-- Comment seems wrong for %Right% -->
				// strValue = this.JComm_HandleNoData(this.JComm_GetRightTextInString(strDispText, strExpTextTemp))
				strValue = JComm_HandleNoData(JComm_GetRightTextInString(sDispText, sExpTextTemp));
				strMethodDetails = `Returned text to the right of '${sExpTextTemp}': '${strValue}'.`;
				break;
			default:
				boolTextIsPresent = false; // Or treat as error
				strMethodDetails = `FAILED!!! Unknown rule '${strTextRule}' provided.`;
				break;
		}
	} else {
		boolTextIsPresent = false;
		strMethodDetails = "FAILED!!! The displayed value '" + sDispText + "' DOES NOT CONTAIN THE VALUE '" + sExpTextTemp + "' !!!";
	}
	mapReturn['boolTextIsPresent'] = boolTextIsPresent.toString();
	mapReturn['strMethodDetails'] = strMethodDetails;
	mapReturn['strValue'] = strValue;
	return mapReturn;
}

/**
* ------------------ JComm_TextLocationInString -----------------------
 * Gets the location of the text in the assigned string value.
 * @param strInputText The value to parse
 * @param strExpectedSubString The value to find and parse to
 *
 * @return intTextPos The interger value of the location of the sting in the text Zero based.
 * @author pkanaris
 * @author Created 4/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static int JComm_TextLocationInString(String strInputText, String strExpectedSubString){
function StringsAndNumbers_JComm_TextLocationInString(strInputText, strExpectedSubString){
	if (strInputText == null || strExpectedSubString == null) return -1;
	// int intTextPos = strInputText.indexOf(strExpectedSubString, 0);
	const intTextPos = String(strInputText).indexOf(String(strExpectedSubString)); // Second arg (fromIndex) defaults to 0
	return intTextPos;
}

/**
* ------------------ JComm_StripLeadingAssignChars -----------------------
 * Remove special non-print characters from the leading value of a string.
 * @param strIn The value to parse
 *
 * @return strOut The string resulting in the check for special leading characters.
 * @author pkanaris
 * @author Created 4/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static String JComm_StripLeadingAssignChars(String strIn){
function StringsAndNumbers_JComm_StripLeadingAssignChars(strIn){
	//Included characters are letters, numbers spaces, and special characters
	//All others will be removed.
	const boolDoDebug = TCEP_boolDoDebug;
	if (strIn == null || strIn === "") return strIn; // Handle null or empty
	if (boolDoDebug == false){
		Tester.Message('StringsAndNumbers_JComm_StripLeadingAssignChars The strIn value: ' + strIn)
	}
	const sIn = String(strIn);
	const intStrLen = sIn.length;
	// def charAscii = GVars.GblUndefined('Value') // Becomes number
	let charAscii;
	let strOut = GblUndefined; // Initialize with a defined value or strIn

	// charAscii = strIn.codePointAt(0)
	charAscii = sIn.codePointAt(0); // Returns undefined if string is empty, handle above

	if (boolDoDebug == true){
		console.log('The Ascii value for char 0 is: ' + charAscii);
	}

	if (charAscii !== undefined && charAscii > 126){ // Check charAscii is a number
		strOut = sIn.substring(1, intStrLen); //Strip the first character off
	}
	else{
		strOut = sIn;
	}

	if (boolDoDebug == true){
		console.log(" StringIn is: '" + sIn + "' and stringout is: '" + strOut + "'.");
	}
	// If GVars.GblUndefined returns actual undefined, ensure a string is returned
	return strOut === undefined ? sIn : strOut;
}

/**
* ------------------ JComm_StringIsBoolean -----------------------
 * Determine if the string presented is a boolean value.
 * @param strInputText The value to parse
 *
 * @return boolIsBoolean The string is a boolean value?
 * @author pkanaris
 * @author Created 4/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static boolean JComm_StringIsBoolean (String strInputText) {
function StringsAndNumbers_JComm_StringIsBoolean (strInputText) {
	let boolIsBoolean = false;
	if (strInputText != null) {
		const lowerInput = String(strInputText).toLowerCase();
		if (lowerInput === 'true' || lowerInput === 'false') {
			boolIsBoolean = true;
		}
	}
	return boolIsBoolean;
}

/**
* ------------------ JComm_SplitString -----------------------
 * Split the string and return a lst object
 * @param strInputText The value to parse
 * @param strDelimiter The text to split the string on.
 *
 * @return lstValues The list of values created from the split
 * @author pkanaris
 * @author Created 4/28/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static def JComm_SplitString (String strInputText, String strDelimiter) {
function StringsAndNumbers_JComm_SplitString (strInputText, strDelimiter) {
	if (strInputText == null) return []; // Return empty array for null input
	let effectiveDelimiter = String(strDelimiter);
	if (strDelimiter == '|') {
		effectiveDelimiter = '\\|'; //Add the escape chars for regex split
									 // However, String.split in JS treats string arg literally, unless it's a regex object
									 // So, for '|', it's fine. This Groovy logic implies str.split uses regex by default.
									 // JS string.split(/regex/) or string.split("string")
									 // If we want to match Groovy's potential regex behavior for delimiter:
									 // const lstValues = String(strInputText).split(new RegExp(effectiveDelimiter));
									 // For simple string delimiter:
	}
	// def lstValues = strInputText.split(strDelimiter)
	const lstValues = String(strInputText).split(effectiveDelimiter);
	return lstValues;
}

/**
* ------------------ JComm_StringToBoolean -----------------------
 * Return the string as a boolean value.
 * ************************** NOTE CHECK IF BOOLEAN BEFORE RUNNING THIS METHOD !!! ***********************
 * @param strInputText The value to parse
 *
 * @return boolIsBoolean The boolean value of the string
 * @author pkanaris
 * @author Created 4/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static boolean JComm_StringToBoolean (String strInputText) {
function StringsAndNumbers_JComm_StringToBoolean (strInputText) {
	//NOTE: WILL ALWAYS RETURN FALSE IF THE STRING LOWER CASE DOES NOT EQUAL 'true'
	let boolValue = false;
	if (strInputText != null) {
		if (String(strInputText).toLowerCase().trim() == 'true') {
			boolValue = true;
		}
	}
	return boolValue;
}

/**
* ------------------ JComm_StringToInteger -----------------------
 * Return the string as an interger value.
 * ************************** NOTE RETURNS 'NaN' IF NOT A NUMBER !!! *********************** <!-- Groovy code returns null -->
 * @param strInputText The value to parse
 *
 * @return intValue The integer value of the string
 * @author pkanaris
 * @author Created 4/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static JComm_StringToInteger (String strInputText){ // Return type in Groovy is int (implicitly Object then cast)
function StringsAndNumbers_JComm_StringToInteger (strInputText){
	let intValue = null; // Match Groovy's potential null return
	if (strInputText != null) {
		const sInputText = String(strInputText);
		//Take the value left of a period for whole number only
		// String strTempValue = this.JComm_GetLeftTextInString(strInputText, ".")
		const strTempValue = JComm_GetLeftTextInString(sInputText, ".");
		//Check if number and return integer
		// intValue = strTempValue.isInteger() ? (strTempValue as int) : null // Groovy specific
		// JS equivalent:
		if (strTempValue != null && /^-?\d+$/.test(String(strTempValue).trim())) { // Check if it's an integer string
			intValue = parseInt(String(strTempValue).trim(), 10);
			if (isNaN(intValue)) intValue = null; // parseInt can return NaN
		} else {
			intValue = null;
		}
	}
	return intValue;
}

/**
* ------------------ JComm_StringToFloat -----------------------
 * Return the string as an float value.
 * ************************** NOTE RETURNS 'NaN' IF NOT A NUMBER !!! *********************** <!-- Groovy code returns null -->
 * @param strInputText The value to parse
 *
 * @return fltValue The float value of the string
 * @author pkanaris
 * @author Created 4/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static JComm_StringToFloat (String strInputText){ // Return type def implies Object, can be null
function StringsAndNumbers_JComm_StringToFloat (strInputText){
	/* returns NaN when the value sent is not a number */ // Original comment refers to JS parseFloat
	// def fltValue = strInputText?.isFloat() ? strInputText.toFloat() : null // Groovy specific
	let fltValue = null;
	if (strInputText != null) {
		const sInputText = String(strInputText).trim();
		// Check if it's a valid float string
		if (sInputText !== "" && !isNaN(Number(sInputText))) { // Number() handles floats
			fltValue = parseFloat(sInputText);
			 if (isNaN(fltValue)) fltValue = null; // Ensure null if parseFloat results in NaN (e.g. from just "+")
		} else {
			fltValue = null;
		}
	}
	return fltValue;
}

/**
* ------------------ JComm_IntegerToString -----------------------
 * Return the integer as a string value.
 * @param intValue The value convert
 *
 * @return strValue The string value of the integer
 * @author pkanaris
 * @author Created 4/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static JComm_IntegerToString (Number intValue) { // Parameter type Number in Groovy
function StringsAndNumbers_JComm_IntegerToString (intValue) {
	let strValue;
	// strValue = this.JComm_HandleNoData(String.valueOf(intValue))
	// String.valueOf(null) in Java is "null". JComm_HandleNoData would then process "null".
	// In JS, String(null) is "null", String(undefined) is "undefined".
	if (intValue === null || intValue === undefined) {
		strValue = GblNull; // Consistent with JComm_HandleNoData logic for null input
	} else {
		strValue = JComm_HandleNoData(String(intValue));
	}
	return strValue;
}

/**
 * ------------------ JComm_ConcatenateTwoStrings -----------------------
 * Concatenate two string and return as a consolidated value
 * @param strValue1 First value
 * @param strValue2 Second value
 *
 * @return strNewValue The Concatenated value of the two strings
 * @author pkanaris
 * @author Created 4/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static JComm_ConcatenateTwoStrings (String strValue1, String strValue2){
function StringsAndNumbers_JComm_ConcatenateTwoStrings (strValue1, strValue2){
	// String strNewValue = strValue1 + strValue2
	// Handle nulls to avoid "null" or "undefined" in the string
	const sVal1 = (strValue1 == null) ? "" : String(strValue1);
	const sVal2 = (strValue2 == null) ? "" : String(strValue2);
	const strNewValue = sVal1 + sVal2;
	return strNewValue;
}

/**
* ------------------ JComm_StringToArray -----------------------
* Count the number of items and return the elements in and array
* @param strValue The value to parse for items
* @param strValue2 The delimiter for the value. <!-- Param name strValue2 seems like a typo, should be strDelimiter -->
*
* @return mapItemCntAndArray The Count of items and the array of values.
* @return intCount the number of items
* @return strNewValue the value if only 1 item <!-- Not explicitly returned like this in map-->
* @return arryOfValue the values in an array to include the value if only one
* @author pkanaris
* @author Created 01/24/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
// public static Map JComm_StringToArray (String strValue, String strDelimiter){
function StringsAndNumbers_JComm_StringToArray (strValue, strDelimiter){
	// Map<String,String> mapReturnValues = [:] // In JS, this is a plain object
	const mapReturnValues = {};
	// int intItemCount // Defined later as intLenArray
	let arryValues;
	let intLenArray;

	if (strValue == null) { // Handle null input gracefully
		arryValues = [];
		intLenArray = 0;
	} else {
		const sValue = String(strValue);
		const sDelimiter = String(strDelimiter);
		// int intTextPos = strValue.indexOf(strDelimiter)
		const intTextPos = sValue.indexOf(sDelimiter);

		if (intTextPos > 0) { // Groovy's split behavior: if delim is not found, array has 1 element.
							  // If delim is found, it splits. >0 ensures it's not at the start.
							  // If delimiter is at start, Groovy split yields empty string as first element.
			//Create array using the splitter
			let strSplitter = sDelimiter;
			if (sDelimiter == '|') {
				// strSplitter = '\\|' //Add the escape chars.
				// This is needed if str.split() in Groovy uses regex.
				// JS string.split() with a string argument is literal.
				// No change needed for JS string.split('|')
			}
			// arryValues = strValue.split(strSplitter)
			arryValues = sValue.split(strSplitter);
			intLenArray = arryValues.length;
		}
		else { // Delimiter not found or at start. Groovy makes a list of one item: the original string.
			arryValues = [sValue];
			intLenArray = arryValues.length;
		}
	}
	mapReturnValues['intItemCount'] = intLenArray.toString();
	mapReturnValues['ArryOfValues'] = arryValues;
	return mapReturnValues;
}

/**
* ------------------ JComm_ReturnStringIndexFromArray -----------------------
* Return the string index within the array.
* @param arryValues   The array of values to check
* @param strValue	 The value to search for in the array
*
* @return intItemIndex  The index of the value in the array
* @author pkanaris
* @author Created 06/11/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
// public static int JComm_ReturnStringIndexFromArray(List<String> arryValues, String strValue) {
function StringsAndNumbers_JComm_ReturnStringIndexFromArray(arryValues, strValue) {
	let intItemIndex = -1;
	if (!Array.isArray(arryValues)) return -1; // Basic type check

	const intArrySize = arryValues.length;
	// String strTempArryValue // Declared in loop
	for (let loopArryValues = 0; loopArryValues < intArrySize; loopArryValues ++) {
		const strTempArryValue = arryValues[loopArryValues];
		if (strTempArryValue === strValue) { // Use === for strict equality in JS
			intItemIndex = loopArryValues;
			break;
		}
	}
	return intItemIndex;
}

//https://www.geeksforgeeks.org/multidimensional-arrays-in-java/
/**
* ------------------ JComm_StringToArray ----------------------- <!-- Title indicates mistake, should be JComm_StringToMultiDimensionArray -->
* NOTE: ONLY ALLOWS 2 or 3 Dimensional arrays
* Convert a string to a multidimensional array based on the following:
* @param strValue		The value to parse for items
* @param strItemDelimter   The value to parse items, usually a semicolon (;)
* @param strDelimiter	  The delimiter for the items individual values, usually a pipe "|" .
*
* @return mapItemCntAndArray  The array and count of items in each dimension of the array.
* @return strDimCounts	  The number of items in each dimension separated by a "|"
* @return arryMultDimOfValues   The values in an array to include the value if only one
* @author pkanaris
* @author Created 06/08/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
// public static Map JComm_StringToMultiDimensionArray (String strValue, String strItemDelimter, String strDelimiter){
function StringsAndNumbers_JComm_StringToMultiDimensionArray (strValue, strItemDelimter, strDelimiter){
	// This method is a stub in Groovy.
	console.warn("JComm_StringToMultiDimensionArray is a stub and not implemented.");
	const mapResults = {};
	mapResults['strDimCounts'] = "0"; // Placeholder
	mapResults['arryMultDimOfValues'] = []; // Placeholder
	return mapResults;
}

//For JTR-17513 will convert the value based on line feeds to a delimited value
// public static JComm_ProcessValueForMultiLine (String strInValue){
function StringsAndNumbers_JComm_ProcessValueForMultiLine (strInValue){
	// This method is a stub in Groovy.
	console.warn("JComm_ProcessValueForMultiLine is a stub and not implemented.");
	// Likely involves replacing \n or \r\n with a specified delimiter.
	// Example: return String(strInValue).replace(/\r\n|\r|\n/g, GVars.GblDelimiter("Value"));
	return strInValue; // Placeholder
}
  //Check if value is multiline
  //Return assignMultiLineValue
  /**
   * Code from the multiline
   //Determine if we have multiline
	//Check if either is contained in the string
	int intGblLF = StrNums.JComm_TextLocationInString(strTempValue, gblLineFeed)
	int intGBLHTMLLF = StrNums.JComm_TextLocationInString(strTempValue, gblHTMLLF)
	List arrayMultiLineValues
	int intArrayCnt = 0
	if (intGblLF > 0) {
		//Split the input value to begin the processing
		Map mapItemsValues = StrNums.JComm_StringToArray(strTempValue, gblLineFeed)
		arrayMultiLineValues = mapItemsValues.ArryOfValues
		intArrayCnt = arrayMultiLineValues.size()
	}
	else if(intGBLHTMLLF > 0) {
		//Split the input value to begin the processing
		Map mapItemsValues = StrNums.JComm_StringToArray(strTempValue, gblHTMLLF)
		arrayMultiLineValues = mapItemsValues.ArryOfValues
		intArrayCnt = arrayMultiLineValues.size()
	}
	if (intArrayCnt > 0) {
		//Return each item and enter the value then send a return
		for (int loopArry = 0; loopArry < intArrayCnt; loopArry++) {
			weEditBox.sendKeys(arrayMultiLineValues[loopArry])
			if (loopArry < intArrayCnt -1) {
				weEditBox.sendKeys(Keys.ENTER)
			}
		}
	}
	else {
		weEditBox.sendKeys(strTempValue)
	}
*/