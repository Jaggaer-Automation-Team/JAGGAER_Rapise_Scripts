/* 	Jaggaer DateTime Functions used in all test 

*/

SeSGlobalObject("DateTime");

/** ---------------------------- Date funtions ------------------------------*/



/**
 * --------------------------------- Date/Time Value -------------------------------------
 *
 */
/**
 * ------------------ datetimeReturnDynamicFormatedDateTime -----------------------
 * Return the date and time in a specific format based on values set.
 *
 * @param strInValue the value to be processed in a format from the GenerateDyanmicDateTime file
 * i.e. =D[Today  ];=T[12:00 p HH:mm], =D[01/23/2016 + 1 Years];=T[14:45 hh:mm:ss a]
 * JTR-17508
 * @return: Formated date/time
 * @created 11/17/2022
 * @author PGKanaris
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static datetimeReturnDynamicFormatedDateTime(String strInValue) {
function DateTime_datetimeReturnDynamicFormatedDateTime(/*string*/strInValue) {
	const gblDelimiter = GblDelimiter;
	let strOutputValue = "";
	// int intCountDTItems = strInValue.count('[')
	const intCountDTItems = (String(strInValue).match(/\[/g) || []).length;
	if (intCountDTItems > 0) {
		//Process the value as potential dynamic datetime
		//Split into and array of values
		// List arryDateTime, arryValues // Declared later where used
		// int intArryDatTimeCnt, intArrayCnt // Declared later where used

		//Split the in value if needed
		// Map mapArryInValues = [:]
		const mapArryInValues = StrNums.JComm_StringToArray(strInValue, ";"); //Always will be split by a semicolon
		const intArryDatTimeCnt = StrNums.JComm_StringToInteger(mapArryInValues.intItemCount);
		const arryDateTime = mapArryInValues.ArryOfValues;

		//Loop through the values
		// String strTempArryValue, strTempOutValue
		for (let intLoopDTValues = 0; intLoopDTValues < intArryDatTimeCnt; intLoopDTValues++) {
			const strTempArryValue = arryDateTime[intLoopDTValues];
			let strTempOutValue;
			//Determine if it is a dynamic date/time
			if (String(strTempArryValue).includes('=D[')) { // strTempArryValue.indexOf('=D[') >= 0
				strTempOutValue = datetimeReturnDynamicFormatedDate(strTempArryValue); // `this.` removed
			}
			else if(String(strTempArryValue).includes('=T[')) { // strTempArryValue.indexOf('=T[') >= 0
				strTempOutValue = datetimeReturnDynamicFormatedTime(strTempArryValue); // `this.` removed
			}
			else {
				strTempOutValue = strTempArryValue;
			}
			//Check if this is the first value or second
			if (intLoopDTValues == 0) {
				strOutputValue = strTempOutValue;
			}
			else {
				strOutputValue = strOutputValue + " " + strTempOutValue; //Added a space since most values data time values are separated by a space
			}
		}
		return strOutputValue;
	}
	else {
		return strInValue; //Return the original value since there does not appear to be a dynamic value
	}
}
/**
 * ------------------ datetimeReturnDynamicFormatedDate -----------------------
 * Return the dynamic date in a specific format based on values set.
 *
 * @param strDateValue the value to be processed in a format from the GenerateDyanmicDateTime file
 * i.e. =D[Today|MM/dd/yyyy], =D[11/23/2022|MMMM dd, yyyy], or =D[01/23/2016 + 1 Years|MM/dd/yyyy]
 * JTR-17508
 * @return: Formated date
 * @created 11/17/2022
 * @author PGKanaris
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static datetimeReturnDynamicFormatedDate(String strDateValue) {
function DateTime_datetimeReturnDynamicFormatedDate(strDateValue) {
	const gblDelimiter = GVars.GblDelimiter('Value');
	let strDateOut;
	// int intArrayCnt // Not needed directly if using arryValues.length
	// List arryValues // Declared below

	//ADD a split for the value to return the date/time format
	// Map mapArray = [:]
	// String strDateFormat, strTempDateValue // Declared below
	const mapArray = StrNums.JComm_StringToArray(strDateValue, gblDelimiter);
	const intArrayCnt = StrNums.JComm_StringToInteger(mapArray.intItemCount); // Or mapArray.ArryOfValues.length
	const arryValues = mapArray.ArryOfValues;

	let strDateFormat;
	let strTempDateValue;
	//Should we fail if more than 2 values?, Should we check the value is a date as well as check if valid new format?
	if (intArrayCnt == 1) {
		strDateFormat = "MM/dd/yyyy";//Standard format
		strTempDateValue = arryValues[0];
	}
	else { // Assumes intArrayCnt >= 2
		strDateFormat = arryValues[1];
		strTempDateValue = arryValues[0];
	}
	strDateOut = calculateDateValue(strTempDateValue, strDateFormat); // `this.` removed
	return strDateOut;
}
/**
 * ------------------ datetimeReturnDynamicFormatedTime -----------------------
 * Return the dynamic time in a specific format based on values set.
 *
 * @param strTimeValue the value to be processed in a format from the GenerateDyanmicDateTime file
 * JTR-17508
 * @return: Formated date <!-- Typo in original comment, should be Formatted time -->
 * @created 11/17/2022
 * @author PGKanaris
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static datetimeReturnDynamicFormatedTime(String strTimeValue) {
function DateTime_datetimeReturnDynamicFormatedTime(strTimeValue) {
	//11/20/2022 Dynamic time has not been defined. Stubbed in for the time being.
	const gblDelimiter = GVars.GblDelimiter('Value');
	let strTimeOut;
	// int intArrayCnt // Not needed directly
	// List arryValues // Declared below
	//ADD a split for the value to return the date/time format
	// Map mapArray = [:]
	// String strTimeFormat, strTempTimeValue // Declared below
	const mapArray = StrNums.JComm_StringToArray(strTimeValue, gblDelimiter);
	const intArrayCnt = StrNums.JComm_StringToInteger(mapArray.intItemCount);
	const arryValues = mapArray.ArryOfValues;

	let strTimeFormat;
	let strTempTimeValue;
	//Should we fail if more than 2 values?, Should we check the value is a time as well as check if valid new format?
	if (intArrayCnt == 1) {
		strTimeFormat = "hh:mm a";//Standard format
		strTempTimeValue = arryValues[0];
	}
	else { // Assumes intArrayCnt >= 2
		strTimeFormat = arryValues[1];
		strTempTimeValue = arryValues[0];
	}
	strTimeOut = returnFormatedTime(strTempTimeValue, strTimeFormat); // `this.` removed
	return strTimeOut;
}
/**
 * ---------------------------------  Date Methods ---------------------------------------
 */
/**
 * ------------------ timeCreateDateTimeValueForFileName -----------------------
 * Return the current date/time in the format "yyyyMMdd_HHmmss".
 *
 * @return: Formated data/time
 * @created 04/02/2021
 * @author PGKanaris
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static timeCreateDateTimeValueForFileName() {
function DateTime_timeCreateDateTimeValueForFileName() {
	const dateFormat = DateTimeFormatterModule.ofPattern("yyyyMMdd_HHmmss");
	const now = LocalDateTimeModule.now();
	return dateFormat.format(now);
}
/**
 * Methods to be used in the calculation of the date and time data required for test cases
 */

/**
 * ---------------------------------  Date Methods ---------------------------------------
 */
/**
 * -------------------------------------- calculateDateValue ------------------------------
 * Calculates the date value based on the starting date (defined or derived from RelTimeFrame), plus/minus time, offset integer and offset time period
 *
 * +++++++++++++++++++++++++++++ NOTE: RETURNS THE ERROR MESSAGE IN THE OUTPUT IF UNABLE TO PROCESS THE INPUT DATE!!!! +++++++++++++++++++++++++++++++
 * @param strInputDate The values to be processed
 * @param strOutputDateFormat The date format to return
 * @return strOutputDate
 * Created: 08/20/2019
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
// public static calculateDateValue (String strInputDate, String strOutputDateFormat){
function DateTime_calculateDateValue (strInputDate, strOutputDateFormat) {
	//Define the global variables used in the static method
	const gblNull = GVars.GblNull('Value');
	const gblInvalidDate = GVars.GblInvalidDate('Value');
	//Define variables
	let strRelTimeFrame, strPlusMinusTime, strOffSetTimeInt, strOffSetTimeperiod, strTmpParse, strTmpDate = "", strTmpValue, strInFormat, strOutputDate;
	let intOffSetTime, intInStr; // `Integer` becomes `number` (or inferred)
	strRelTimeFrame = gblNull;
	strTmpParse = gblNull; // Keep initial value
	// strTmpValue declared above
	// strInFormat delcared above

	// boolean boolIsDateValue, boolProcessOffsetTime // Used later
	const dateCurrent = new Date(); // JS Date
	let dateTemp = null; // JS Date | null
	// SimpleDateFormat dateFormat = new SimpleDateFormat('MMddyyyy') // Used for default formatting
	const defaultInternalSDF = new SimpleDateFormatImplMock('MMddyyyy');
	strOutputDate = gblNull;

	//check if the value is null
	if (StrNums.JComm_HandleNoData(strInputDate) == gblNull) {
		strOutputDate = 'FAILED!!! The date value of: ' + strInputDate + ' is ' + gblNull + '.';
	} else {
		//Check if the value contains the specified expression start
		intInStr = StrNums.JComm_TextLocationInString(strInputDate, '='); //Calculated dates start with '=' otherwise normal date which will can check formating and return
		if (intInStr >= 0) {
			//Parse the date into separate items
			strTmpParse = '=D[';
			strTmpValue = StrNums.JComm_GetRightTextInString(strInputDate, strTmpParse);
			//Check if a space is present
			intInStr = StrNums.JComm_TextLocationInString(strTmpValue, ' ');
			if (intInStr > 0 && intInStr < String(strTmpValue).indexOf(']')) { // Ensure space is meaningful part of =D[...]
				strTmpParse = ' ';
			} else {
				strTmpParse = ']';
			}
			strRelTimeFrame = StrNums.JComm_GetLeftTextInString(strTmpValue, strTmpParse);
			switch (strRelTimeFrame) {
				case 'Today':
					strTmpDate = defaultInternalSDF.format(dateCurrent);
					break;
				case 'Yesterday':
					// use(TimeCategory) { strTmpDate = dateFormat.format(dateCurrent - 1.days) }
					strTmpDate = defaultInternalSDF.format(TimeCategoryOps.addOrSubtract(dateCurrent, -1, 'days'));
					break;
				case 'Tomorrow':
					// use(TimeCategory) { strTmpDate = dateFormat.format(dateCurrent + 1.days) }
					strTmpDate = defaultInternalSDF.format(TimeCategoryOps.addOrSubtract(dateCurrent, 1, 'days'));
					break;
				case 'WeekAgo':
					strTmpDate = defaultInternalSDF.format(TimeCategoryOps.addOrSubtract(dateCurrent, -1, 'weeks'));
					break;
				case 'WeekFromToday':
					strTmpDate = defaultInternalSDF.format(TimeCategoryOps.addOrSubtract(dateCurrent, 1, 'weeks'));
					break;
				case 'MonthAgo':
					strTmpDate = defaultInternalSDF.format(TimeCategoryOps.addOrSubtract(dateCurrent, -1, 'months'));
					break;
				case 'MonthFromToday':
					strTmpDate = defaultInternalSDF.format(TimeCategoryOps.addOrSubtract(dateCurrent, 1, 'months'));
					break;
				case 'YearAgo':
					strTmpDate = defaultInternalSDF.format(TimeCategoryOps.addOrSubtract(dateCurrent, -1, 'years'));
					break;
				case 'YearFromToday':
					strTmpDate = defaultInternalSDF.format(TimeCategoryOps.addOrSubtract(dateCurrent, 1, 'years'));
					break;
				default:
					strTmpDate = strRelTimeFrame; //We assume the first portion is a date if no match has been found. Will test for date next.
					break;
			}
		} else {
			strTmpDate = strInputDate;
		}
		//Check if the value is indeed a date
		const boolIsDateValue = isValidDate(strTmpDate); // `this.` removed
		let currentOutputDateIsNull = (strOutputDate === gblNull); // Track if strOutputDate gets an error

		if (boolIsDateValue == true ) {
			if (strTmpValue !== null && StrNums.JComm_HandleNoData(strTmpValue) != gblNull) { // strTmpValue is from =D[...]
				// const boolProcessOffsetTime = true; // Not strictly needed in JS logic flow here
				//Next we must determine if we are adding or subtracting values from the base date
				strTmpParse = ' '; // Default delimiter
				//remove the strRelTimeFrame
				let remainingTmpValue = StrNums.JComm_GetRightTextInString(strTmpValue, strRelTimeFrame || ""); // Ensure strRelTimeFrame is not null
				remainingTmpValue = remainingTmpValue ? String(remainingTmpValue).trimStart() : "";

				//Return the +- value if present
				strPlusMinusTime = StrNums.JComm_GetLeftTextInString(remainingTmpValue, strTmpParse);
				if (strPlusMinusTime == '+' || strPlusMinusTime == '-') {
					//Get the offsetTimeInt
					//remove the strPlusMinusTime
					remainingTmpValue = StrNums.JComm_GetRightTextInString(remainingTmpValue, strPlusMinusTime).trimStart();
					//return the integer for offsetTime
					strOffSetTimeInt = StrNums.JComm_GetLeftTextInString(remainingTmpValue, strTmpParse);
					//Convert to a number if possible
					if (StringUtilsModule.isNumeric(strOffSetTimeInt) == true) {
						intOffSetTime = StrNums.JComm_StringToInteger(strOffSetTimeInt);
						//Get the offsetTimePeriod
						//remove the strOffSetTimeInt
						remainingTmpValue = StrNums.JComm_GetRightTextInString(remainingTmpValue, strOffSetTimeInt).trimStart();
						strTmpParse = ']';
						strOffSetTimeperiod = StrNums.JComm_GetLeftTextInString(remainingTmpValue, strTmpParse);
						//Check for the OffSet Time Period and calculate the date
						//Remove / or - from the string
						const cleanedStrTmpDate = String(strTmpDate).replace(/[\/\-]/g, ''); // Ensure strTmpDate is a string
						strInFormat = returnDateformat(cleanedStrTmpDate); // `this.` removed
						if (strInFormat == gblNull) {
							strOutputDate = (strOutputDate === gblNull ? "" : strOutputDate) + 'FAILED!!! The date: ' + cleanedStrTmpDate + ' DOES NOT MATCH ASSIGNED DATE FORMATS!!!';
							currentOutputDateIsNull = false;
						} else {
							const dateTempSDF = new SimpleDateFormatImplMock(strInFormat);
							dateTempSDF.setLenient(false);
							try {
								dateTemp = dateTempSDF.parse(cleanedStrTmpDate.trim());
							} catch (error) {
								strOutputDate = (strOutputDate === gblNull ? "" : strOutputDate) + 'FAILED!!! UNABLE TO PARSE THE DATE ' + cleanedStrTmpDate + ' INTO THE DATE OBJECT!!! ' + error.message;
								currentOutputDateIsNull = false;
							}
							if (currentOutputDateIsNull && dateTemp){ // Only if no error so far
								let unit = null; // For TimeCategoryOps
								switch (strOffSetTimeperiod){
									case 'Days': unit = 'days'; break;
									case 'Weeks': unit = 'weeks'; break;
									case 'Months': unit = 'months'; break;
									case 'Years': unit = 'years'; break;
									default:
										strOutputDate = (strOutputDate === gblNull ? "" : strOutputDate) + 'FAILED!!! The OffSet Period Specified of: ' + strOffSetTimeperiod + ' DOES NOT MATCH expected time periods!!!';
										currentOutputDateIsNull = false;
										break;
								}
								if (unit && dateTemp) {
									const offsetAmount = (strPlusMinusTime === '+') ? intOffSetTime : -intOffSetTime;
									const calculatedDate = TimeCategoryOps.addOrSubtract(dateTemp, offsetAmount, unit);
									strTmpDate = defaultInternalSDF.format(calculatedDate); // Update strTmpDate
								}
							}
						}
					} else {
						strOutputDate = (strOutputDate === gblNull ? "" : strOutputDate) + 'FAILED!!! The string: "' + strOffSetTimeInt + '" IS NOT NUMERIC!!! We cannont continue processing!!!';
						currentOutputDateIsNull = false;
					}
				}
			}
			//Process the date based on the assigned format
			//Check if null
			if (StrNums.JComm_HandleNoData(strTmpDate) != gblNull) {
				if (currentOutputDateIsNull) { // If no error so far
					//Recheck the date to be sure current date is valid
					//Check if the value is indeed a date
					const boolIsFinalDateValue = isValidDate(strTmpDate); // `this.` removed
					if (boolIsFinalDateValue == true) {
						strOutputDate = returnFormatedDate(strTmpDate, strOutputDateFormat); // `this.` removed
					} else if (strOutputDate === gblNull) { // if returnFormatedDate didn't set an error but it's not valid
						strOutputDate = gblInvalidDate;
					}
				}
			}
		} else if (currentOutputDateIsNull) { // Original strTmpDate was not valid, and strOutputDate still gblNull
			strOutputDate = gblInvalidDate;
		}
	}
	return strOutputDate === gblNull ? gblInvalidDate : strOutputDate; // Ensure a non-null string is returned
}

//Return Standard MM/dd/yyyy format for date
// public static String returnStdDateFormat (String strInputDate) {
function DateTime_returnStdDateFormat (strInputDate) {
	let strNewDate;
	strNewDate = returnFormatedDate(strInputDate, "MM/dd/yyyy"); // `this.` removed
	return strNewDate;
}

//Return UTC with Offset format for date
// public static String returnDateTimeWithUTCOffset (LocalDateTime now) {
function DateTime_returnDateTimeWithUTCOffset (now) { // `now` expected to be LocalDateTime-like mock
	//LocalDateTime now = LocalDateTime.now() // Parameter `now` is used
	const strDatePattern = "yyyy-MM-dd'T'HH:mm:ss";
	const dateFormat = DateTimeFormatterModule.ofPattern(strDatePattern);
	const strDateTime = dateFormat.format(now); // Format the passed 'now'
	//Calculate the UTC offset
	//https://www.baeldung.com/java-zone-offset
	//Return the local zoneID
	const tz = TimeZoneModule.getDefault();
	let id = tz.toZoneId();
	//TODO set to only use ET due to Jira Xray issue. Verifying if EC2 will post if ET. PGK 10/23/2023
	//id = ZoneIdModule.of("America/New_York"); // Example if hardcoding
	const zonedDateTime = now.atZone(id); // `now` needs to have `atZone`
	const zoneOffset = zonedDateTime.getOffset();
	const strUTCDateTime = strDateTime + zoneOffset.toString();
	//ZoneOffset zoneOffset =
	return strUTCDateTime;
}

//Return UTC with Offset format for date with milliseconds
// public static String returnDateTimeMilliSecsWithUTCOffset (LocalDateTime now, boolean boolRemoveColonInUTC) {
function DateTime_returnDateTimeMilliSecsWithUTCOffset (now, boolRemoveColonInUTC) { // `now` expected to be LocalDateTime-like mock
	//LocalDateTime now = LocalDateTime.now() // Parameter `now` is used
	const strDatePattern = "yyyy-MM-dd'T'HH:mm:ss.SSS";
	const dateFormat = DateTimeFormatterModule.ofPattern(strDatePattern);
	let strDateTime = dateFormat.format(now);
	let strValueOffset;
	//Calculate the UTC offset
	//https://www.baeldung.com/java-zone-offset
	//Return the local zoneID
	const tz = TimeZoneModule.getDefault();
	let id = tz.toZoneId();
	//TODO set to only use ET due to Jira Xray issue. Verifying if EC2 will post if ET. PGK 10/23/2023
	//id = ZoneIdModule.of("America/New_York"); // Example if hardcoding
	const zonedDateTime = now.atZone(id);
	const zoneOffset = zonedDateTime.getOffset();

	strValueOffset = zoneOffset.toString();
	if (boolRemoveColonInUTC == true) {
		//TODO Create new string function DateTime_to replace char in string (original comment)
		// Groovy: strValueOffset = zoneOffset.toString().replace(':','')
		if (strValueOffset !== "Z") { // "Z" for UTC has no colon
			 strValueOffset = strValueOffset.replace(':', '');
		}
	}
	// else { // strValueOffset already has colon or is "Z"
	//  strValueOffset = zoneOffset.toString()
	// }

	if (id.toString() == "Etc/UTC" || strValueOffset === "Z") { // If system is UTC or offset is Z
		if (boolRemoveColonInUTC == true) {
			strValueOffset ="+0000";
		}
		else {
			strValueOffset = "+00:00";
		}
	}
	let strUTCDateTime = strDateTime + strValueOffset;
	//Added 10/23/2023 PGK to remove the "Z" due to UTC
	// This logic seems to want to remove a literal "Z" if it's at the end of strUTCDateTime,
	// which might happen if zoneOffset.toString() was "Z" and boolRemoveColonInUTC was false.
	strUTCDateTime = StrNums.JComm_GetLeftTextInString(strUTCDateTime, "Z"); // This will remove Z and anything after it.
																		  // If Z is the last char, it becomes empty.
																		  // If Z is not present, original string is returned.
																		  // If strUTCDateTime was "datetimeZ", it becomes "datetime".
																		  // If strValueOffset became "Z", then strUTCDateTime = "datetimeZ" before this.
	//ZoneOffset zoneOffset = // Comment in original
	return strUTCDateTime;
}

//Returns the Month Number, Day Number and Year Number
// public static returnMonthDayYearMapper (String strInputDate){
function DateTime_returnMonthDayYearMapper (strInputDate) {
	// Stub in Groovy
	console.warn("returnMonthDayYearMapper is a stub.");
	return {}; // Or some default map/object
}

// public static returnFullMonthNameFromDate (String strInputDate) {
function DateTime_returnFullMonthNameFromDate (strInputDate) {
	let strNewDate;
	strNewDate = returnFormatedDate(strInputDate, "MMMM"); // `this.` removed
	return strNewDate;
}

// public static returnAbbrevMonthNameFromDate (String strInputDate) {
function DateTime_returnAbbrevMonthNameFromDate (strInputDate) {
	let strNewDate;
	strNewDate = returnFormatedDate(strInputDate, "MMM"); // `this.` removed
	return strNewDate;
}

// public static returnMonthNumberFromDate (String strInputDate) {
function DateTime_returnMonthNumberFromDate (strInputDate) {
	let strNewDate;
	strNewDate = returnFormatedDate(strInputDate, "MM"); // `this.` removed
	return strNewDate;
}

// public static returnDayValueFromDate (String strInputDate) {
function DateTime_returnDayValueFromDate (strInputDate) {
	let strNewDate;
	strNewDate = returnFormatedDate(strInputDate, "dd"); // `this.` removed
	return strNewDate;
}

// public static returnFullDayNameFromDate (String strInputDate) {
function DateTime_returnFullDayNameFromDate (strInputDate) {
	let strNewDate;
	strNewDate = returnFormatedDate(strInputDate, "EEEE"); // `this.` removed
	return strNewDate;
}

// public static returnAbbrevDayNameFromDate (String strInputDate) {
function DateTime_returnAbbrevDayNameFromDate (strInputDate) {
	let strNewDate;
	strNewDate = returnFormatedDate(strInputDate, "EEE"); // `this.` removed
	return strNewDate;
}

// public static returnYearValueFromDate (String strInputDate) {
function DateTime_returnYearValueFromDate (strInputDate) {
	let strNewDate;
	strNewDate = returnFormatedDate(strInputDate, "yyyy"); // `this.` removed
	return strNewDate;
}

// def returnMonthNumToFullName (Number intMonthNumber){
function DateTime_returnMonthNumToFullName (intMonthNumber) {
	// Stub in Groovy
	console.warn("returnMonthNumToFullName is a stub.");
	// Example: const months = ["January", ...]; return months[intMonthNumber-1];
	return `FullMonth_for_${intMonthNumber}`;
}

// def returnMonthNumToAbbreviation (Number intMonthNumber){
function DateTime_returnMonthNumToAbbreviation (intMonthNumber) {
	// Stub in Groovy
	console.warn("returnMonthNumToAbbreviation is a stub.");
	return `AbbrMonth_for_${intMonthNumber}`;
}

// def returnWkDayNumToDayName (Number intWkDayNumber){
function DateTime_returnWkDayNumToDayName (intWkDayNumber) {
	// Stub in Groovy
	console.warn("returnWkDayNumToDayName is a stub.");
	return `DayName_for_WkDay_${intWkDayNumber}`;
}

// def returnDateDayNames (String strInputDate){
function DateTime_returnDateDayNames (strInputDate) {
	// Stub in Groovy
	console.warn("returnDateDayNames is a stub.");
	return `DayNames_for_Date_${strInputDate}`;
}

// public static boolean isValidDate(String strInDate) {
function DateTime_isValidDate(strInDate) {
	const gblNull = GVars.GblNull('Value');
	// Integer intLenStr // Not strictly needed thanks to JS string .length
	// boolean boolHasDash, boolHasDivide // Inferred from .includes()
	if (strInDate === null || strInDate === undefined) return false; // Handle null/undefined input
	const sInDate = String(strInDate); // Ensure it's a string

	const intLenStr = sInDate.length;
	const boolHasDash = sInDate.includes('-');
	const boolHasDivide = sInDate.includes('/');
	let strTestFormat = gblNull;
	// String strFormatedDate // Not used for boolean return in Groovy logic

	//Formats are based on the input date alway being in the monthdayyear order and should be 8 or 10 characters
	if (intLenStr == 8 && boolHasDash == false && boolHasDivide == false) {
		strTestFormat = "MMddyyyy";
	} else if (intLenStr == 10 && boolHasDash == true) {
		strTestFormat = "MM-dd-yyyy";
	} else if (intLenStr == 10 && boolHasDivide == true) {
		strTestFormat = "MM/dd/yyyy";
	}

	if (strTestFormat != gblNull && strTestFormat !== null) { // Check strTestFormat itself is not null as returned by GVars
		const dateFormat = new SimpleDateFormatImplMock(strTestFormat);
		dateFormat.setLenient(false);
		try {
			//def date = Date.parse("MM/dd/yyyy",strInDate) // Groovy specific if format was static
			dateFormat.parse(sInDate.trim()); // strFormatedDate (result of parse) not used for return
			return true;
		} catch (pe) { // ParseException
			return false;
		}
	} else {
		return false;
	}
}

// public static returnDateformat (String strInDate){
function DateTime_returnDateformat (strInDate) {
	const gblNull = GVars.GblNull('Value');
	// Integer intLenStr // Not needed
	// boolean boolHasDash, boolHasDivide // Inferred
	const sInDate = String(strInDate); // Ensure string

	const intLenStr = sInDate.length;
	const boolHasDash = sInDate.includes('-');
	const boolHasDivide = sInDate.includes('/');
	let strTestFormat = gblNull;
	//Formats are based on the input date alway being in the monthdayyear order and should be 8 or 10 characters
	if (intLenStr == 8 && boolHasDash == false && boolHasDivide == false) {
		strTestFormat = "MMddyyyy";
	} else if (intLenStr == 10 && boolHasDash == true) {
		strTestFormat = "MM-dd-yyyy";
	} else if (intLenStr == 10 && boolHasDivide == true) {
		strTestFormat = "MM/dd/yyyy";
	}
	return strTestFormat;
}

// public static returnFormatedDate (String strInDate, String strOutFormat) {
function DateTime_returnFormatedDate (strInDate, strOutFormat) {
	const gblNull = GVars.GblNull('Value');
	//TODO should we add a map for error handling? (Original comment)
	//TODO should we have a switch to validate the format? (Original comment)
	//Reformat the date to the specified format
	let strOutFormatedText = gblNull;
	// Date dateTemp // Not directly used in this refactoring path
	let strTempOutFormat = String(strOutFormat); // Ensure string

	// const boolIsMMM = StrNums.JComm_TextLocationInString(strOutFormat, 'mmm/') != -1; // Groovy version
	// Replicating the logic of replacing 'mm' with 'MM' if 'mm' is likely month
	const isMMM = strTempOutFormat.includes('mmm/'); // 'mmm/' is non-standard for Java SDF/DTF (MMM)
	if (!isMMM) {
		// If format doesn't contain time parts (H,h,s,a) and has 'mm', assume 'mm' was intended as month
		if (!/[Hhmsa]/.test(strTempOutFormat) && strTempOutFormat.includes('mm')) {
			strTempOutFormat = strTempOutFormat.replace(/mm/g,"MM"); // replace all 'mm'
		}
	}
	// const dateOutFormat = DateTimeFormatterModule.ofPattern(strTempOutFormat); // Java.time version

	//Remove / or - from the string
	const cleanedInDate = String(strInDate).replace(/[\/\-]/g, '');
	const strInDateFormatPattern = returnDateformat(cleanedInDate); // `this.` removed

	if (strInDateFormatPattern == gblNull) {
		strOutFormatedText = (strOutFormatedText === gblNull ? "" : strOutFormatedText) + 'FAILED!!! The date: ' + cleanedInDate + ' DOES NOT MATCH ASSIGNED DATE FORMATS!!!';
	} else {
		// Using java.time.LocalDate approach (mocked)
		// DateTimeFormatter dtFormat = DateTimeFormatter.ofPattern(strDateFormat)
		// try {
		//  LocalDate newDate = LocalDate.parse(cleanedInDate, dtFormat)
		//  strOutFormatedText = newDate.format(dateOutFormat)
		// }
		// catch (IllegalArgumentException e) { strOutFormatedText = ... }
		// catch (DateTimeParseException e) { strOutFormatedText = ... }

		// Using SimpleDateFormatImplMock as it's more aligned with original Groovy if it used SDF implicitly
		try {
			const sdfIn = new SimpleDateFormatImplMock(strInDateFormatPattern);
			const parsedDate = sdfIn.parse(cleanedInDate); // Parse to JS Date

			const sdfOut = new SimpleDateFormatImplMock(strTempOutFormat);
			strOutFormatedText = sdfOut.format(parsedDate); // Format JS Date
		} catch (error) {
			let errorType = "Error";
			if (error.name === "DateTimeParseException" || error.message.includes("Unparseable date")) errorType = "DateTimeParseException";
			else if (error.name === "IllegalArgumentException") errorType = "IllegalArgumentException"; // Less likely from SDF mock

			strOutFormatedText = (strOutFormatedText === gblNull ? "" : strOutFormatedText) + `FAILED!!! ${errorType}: ` + error.message;
		}
	}
	return strOutFormatedText;
}

// public static verifyAssignDateFormatValid(String strInDateFormat) {
function DateTime_verifyAssignedDateFormatValid(strInDateFormat) {
	let strActualResults;
	let boolMethodPassed = false;
	// Map<String,String> mapMethResults = new HashMap<String,String>() // JS object
	const mapMethResults = {};
	// Number intMethStatus = -1 // Becomes number
	let intMethStatus = -1;

	switch (strInDateFormat){
		case 'MM/dd/yyyy':
		case 'M/d/yyyy':
		case 'MMddyyyy':
		case 'Mdyyyy':
		case 'MM-dd-yyyy':
		case 'MMM/dd/yyyy':
		case 'MMMM dd, yyyy': // Groovy allows duplicate case, effectively first one wins
		// case 'MMMM dd, yyyy': // Duplicate in Groovy, already covered
			boolMethodPassed = true;
			break;
	}

	if (boolMethodPassed == true) {
		strActualResults = 'The date format assigned of: ' + strInDateFormat + ' is valid';
	} else {
		strActualResults = 'FAILED!!! The date format assigned of: ' + strInDateFormat + ' is NOT VALID';
	}
	//Set the mapMethResults
	if (boolMethodPassed == true) {
		intMethStatus = 0;
	} else {
		intMethStatus = 1;
	}
	mapMethResults['methStatus'] = TestExecReporting.testExecTestStepStatus(intMethStatus);
	mapMethResults['methDetails'] = strActualResults;
	//return mapMethResults
	return mapMethResults;
}

/**
 * --------------------------------- Google Calendar Methods ------------------------------------
 */
/**
 * Returns the number of months difference between the two dates
 * strDate1 should be the start date and strDate2 should be the end date.
 * @param strDate1 start date
 * @param strDate2 end date
 * @return int totalMonths
 * Created: 01/29/2020
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
// public static returnFormatedDateFromMMddYYYYToYYYYMMdd (String strInDate) {
function DateTime_returnFormatedDateFromMMddYYYYToYYYYMMdd (strInDate) {
	const gblNull = GVars.GblNull('Value');
	//TODO should we add a map for error handling? (Original comment)
	//TODO should we have a switch to validate the format? (Original comment)
	//Reformat the date to the specified format
	// String strTempDate // Not used for the final return path of formatted string
	let strOutFormatedText = gblNull;
	let dateTemp;
	const strInDateFormat = 'MMddyyyy'; // Assumes input is always without separators like Groovy
	const strOutDateFormat = 'yyyyMMdd';
	let boolPassed = true;

	//Check if the date is valid against the assigned format
	const sdfIn = new SimpleDateFormatImplMock(strInDateFormat);
	sdfIn.setLenient(false);

	try {
		//def date = Date.parse("MM/dd/yyyy",strInDate) // Groovy specific parsing
		// strTempDate = dateFormat.parse(strInDate.trim()) // In Groovy, if dateFormat is SDF, parse returns Date.
														 // The assignment to strTempDate is unusual if it expected a String.
														 // More likely, it was: dateTemp = dateFormat.parse(strInDate.trim())
		dateTemp = sdfIn.parse(String(strInDate).trim());
	} catch (parseException) { // ParseException
		boolPassed = false;
		strOutFormatedText = (strOutFormatedText === gblNull ? "" : strOutFormatedText) + 'FAILED!!! UNABLE TO PARSE THE DATE ' + strInDate + ' INTO THE DATE OBJECT Using format: ' + strInDateFormat + '!!! ' + parseException.message;
	}

	if (boolPassed == true && dateTemp){ // Redundant check for dateTemp if boolPassed is true
		// SimpleDateFormat dateTempFormat = new SimpleDateFormat(strInDateFormat) // Not needed again, already used sdfIn
		// dateTempFormat.setLenient(false)
		// try {
		//  dateTemp = dateTempFormat.parse(strInDate.trim()) // Already parsed into dateTemp
		// } catch (ParseException pe) { ... } // This inner try-catch seems redundant if outer parse succeeded.

		// if (strOutFormatedText == gblNull){ // Only format if no parse error on dateTemp
		const sdfOut = new SimpleDateFormatImplMock(strOutDateFormat);
		try {
			strOutFormatedText = sdfOut.format(dateTemp); // Format the parsed Date object
		} catch (formatEx) {
			boolPassed = false; // Mark as failed if formatting fails
			strOutFormatedText = (strOutFormatedText === gblNull ? "" : strOutFormatedText) + 'FAILED!!! UNABLE TO FORMAT THE PARSED DATE into ' + strOutDateFormat + '!!! ' + formatEx.message;
		}
		// }
	}
	//Comment out when not in debug
	/**
	 if (boolPassed == true) {
	 KeywordUtil.markPassed('The converted values in format:' + strOutDateFormat + ' is: ' + strOutFormatedText)
	 }
	 */
	if (boolPassed == false && strOutFormatedText !== gblNull) { // Only print if an error message was set
		console.log(strOutFormatedText);
	}
	return strOutFormatedText;
}

// public static returnFormatedDateFromYYYYMMddToMMddYYYY (String strInDate) {
function DateTime_returnFormatedDateFromYYYYMMddToMMddYYYY (strInDate) {
	const gblNull = GVars.GblNull('Value');
	//TODO should we add a map for error handling? (Original comment)
	//TODO should we have a switch to validate the format? (Original comment)
	//Reformat the date to the specified format
	// String strTempDate // Not used
	let strOutFormatedText = gblNull;
	let dateTemp;
	const strInDateFormat = 'yyyyMMdd';
	const strOutDateFormat = 'MMddyyyy';
	let boolPassed = true;

	//Check if the date is valid against the assigned format
	const sdfIn = new SimpleDateFormatImplMock(strInDateFormat);
	sdfIn.setLenient(false);
	try {
		dateTemp = sdfIn.parse(String(strInDate).trim());
	} catch (parseException) { // ParseException
		boolPassed = false;
		strOutFormatedText = (strOutFormatedText === gblNull ? "" : strOutFormatedText) + 'FAILED!!! UNABLE TO PARSE THE DATE ' + strInDate + ' INTO THE DATE OBJECT Using format: ' + strInDateFormat + '!!! ' + parseException.message;
	}

	if (boolPassed == true && dateTemp){
		// SimpleDateFormat dateTempFormat = new SimpleDateFormat(strInDateFormat) // Redundant
		// ...
		// if (strOutFormatedText == gblNull){
		const sdfOut = new SimpleDateFormatImplMock(strOutDateFormat);
		try {
			strOutFormatedText = sdfOut.format(dateTemp);
		} catch (formatEx) {
			boolPassed = false;
			strOutFormatedText = (strOutFormatedText === gblNull ? "" : strOutFormatedText) + 'FAILED!!! UNABLE TO FORMAT THE PARSED DATE into ' + strOutDateFormat + '!!! ' + formatEx.message;
		}
		// }
	}
	//Comment out when not in debug
	/**
	 if (boolPassed == true) {
	 KeywordUtil.markPassed('The converted values in format:' + strOutDateFormat + ' is: ' + strOutFormatedText)
	 }
	 */
	if (boolPassed == false && strOutFormatedText !== gblNull) {
		console.log(strOutFormatedText);
	}
	return strOutFormatedText;
}
/**
 * ---------------------------------  Time Methods ---------------------------------------
 */
/**
 * ------------------ WaitSecs -----------------------
 * Wait the assigned number of seconds
 * @param numSec The number of seconds to wait
 *
 * @return: NONE
 * @created 04/02/2021
 * @author PGKanaris
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static void WaitSecs (int numSecs) {
function DateTime_WaitSecs (numSecs) { // void return type is implicit
	const intNumMillSecs = numSecs * 1000;
	const currentTimeInit = new Date().getTime(); // Renamed from currentTime
	const endTime = currentTimeInit + intNumMillSecs;
	let newTime = new Date().getTime();
	while (endTime >= newTime) {
		newTime = new Date().getTime();
	}
}
/*
	* DateTime_WaitMilliSecs
	* Wait the assigned number of milliseconds
	* @ param numMillisec The number of milliseconds to wait
	*
	* return: NONE
	* Created 10/15/2020
	* Author PGKanaris
	* Last Edited:
	* Last Edited By:
	* Edit Comments: (Include email, date and details)
*/
function DateTime_WaitMilliSecs (/*number*/ numMilliSecs) {
//TODO should we check if this is a number and some portion of milliseconds?
	var currentTime = new Date().getTime();
	var newTime = new Date().getTime();
	endTime = currentTime + numMilliSecs
	//Tester.Message('Current Time: ' + currentTime + ' wait ms: ' + numMillisec + ' end time including wait: ' + endTime)
	while (endTime >= newTime) {
		newTime = new Date().getTime();
		//Tester.Message('New Time: ' + newTime )
   	}
	//Global.DoSleep(numMillisec); //In Milliseconds
	//Tester.Message('New Time in millisecs: ' + newTime )
}
/**
  * ------------------ myTimeDifferanceKeyword -----------------------
 * Verifies the code works to properly calculate the time difference.
 *
 * @return: strOutputMessage A message containing the time differance calculated.
 * @created 04/02/2021
 * @author PGKanaris
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static myTimeDifferanceKeyword(){
function DateTime_myTimeDifferanceKeyword(){
	const myStartTime = Date.now(); // System.currentTimeMillis()
	WaitSecs(1); // `this.` removed
	const myEndTime = Date.now();
	const myTimeDifferance = timeGetStrDiffMillSecsInMinSeconds(myStartTime, myEndTime); // `this.` removed
	const mySecondDifferance = timeGetStrDiffMillSecsInToSeconds(myStartTime, myEndTime); // `this.` removed
	const strOutputMessage = 'Date Time Diff in min-secs: ' + myTimeDifferance + GVars.GblLineFeed('Value') + 'Differance in seconds: ' + mySecondDifferance;
	console.log(strOutputMessage);
	return strOutputMessage;
}
/**
  * ------------------ timeGetStrDiffMillSecsInToSeconds -----------------------
 * Calculate the number of seconds difference from milliseconds start and end time values.
 *
 * @return: strOutputMessage A message containing the time difference calculated. <!-- Comment seems to be for myTimeDifferanceKeyword -->
 * @created 04/02/2021
 * @author PGKanaris
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static timeGetStrDiffMillSecsInToSeconds (long startTimeMillSecs, long endTimeMillSecs){
function DateTime_timeGetStrDiffMillSecsInToSeconds (startTimeMillSecs, endTimeMillSecs) {
	// int secs = (int)(endTimeMillSecs-startTimeMillSecs) /1000
	const secs = Math.floor((endTimeMillSecs - startTimeMillSecs) / 1000); // JS equivalent of int division (truncate)
	return secs;
}
/**
 * ------------------ timeGetStrDiffMillSecsInMinSeconds -----------------------
 * Calculate the number of seconds difference from milliseconds start and end time values.
 *
 * @return: String value mins + ' min(s), ' + secs + ' sec(s)'
 * @created 04/02/2021
 * @author PGKanaris
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
*/
// public static String timeGetStrDiffMillSecsInMinSeconds (long startTimeMillSecs, long endTimeMillSecs){
function DateTime_timeGetStrDiffMillSecsInMinSeconds (startTimeMillSecs, endTimeMillSecs) {
	// long duration; // Not strictly needed, can be inline
	// double minutes; // Not strictly needed
	// double seconds; // Not strictly needed
	const duration = endTimeMillSecs - startTimeMillSecs;
	// minutes = duration / (60 * 1000)
	// int mins = minutes
	const totalSeconds = Math.floor(duration / 1000);
	const mins = Math.floor(totalSeconds / 60);
	// seconds = (duration / 1000) - (mins * 60)
	const secs = totalSeconds % 60;
	// int secs = seconds
	return mins + ' min(s), ' + secs + ' sec(s)';
}
// public static boolean isValidTime(String strInTime) {
function DateTime_isValidTime_Original(strInTime) { // Renamed to avoid clash with the one defined earlier if running in same scope
	const gblNull = GVars.GblNull('Value'); // Not used in Groovy
	let boolValidTime = false;
	// String strFormatedTime // Not used for return
	// Date time // Variable to hold successfully parsed date
	let parsedTime; // To hold the result of a successful parse

	// Groovy's Date.parse(format, string) is powerful. JS needs SDF mock.
	try {
		parsedTime = groovyDateParse("hh:mm a", String(strInTime)); // Using mocked helper
		boolValidTime = true;
	} catch (pe) {
		//fall through
	}
	if (boolValidTime == false){
		try {
			parsedTime = groovyDateParse("hh:mm:ss a", String(strInTime));
			boolValidTime = true;
		} catch (pe) {
			//fall through
		}
	}
	if (boolValidTime == false){
		try {
			parsedTime = groovyDateParse("HH:mm:ss", String(strInTime));
			boolValidTime = true;
		} catch (pe) {
			//fall through
		}
	}
	if (boolValidTime == false){
		try {
			parsedTime = groovyDateParse("HH:mm", String(strInTime));
			boolValidTime = true;
		} catch (pe) {
			//fall through
		}
	}
	if (boolValidTime == true && parsedTime) { // parsedTime will be set if any try succeeded
		const stdTimeFormat = new SimpleDateFormatImplMock("HH:mm:ss");
		stdTimeFormat.setLenient(false);
		// strFormatedTime = stdTimeFormat.format(parsedTime); // Result not used by Groovy for boolean return
	}
	return boolValidTime;
}

// public static returnTimeFormat(String strInTime) {
function DateTime_returnTimeFormat_Original(strInTime) { // Renamed to avoid clash
	const gblNull = GVars.GblNull('Value');
	let strTimeFormat = gblNull;
	// String strFormatedTime; // Not used for return
	let parsedTime; // To hold result of a successful parse/check

	//Return the first format that matches the string
	try {
		parsedTime = groovyDateParse("hh:mm a", String(strInTime));
		strTimeFormat = 'hh:mm a';
	} catch (pe) {
		//fall through
	}
	if (strTimeFormat_is_null(strTimeFormat)) { // Using helper to check gblNull or actual null
		try {
			parsedTime = groovyDateParse("hh:mm:ss a", String(strInTime));
			strTimeFormat = 'hh:mm:ss a';
		} catch (pe) {
			//fall through
		}
	}
	if (strTimeFormat_is_null(strTimeFormat)) {
		try {
			parsedTime = groovyDateParse("HH:mm:ss", String(strInTime));
			strTimeFormat = 'HH:mm:ss';
		} catch (pe) {
			//fall through
		}
	}
	if (strTimeFormat_is_null(strTimeFormat)) {
		try {
			parsedTime = groovyDateParse("HH:mm", String(strInTime));
			strTimeFormat = 'HH:mm';
		} catch (pe) {
			//fall through
		}
	}
	if (!strTimeFormat_is_null(strTimeFormat) && parsedTime) {
		const stdTimeFormat = new SimpleDateFormatImplMock("HH:mm:ss");
		stdTimeFormat.setLenient(false);
		// strFormatedTime = stdTimeFormat.format(parsedTime); // Result not used for return
	}
	return strTimeFormat;
}
// public static returnStdTimeFormat (String strInTime){
function DateTime_returnStdTimeFormat_Original (strInTime) { // Renamed to avoid clash
	const gblNull = GVars.GblNull('Value');
	let strOutTime = gblNull;
	const strStdFormat = 'HH:mm:ss';
	//Return the strInTime Format
	const strInTimeFormat = returnTimeFormat_Original(strInTime); // Using the renamed version
	if (!strTimeFormat_is_null(strInTimeFormat)) {
		//Convert to date
		const timeIn = groovyDateParse(strInTimeFormat, String(strInTime)); // strInTimeFormat is not null
		const stdTimeFormat = new SimpleDateFormatImplMock(strStdFormat);
		stdTimeFormat.setLenient(false);
		strOutTime = stdTimeFormat.format(timeIn);
	}
	return strOutTime;
}

// Helper for time part extractors
function DateTime_returnTimePart(strInTime, partFormat) {
	const gblNull = GVars.GblNull('Value');
	let strOutTime = gblNull;
	const strInTimeFormat = returnTimeFormat_Original(strInTime); // Use original version
	if (!strTimeFormat_is_null(strInTimeFormat)) {
		try {
			const timeIn = groovyDateParse(strInTimeFormat, String(strInTime)); // Parse with determined format
			const partFormatter = new SimpleDateFormatImplMock(partFormat);
			partFormatter.setLenient(false);
			strOutTime = partFormatter.format(timeIn);
		} catch (error) {
			console.error(`Error in returnTimePart ('${partFormat}') for '${strInTime}': ${error.message}`);
			strOutTime = gblNull;
		}
	}
	return strOutTime;
}

// public static return24HrValueFromTime (String strInTime) {
function DateTime_return24HrValueFromTime_Original (strInTime) { // Renamed
	// Use 'kk' for 1-24 hour format as in Java's SimpleDateFormat
	return returnTimePart(strInTime, 'kk');
}

// public static return23HrValueFromTime (String strInTime) {
function DateTime_return23HrValueFromTime_Original (strInTime) { // Renamed
	// Use 'HH' for 0-23 hour format
	return returnTimePart(strInTime, 'HH');
}

// public static return12HrValueFromTime (String strInTime) {
function DateTime_return12HrValueFromTime_Original (strInTime) { // Renamed
	// Use 'hh' for 1-12 hour format
	return returnTimePart(strInTime, 'hh');
}

// public static returnMinuteValueFromTime (String strInTime) {
function DateTime_returnMinuteValueFromTime_Original (strInTime) { // Renamed
	return returnTimePart(strInTime, 'mm');
}

// public static returnSecondValueFromTime (String strInTime) {
function DateTime_returnSecondValueFromTime_Original (strInTime) { // Renamed
	return returnTimePart(strInTime, 'ss');
}

// public static returnAMPMValueFromTime (String strInTime) {
function DateTime_returnAMPMValueFromTime_Original (strInTime) { // Renamed
	return returnTimePart(strInTime, 'a');
}
// public static verifyAssignedTimeFormatValid(String strInDateFormat) {
function DateTime_verifyAssignedTimeFormatValid(strInDateFormat) { // `strInDateFormat` is actually a time format
	let strActualResults;
	let boolMethodPassed = false;
	// Map<String,String> mapMethResults = new HashMap<String,String>()
	const mapMethResults = {};
	// Number intMethStatus = -1
	let intMethStatus = -1;
	switch (strInDateFormat){ // Original Groovy used strInDateFormat as variable name
		case "hh:mm a":
		case "hh:mm:ss a":
		case "HH:mm:ss":
		case "HH:mm":
			boolMethodPassed = true;
			break;
	}
	if (boolMethodPassed == true) {
		strActualResults = 'The time format assigned of: ' + strInDateFormat + ' is valid';
	} else {
		strActualResults = 'FAILED!!! The time format assigned of: ' + strInDateFormat + ' is NOT VALID';
	}
	//Set the mapMethResults
	if (boolMethodPassed == true) {
		intMethStatus = 0;
	} else {
		intMethStatus = 1;
	}
	// mapMethResults.put('methStatus', TestExecReporting.GetTestStepStatus(intMethStatus)) // Original used GetTestStepStatus
	mapMethResults['methStatus'] = TestExecReporting.GetTestStepStatus(intMethStatus); // GetTestStepStatus from Groovy
	mapMethResults['methDetails'] = strActualResults;
	//return mapMethResults
	return mapMethResults;
}
// public static returnFormatedTime (String strInTime, String strAssignedFormat) {
function DateTime_returnFormatedTime_Original (strInTime, strAssignedFormat) { // Renamed to avoid clash
	const gblNull = GVars.GblNull('Value');
	//TODO should we add a map for error handling? (Original comment)
	//TODO should we have a switch to validate the format? (Original comment)
	//Reformat the date to the specified format
	let strOutFormatedText = gblNull;
	// Date dateTemp // Not used directly

	// String strTimeFormat = this.returnTimeFormat(strInTime)
	const strTimeFormat = returnTimeFormat_Original(strInTime); // Using renamed

	if (strTimeFormat_is_null(strTimeFormat)) { // Using helper to check for gblNull and actual null
		strOutFormatedText = (strOutFormatedText === gblNull || strOutFormatedText === null ? "" : strOutFormatedText) + 'FAILED!!! The time: ' + strInTime + ' DOES NOT MATCH ASSIGNED TIME FORMATS!!!';
	} else {
		//Convert to Time
		// Date timeIn = Date.parse(strTimeFormat,strInTime)
		const timeIn = groovyDateParse(strTimeFormat, String(strInTime)); // Using mocked helper

		// SimpleDateFormat stdTimeFormat = new SimpleDateFormat(strAssignedFormat)
		const outTimeFormat = new SimpleDateFormatImplMock(strAssignedFormat);
		outTimeFormat.setLenient(false);
		try {
			strOutFormatedText = outTimeFormat.format(timeIn);
		} catch (error) {
			 strOutFormatedText = (strOutFormatedText === gblNull || strOutFormatedText === null ? "" : strOutFormatedText) + 'FAILED!!! Could not format time: ' + error.message;
		}
	}
	return strOutFormatedText;
}

// Helper to check if a format string from returnTimeFormat is effectively null
function DateTime_strTimeFormat_is_null(formatStr) {
	const gblNull = GVars.GblNull('Value');
	return formatStr === gblNull || formatStr === null;
}