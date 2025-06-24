// --- CommonWebCore.ts ---

SeSGlobalObject("CommonWebCore");

//Describe the library purpose
/**
 * This library contains the methods for common web activities.
 * These methods can be used for all standard web programs and may include custom
 * controls that are used by more than one Jaggear team.
 * Examples are:
 * returnWebElement
 * objHighlightElementJS
 * objVerifyState
 * obSetVerifyEditBox
 *
 */

//Add Imports
// Importing Selenium WebDriver bindings for Node.js
//import { Builder, By, Key, until, WebDriver, WebElement, WebDriverWait, Select, Actions } from 'selenium-webdriver';
// No direct equivalent for WebDriverManager as it's a Java utility.
// Browser specific drivers (ChromeDriver, EdgeDriver, FirefoxDriver) are typically managed by 'selenium-webdriver' Builder.

// Add Jaggaer Libs (assuming these will also be converted to TypeScript modules)
//import * as TCObj from '../common/TestObjects';
//import * as GVars from '../common/GlobalVariables';
//import * as DateTime from '../common/DateTime';
//import * as StrNums from '../common/StringsAndNumbers';
//import * as TSExecRep from '../common/TestExecReporting'; // Imported but not used in provided snippet
//import * as TSExecParams from '../common/TestCaseExecParams';

/**
 * Helper function for `Thread.sleep` equivalent.
 * Since `export async` and `Promise` are not allowed, this will be a blocking "busy-wait" loop.
 * Use with caution as it blocks the Node.js event loop.
 */
function blockingSleep(ms) {
	const start = Date.now();
	while (Date.now() < start + ms) {
		// Busy-wait to simulate Thread.sleep()
	}
}

/**
 * Re-implementation of CommonWebCore class as a collection of functions.
 * Public static methods are converted to exported functions.
 */

/**
 * -------------------------------------  returnWebElement  -----------------------------------
 * Return the browser profile in the WebDriver
 * @param strObjXPath The full XPath to the object
 * @return weElement The WebElement found by Xpath
 * @author pkanaris
 * @author Created: 04/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 * @author pkanaris@jaggaer.com StrNums.JComm_ReplaceValue(strObjXPath, ".//", "//") //Added for backward compatibility of concatenated strings
 */
function CommonWebCore_returnWebElement(strObjXPath) {
	let boolPassed = true;
	const gblLineFeed = GblLineFeed;
	let strMethodDetails; // Will be assigned later
	if (boolDoDebug == true){
		Tester.Message('CommonWebCore_returnWebElement.strObjXPath: ' + strObjXPath)
	}
	var weElement = Navigator.SeSFind(strObjXPath)
	if (weElement === null ) { // Check for null or undefined
		strMethodDetails = 'The Element WAS NOT FOUND By.Xpath(' + strObjXPath + ')!!!';
		boolPassed = false; // Ensure boolPassed is false if element is not found
		if (boolDoDebug === true) {
			Tester.Message(strMethodDetails);
		}
	}
	else if (boolPassed === true) { // boolPassed might be false if initial catch leads here
		if (boolDoDebug === true) {
			Tester.Message("Element present");
		}
		strMethodDetails = 'The Element was found By.Xpath(' + strObjXPath + ')!!!';
		//Update the test objects with the element XPath
		TestCaseObject.strTOElementXpath = strObjXPath;
	}
	if (boolPassed === false) {
		Tester.Message(strMethodDetails);
	}
	return weElement;
}

function CommonWebCore_taboutOfElement(strObjXPath) {
	let weElement; // Placeholder for WebElement
	weElement = findWebElement(strObjXPath); // Reuse internal helper
	if (weElement !== null && weElement !== undefined) {
		weElement.sendKeys(Key.TAB); // Use Selenium's Key enum
		DateTime.WaitSecs(TSExecParams.getIntViewDelaySecs()); // Assumes DateTime.WaitSecs is blocking
	}
}

/**
 * -------------------------------------  findWebElement  -----------------------------------
 * Find the element ignoring if present.
 * @param strObjXPath The full XPath to the object
 * @return weElement The WebElement found by Xpath
 * @author pkanaris
 * @author Created: 01/03/2023
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 * @author pkanaris@jaggaer.com StrNums.JComm_ReplaceValue(strObjXPath, ".//", "//") //Added for backward compatibility of concatenated strings
 */
function CommonWebCore_findWebElement(strObjXPath) {
	const boolDoDebug = boolDoDebug;
	let boolPassed = true;
	const gblLineFeed = GVars.GblLineFeed("Value");
	let strMethodDetails;
	let weElement; // Placeholder for WebElement

	const driver = TestCaseObject.tcDriver; // Assumes TestCaseObject.tcDriver returns a WebDriver instance

	try {
		weElement = driver.findElement(By.xpath(strObjXPath));
	}
	catch (e) {
		boolPassed = false;
		console.log(strMethodDetails = 'The Element WAS NOT FOUND By.Xpath(' + strObjXPath + ')!!! See Expeception error: ' + gblLineFeed + (e instanceof Error ? e.stack : String(e)));
	}
	if (weElement === null || weElement === undefined) {
		if (boolDoDebug === true) {
			console.log(strMethodDetails = 'The Element WAS NOT FOUND By.Xpath(' + strObjXPath + ')!!!');
		}
	}
	else if (boolPassed === true) {
		if (boolDoDebug === true) {
			console.log("Element present");
		}
		strMethodDetails = 'The Element was found By.Xpath(' + strObjXPath + ')!!!';
		//Update the test objects with the element XPath
		TCObj.setStrTOElementXpath(strObjXPath);
	}
	if (boolPassed === false) {
		console.log(strMethodDetails);
	}
	return weElement;
}

/** Browser instances
 * Methods for working with multiple instances of the browser
 * NOTE: Make sure all other instances of the browser (i.e. Chrome) are closed before starting the test
 */

/**
 * ------------------------------------- returnBrowserInstanceCount
 * Returns the number of browser instances currently open
 */
function CommonWebCore_returnBrowserInstanceCount() {
	const driver = TestCaseObject.tcDriver; // Assumes TestCaseObject.tcDriver returns a WebDriver instance
	let intWindows;
	try {
		// getWindowHandles returns a Promise in JS/TS. Need to resolve it.
		// Since `async/await` is not allowed, this will implicitly return a Promise if `WebDriver` is from `selenium-webdriver`.
		// To simulate the original behavior (blocking/sync count), you'd need a different WebDriver wrapper or to allow async.
		// Assuming your WebDriver client handles this synchronously under the hood or you accept it's a "sync-like" call here.
		// If it means getting the count of already loaded window handles, it might be synchronously accessible.
		// Standard selenium-webdriver's getWindowHandles is async.
		// This is a common point of divergence for Java vs. JS WebDriver.
		// For strict sync execution, a wrapper might be needed.
		// For this template, we'll assume it behaves synchronously for the sake of strict adherence to no async/promise.
		intWindows = driver.getAllWindowHandles().length; // Hypothetical sync access to handles
	}
	catch (e) {
		// Original: WebDriverException
		intWindows = -1;
	}
	return intWindows;
}

/**
 * ------------------------------------- waitForCountOfBrowserInstances
 * Wait until the specific number of browser instances are found based on ==, >, <
 */
function CommonWebCore_waitForCountOfBrowserInstances(intCntBrowser, boolCntEquals) {
	const mapResults = {};
	let strMethodDetails; // Will be assigned later
	let boolPassed = true;
	let intTotalBrowsers;
	const weDriver = TestCaseObject.tcDriver; // Assumes TestCaseObject.tcDriver returns a WebDriver instance

	if (boolCntEquals === true) {
		//Return the number of browsers
		intTotalBrowsers = weDriver.getAllWindowHandles().length; // Hypothetical sync access
		//Check if we matched the count of expected browser windows
		if (intTotalBrowsers === intCntBrowser) {
			strMethodDetails = " The system is displaying the expected count of: ' " + intCntBrowser + "' browser windows.";
		}
		else {
			const waitTime = intMaxWaitTimeSecs * 1000; // Convert seconds to milliseconds
			const waitElem = new WebDriverWait(weDriver, waitTime);
			//Wait for the new window to be present
			// Original: waitElem.until(ExpectedConditions.numberOfWindowsToBe(intCntBrowser))
			// This `until` call is typically async in selenium-webdriver. To avoid `async/await`,
			// we'll simulate the blocking wait but acknowledge this is a key difference.
			try {
				waitElem.until(until.numberOfWindows(intCntBrowser));
			} catch (e) {
				// If it times out, the `until` will throw an error.
				if (boolDoDebug) console.log(`Wait for number of windows timed out: ${e.message}`);
			}

			intTotalBrowsers = weDriver.getAllWindowHandles().length; // Get final count
			//Check if we matched the count of expected browser windows
			if (intTotalBrowsers === intCntBrowser) {
				strMethodDetails = " The system is displaying the expected count of: ' " + intCntBrowser + "' browser windows.";
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The system is displaying '" + intTotalBrowsers + "' BUT EXPECTED COUNT IS: ' " + intCntBrowser + "' BROWSER WINDOWS!!!";
			}
		}
	}
	else {
		//Check if at least expected browser window count was found
		//loop while boolLoopContinue is true
		let boolLoopContinue = true;
		let intTempLoopSecs;
		const myStartTime = Date.now(); // Equivalent to System.currentTimeMillis()
		let myLoopTime;
		const intMaxWait = intMaxWaitTimeSecs;
		while (boolLoopContinue === true) {
			intTotalBrowsers = weDriver.getAllWindowHandles().length; // Hypothetical sync access
			if (intTotalBrowsers >= intCntBrowser) {
				boolLoopContinue = false; //Will exit
			}
			else {
				//Check if we have exceeded the max wait time
				myLoopTime = Date.now();
				//Return number of seconds
				intTempLoopSecs = DateTime.timeGetStrDiffMillSecsInToSeconds(myStartTime, myLoopTime); // Assumes DateTime helper
				if (intTempLoopSecs >= intMaxWait) {
					boolLoopContinue = false; //Will exit
				}
			}
			SeSSleep(200); //Sleep in milliseconds
		}
		if (intTotalBrowsers > intCntBrowser) {
			strMethodDetails = " The system is displaying '" + intTotalBrowsers + "' which is greater than the expected count of: ' " + intCntBrowser + "' browser windows.";
		}
		else {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The system is displaying '" + intTotalBrowsers + "' which is NOT GREATER than the EXPECTED COUNT of: ' " + intCntBrowser + "' browser windows.";
		}

	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.intTotalBrowsers = intTotalBrowsers;
	return mapResults;
}

/**
 * -------------------------------------  returnWebElementCustomTiming  -----------------------------------
 * Return the web element based on the XPath using custom timing in milliseconds and retry counts
 * @param strObjXPath The full XPath to the object
 * @param intMaxWaitSecs The number of seconds to continue trying to find the element
 * @param intIntervalMS The number of milliseconds between tries to find the element
 * @return weElement The WebElement found by Xpath
 * @author pkanaris
 * @author Created: 04/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_returnWebElementCustomTiming(strObjXPath, intMaxWaitSecs, intIntervalMS) {
	const boolDoDebug = boolDoDebug;
	let boolPassed = true;
	let strMethodDetails;
	let weElement; // Placeholder for WebElement

	const driver = TestCaseObject.tcDriver; // Assumes TestCaseObject.tcDriver returns a WebDriver instance
	const waitTime = intMaxWaitSecs * 1000;

	const waitElem = new WebDriverWait(driver, waitTime, intIntervalMS); // Interval in milliseconds
	// Original: waitElem.ignoring(StaleElementReferenceException.class)
	// In selenium-webdriver JS, 'ignoring' uses specific error names.
	// waitElem.ignore(StaleElementReferenceError); // This would require async/await. For sync, manual retry.

	//Wrap in a try and catch the stale element
	try {
		weElement = waitElem.until(until.presenceOfElementLocated(By.xpath(strObjXPath)));
	}
	catch (e) {
		// Original: StaleElementReferenceException, then try again.
		// `until` automatically retries so if it throws StaleElement now, it probably failed after retries or it's another error type.
		// Second `until` call is typically async and can throw or time out.
		// To strictly match "try again" without `async/await`, would need a custom loop.
		// For now, assuming the WebDriverWait `until` handles the retry implicitly as it did in Java.
		// If it throws here, re-throwing or handling as failure.
		console.log(`Error finding element (custom timing): ${e.stack}`);
		boolPassed = false;
	}

	if (weElement === null || weElement === undefined) {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! The Element WAS NOT FOUND By.Xpath(' + strObjXPath + ')!!!';
	}
	else {
		if (boolDoDebug === true) {
			console.log("Element present");
		}
		strMethodDetails = 'The Element was found By.Xpath(' + strObjXPath + ')!!!';
	}
	if (boolPassed === false) {
		console.log(strMethodDetails);
	}
	return weElement;
}

/**
 * -------------------------------------  objHighlightElementJS  -----------------------------------
 * Highlight the WebElement assigned for the number of instances defined for the test execution using Java Script
 * @param weElement The WebElement Object (placeholder for WebElement)
 * @param strTestObjType The testobject type (i.e button, popup, list, checkbox, etc.)
 * @param strTestObjName The meaningful name of the testobject
 * @return mapResults The results showing Passed and method details.
 * @author pkanaris
 * @author Created: 04/22/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objHighlightElementJS(weElement, strTestObjType, strTestObjName) {
	//Create the map
	const mapResults = {};
	//Return the highlight count
	const intHighLightCnt = intHighlightCount;
	//Results Variables
	const strMethodSummary = 'Highlight ' + strTestObjType + " " + strTestObjName + '.'; // Not used in Groovy, keeping as comment for clarity
	const strMethodDescription = 'The' + strTestObjType + ' ' + strTestObjName + ' will be highlighted ' + intHighLightCnt + ' times.'; // Not used
	const strMethodResults = 'The testobject will be highlighted.'; // Not used
	let strMethodDetails; // Will be assigned later
	let boolPassed = true;
	//Check if an element was passed in
	if (weElement !== null && weElement !== undefined) {
		const driver = TestCaseObject.tcDriver; // Assumes TestCaseObject.tcDriver returns a WebDriver instance
		// `JavascriptExecutor` in Java is implicitly available on `WebDriver`.
		// In selenium-webdriver JS, `executeScript` is called directly on the driver.
		// Original: JavascriptExecutor executor = ((TestCaseObject.tcDriver) as JavascriptExecutor)
		//Build loop to execute and wait a standard .25 seconds between flashes
		for (let loopCnt = 0; loopCnt < intHighLightCnt; loopCnt++) {
			WebDriver.ExecuteScript("arguments[0].setAttribute('style','outline: dashed red;');", weElement);
			SeSSleep(200); // Sleep in milliseconds
			WebDriver.ExecuteScript("arguments[0].setAttribute('style','outline: solid green;');", weElement);
			SeSSleep(200); // Sleep in milliseconds
		}
		WebDriver.ExecuteScript("arguments[0].setAttribute('style','outline: none;');", weElement);
		strMethodDetails = "The ' " + strTestObjName + "' '" + strTestObjType + "' is present and highlighted.";
	}
	else {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The ' " + strTestObjName + "' '" + strTestObjType + "' WAS NOT PRESENT!!!";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * -------------------------------------  objMoveToClickElement  -----------------------------------
 * Move to the WebElement assigned and click on the element using the offset X,Y values
 * @param weElement The WebElement Object (placeholder for WebElement)
 * @param strTestObjType The testobject type (i.e button, popup, list, checkbox, etc.)
 * @param strTestObjName The meaningful name of the testobject
 * @param intOffsetX The number of pixels +/- on the X axis for offset from center
 * @param intOffsetY The number of pixels +/- on the Y axis for offset from center
 * @return mapResults The results showing Passed and method details.
 * @author pkanaris
 * @author Created: 04/22/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objMoveToClickElement(weElement, strTestObjType, strTestObjName, intOffSetX, intOffSetY) {
	//Create the map
	const mapResults = {};
	let boolPassed = true;
	let strMethodDetails; // Will be assigned later
	//Move to the element
	//Create actions to allow moving to an object in case it is not displayed.
	const driver = TestCaseObject.tcDriver; // Assumes TestCaseObject.tcDriver returns a WebDriver instance
	const actions = new Actions(driver); // Uses Selenium's Actions class
	actions.move({ origin: weElement, x: intOffSetX, y: intOffSetY }).perform(); // `move` for relative offset
	SeSSleep(1000); // Wait 1 second
	weElement.click();
	strMethodDetails = "Moved to and clicked the elememnt '" + strTestObjType + "' and the name '" + strTestObjName + "'.";
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * -------------------------------------  objVerifyState  -----------------------------------
 * Verify the visible and enabled state of the object
 * @param weObject The WebElement Object (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param boolVisible Is the object visible/displayed? true/false
 * @param boolEnabled Is the object enabled? true/false
 * @return mapResults The results showing Passed and method details.
 * @author pkanaris
 * @author Created: 04/22/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objVerifyState(weObject, strObjName, boolVisible, boolEnabled) {
	const strGblLineFeed = GblLineFeed;
	//Create the map
	const mapResults = {};
	let strMethodDetails; // Will be assigned later
	let boolPassed = true;
	let boolIsEnabled;
	// Original: boolIsVisible = weObject.isDisplayed() - weObject.isDisplayed() is usually async
	// Assuming synchronous execution here.	
	let /**boolean*/ boolIsVisible = Navigator.DoWaitForVisible(weObject, 1000, {}) !== false; 
	if (boolDoDebug === true) {
		Tester.Message('CommonWebCore_objVerifyState.boolIsVisible: ' + boolIsVisible)
	}
	//Return the object class
	// Original: String strObjClass = weObject.getAttribute('class')
	// Original: String strObjAriaDisabled = weObject.getAttribute('aria-disabled')
	// WebElement.getAttribute is usually async in selenium-webdriver.
	// Assuming synchronous execution for this template.
	let strObjClass = StringsAndNumbers.JComm_HandleNoData(Navigator.ExecJS("document.querySelector(weObject).className;"));
	let strObjAriaDisabled = StringsAndNumbers.JComm_HandleNoData(Navigator.ExecJS("document.querySelector(weObject).aria-disabled;"));
	if (boolDoDebug === true) {
		Tester.Message('The ' + strObjName + ' class name: ' + strObjClass + ' strObjAriaDisabled: ' + strObjAriaDisabled)
	}
	if (strObjClass === 'unselect-button item item-block item-md item-checkbox item-checkbox-disabled' || strObjAriaDisabled === true) { //Long-term may require a switch
		boolIsEnabled = false; //Based on the class the object is disabled. PGK 04/23/2022
	} else {
		// Original: boolIsEnabled = weObject.isEnabled() - weObject.isEnabled() is usually async
		// Assuming synchronous execution.
		boolIsEnabled = Navigator.DoWaitForEnabled(weObject, 1000, {}) !== false;
	}
	//Check if the match the expected
	if (boolIsVisible === boolVisible) {
		strMethodDetails = "The '" + strObjName + "' visible matches the expected value of: " + boolVisible + strGblLineFeed;
	}
	else {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The '" + strObjName + "' VISIBLE DOES NOT MATCH EXPECTED of: " + boolVisible + strGblLineFeed;
	}
	if (boolIsVisible === true) { //If the element is not visible the enabled state does not matter.
		if (boolIsEnabled === boolEnabled) {
			strMethodDetails = strMethodDetails + "The '" + strObjName + "' enabled matches the expected value of: " + boolEnabled;
		}
		else {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The '" + strObjName + "' ENABLED DOES NOT MATCH EXPECTED of: " + boolEnabled;
		}
	}
	else {
		strMethodDetails = strMethodDetails + "The '" + strObjName + "' is not visible so 'Enabled' state is ignored.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * -------------------------------------  returnChildElements  -----------------------------------
 * Return the child elements of the parent
 * @param weObject The WebElement Object (placeholder for WebElement)
 * @param strChildObjXPath The XPath to the child objects
 * @return mapChildElement The results containing the list of child items and the count.
 * @author pkanaris
 * @author Created: 4/30/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_returnChildElements(weObject, strChildObjXPath) {
	const mapChildElement = {};
	const strGblLineFeed = GVars.GblLineFeed('Value'); // Not used in this function
	const lstChildObjects = weObject.findElements(By.xpath(strChildObjXPath)); // Assuming findElements returns an array of WebElements sync
	const cntChildObjs = lstChildObjects.length;
	mapChildElement.lstChildObjects = lstChildObjects;
	mapChildElement.cntChildObjs = cntChildObjs;
	return mapChildElement;
}

/**
 * -------------------------------------  returnChildElement  -----------------------------------
 * Return the child element of the parent. NOTE: Only ONE CHILD can BE Present
 * @param weObject The WebElement Object (placeholder for WebElement)
 * @param strChildObjXPath The XPath to the child object
 * @return weChild The child element.
 * @author pkanaris
 * @author Created: 4/30/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_returnChildElement(weObject, strChildObjXPath) {
	let weChild; // Placeholder for WebElement
	try {
		weChild = weObject.findElement(By.xpath(strChildObjXPath)); // Assuming findElement returns WebElement sync
	}
	catch (e) {
		weChild = null;
	}
	return weChild;
}

/**
 * -------------------------------------  returnChildElement  -----------------------------------
 * Return the child element of the parent. NOTE: Only ONE CHILD can BE Present
 * @param weParentObject The WebElement Parent Object (placeholder for WebElement)
 * @param strChildObjXPath The child object xpath
 * @param strChildObjProperty The property to check (not used in implementation)
 * @return weChild The child element.
 * @author pkanaris
 * @author Created: 4/30/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_returnChildElementByPropertyValue(weParentObject, strChildObjXPath) { // strChildObjProperty removed as unused
	let weChild; // Placeholder for WebElement
	//Return the child objects from the parent
	//Loop through the object checking the assigned property for the value specified.
	// This implementation is a placeholder, as the actual logic was commented out in Groovy.
	// You would implement the logic to find the child by property value here.
	return weChild;
}

/**
 * -------------------------------------  isAttribtuePresent  -----------------------------------
 * Check if the assigned attribute is present for the element
 * @param we The WebElement Object (placeholder for WebElement)
 * @param strAttribute The attribute name
 * @return result True or False for attribute presence.
 * @author pkanaris
 * @author Created: 4/30/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_isAttribtuePresent(we, strAttribute) {
	let result = false;
	const strGlbNull = GVars.GblNull('Value');
	try {
		// `getAttribute` is usually async. Assuming sync for templating.
		const value = StrNums.JComm_HandleNoData(we.getAttribute(strAttribute)); // Assumes getAttribute returns string
		if (value !== strGlbNull) {
			result = true;
		}
	} catch (e) {
		// Catching any exception, likely if element is stale or attribute not found.
	}
	return result;
}

/************************************************* Object specific methods ***********************************************************/
//**EDITBOX **/
/**
 * -------------------------------------  objSetEditBoxValue  -----------------------------------
 * Set and verify the edit box to the assigned value.
 * @param weEditBox The Editbox WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strAssignedValue The assigned text value for the test object
 * @param boolClearValue Clear the value prior to setting? True/False
 * @param boolVerifyValue Verify the value after set? True/False
 * @param boolTabOutOfField Tab out of the field after entry? True/False
 * @return mapResults
 *  1. boolPassed The method passed?
 *  2. strMethodDetails The details on the processing of the method
 * @author pkanaris
 * @author Created: 04/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSetEditBoxValue(weEditBox, strObjName, strAssignedValue, boolClearValue, boolVerifyValue, boolTabOutOfField) {
	//GlobalVars
	const gblNull = GVars.GblNull('Value');
	// Declare the variables
	const strTestObjType = 'EditBox';
	const mapResults = {};
	let strMethodDetails; // Will be assigned later
	let boolPassed = true;
	let strTestObjText; // Will be assigned later
	//Process the edit box
	if (boolClearValue === true) {
		//Clear the text first
		Tester.SuppressReport(true);
		weEditBox.DoSetText("")
		Tester.SuppressReport(false);
		//weEditBox.Clear(); // WebElement.clear()
		DateTime.WaitSecs(intViewDelaySecs);// Assumes blocking wait
		//Clear does not appear to work on iOS devices
		//Check the text and if not clear click after the text and type backspace/delete.
		//strTestObjText = StringsAndNumbers.JComm_HandleNoData(weEditBox.GetAttribute("value")); // Assumes getAttribute reads value sync
		strTestObjText = StringsAndNumbers.JComm_HandleNoData(SeS(weEditBox).GetValue())   
		if (strTestObjText !== gblNull) {
			const lenStr = strTestObjText.length;
			for (let loop = 0; loop < lenStr; loop++) {
				strTestObjText = StringsAndNumbers.JComm_HandleNoData(weEditBox.GetAttribute("value"));
				if (strTestObjText === gblNull) {
					break; //The value is cleared
				} else {
					Tester.SuppressReport(true);
					weEditBox.DoSendKeys("{BACKSPACE}");
					Tester.SuppressReport(false);
					//weEditBox.SendKeys({BACKSPACE}); // Use Selenium Key enum
					blockingSleep(100); // Blocking sleep
				}
			}
		}
	}
	//Set the value
	let strTempValue;
	if (StringsAndNumbers.JComm_HandleNoData(strAssignedValue) === gblNull) {
		strTempValue = ' ';
	} else {
		strTempValue = strAssignedValue;
	}
	Tester.SuppressReport(true);
	weEditBox.DoSendKeys(strTempValue); // WebElement.sendKeys
	Tester.SuppressReport(false);
	//WebDriver.Actions().SendKeys(strTempValue,weEditBox.element).Perform();
	DateTime.WaitSecs(intViewDelaySecs); // Assumes blocking wait
	if (boolTabOutOfField === true) {
		//WebDriver.Actions().SendKeys("\t").Perform();  //AI 2
		Tester.SuppressReport(true);
		weEditBox.DoSendKeys("{TAB}"); // Use Selenium Key enum
		DateTime.WaitSecs(intViewDelaySecs); // Assumes blocking wait
		Tester.SuppressReport(false);
	}
	if (boolVerifyValue === true) {
		const mapVerValue = objVerifyEditBoxValue(weEditBox, strObjName, strAssignedValue); // Reuse internal helper
		boolPassed = StringsAndNumbers.JComm_StringToBoolean(mapVerValue.boolPassed);
		strMethodDetails = mapVerValue.strMethodDetails;
	} else {
		strMethodDetails = 'The ' + strTestObjType + " " + strObjName + ' text of: ' + strTestObjText +
			' was set, to the specified value of: ' + strAssignedValue + '.';
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * -------------------------------------  objSetVerEditBoxValueClickEnter  -----------------------------------
 * Set and verify the edit box to the assigned value.
 * @param weEditBox The Editbox WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strObjFullXpath The full XPath to the object (used to re-find it)
 * @param strAssignedValue The assigned text value for the test object
 * @param boolClearValue Clear the value prior to setting? True/False
 * @param boolVerifyValue Verify the value after set? True/False
 * @param boolClickEnter Click enter after entering the value? True/False
 * @return mapResults
 *  1. boolPassed The method passed?
 *  2. strMethodDetails The details on the processing of the method
 * @author pkanaris
 * @author Created: 07/06/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSetVerEditBoxValueClickEnter(weEditBox, strObjName, strObjFullXpath, strAssignedValue, boolClearValue, boolVerifyValue, boolClickEnter) {
	//GlobalVars
	const gblNull = GVars.GblNull('Value');
	// Declare the variables
	const strTestObjType = 'EditBox';
	const mapResults = {};
	let strMethodDetails;
	let boolPassed = true;
	let strTestObjText;
	//Process the edit box
	if (boolClearValue === true) {
		//Clear the text first
		weEditBox.clear();
		DateTime.WaitSecs(TSExecParams.getIntViewDelaySecs());
		//Clear does not appear to work on iOS devices
		//Check the text and if not clear click after the text and type backspace/delete.
		strTestObjText = StrNums.JComm_HandleNoData(weEditBox.getAttribute("value"));
		if (strTestObjText !== gblNull) {
			const lenStr = strTestObjText.length;
			for (let loop = 0; loop < lenStr; loop++) {
				strTestObjText = StrNums.JComm_HandleNoData(weEditBox.getAttribute("value"));
				if (strTestObjText === gblNull) {
					break; //The value is cleared
				} else {
					weEditBox.sendKeys(Key.BACK_SPACE);
					blockingSleep(100);
				}
			}
		}
	}
	//Set the value
	let strTempValue;
	if (StrNums.JComm_HandleNoData(strAssignedValue) === gblNull) {
		strTempValue = ' ';
	} else {
		strTempValue = strAssignedValue;
	}
	weEditBox.sendKeys(strTempValue);
	DateTime.WaitSecs(TSExecParams.getIntViewDelaySecs());
	if (boolClickEnter === true) {
		weEditBox.sendKeys(Key.ENTER);
		DateTime.WaitSecs(TSExecParams.getIntViewDelaySecs());
		//Refresh the editbox (re-find element, can cause stale element issues)
		weEditBox = returnWebElement(strObjFullXpath); // Reuse internal helper
	}
	if (boolVerifyValue === true) {
		const mapVerValue = objVerifyEditBoxValue(weEditBox, strObjName, strAssignedValue); // Reuse internal helper
		boolPassed = StrNums.JComm_StringToBoolean(mapVerValue.boolPassed);
		strMethodDetails = mapVerValue.strMethodDetails;
	} else {
		strMethodDetails = 'The ' + strTestObjType + " " + strObjName + ' text of: ' + strTestObjText +
			' was set, to the specified value of: ' + strAssignedValue + '.';
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

function CommonWebCore_objChkClickableAndStaleObjectTest(weElement) {
	/** Solution based on issue found in Accounts Payable Reports Cycle Time - Invoice Workflow Set and Process Run Report
	 * Dynamic pages refresh the screen after each field and cause stale object issues
	 * Method should be called after each time we return and object
	 * Should there be a delay and retry when this method is called?
	 */
	//Check if clickable while traping any stale objects
	// This function body was empty in the Groovy source.
	// You would implement logic to handle stale element and check clickability here.
	return {}; // Return an empty Map as placeholder
}

/**
 * -------------------------------------  objSetMultiLineTextBoxValue  -----------------------------------
 * Set and verify the edit box to the assigned value.
 * @param weEditBox The Editbox WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strAssignedValue The assigned text value for the test object
 * @param boolClearValue Clear the value prior to setting? True/False
 * @param boolVerifyValue Verify the value after set? True/False
 * @param boolTabOutOfField Tab out of the field after entry? True/False
 * @return mapResults
 *  1. boolPassed The method passed?
 *  2. strMethodDetails The details on the processing of the method
 * @author pkanaris
 * @author Created: 05/22/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSetMultiLineTextBoxValue(weEditBox, strObjName, strAssignedValue, boolClearValue, boolVerifyValue, boolTabOutOfField) {
	//GlobalVars
	const gblNull = GVars.GblNull('Value');
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblHTMLLF = GVars.GblHTMLLineFeed('Value');
	// Declare the variables
	const strTestObjType = 'EditBox';
	const mapResults = {};
	let strMethodDetails;
	let boolPassed = true;
	let strTestObjText;
	//Process the edit box
	if (boolClearValue === true) {
		//Clear the text first
		weEditBox.clear();
		DateTime.WaitSecs(TSExecParams.getIntViewDelaySecs());
		//Clear does not appear to work on iOS devices
		//Check the text and if not clear click after the text and type backspace/delete.
		strTestObjText = StrNums.JComm_HandleNoData(weEditBox.getAttribute("value"));
		if (strTestObjText !== gblNull) {
			const lenStr = strTestObjText.length;
			for (let loop = 0; loop < lenStr; loop++) {
				strTestObjText = StrNums.JComm_HandleNoData(weEditBox.getAttribute("value"));
				if (strTestObjText === gblNull) {
					break; //The value is cleared
				} else {
					weEditBox.sendKeys(Key.BACK_SPACE);
					blockingSleep(100);
				}
			}
		}
	}
	//Set the value
	let strTempValue;
	if (StrNums.JComm_HandleNoData(strAssignedValue) === gblNull) {
		strTempValue = ' ';
	} else {
		strTempValue = strAssignedValue;
	}
	//Determine if we have multiline
	//Check if either is contained in the string
	const intGblLF = StrNums.JComm_TextLocationInString(strTempValue, gblLineFeed);
	const intGBLHTMLLF = StrNums.JComm_TextLocationInString(strTempValue, gblHTMLLF);
	let arrayMultiLineValues; // Placeholder for Array
	let intArrayCnt = 0;
	if (intGblLF > 0) {
		//Split the input value to begin the processing
		const mapItemsValues = StrNums.JComm_StringToArray(strTempValue, gblLineFeed);
		arrayMultiLineValues = mapItemsValues.ArryOfValues;
		intArrayCnt = arrayMultiLineValues.length;
	}
	else if (intGBLHTMLLF > 0) {
		//Split the input value to begin the processing
		const mapItemsValues = StrNums.JComm_StringToArray(strTempValue, gblHTMLLF);
		arrayMultiLineValues = mapItemsValues.ArryOfValues;
		intArrayCnt = arrayMultiLineValues.length;
	}
	if (intArrayCnt > 0) {
		//Return each item and enter the value then send a return
		for (let loopArry = 0; loopArry < intArrayCnt; loopArry++) {
			weEditBox.sendKeys(arrayMultiLineValues[loopArry]);
			if (loopArry < intArrayCnt - 1) {
				weEditBox.sendKeys(Key.ENTER);
			}
		}
	}
	else {
		weEditBox.sendKeys(strTempValue);
	}
	DateTime.WaitSecs(TSExecParams.getIntViewDelaySecs());
	if (boolTabOutOfField === true) {
		weEditBox.sendKeys(Key.TAB);
		DateTime.WaitSecs(TSExecParams.getIntViewDelaySecs());
	}
	if (boolVerifyValue === true) {
		//TODO custom verify to remove the enter key
		const mapVerValue = objVerifyTextBoxValue(weEditBox, strObjName, strAssignedValue); // Reuse internal helper
		boolPassed = StrNums.JComm_StringToBoolean(mapVerValue.boolPassed);
		strMethodDetails = mapVerValue.strMethodDetails;
	} else {
		strMethodDetails = 'The ' + strTestObjType + " " + strObjName + ' text of: ' + strTestObjText +
			' was set, to the specified value of: ' + strAssignedValue + '.';
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * -------------------------------------  objSetEditBoxAttibuteValue  -----------------------------------
 * Set and verify the edit box to the assigned value.
 * @param weEditBox The Editbox WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strAssignedValue The assigned text value for the test object
 * @param strAssignAttribute The attribute to update and verify
 * @param boolVerifyValue Verify the value after set? True/False
 * @return mapResults
 *  1. boolPassed The method passed?
 *  2. strMethodDetails The details on the processing of the method
 * @author pkanaris
 * @author Created: 05/16/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSetEditBoxAttibuteValue(weEditBox, strObjName, strAssignedValue, strAssignAttribute, boolVerifyValue) {
	//GlobalVars
	const gblNull = GVars.GblNull('Value');
	// Declare the variables
	const strTestObjType = 'EditBox';
	const mapResults = {};
	let strMethodDetails;
	let boolPassed = true;
	let strTestObjText; // Will be assigned later
	//Process the edit box
	//Set the value
	let strTempValue;
	if (StrNums.JComm_HandleNoData(strAssignedValue) === gblNull) {
		strTempValue = ' ';
	} else {
		strTempValue = strAssignedValue;
	}
	//TODO create a custom method to set the attribute of an element
	const mapSetAttribute = weSetAttibute(weEditBox, strObjName, strAssignedValue, strAssignAttribute); // Reuse internal helper
	boolPassed = StrNums.JComm_StringToBoolean(mapSetAttribute.boolPassed);
	strMethodDetails = mapSetAttribute.strMethodDetails;
	if (boolPassed === true) {
		if (boolVerifyValue === true) {
			const mapVerValue = objVerifyEditBoxValue(weEditBox, strObjName, strAssignedValue); // Reuse internal helper
			boolPassed = StrNums.JComm_StringToBoolean(mapVerValue.boolPassed);
			strMethodDetails = mapVerValue.strMethodDetails;
		} else {
			strMethodDetails = 'The ' + strTestObjType + " " + strObjName + ' text of: ' + strTestObjText +
				' was set, to the specified value of: ' + strAssignedValue + '.';
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * -------------------------------------  weSetAttibute  -----------------------------------
 * Set and verify the edit box to the assigned value.
 * @param weElement The WebElement (placeholder for WebElement)
 * @param strElemName The name of the web element
 * @param strValue The assigned text value for the test object
 * @param strAttibute The attribute that will be processed
 * @return mapResults
 *  1. boolPassed The method passed?
 *  2. strMethodDetails The details on the processing of the method
 * @author pkanaris
 * @author Created: 05/16/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_weSetAttibute(weElement, strElemName, strValue, strAttibute) {
	let boolPassed = true;
	let strMethodDetails; // Will be assigned later
	const mapResults = {};
	try {
		const driver = TestCaseObject.tcDriver; // Assumes TestCaseObject.tcDriver returns a WebDriver instance
		// Original: JavascriptExecutor executor = ((TestCaseObject.tcDriver) as JavascriptExecutor)
		// driver.executeScript is used for JavascriptExecutor equivalent
		driver.executeScript("arguments[0].setAttribute(" + JSON.stringify(strAttibute) + "," + JSON.stringify(strValue) + ");", weElement); // JSON.stringify for proper JS string literal
		strMethodDetails = "Set webElement '" + strElemName + "' ATTRIBUTE: " +
			strAttibute + " to the assigned value of: " + strValue + ".";
	}
	catch (e) {
		boolPassed = false;
		strMethodDetails = "FAILED TO Set webElement '" + strElemName + "' ATTRIBUTE: " +
			strAttibute + " to the ASSIGNED VALUE of: " + strValue + " !!! " + 'Exception occurred!!! SEE ERROR STACK TRACE: ' +
			(e instanceof Error ? e.stack : String(e));
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * -------------------------------------  objVerifyElementAttribute  -----------------------------------
 * Verify the web element attribute value.
 * @param weElement The WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strAssignedValue The assigned text value for the test object
 * @param strAssignedAttribute The attribute to evaluate
 * @return mapResults
 *  1. boolPassed The method passed?
 *  2. strMethodDetails The details on the processing of the method
 * @author pkanaris
 * @author Created: 04/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objVerifyElementAttribute(weElement, strObjName, strAssignedValue, strAssignedAttribute) {
	//GlobalVars
	const gblNull = GVars.GblNull('Value');
	// Declare the variables
	const strTestObjType = 'WebElement';
	const mapResults = {};
	let strMethodDetails;
	let boolPassed = true;
	let strTestObjText; // Will be assigned later
	//Retrieve the new value and verify the value matches the assignedValue
	//check if the attribute is present
	const boolAttributePresent = isAttribtuePresent(weElement, strAssignedAttribute); // Reuse internal helper
	if (boolAttributePresent === true) {
		strTestObjText = StrNums.JComm_HandleNoData(weElement.getAttribute("value")); // Assumes getAttribute reads value sync
		if (StrNums.JComm_HandleNoData(strAssignedValue) === strTestObjText) {
			strMethodDetails = 'The ' + strTestObjType + ' ' + strObjName + ' text values was set and matched the specified value of: ' + strAssignedValue + '.';
		} else {
			boolPassed = false;
			strMethodDetails = 'FAILED!!! The ' + strTestObjType + " " + strObjName + ' text of: ' + strTestObjText +
				' was set, but DOES NOT MATCH the specified value of: ' + strAssignedValue + '!!!';
		}
	}
	else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! The ' + strTestObjType + " " + strObjName + " DOES NOT HAVE THE ATTRIBUTE: " + strAssignedAttribute + "!!!";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * -------------------------------------  objVerifyEditBoxValue  -----------------------------------
 * Verify the edit box value.
 * @param weEditBox The Editbox WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strAssignedValue The assigned text value for the test object
 * @return mapResults
 *  1. boolPassed The method passed?
 *  2. strMethodDetails The details on the processing of the method
 * @author pkanaris
 * @author Created: 04/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objVerifyEditBoxValue(weEditBox, strObjName, strAssignedValue) {
	//GlobalVars
	const gblNull = GVars.GblNull('Value');
	// Declare the variables
	const strTestObjType = 'EditBox';
	const mapResults = {};
	let strMethodDetails;
	let boolPassed = true;
	let strTestObjText;
	//Retrieve the new value and verify the value matches the assignedValue
	strTestObjText = StringsAndNumbers.JComm_HandleNoData(SeS(weEditBox).GetValue())
	if (StringsAndNumbers.JComm_HandleNoData(strAssignedValue) === strTestObjText) {
		strMethodDetails = 'The ' + strTestObjType + ' ' + strObjName + ' text values was set and matched the specified value of: ' + strAssignedValue + '.';
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! The ' + strTestObjType + " " + strObjName + ' text of: ' + strTestObjText +
			' was set, but DOES NOT MATCH the specified value of: ' + strAssignedValue + '!!!';
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * -------------------------------------  objVerifyTextBoxValue  -----------------------------------
 * Verify the edit box value (specifically for multi-line text boxes, using getText()).
 * @param weEditBox The Editbox WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strAssignedValue The assigned text value for the test object
 * @return mapResults
 *  1. boolPassed The method passed?
 *  2. strMethodDetails The details on the processing of the method
 * @author pkanaris
 * @author Created: 04/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objVerifyTextBoxValue(weEditBox, strObjName, strAssignedValue) {
	//GlobalVars
	const gblNull = GVars.GblNull('Value');
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblHTMLLF = GVars.GblHTMLLineFeed('Value');
	// Declare the variables
	const strTestObjType = 'EditBox';
	const mapResults = {};

	let strMethodDetails;
	let boolTextMatched = true;
	let boolPassed = true;
	let strTestObjText; // Will be assigned later
	let strTempErrorText = ''; // Initialize error text

	//Retrieve the new value and verify the value matches the assignedValue
	strTestObjText = StrNums.JComm_HandleNoData(weEditBox.getText()); // Assumes getText() reads text sync
	//Split the assigned and the returned text
	const intGblLF = StrNums.JComm_TextLocationInString(strAssignedValue, gblLineFeed);
	const intGBLHTMLLF = StrNums.JComm_TextLocationInString(strAssignedValue, gblHTMLLF);
	let arrayInputValues; // Placeholder for Array
	let intArrayInputCnt = 0;
	if (intGblLF > 0) {
		//Split the input value to begin the processing
		const mapItemsValues = StrNums.JComm_StringToArray(strAssignedValue, gblLineFeed);
		arrayInputValues = mapItemsValues.ArryOfValues;
		intArrayInputCnt = arrayInputValues.length;
	}
	else if (intGBLHTMLLF > 0) {
		//Split the input value to begin the processing
		const mapItemsValues = StrNums.JComm_StringToArray(strAssignedValue, gblHTMLLF);
		arrayInputValues = mapItemsValues.ArryOfValues;
		intArrayInputCnt = arrayInputValues.length;
	}
	//Split the input value to begin the processing '\n' is the value returned by get text for line breaks.
	const mapItemsTextValues = StrNums.JComm_StringToArray(strTestObjText, '\n');
	const arrayDisplayedValues = mapItemsTextValues.ArryOfValues;
	const intArrayTextCnt = arrayDisplayedValues.length;
	if (intArrayInputCnt === intArrayTextCnt) {
		let strTempInput, strTempDisplayed;
		//Process the arrays and verify they match
		for (let loopText = 0; loopText < intArrayInputCnt; loopText++) {
			strTempInput = arrayInputValues[loopText];
			strTempDisplayed = arrayDisplayedValues[loopText];
			if (strTempInput !== strTempDisplayed) {
				strTempErrorText = strTempErrorText + "MISMATCHED TEXT ITEM: " + loopText + "Displayed: " +
					strTempDisplayed + " Expected: " + strTempInput + gblHTMLLF;
				boolTextMatched = false;
				break;
			}
		}
	}
	else {
		boolTextMatched = false;
	}
	if (boolTextMatched === true) {
		strMethodDetails = 'The ' + strTestObjType + ' ' + strObjName + ' text values was set and matched the specified value of: ' + strAssignedValue + '.';
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! The ' + strTestObjType + " " + strObjName + ' text of: ' + strTestObjText +
			' was set, but DOES NOT MATCH the specified value of: ' + strAssignedValue + '!!! See details below: ' + gblHTMLLF + strTempErrorText;
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

//** MASKEDIT **/
/**
 * -------------------------------------  objSetMaskEditBoxValue  -----------------------------------
 * Set the masked edit value. Do not report the actual text in the field.
 * @param weEditBox The Editbox WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strAssignedValue The assigned text value for the test object
 * @param boolClearValue Clear the value prior to setting? True/False
 * @param boolInputEncrypted Is the input value encrypted? True/False (Not implemented logic)
 * @param boolVerifyValue Verify the value after set? True/False
 * @return mapResults
 *  1. boolPassed The method passed?
 *  2. strMethodDetails The details on the processing of the method
 * @author pkanaris
 * @author Created: 04/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSetMaskEditBoxValue(weEditBox, strObjName, strAssignedValue, boolClearValue, boolInputEncrypted, boolVerifyValue) {
	//GlobalVars
	const gblNull = GblNull;
	// Declare the variables
	const strTestObjType = 'EditBox';
	const mapResults = {};
	let strMethodDetails;
	let boolPassed = true;
	let strTestObjText; // Will be assigned later
	//Process the edit box
	if (boolClearValue === true) {
		//Clear the text first
		Tester.SuppressReport(true); 
		weEditBox.DoSetText("");
		Tester.SuppressReport(false);
		DateTime.WaitSecs(intViewDelaySecs);
		//Clear does not appear to work on iOS devices
		//Check the text and if not clear click after the text and type backspace/delete.
		strTestObjText = StringsAndNumbers.JComm_HandleNoData(SeS(weEditBox).GetValue());
		if (strTestObjText !== gblNull) {
			const lenStr = strTestObjText.length;
			for (let loop = 0; loop < lenStr; loop++) {
				strTestObjText = StringsAndNumbers.JComm_HandleNoData(SeS(weEditBox).GetValue());
				if (strTestObjText === gblNull) {
					break; //The value is cleared
				} else {
					Tester.SuppressReport(true);
					weEditBox.DoSendKeys("{BACKSPACE}");
					Tester.SuppressReport(false);
					blockingSleep(100);
				}
			}
		}
		//Alternate will be change the text and trim off the https:// which already present
	}
	//Set the value when encrypted
	let strTempPasswordValue;
	if (boolInputEncrypted === true) {
		//Assign the value to strTempValue. NOTE: the decrypted value is available in the code
		strTempPasswordValue = Global.DoDecrypt(strAssignedValue);
	}
	else {
		strTempPasswordValue = strAssignedValue;
	}
	//Set the value
	let strTempValue;
	if (StringsAndNumbers.JComm_HandleNoData(strTempPasswordValue) === gblNull) {
		strTempValue = '';
	} else {
		strTempValue = StringsAndNumbers.JComm_HandleNoData(strTempPasswordValue);
	}
	Tester.SuppressReport(true);
	weEditBox.DoSendKeys(strTempValue);
	DateTime.WaitSecs(intViewDelaySecs);
	weEditBox.DoSendKeys("{TAB}");
	Tester.SuppressReport(false);
	//WebDriver.Actions().SendKeys(String.fromCharCode(0xE004),weEditBox.element).Perform();
	DateTime.WaitSecs(intViewDelaySecs);
	if (boolVerifyValue === true) {
		//Retrieve the new value and verify the value matches the assignedValue
		strTestObjText = StringsAndNumbers.JComm_HandleNoData(SeS(weEditBox).GetValue());
		if (strTempValue === strTestObjText) {
			strMethodDetails = "The " + strTestObjType + " " + strObjName + " text values was set and matched the specified value and is displayed as encyrpted text.";
		} else {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The " + strTestObjType + " " + strObjName + " text of: " + strTestObjText +
				" was set, but DOES NOT MATCH the specified value of: " + strTempValue + "!!!";
		}
	} else {
		strMethodDetails = "The " + strTestObjType + " " + strObjName + " text of: " + strTestObjText +
			" was set, to the specified value of: " + strAssignedValue + ".";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Should ensure strMethodDetails is set before returning
	return mapResults;
}

/* MSGBOX */
/**
 * -------------------------------------  objVerifyMsgBoxValue  -----------------------------------
 * Verify the value displayed in the message box text.
 * @param weMsgBox The message box WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strAssignedValue The assigned text value for the test object
 * @return mapResults
 *  1. boolPassed The method passed?
 *  2. strMethodDetails The details on the processing of the method
 * @author pkanaris
 * @author Created: 04/23/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objVerifyMsgBoxValue(weMsgBox, strObjName, strAssignedValue) {
	//GlobalVars
	const gblNull = GblNull;
	// Declare the variables
	const strTestObjType = 'msgBox';
	const mapResults = {};
	let strMethodDetails;
	let boolPassed = true;
	let strTestObjText;
	//Retrieve the new value and verify the value matches the assignedValue
	strTestObjText = StringsAndNumbers.JComm_StripLeadingAssignChars(StringsAndNumbers.JComm_HandleNoData(SeS(weMsgBox).GetText().trim()))// Assumes getText() reads text sync
	if (StringsAndNumbers.JComm_HandleNoData(strAssignedValue) === strTestObjText) {
		strMethodDetails = 'The ' + strTestObjType + ' ' + strObjName + ' text values matched the specified value of: ' + strAssignedValue;
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! The ' + strTestObjType + " " + strObjName + ' text of: ' + strTestObjText +
			' DOES NOT MATCH the specified value of: ' + strAssignedValue + '!!!';
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

//**STANDARD CHECKBOX **/
/**
 * -------------------------------------  getWebElementCheckboxChecked  -----------------------------------
 * Return the checked state of the assigned checkbox element.
 * @param weCheckBox The checkbox WebElement (placeholder for WebElement)
 * @return boolIsChecked The current checked state
 * @author pkanaris
 * @author Created: 12/31/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objGetCheckboxChecked(weCheckbox) {
	let boolIsChecked;
	boolIsChecked = weCheckbox.isSelected(); // Assumes isSelected() reads state sync
	return boolIsChecked;
}

/**
 * -------------------------------------  verifyWebElementCheckBoxChecked  -----------------------------------
 * Return the checked state of the assigned checkbox element.
 * @param weCheckBox The checkbox WebElement (placeholder for WebElement)
 * @return boolPassed The current checked state matches the expected?
 * @author pkanaris
 * @author Created: 12/31/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objVerifyCheckBoxChecked(weCheckbox, boolChecked) {
	let boolPassed = true;
	if (objGetCheckboxChecked(weCheckbox) !== boolChecked) { // Reuse internal helper
		boolPassed = false;
	}
	//report the details if verify fails
	return boolPassed;
}

/**
 * -------------------------------------  objSetCheckBox  -----------------------------------
 * Set and checkbox to the assigned value.
 * @param weCheckBox The checkbox WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param boolIsChecked The assigned checked state
 * @return mapResults
 *  1. strMethodDetails The details on the processing of the method
 * 2. webElement The checkbox after state has changed.
 * @author pkanaris
 * @author Created: 12/31/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSetCheckBox(weCheckBox, strObjName, boolIsChecked) {
	//GlobalVars (not used here)
	const mapResults = {};
	let strMethodDetails;
	//Process the check box
	//Return the current state
	const currentChecked = objGetCheckboxChecked(weCheckBox); // Reuse internal helper
	//Click if not in the correct state
	if (currentChecked !== boolIsChecked) {
		weCheckBox.click();
		blockingSleep(100);
		strMethodDetails = "Clicked the '" + strObjName + "' checkbox to set the state to: '" + boolIsChecked + "'.";
	}
	else {
		strMethodDetails = "The '" + strObjName + "' checkbox was already set the state to: '" + boolIsChecked + "'.";
	}
	//Return the web element since it will be modified and the prior state would be a stale object
	//update the map
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.webElement = weCheckBox; // Return the same WebElement reference
	return mapResults;
}

/**
 * -------------------------------------  objSetCheckBoxJavaScript  -----------------------------------
 * Set the checkbox to the assigned checked state using Javascript.
 * @param weCheckBox The checkbox WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param boolIsChecked The assigned checked state
 * @return mapResults
 *  1. strMethodDetails The details on the processing of the method
 * 2. webElement The checkbox after state has changed.
 * @author pkanaris
 * @author Created: 12/31/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSetCheckBoxJavaScript(weCheckBox, strObjName, boolIsChecked) {
	//GlobalVars (not used here)
	const mapResults = {};
	let strMethodDetails;
	//Process the check box
	//Return the current state
	const currentChecked = objGetCheckboxChecked(weCheckBox); // Reuse internal helper
	//Click if not in the correct state
	if (objGetCheckboxChecked(weCheckBox) !== boolIsChecked) { // Reuse internal helper
		const driver = TestCaseObject.tcDriver; // Assumes TestCaseObject.tcDriver returns a WebDriver instance
		driver.executeScript("arguments[0].click();", weCheckBox);
		blockingSleep(100);
		strMethodDetails = "Clicked the '" + strObjName + "' checkbox to set the state to: '" + boolIsChecked + "'.";
	}
	else {
		strMethodDetails = "The '" + strObjName + "' checkbox was already set the state to: '" + boolIsChecked + "'.";
	}
	//Return the web element since it will be modified and the prior state would be a stale object
	//update the map
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.webElement = weCheckBox; // Return the same WebElement reference
	return mapResults;
}

/**
 * -------------------------------------  objSetGraphicCheckBox  -----------------------------------
 * Set and checkbox to the assigned value.
 * @param weCheckSelector The webelement that must be clicked to change the state (placeholder for WebElement)
 * @param weCheckBox The checkbox element that will reflect the checked state (placeholder for WebElement)
 * @param strAssocText The text associated if any to the checkbox
 * @param boolIsChecked The assigned checked state
 * @return mapResults
 *  1. strMethodDetails The details on the processing of the method
 * 2. webElement The checkbox after state has changed.
 * @author pkanaris
 * @author Created: 03/12/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSetGraphicCheckBox(weCheckSelector, weCheckBox, strAssocText, boolIsChecked) {
	//GlobalVars (not used here)
	const mapResults = {};
	let strMethodDetails;
	//Process the check box
	//Return the current state
	const currentChecked = objGetCheckboxChecked(weCheckBox); // Reuse internal helper
	//Click if not in the correct state
	if (objGetCheckboxChecked(weCheckBox) !== boolIsChecked) { // Reuse internal helper
		weCheckSelector.click();
		blockingSleep(100);
		strMethodDetails = "Clicked the '" + strAssocText + "' checkbox to set the state to: '" + boolIsChecked + "'.";
	}
	else {
		strMethodDetails = "The '" + strAssocText + "' checkbox was already set the state to: '" + boolIsChecked + "'.";
	}
	//Return the web element since it will be modified and the prior state would be a stale object
	//update the map
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.webElement = weCheckBox; // Return the same WebElement reference
	return mapResults;
}

//Radio Group
/**
 * -------------------------------------  objSetRadioGroupOption  -----------------------------------
 * Select the Radio Group option specified.
 * @param weRadioGroup The Radio Group WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strRadioOptXpath The XPath for the radio button/option elements
 * @param strRadioGrpXPath The XPath for the group itself (used for re-finding elements from entire group)
 * @param strRadioBtnXPath The XPath for the actual radio button element within item
 * @param strRadioOptext The text of the assigned radio button/option
 * @param boolVerifyValue Verify the value after set? True/False
 * @return mapResults
 *  1. boolMethodPassed The method passed? True/False
 * 2. strMethodDetails The details on the processing of the method
 * @author pkanaris
 * @author Created: 02/12/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSetRadioGroupOption(weRadioGroup, strObjName, strRadioOptXpath, strRadioGrpXPath, strRadioBtnXPath, strRadioOptext, boolVerifyValue) {
	//GlobalVars
	const gblNull = GVars.GblNull('Value');
	strRadioOptext = StrNums.JComm_HandleNoData(strRadioOptext);
	// Declare the variables
	const strTestObjType = 'RadioGroup';
	const mapResults = {};
	let boolMethodPassed = true;
	let strMethodDetails; // Will be assigned later
	//Process the radio group
	//Return the child items which are the available radio buttons/options
	//Return the option list
	const lstChildObjects = weRadioGroup.findElements(By.xpath(strRadioOptXpath)); // Assuming findElements returns array sync
	const intCntChildObj = lstChildObjects.length;
	if (intCntChildObj > 0) {
		//Process the child objects
		let weRadioOpt; // Placeholder for WebElement
		let strOptInstanceText;
		let strOptItemsText = ''; // Initialize to empty string
		let boolRadioOptionFound = false;
		for (let intItemInstance = 0; intItemInstance < intCntChildObj; intItemInstance++) {
			weRadioOpt = lstChildObjects[intItemInstance];
			//Highlight
			if (TSExecParams.getBoolDoHighlight() === true) {
				const mapHighlight = objHighlightElementJS(weRadioOpt, 'RadioOption', strObjName + ' Radio Group'); // Reuse internal helper
			}
			strOptInstanceText = weRadioOpt.getText(); // Assuming getText() reads text sync
			if (intItemInstance === 0) {
				strOptItemsText = strOptInstanceText;
			}
			else {
				strOptItemsText = strOptItemsText + '; ' + strOptInstanceText;
			}
			//Check if the text matches the specified button text
			if (strRadioOptext === strOptInstanceText) {
				boolRadioOptionFound = true;
				try {
					weRadioOpt.click();
				}
				catch (e) {
					strMethodDetails = (strMethodDetails || '') + "FAILED TO MAKE SELECTION BY INDEX!!! " + 'Exception occurred!!! SEE ERROR STACK TRACE: ' +
						(e instanceof Error ? e.stack : String(e));
				}
				//wait after the selection
				DateTime.WaitSecs(TSExecParams.getIntViewDelaySecs()); // Assumes blocking wait
				//Verify the radio button is selected
				if (boolVerifyValue === true) {
					//Return the element
					const strTmpRadioOpt = strRadioBtnXPath.replace("./", "//") + "[" + (intItemInstance + 1) + "]";
					const weRadioOptBtn = returnWebElement(strRadioGrpXPath + strTmpRadioOpt); // Reuse internal helper
					if (weRadioOptBtn !== null && weRadioOptBtn !== undefined) {
						//Highlight
						if (TSExecParams.getBoolDoHighlight() === true) {
							const mapHighlight = objHighlightElementJS(weRadioOptBtn, 'RadioOptionButton', strObjName + ' Radio Group'); // Reuse internal helper
						}
						const boolSelected = weRadioOptBtn.getAttribute("selected"); // Assumes getAttribute reads sync
						if (boolSelected === true) {
							strMethodDetails = "Successfully selected the radio option: " + strRadioOptext +
								" from the Radio Group '" + strObjName + "'.";
						}
						else {
							boolMethodPassed = false;
							strMethodDetails = "FAILED!!! to SELECT the radio option: " + strRadioOptext +
								" from the Radio Group '" + strObjName + "'!!!";
						}
					}
					else {
						boolMethodPassed = false;
						strMethodDetails = "FAILED!!! the radio option CHILD RADIO BUTTON WAS NOT FOUND!!!";
					}
				}
				break; //Exit the for loop for finding the item in the options
			}
		}
		if (boolRadioOptionFound === false) {
			boolMethodPassed = false;
			strMethodDetails = "FAILED!!! the element '" + strObjName + "' DOES NOT CONTAIN THE RADIO BUTTONS/OPTION MATCHING THE ASSIGNED VALUE OF: " +
				strRadioOptext + "!!!";
		}
	}
	else {
		boolMethodPassed = false;
		strMethodDetails = "FAILED!!! the element '" + strObjName + "' DOES NOT HAVE ANY CHILD RADIO BUTTONS/OPTIONS!!!";
	}

	//Update the map
	mapResults.boolMethodPassed = boolMethodPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Should ensure strMethodDetails is set before returning
	return mapResults;
}

/**
 * -------------------------------------  objSetRadioGroupOptionJavaScript  -----------------------------------
 * Select the Radio Group option specified. The Label contains ::before and ::after and must be clicked using Java Script
 * @param weOptionParent The Radio Group parent WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strRadioOptXpath The XPath for the radio button/option elements (e.g., label text)
 * @param strRadioGrpXPath The XPath for the radio group itself (not directly used in function, but from signature)
 * @param strRadioBtnXPath The XPath for the actual clickable radio button element (e.g., input)
 * @param strRadioSelXPath The XPath for verification (selected state)
 * @param strRadioOptext The text of the assigned radio button/option
 * @param boolVerifyValue Verify the value after set? True/False
 * @return mapResults
 *  1. boolMethodPassed The method passed? True/False
 * 2. strMethodDetails The details on the processing of the method
 * @author pkanaris
 * @author Created: 02/12/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSetRadioGroupOptionJavaScript(weOptionParent, strObjName, strRadioOptXpath, strRadioGrpXPath, strRadioBtnXPath, strRadioSelXPath, strRadioOptext, boolVerifyValue) {
	//GlobalVars
	const gblNull = GVars.GblNull('Value');
	// Declare the variables
	const strTestObjType = 'RadioGroup';
	const mapResults = {};
	let boolMethodPassed = true;
	let strMethodDetails; // Will be assigned later
	let intItemInstance;
	//Process the radio group
	//Return the child items which are the available radio buttons/options
	//Return the option list (labels/text)
	const lstChildRDOObjects = weOptionParent.findElements(By.xpath(strRadioOptXpath)); // Assumes findElements returns array sync
	const intCntRDOChildObj = lstChildRDOObjects.length;
	// Return the actual button elements
	const lstChildRDOBtnsObjects = weOptionParent.findElements(By.xpath(strRadioBtnXPath)); // Assumes findElements returns array sync
	const intCntRDOBtnsChildObj = lstChildRDOBtnsObjects.length;
	if (intCntRDOChildObj > 0) {
		//Process the child objects
		let weRadioOpt; // Placeholder for WebElement (label/text)
		let weRadioOptBtn; // Placeholder for WebElement (button)
		let strOptInstanceText;
		let strOptItemsText = ''; // Initialize to empty string
		let boolRadioOptionFound = false;
		const driver = TestCaseObject.tcDriver; // Assumes WebDriver instance
		for (intItemInstance = 0; intItemInstance < intCntRDOChildObj; intItemInstance++) {
			weRadioOpt = lstChildRDOObjects[intItemInstance];
			weRadioOptBtn = lstChildRDOBtnsObjects[intItemInstance];
			try {
				//Highlight
				if (TSExecParams.getBoolDoHighlight() === true) {
					const mapHighlight = objHighlightElementJS(weRadioOpt, 'RadioOption', strObjName + ' Radio Group'); // Reuse internal helper
				}
				strOptInstanceText = weRadioOpt.getText(); // Assumes getText() reads text sync
			}
			catch (e) {
				const strException = (e instanceof Error ? e.stack : String(e));
				console.log("RDO Option raised an Exception: " + strException);
			}
			if (StrNums.JComm_HandleNoData(strOptItemsText) === gblNull) { // This check seems flawed - checking strOptItemsText (accumulator) not strOptInstanceText
				//Try to return the option button text since we appear to not find one for the option element
				try {
					//Highlight
					if (TSExecParams.getBoolDoHighlight() === true) {
						const mapHighlight = objHighlightElementJS(weRadioOptBtn, 'RadioOptionBtn', strObjName + ' Radio Group'); // Reuse internal helper
					}
					strOptInstanceText = weRadioOptBtn.getText(); // Assumes getText() reads text sync
				}
				catch (e) {
					const strException = (e instanceof Error ? e.stack : String(e));
					console.log("RDO Option raised an Exception: " + strException);
				}
			}
			if (intItemInstance === 0) {
				strOptItemsText = strOptInstanceText;
			}
			else {
				strOptItemsText = strOptItemsText + '; ' + strOptInstanceText;
			}
			//Check if the text matches the specified button text
			if (strRadioOptext === strOptInstanceText) {
				boolRadioOptionFound = true;
				try {
					driver.executeScript("arguments[0].click();", weRadioOptBtn);
					blockingSleep(100);
				}
				catch (e) {
					strMethodDetails = (strMethodDetails || '') + "FAILED TO MAKE SELECTION BY INDEX!!! " + 'Exception occurred!!! SEE ERROR STACK TRACE: ' +
						(e instanceof Error ? e.stack : String(e));
				}
				//wait after the selection
				DateTime.WaitSecs(TSExecParams.getIntViewDelaySecs()); // Assumes blocking wait
				//Verify the radio button is selected
				if (boolVerifyValue === true) {
					//Return the element
					const lstChildVerifyOptSel = weOptionParent.findElements(By.xpath(strRadioSelXPath)); // Assumes findElements returns array sync
					const intVerifyCntOptions = lstChildVerifyOptSel.length;
					const weRadioOptSel = lstChildVerifyOptSel[intItemInstance];
					if (weRadioOptSel !== null && weRadioOptSel !== undefined) {
						//Highlight
						if (TSExecParams.getBoolDoHighlight() === true) {
							const mapHighlight = objHighlightElementJS(weRadioOptSel, 'RadioOptionSelection', strObjName + ' Radio Group'); // Reuse internal helper
						}
						const boolSelected = weRadioOpt.isSelected(); // Assumes isSelected() reads state sync
						/**
						//TODO Return the item class and check if contains selected
						String strRDOClass = StrNums.JComm_HandleNoData(weRadioOptSel.getAttribute("class"))
						//Check if selected is in the class name
						boolSelected = StrNums.JComm_VerifyTextPresent(strRDOClass, "selected", '%Con%')
						Not used for this method. If we need to implement we will need to do something different .*/
						if (boolSelected === true) {
							strMethodDetails = "Successfully selected the radio option: " + strRadioOptext +
								" from the Radio Group '" + strObjName + "'.";
						}
						else {
							boolMethodPassed = false;
							strMethodDetails = "FAILED!!! to SELECT the radio option: " + strRadioOptext +
								" from the Radio Group '" + strObjName + "'!!!";
						}
					}
					else {
						boolMethodPassed = false;
						strMethodDetails = "FAILED!!! the radio option CHILD RADIO BUTTON WAS NOT FOUND!!!";
					}
				}
				break; //Exit the for loop for finding the item in the options
			}
		}
		if (boolRadioOptionFound === false) {
			boolMethodPassed = false;
			strMethodDetails = "FAILED!!! the element '" + strObjName + "' DOES NOT CONTAIN THE RADIO BUTTONS/OPTION MATCHING THE ASSIGNED VALUE OF: "
				+ strRadioOptext + "!!!";
		}
	}
	else {
		boolMethodPassed = false;
		strMethodDetails = "FAILED!!! the element '" + strObjName + "' DOES NOT HAVE ANY CHILD RADIO BUTTONS/OPTIONS!!!";
	}

	//Update the map
	mapResults.boolMethodPassed = boolMethodPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure strMethodDetails is set before returning
	return mapResults;
}

//**STANDARD Single Select Combo/Dropdown Box **/
/**
 * -------------------------------------  objSelectPickListItemByOption  -----------------------------------
 * Select the specified item from the list by available options.
 * @param weDropDown The dropdown WebElement (placeholder for WebElement)
 * @param strDropDownXPath The full xpath to the dropdown (used for re-finding it)
 * @param strObjName The meaningful name of the testobject
 * @param strItemText The item text to be selected
 * @param boolMatchText Must the item text must be exact match? true/false
 * @param boolByIndex Select the item from the options by index? true/false (Not used in selection logic)
 * @return mapResults
 *  1. boolMethodPassed The pass/fail status. true/false
 *  2. strMethodDetails The details on the processing of the method
 *  3. weNewDropDown The listbox after state has changed.
 * @author pkanaris
 * @author Created: 12/31/2021
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSelectPickListItemByOption(weDropDown, strDropDownXPath, strObjName, strItemText, boolMatchText, boolByIndex) { // boolByIndex is not used in actual select logic
	//Create variables
	const gblLineFeed = GVars.GblLineFeed('Value');
	let boolMethodPassed = true;
	let strMethodDetails; // Will be assigned later
	const mapMethodResults = {};
	let objDropdown = null;
	try {
		objDropdown = new Select(weDropDown); // Selenium's Select class
	} catch (e) {
		// If Select constructor fails (e.g., not a <select> element)
		boolMethodPassed = false;
		strMethodDetails = "FAILED!!!! the '" + strObjName + "' did not return a Select Object!!! Error: " + (e instanceof Error ? e.message : String(e));
	}

	if (objDropdown !== null && boolMethodPassed) { // Only proceed if Select object was successfully created
		//Process the object
		let boolValueFound = false;
		let boolOptionSelected = false; // Initialize to false
		//Return the option list
		const drpDownOptions = objDropdown.getOptions(); // Assumes getOptions() returns array sync
		//Count the options
		const intCntOptions = drpDownOptions.length;
		if (intCntOptions > 0) {
			//Process the values
			let weOption; // Placeholder for WebElement
			let strOptInstanceText;
			let strOptItemsText = ''; // Initialize to empty string
			for (let intItemInstance = 0; intItemInstance < intCntOptions; intItemInstance++) {
				weOption = drpDownOptions[intItemInstance];
				strOptInstanceText = StrNums.JComm_HandleNoData(weOption.getText()); // Assumes getText() reads text sync
				if (intItemInstance === 0) {
					strOptItemsText = strOptInstanceText;
				}
				else {
					strOptItemsText = strOptItemsText + '; ' + strOptInstanceText;
				}
				//Check if the value is present
				if (boolMatchText === true) {
					if (strOptInstanceText === strItemText) {
						boolValueFound = true;
					}
				} else {
					//Check if strAssignedValue is contained in the value
					boolValueFound = StrNums.JComm_VerifyTextPresent(strOptInstanceText, strItemText, '%Con%');
				}
				if (boolValueFound === true) {
					boolOptionSelected = true;
					//Are we selecting by index? (Original `boolByIndex` is not used in actual select logic for Selenium `Select` class.
					// Instead, it uses `selectByIndex` or `selectByVisibleText`. The original code just has `selectByVisibleText`).
					// The `boolByIndex` parameter in the argument list implies it should be used for index selection.
					// Adhering to the original Groovy, it always uses `selectByVisibleText` or `selectByIndex` if `boolByIndex` is used.
					// Given the Groovy's switch, it implies `boolByIndex` is to *control* the selection mechanism.
					// However, the Groovy implementation always calls selectByVisibleText or selectByIndex based on the parameter boolByIndex.
					// Replicated original error handling with direct calls to Selenium Select methods.
					try {
						if (boolByIndex === true) { // Use index for selection
						   objDropdown.selectByIndex(intItemInstance);
						} else { // Default to visible text
						   objDropdown.selectByVisibleText(strItemText);
						}
					}
					catch (e) {
						 strMethodDetails = (strMethodDetails || '') + "FAILED TO MAKE SELECTION!!! " + 'Exception occurred!!! SEE ERROR STACK TRACE: ' +
							(e instanceof Error ? e.stack : String(e));
					}

					//wait after the selection
					DateTime.WaitSecs(TSExecParams.getIntViewDelaySecs()); // Assumes blocking wait
					//reset the weDropdown
					const weNewDropDown = returnWebElement(strDropDownXPath); // Reuse internal helper (re-find element)
					mapMethodResults.objNewDropBox = weNewDropDown; // Add to results map
					break; //Exit the for loop for finding the item in the options
				}
			}
			if (boolOptionSelected === true) {
				strMethodDetails = "Selected the value of: " + strItemText + " from the dropdown '" + strObjName + ".";
			} else {
				boolMethodPassed = false;
				strMethodDetails = "FAILED!!! The dropdown element : '" + strObjName + "' DOES NOT contain the value: " + strItemText + gblLineFeed;
				strMethodResults = strMethodDetails + "Checked ALL " + intCntOptions + " options text consisting of: " + strOptItemsText + "!!!"; // Use strMethodResults as the Groovy does
			}
		}
		else {
			boolMethodPassed = false;
			strMethodDetails = 'FAILED!!! The dropdown element : ' + strObjName + ' CONTAINS ZERO OPTIONS!!!';
		}
	}
	mapMethodResults.boolPassed = boolMethodPassed.toString();
	mapMethodResults.strMethodDetails = strMethodDetails; // Ensure strMethodDetails is set before returning
	return mapMethodResults;
}

/**
 * -------------------------------------  objSelectVerifyItemSelectedByOption  -----------------------------------
 * Verify the value selected in the dropdown by the select option.
 * @param weDropDown The dropdown WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strItemText The item text to be selected
 * @param boolMatchText Must the item text must be exact match? true/false
 * @return mapResults
 *  1. boolMethodPassed The pass/fail status. true/false
 *  2. strMethodDetails The details on the processing of the method
 * @author pkanaris
 * @author Created: 01/19/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSelectVerifyItemSelectedByOption(weDropDown, strObjName, strItemText, boolMatchText) {
	//Create variables
	const gblLineFeed = GVars.GblLineFeed('Value'); // Not used
	let boolMethodPassed = true;
	let strMethodDetails; // Will be assigned later
	let objDropdown = null;
	try {
		objDropdown = new Select(weDropDown); // Selenium's Select class
	} catch (e) {
		boolMethodPassed = false;
		strMethodDetails = "FAILED!!!! the '" + strObjName + "' did not return a Select Object!!! Error: " + (e instanceof Error ? e.message : String(e));
	}
	const mapMethodResults = {};

	if (objDropdown !== null && boolMethodPassed) { // Only proceed if Select object was successfully created
		//Process the object
		const objOption = objDropdown.getFirstSelectedOption(); // Assumes getFirstSelectedOption() returns WebElement sync
		const strCurrentSelected = StrNums.JComm_HandleNoData(objOption.getText()); // Assumes getText() reads text sync
		//Check if the value is correct
		if (boolMatchText === true) {
			if (strCurrentSelected === strItemText) {
				boolMethodPassed = true;
			} else {
				boolMethodPassed = false; // Add explicit fail for strict match if no match
			}
		} else {
			//Check if strAssignedValue is contained in the value
			boolMethodPassed = StrNums.JComm_VerifyTextPresent(strCurrentSelected, strItemText, '%Con%');
		}
		if (boolMethodPassed === true) {
			strMethodDetails = "The select '" + strObjName + "' is displaying the value of: '" + strItemText + "' as expected.";
		}
		else {
			strMethodDetails = "FAILED The select '" + strObjName + "' is displaying the value of: '" + strCurrentSelected +
				"' which DOES NOT MATCH the value expected of: '" + strItemText + "'!!!";
		}
	}
	mapMethodResults.boolPassed = boolMethodPassed.toString();
	mapMethodResults.strMethodDetails = strMethodDetails; // Ensure strMethodDetails is set before returning
	return mapMethodResults;
}

//** List Box **/
/**
 * -------------------------------------  objSelectListBoxItemByChild  -----------------------------------
 * Select the specified item from the list box by available child elements XPath.
 * @param weListBox The list box WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strItemXPath The XPath for the child items (e.g., li, option)
 * @param strItemText The item text to be selected
 * @return mapResults
 *  1. boolMethodPassed The pass/fail status. true/false
 *  2. strMethodDetails The details on the processing of the method
 *  3. weNewDropDown The listbox after state has changed.
 * @author pkanaris
 * @author Created: 02/03/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSelectListBoxItemByChild(weListBox, strObjName, strItemXPath, strItemText) {
	//Create variables
	const gblLineFeed = GVars.GblLineFeed('Value');
	let boolMethodPassed = true;
	let strMethodDetails; // Will be assigned later
	const mapMethodResults = {};
	//Process the object
	let boolValueFound = false;
	let lstChildObjects; // Placeholder for Array<WebElement>
	let boolOptionSelected = false; // Initialize to false
	let intCntOptions;

	//Check to see if items are returned and use a timer to retry
	// The original loop condition `retrycnt = 1` means it will always re-assign to 1.
	// It should be `retrycnt > 0` or similar for proper retry logic.
	// Assuming it's meant to retry `5` times, `for (let retrycnt = 0; retrycnt < 5; retrycnt++)`
	// but replicating the original structure with `while (intCntOptions === 0 && retrycnt < MAX_RETRIES)`
	for (let retrycnt = 5; retrycnt >= 1; retrycnt--) { // Adjusted logic for a standard retry loop
		//Return the option list
		lstChildObjects = weListBox.findElements(By.xpath(strItemXPath)); // Assuming findElements returns array sync
		//Count the options
		intCntOptions = lstChildObjects.length;
		if (intCntOptions > 0) {
			break;
		}
		else {
			blockingSleep(100); // Sleep 100ms between retries
		}
	}
	if (intCntOptions > 0) {
		//Process the values
		let weListItem; // Placeholder for WebElement
		let strOptInstanceText;
		let strOptItemsText = ''; // Initialize to empty string
		for (let intItemInstance = 0; intItemInstance < intCntOptions; intItemInstance++) {
			weListItem = lstChildObjects[intItemInstance];
			//Highlight
			if (TSExecParams.getBoolDoHighlight() === true) {
				const mapHighlight = objHighlightElementJS(weListItem, 'ListItem', strObjName + ' List Item'); // Reuse internal helper
			}
			strOptInstanceText = StrNums.JComm_HandleNoData(StrNums.JComm_ReplaceValue(weListItem.getText(), '\n', ' ')); // Assumes getText() reads text sync
			if (intItemInstance === 0) {
				strOptItemsText = strOptInstanceText;
			}
			else {
				strOptItemsText = strOptItemsText + '; ' + strOptInstanceText;
			}
			//Check if the value is present
			if (strOptInstanceText === strItemText) {
				boolValueFound = true;
			}
			if (boolValueFound === true) {
				boolOptionSelected = true;
				try {
					weListItem.click();
				}
				catch (e) {
					// Original: ElementClickInterceptedException then generic e
					if (e && e.name === 'ElementClickInterceptedError') { // Catch specific Selenium error
						const driver = TestCaseObject.tcDriver; // Assumes WebDriver instance
						driver.executeScript("arguments[0].click();", weListItem); // Attempt click with JS
					} else {
						strMethodDetails = (strMethodDetails || '') + "FAILED TO MAKE SELECTION Exception occurred!!! SEE ERROR STACK TRACE: " +
							(e instanceof Error ? e.stack : String(e));
					}
				}
				//wait after the selection
				DateTime.WaitSecs(TSExecParams.getIntViewDelaySecs()); // Assumes blocking wait
				break; //Exit the for loop for finding the item in the options
			}
		}
		if (boolOptionSelected === true) {
			strMethodDetails = "Selected the value of: " + strItemText + " from the dropdown '" + strObjName + ".";
		} else {
			boolMethodPassed = false;
			strMethodDetails = "FAILED!!! The dropdown element : '" + strObjName + "' DOES NOT contain the value: " + strItemText + gblLineFeed;
			strMethodDetails = strMethodDetails + "Checked ALL " + intCntOptions + " options text consisting of: " + strOptItemsText + "!!!";
		}
	}
	else {
		boolMethodPassed = false;
		strMethodDetails = 'FAILED!!! The dropdown element : ' + strObjName + ' CONTAINS ZERO OPTIONS!!!';
	}
	mapMethodResults.boolPassed = boolMethodPassed.toString();
	mapMethodResults.strMethodDetails = strMethodDetails; // Ensure strMethodDetails is set before returning
	return mapMethodResults;
}

/**
 * -------------------------------------  objVerifyListBoxItem  -----------------------------------
 * Verify the list box item.
 * @param mapVerLstItemData Map containing verification data: weListBox, strObjName, strItemXPath, strItemText, intItemIndex, intItemCnt
 * @return mapResults
 *  1. boolMethodPassed The pass/fail status. true/false
 *  2. strMethodDetails The details on the processing of the method
 * @author pkanaris
 * @author Created: 11/03/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objVerifyListBoxItem(mapVerLstItemData) {
	//Create variables
	const gblLineFeed = GVars.GblLineFeed('Value');
	let boolMethodPassed = true;
	let strMethodDetails; // Will be assigned later
	const mapMethodResults = {};
	//Process the object
	let boolValueFound = false;
	let lstChildObjects; // Placeholder for Array<WebElement>
	let intCntOptions;
	let intLoopStart;
	let intLoopEnd;
	let intItemInstance; // Declared here to be accessible after loop

	//Return the values from the map
	const weListBox = mapVerLstItemData.weListBox;
	const strObjName = mapVerLstItemData.strObjName;
	const strItemXPath = mapVerLstItemData.strItemXPath;
	const strItemText = mapVerLstItemData.strItemText;
	const intItemIndex = mapVerLstItemData.intItemIndex;
	const intItemCnt = mapVerLstItemData.intItemCnt;

	//Check to see if items are returned and use a timer to retry
	for (let retrycnt = 5; retrycnt >= 1; retrycnt--) { // Adjusted retry loop logic
		//Return the option list
		lstChildObjects = weListBox.findElements(By.xpath(strItemXPath)); // Assuming findElements returns array sync
		//Count the options
		intCntOptions = lstChildObjects.length;
		if (intCntOptions > 0) {
			break;
		}
		else {
			blockingSleep(100);
		}
	}
	if (intCntOptions === 0) {
		//Report no items found
		boolMethodPassed = false;
		strMethodDetails = "FAILED!!! The list box DID NOT RETURN ANY ITEMS!!!";
	}
	else if (intCntOptions !== intItemCnt && intItemCnt !== -1) {
		//Report the count does not match
		boolMethodPassed = false;
		strMethodDetails = "FAILED!!! The displayed count of: " + intCntOptions + " DOES NOT MATCH THE EXPECTED COUNT OF: " + intItemCnt + "!!!";
	}
	else {
		//Process the values
		let weListItem; // Placeholder for WebElement
		let strOptInstanceText;
		let strOptItemsText = ''; // Initialize to empty string
		//Check if we are verifying the item by index
		if (intItemIndex === -1) {
			intLoopStart = 0;
			intLoopEnd = intCntOptions - 1; // Loop until last index inclusive
		}
		else if (intItemIndex >= intCntOptions) { // Changed condition due to intCntOptions represents 1-based size
			boolMethodPassed = false;
			strMethodDetails = "FAILED!!! The list box element : '" + strObjName + "' CONTAINS '" + intCntOptions + "' WHICH IS LESS THEN THE ASSIGNED INDEX OF: " + intItemIndex + "!!!";
		}
		else {
			intLoopStart = intItemIndex;
			intLoopEnd = intItemIndex;
		}
		if (boolMethodPassed === true) {
			for (intItemInstance = intLoopStart; intItemInstance <= intLoopEnd; intItemInstance++) {
				weListItem = lstChildObjects[intItemInstance];
				//Highlight
				if (TSExecParams.getBoolDoHighlight() === true) {
					const mapHighlight = objHighlightElementJS(weListItem, 'ListItem', strObjName + ' List Item'); // Reuse internal helper
				}
				strOptInstanceText = weListItem.getText(); // Assumes getText() reads text sync
				if (intItemInstance === 0) {
					strOptItemsText = strOptInstanceText;
				}
				else {
					strOptItemsText = strOptItemsText + '; ' + strOptInstanceText;
				}
				//Check if the value is present
				if (strOptInstanceText === strItemText) {
					boolValueFound = true;
					break;
				}
			}
		}
		if (boolValueFound === true) {
			strMethodDetails = "The value of: " + strItemText + " is present in the list box: '" + strObjName + " at index: " + intItemInstance + ".";
		} else if (boolValueFound === false && intItemIndex === -1) {
			boolMethodPassed = false;
			strMethodDetails = "FAILED!!! The listbox element : '" + strObjName + "' DOES NOT contain the value: " + strItemText + gblLineFeed;
			strMethodDetails = strMethodDetails + "Checked ALL " + intCntOptions + " options text consisting of: " + strOptItemsText + "!!!";
		}
		else {
			boolMethodPassed = false;
			strMethodDetails = "FAILED!!! The listbox element : '" + strObjName + "' DOES NOT contain the value: " + strItemText + " at INDEX: " + intItemIndex + "!!!";
		}
	}
	mapMethodResults.boolPassed = boolMethodPassed.toString();
	mapMethodResults.strMethodDetails = strMethodDetails; // Ensure strMethodDetails is set before returning
	mapMethodResults.intItemIndex = intItemInstance;
	return mapMethodResults;
}

//** Multi Select List Box **/
/**
 * -------------------------------------  objSelectListBoxItemByChildContains  -----------------------------------
 * Select the specified item from the list box by available child elements XPath that contains the assigned value.
 * @param weListBox The dropdown WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strItemXPath The XPath for the child items
 * @param strItemText The item text to be selected (will be checked for partial match)
 * @return mapResults
 *  1. boolMethodPassed The pass/fail status. true/false
 *  2. strMethodDetails The details on the processing of the method
 *  3. weNewDropDown The listbox after state has changed.
 * @author pkanaris
 * @author Created: 04/11/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSelectListBoxItemByChildContains(weListBox, strObjName, strItemXPath, strItemText) {
	//Create variables
	const gblLineFeed = GVars.GblLineFeed('Value');
	let boolMethodPassed = true;
	let strMethodDetails; // Will be assigned later
	const mapMethodResults = {};
	//Process the object
	let boolValueFound = false;
	let boolOptionSelected = false; // Initialize to false
	//Return the option list
	const lstChildObjects = weListBox.findElements(By.xpath(strItemXPath)); // Assuming findElements returns array sync
	//Count the options
	const intCntOptions = lstChildObjects.length;
	if (intCntOptions > 0) {
		//Process the values
		let weListItem; // Placeholder for WebElement
		let strOptInstanceText;
		let strOptItemsText = ''; // Initialize to empty string
		for (let intItemInstance = 0; intItemInstance < intCntOptions; intItemInstance++) {
			weListItem = lstChildObjects[intItemInstance];
			strOptInstanceText = weListItem.getText(); // Assumes getText() reads text sync
			if (intItemInstance === 0) {
				strOptItemsText = strOptInstanceText;
			}
			else {
				strOptItemsText = strOptItemsText + '; ' + strOptInstanceText;
			}
			//Check if the value is present (using contains logic)
			boolValueFound = StrNums.JComm_VerifyTextPresent(strOptInstanceText, strItemText, '%Con%');
			if (boolValueFound === true) {
				boolOptionSelected = true;
				//Highlight
				if (TSExecParams.getBoolDoHighlight() === true) {
					const mapHighlight = objHighlightElementJS(weListItem, 'ListItem', strObjName + ' List Item'); // Reuse internal helper
				}
				try {
					weListItem.click();
				}
				catch (e) {
					strMethodDetails = (strMethodDetails || '') + "FAILED TO MAKE SELECTION BY INDEX!!! " + 'Exception occurred!!! SEE ERROR STACK TRACE: ' +
						(e instanceof Error ? e.stack : String(e));
				}
				//wait after the selection
				DateTime.WaitSecs(TSExecParams.getIntViewDelaySecs()); // Assumes blocking wait
				break; //Exit the for loop for finding the item in the options
			}
		}
		if (boolOptionSelected === true) {
			strMethodDetails = "Selected the value of: " + strItemText + " from the dropdown '" + strObjName + ".";
		} else {
			boolMethodPassed = false;
			strMethodDetails = "FAILED!!! The dropdown element : '" + strObjName + "' DOES NOT contain the value: " + strItemText + gblLineFeed;
			strMethodDetails = strMethodDetails + "Checked ALL " + intCntOptions + " options text consisting of: " + strOptItemsText + "!!!";
		}
	}
	else {
		boolMethodPassed = false;
		strMethodDetails = 'FAILED!!! The dropdown element : ' + strObjName + ' CONTAINS ZERO OPTIONS!!!';
	}
	mapMethodResults.boolPassed = boolMethodPassed.toString();
	mapMethodResults.strMethodDetails = strMethodDetails; // Ensure strMethodDetails is set before returning
	return mapMethodResults;
}

/**
 * -------------------------------------  objSelectListBoxItemSeparateSelectElemByChild  -----------------------------------
 * Select the specified item from the list box by finding text in the item path.
 * This version allows a separate XPath for the clickable element within the list item.
 * @param weListBox The dropdown WebElement (placeholder for WebElement)
 * @param strObjName The meaningful name of the testobject
 * @param strItemXPath The XPath for the child items (the containers)
 * @param strSelectItemXpath The XPath for the clickable element *within* the item container
 * @param strItemText The item text to be selected
 * @return mapResults
 *  1. boolMethodPassed The pass/fail status. true/false
 *  2. strMethodDetails The details on the processing of the method
 *  3. weNewDropDown The listbox after state has changed.
 * @author pkanaris
 * @author Created: 06/30/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_objSelectListBoxItemSeparateSelectElemByChild(weListBox, strObjName, strItemXPath, strSelectItemXpath, strItemText) {
	//Create variables
	const gblLineFeed = GVars.GblLineFeed('Value');
	const gblNull = GVars.GblNull('Value');
	let boolMethodPassed = true;
	let strMethodDetails; // Will be assigned later
	const mapMethodResults = {};
	//Process the object
	let boolValueFound = false;
	let boolOptionSelected = false; // Initialize to false
	//Return the option list
	const lstChildObjects = weListBox.findElements(By.xpath(strItemXPath)); // Assuming findElements returns array sync
	//Count the options
	const intCntOptions = lstChildObjects.length;
	if (intCntOptions > 0) {
		//Process the values
		let weListItem; // Placeholder for WebElement (the container)
		let strOptInstanceText;
		let strOptItemsText = ''; // Initialize to empty string
		for (let intItemInstance = 0; intItemInstance < intCntOptions; intItemInstance++) {
			weListItem = lstChildObjects[intItemInstance];
			strOptInstanceText = weListItem.getText(); // Assumes getText() reads text sync
			if (intItemInstance === 0) {
				strOptItemsText = strOptInstanceText;
			}
			else {
				strOptItemsText = strOptItemsText + '; ' + strOptInstanceText;
			}
			//Check if the value is present
			if (strOptInstanceText === strItemText) {
				boolValueFound = true;
			}
			if (boolValueFound === true) {
				boolOptionSelected = true;
				//Highlight
				if (TSExecParams.getBoolDoHighlight() === true) {
					const mapHighlight = objHighlightElementJS(weListItem, 'ListItem', strObjName + ' List Item'); // Reuse internal helper
				}
				try {
					let weItemClick; // Placeholder for WebElement (the clickable part)
					if (strSelectItemXpath === gblNull) {
						weItemClick = weListItem; // Click the container if no specific selector
					}
					else {
						weItemClick = weListItem.findElement(By.xpath(strSelectItemXpath)); // Find specific clickable element within container
					}
					if (weItemClick === null || weItemClick === undefined) { // Check both null and undefined
						boolMethodPassed = false;
						strMethodDetails = "FAILED!!! Unable to find the child object for selection!!!";
					}
					else {
						weItemClick.click();
					}
				}
				catch (e) {
					strMethodDetails = (strMethodDetails || '') + "FAILED TO MAKE SELECTION Exception occurred!!! SEE ERROR STACK TRACE: " +
						(e instanceof Error ? e.stack : String(e));
				}
				//wait after the selection
				DateTime.WaitSecs(TSExecParams.getIntViewDelaySecs()); // Assumes blocking wait
				break; //Exit the for loop for finding the item in the options
			}
		}
		if (boolOptionSelected === true) {
			strMethodDetails = "Selected the value of: " + strItemText + " from the dropdown '" + strObjName + ".";
		} else {
			boolMethodPassed = false;
			strMethodDetails = "FAILED!!! The dropdown element : '" + strObjName + "' DOES NOT contain the value: " + strItemText + gblLineFeed;
			strMethodDetails = strMethodDetails + "Checked ALL " + intCntOptions + " options text consisting of: " + strOptItemsText + "!!!";
		}
	}
	else {
		boolMethodPassed = false;
		strMethodDetails = 'FAILED!!! The dropdown element : ' + strObjName + ' CONTAINS ZERO OPTIONS!!!';
	}
	mapMethodResults.boolPassed = boolMethodPassed.toString();
	mapMethodResults.strMethodDetails = strMethodDetails; // Ensure strMethodDetails is set before returning
	return mapMethodResults;
}

function waitForPageLoad(driver, loadingXPath, maxWaitTimeSecs = 30) {
	const mapMethodResults = {};
	let strMethodDetails = '';
	let boolMethodPassed = false;
	let weLoading = null;

	try {
		weLoading = WebDriver.FindElement(loadingXPath);
	} catch (e) {
		weLoading = null;
	}

	if (weLoading !== null) {
		const intLoopCnt = maxWaitTimeSecs === 0 ? 120 : maxWaitTimeSecs * 4;

		for (let loopWait = 0; loopWait < intLoopCnt; loopWait++) {
			const readyState = WebDriver.ExecuteScript('return document.readyState');
			//const readyState = await driver.executeScript('return document.readyState');
			if (readyState === 'complete') {
				try {
					const isDisplayed = weLoading.isDisplayed();
					if (!isDisplayed) {
						boolMethodPassed = true;
						break;
					}
				} catch (e) {
					// Element might have gone stale or disappeared
					boolMethodPassed = true;
					break;
				}
			}
			DateTime.WaitMilliSecs(250);
		}
	} else {
		boolMethodPassed = true;
	}

	if (boolMethodPassed) {
		strMethodDetails = "Page finished loading.";
	} else {
		strMethodDetails = "FAILED!!! THE PAGE HAS NOT COMPLETED LOADING!!!!";
	}

	mapMethodResults['boolPassed'] = boolMethodPassed.toString();
	mapMethodResults['strMethodDetails'] = strMethodDetails;

	return mapMethodResults;
}
/**
 * -------------------------------------  returnWebElemChildElementCount  -----------------------------------
 * Returns the number of element are displayed based on the child element XPath
 * @param weParent The parent WebElement (placeholder for WebElement)
 * @param strChildXPath The XPath of the child object
 * @return intChildCount The number of the child objects present.
 * @author pkanaris
 * @author Created: 04/06/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_returnWebElemChildElementCount(weParent, strChildXPath) {
	//GlobalVars (not used here besides for comments)
	let intChildCount = 0; // Initialize to 0
	if (weParent !== null && weParent !== undefined) {
		//Return the count of child elements
		const lstChildObjs = weParent.findElements(By.xpath(strChildXPath)); // Assumes findElements returns array sync
		intChildCount = lstChildObjs.length;
	}
	return intChildCount;
}

/**
 * -------------------------------------  returnChildElementCount  -----------------------------------
 * Returns the number of element are displayed based on the child element XPath
 * @param strParentXpath The XPath of the parent
 * @param strChildXPath The XPath of the child object
 * @return intChildCount The number of the child objects present.
 * @author pkanaris
 * @author Created: 04/06/2022
 * @author Last Edited:
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 */
function CommonWebCore_returnChildElementCount(strParentXpath, strChildXPath) {
	//GlobalVars (not used here besides for comments)
	let intChildCount = 0; // Initialize to 0
	//Return the element
	const weParent = returnWebElement(strParentXpath); // Reuse internal helper (re-find parent element)
	if (weParent !== null && weParent !== undefined) {
		//Return the count of child elements
		const driver = TestCaseObject.tcDriver; // Assumes TestCaseObject.tcDriver returns a WebDriver instance
		// Original: TestCaseObject.tcDriver.findElements(By.xpath(strParentXpath + strChildXPath))
		const lstChildObjs = driver.findElements(By.xpath(strParentXpath + strChildXPath)); // Search from driver for combined XPath
		intChildCount = lstChildObjs.length;
	}
	return intChildCount;
}

