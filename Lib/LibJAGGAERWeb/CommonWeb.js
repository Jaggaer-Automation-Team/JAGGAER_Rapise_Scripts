// --- CommonWeb.groovy --- (Original file name preserved for reference)

SeSGlobalObject("CommonWeb");

// package resources.web (Equivalent in TS is module structure or just top-level functions)

//Describe the library purpose
/** This package contains the main methods used as keywords for common web application steps
 * The actions to take on a given element to include returning results in most cases.
 * Exception for results is the returning of a New Browser Profile
 */

//Add Imports (These are represented by global `declare const` for external/internal modules)

// Global declarations for external dependencies (these would normally be imported from npm packages)
/* Review and determine how to address.
declare const WebDriver;
declare const WebElement;
declare const Alert;
declare const FluentWait; // Imported but not used
declare const WebDriverWait;
declare const ExpectedConditions;
declare const StaleElementReferenceException; // Original Java exception type
declare const NoAlertPresentException;	 // Original Java exception type
declare const ChromeDriver;
declare const EdgeDriver;
declare const FirefoxDriver;
declare const WebDriverManager;
declare const This; // Imported but not used
declare const FileNotFoundException; // Imported but not used
declare const Duration; // Imported but not used
declare const TimeDuration; // Imported but not used
declare const TimeUnit;
declare const Select; // Imported but not used
declare const Actions;
declare const By;
declare const JavascriptExecutor; // Aliased, imported but not used
declare const Keys; // Imported but not used
declare const OutputType;
declare const FileUtils;
declare const ExceptionUtils;
declare const XSSFWorkbook; // Imported but not used directly for new objects here
declare const XSSFSheet;	// Imported but not used directly for new objects here

// Global declarations for internal/project-specific modules (these would normally be imported from local files)
declare const TCObj;
declare const GVars;
declare const DateTime;
declare const ExcelData;
declare const StrNums;
declare const TCExecRep; // Imported but not used
declare const TCExecParams;
declare const CWCore;
*/

// A global System object to mimic Java's System.out.println
/* Review all code with system out for println and remove it.
declare const System;
System.out = { // Mimic System.out for println
	println: (message) => console.log(message),
	print: (message) => process.stdout.write(message) // For nested loops in later sections
};
*/


/**
* -------------------------------------  getNewBrowserProfile  -----------------------------------
* Return the browser profile in the WebDriver
* @param assgBroswer The browser that will be used for the test execution
* @return newDriver The WebDriver for the assigned profile.
*
* @author pkanaris
* @author Created: 04/15/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_getNewBrowserProfile (assgBroswer) { // Return type `any` as WebDriver is not defined via interface
	let newdriver = null;
	//String assgBroswer = "Edge"
	//Implement WebDriverManager
	switch (assgBroswer) {
		case "Chrome":
			WebDriverManager.chromedriver().clearDriverCache().setup();
			WebDriverManager.chromedriver().clearResolutionCache().setup();
			WebDriverManager.chromedriver().setup();
			newdriver = new ChromeDriver();
			break;
		case "Edge":
			WebDriverManager.edgedriver().clearDriverCache().setup();
			WebDriverManager.edgedriver().clearResolutionCache().setup();
			WebDriverManager.edgedriver().setup();
			newdriver = new EdgeDriver(); // Removed (WebDriver) cast as it implies an interface
			break;
		case "Firefox":
			WebDriverManager.firefoxdriver().clearDriverCache().setup();
			WebDriverManager.firefoxdriver().clearResolutionCache().setup();
			WebDriverManager.firefoxdriver().setup();
			newdriver = new FirefoxDriver();
			break;
	}
	return newdriver;
}
/*
	* returnBrowserProfileName
	* Returns the Rapise Browers Profile name based on the browser assigned
	* @param strAssgBrowser The assigned browser
	*
	* return strBrowserProfile
	* Created 10/02/2020
	* Author PGKanaris
	* Last Edited:
	* Last Edited By:
	* Edit Comments: (Include email, date and details)
 */
function CommonWeb_ReturnBrowserProfileName (strAssgBrowser) {
	if (boolDoDebug === true) {
		Tester.Message('From function. The assigned browser profiles is: ' + strAssgBrowser)
	}
	switch(strAssgBrowser)
	{
		case 'Chrome':
			strBrowserProfile = 'Selenium - Chrome';
			GetWebDriverNonProfileCapabilities(strBrowserProfile);
			break;
		case 'Edge':
			strBrowserProfile = 'Selenium - Edge';
			break;
		case 'Firefox':
			strBrowserProfile = 'Selenium - Firefox';
			break;
		case 'Internet Explorer':
			strBrowserProfile = 'Internet Explorer HTML';
			break;
		default:
			strBrowserProfile = TComm_Undefined
	}
	if (boolDoDebug === false) {
		Tester.Message('From CommonWeb_ReturnBrowserProfileName. The specified browser profiles is: ' + strBrowserProfile)
	}
	return strBrowserProfile
}

function GetWebDriverNonProfileCapabilities(profile)
{
    var caps = {};
    caps['browserName'] = 'chrome';
    caps['goog:chromeOptions'] = {
        args: [
            '--disable-features=PasswordChangeDetection',
            '--disable-save-password-bubble'
        ]
    };
    return caps;
}

/*
* -----------------------------  CommonWeb_OpenBrowserToURL ---------------------------
	* Opens the specified browser to the assigned URL and optional organization. Processes the DoMaximize
	* @param strBrowser The browser to use
	* @param strURL The URL assigned
	* @param strOrgName The optional organization to enter
	* @boolDoMaximize Maximize the browser? true/false
	*
	* return NONE
	* Created 01142021
	* Author PGKanaris
	* Last Edited:
	* Last Edited By:
	* Edit Comments: (Include email, date and details)
 */
function CommonWeb_OpenBrowserToURL (/**string*/strBrowser, /**string*/strURL, /**string*/strOrgName, /**boolean*/boolDoMaximize) {
	//Return the updated global values
	//TestEnvironment_GetCurrentExecutionValues();
	//Create the URL
	strFullURL = strURL + strOrgName
	//Result Variables
	strStepAction = "Open Browser to URL";
	strStepDescription = "Open the browser: <" + strBrowser + "> to the URL:<" + strFullURL + ">";
	strStepExpectedResult = "The system will open the the assigned browser to the URL specified.";
	strMethodDetails = '';
	strSampleData = strBrowser + GblLineFeed + strURL + GblLineFeed + strOrgName;
	boolPassed = true;
	//select the browser profile to load
	strBrowserProfile = CommonWeb.ReturnBrowserProfileName(strBrowser);
	if (boolDoDebug === true) {
		Tester.Message('From CommonWeb_OpenBrowserToURL. The specified browser profiles is: ' + strBrowserProfile)
	}
	/** WebDriver Methods we will replace with Navigator
	WebDriver.CreateDriver(strBrowser)
	WebDriver.Navigate().GoToUrl(strURL);
	WebDriver.Window().Maximize();
	*/
	//Select the browser profile
	Navigator.SelectBrowserProfile(strBrowserProfile); 
	//Navigator.Browser = strBrowser;
	//Navigate the open browser to the specified URL
	Navigator.Navigate(strFullURL);
	//Process Maximize
	if (boolDoMaximize == true){
		Navigator.Maximize();
	}
	strMethodDetails = "Opened the browser: <" + strBrowser + "> to URL: <"
	+ strFullURL + ">, and the maximize is:<" + boolDoMaximize + ">"; 	
	//Report the results
	 if (boolDoDebug === true){
	 	//Output the report message variables
	 	msg = 'Output the location function report message variables.';
	 	msg+="<br/><b>"+"strStepDescription"+"</b>: "+strStepDescription;
	 	msg+="<br/><b>"+"boolPassed"+"</b>: "+boolPassed;
	 	msg+="<br/><b>"+"strStepExpectedResult"+"</b>: "+strStepExpectedResult;
	 	msg+="<br/><b>"+"strMethodDetails"+"</b>: "+strMethodDetails;
	 	Tester.Message(msg)
	 }
	 	//Update the TestStepResults
	TestStepResults.StepAction = strStepAction;
	TestStepResults.StepDesc = strStepDescription;
	TestStepResults.StepExpected = strStepExpectedResult;
	TestStepResults.StepActual = strMethodDetails;
	TestStepResults.StepData = strSampleData; 
	TestStepResults.StepPassed = boolPassed;
	TestExecReporting.ReportStepResults();
}
/*
* -----------------------------  CommonWeb_CloseBrowser ---------------------------
	* Close the active browser.
	* @param boolCloseAllInstances Close all instance of the current browser? true/false
	**************************** USE CAUTION AS IT MAY CLOSE BROWSERS NOT IN THE TEST **********************
	*
	* return NONE
	* Created 01142021
	* Author PGKanaris
	* Last Edited:
	* Last Edited By:
	* Edit Comments: (Include email, date and details)
 */
function CommonWeb_CloseBrowser (/**boolean*/boolCloseAllInstances) {
	//Result Variables
	strStepAction = "Close Browser";
	strStepDescription = "Close the active browser.";
	strStepExpectedResult = "The Active Browser is closed";
	strMethodDetails = '';
	strSampleData = "*N/A*";
	boolPassed = true;
	//Phase 1 is to close the active broswer only
	Navigator.Close();
	strMethodDetails = "The browser was closed.";
	//Phase 2 we will need to look at tabs/windows or look for popups

	//Update the TestStepResults
	TestStepResults.StepAction = strStepAction;
	TestStepResults.StepDesc = strStepDescription;
	TestStepResults.StepExpected = strStepExpectedResult;
	TestStepResults.StepActual = strMethodDetails;
	TestStepResults.StepData = strSampleData; 
	TestStepResults.StepPassed = boolPassed;
	TestExecReporting.ReportStepResults();
}
/*
* -----------------------------  LibJaggaerWeb_HighlightElement ---------------------------
	* Highlights the specified element
	* @param objAssgElement The element to highlight
	*
	* return NONE
	* Created 10/22/2020
	* Author PGKanaris
	* Last Edited:
	* Last Edited By:
	* Edit Comments: (Include email, date and details)
 */
function CommonWeb_HighlightElement (objAssgElement) {
	//TODO Should we get the original outline? Try this when we have elements with an outline otherwise none works.
	//var borderStyle = WebDriver.executeScript("return window.getComputedStyle(arguments[0]).border;", objAssgElement);
	//Tester.Message('Style is: ' + borderStyle)
	Tester.Message('Entered the CommonWeb_HighlightElemeent with the intHighlightCount:' + intHighlightCount)
	for (cntFlash = 0; cntFlash < intHighlightCount; cntFlash++){
		Navigator.ExecJS("arguments[0].style.border='5px dashed red';", objAssgElement);
		DateTime.WaitMilliSecs(300); //In Milliseconds
		Navigator.ExecJS("arguments[0].style.border='3px solid green';", objAssgElement);
		DateTime.WaitMilliSecs(300); //In Milliseconds	
	}
	//Reset the object back to the original outline style
		Navigator.ExecJS("arguments[0].style.border='none';", objAssgElement);
}

/*
* -----------------------------  LibJaggaerWeb_HighlightElement ---------------------------
	* Highlights the specified element
	* @param objAssgElement The element to highlight
	*
	* return NONE
	* Created 10/22/2020
	* Author PGKanaris
	* Last Edited:
	* Last Edited By:
	* Edit Comments: (Include email, date and details)
 */
function CommonWeb_HighlightElement_Navigator (strObjFullXpath) {
	//TODO Should we get the original outline? Try this when we have elements with an outline otherwise none works.
	//var borderStyle = WebDriver.executeScript("return window.getComputedStyle(arguments[0]).border;", objAssgElement);
	var /**WebElementWrapper*/element = WebDriver.FindElementByXPath(strObjFullXpath)
	var borderStyle = StringsAndNumbers.JComm_HandleNoData(element.GetAttribute("style"));
	Tester.Message('Start HighlightElement.Style is: ' + borderStyle)
	var objAssgnElement = Navigator.DOMFindByXPath(strObjFullXpath)
	for (cntFlash = 0; cntFlash < intHighlightCount; cntFlash++){
		Navigator.ExecJS("arguments[0].style.border='5px dashed red';", objAssgnElement);
		DateTime.WaitMilliSecs(300); //In Milliseconds
		Navigator.ExecJS("arguments[0].style.border='3px solid green';", objAssgnElement);
		DateTime.WaitMilliSecs(300); //In Milliseconds	
	}
	//Reset the object back to the original outline style
	Navigator.ExecJS("arguments[0].style.border='none';", objAssgnElement);
}
/*-------------------------------Web Page Alert Processing ----------------------------------------------*/
/**
* -------------------------------------  ProcessAlertPopup  -----------------------------------
* Process the alert popup to include is alert present, displayed text, enter text, enter password, accept, and dismiss.
* NOTE: We cannot get the server info from the alert. The server info is automatically added to the object. Is not available to selenium
* **** MUST BE THE STEP FOLLOWING THE ACTION that caused the alert.
*
* @param mapDataValues			  The Map containing the input data.
*
* @return mapResults				The Map containing the results.
* @return boolMethodPassed		  The string value of the pass/fail for the method.
* @return boolAlertIsDisplayed	  The string value for it the alert is displayed.
*
* @author pkanaris
* @author Created: 04/15/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_ProcessAlertPopup (mapDataValues) { // mapDataValues and return type are `any` as Map is not defined via interface
	//Setup global variables
	let gblNA = GVars.GblNotApplicable('Value');
	let gblSkip = GVars.GblSkip('Value');
	//Create outputs
	let mapResults = {}; // Empty object literal to mimic Groovy's Map initialization
	let boolPassed = true;
	let boolAlertIsDisplayed = false;
	let strMethodDetails; // Type inferred from assignment
	let objAlertPU;	 // Type inferred from assignment
	let strAlertText;   // Type inferred from assignment
	//Return the map values
	let strMessage = mapDataValues.strMessage;
	let boolAcceptAlert = mapDataValues.boolAcceptAlert;
	let boolDismissAlert = mapDataValues.boolDismissAlert;
	//Check if Alert message is present
	//See if the alert is present
	try {
		//Wait 10 seconds for alert
		let wait = new WebDriverWait(TCObj.tcDriver, 10);
		objAlertPU = wait.until(ExpectedConditions.alertIsPresent());
		boolAlertIsDisplayed = true;
		strMethodDetails = "Alert Popup Message is present.";
	}
	catch (error) { // Catching a generic error as NoAlertPresentException is not typed via interface
		boolPassed = false;
		strMethodDetails = 'FAILED TO FIND ALERT!!! SEE ERROR STACK TRACE: ' + ExceptionUtils.getStackTrace(error);
	}
	if (objAlertPU != null) {
		boolAlertIsDisplayed = true;
		//https://www.softwaretestingmaterial.com/javascript-alerts-popups-selenium/
		//Items we can do with alert popup
		//Verify the message if not gblSkip or gblNA
		if (strMessage == gblNA || strMessage == gblSkip) {
			strMethodDetails = "The user has specified to skip the Alert message verification.";
		}
		else {
			strAlertText = StringsAndNumbers.JComm_HandleNoData(objAlertPU.getText());
			if (strMessage == strAlertText) {
				strMethodDetails = "The Alert message was correctly displayed of '" + strMessage + "'.";
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The DISPLAYED Alert message of '" + strAlertText + "' DID NOT MATCH the expected '" + strMessage + "'!!!";
			}
		}
		//Send values to enter into a field
		//Send credentials
		//If both accept and dismiss are the same fail and return the error message
		if(boolAcceptAlert == boolDismissAlert) {
			if (boolAcceptAlert == true) {
				boolPassed = false;
				strMethodDetails = strMethodDetails + "FAILED!!! The user specifed TRUE for BOTH Accept and Dismiss Alert. We can ONLY perform ONE ACTION!!!";
			}
			else {
				boolPassed = false;
				strMethodDetails = strMethodDetails + "FAILED!!! The user specifed FALSE for BOTH Accept and Dismiss Alert. We MUST perform ONE OF THE ACTIONS!!!";
			}
		}
		else {
			//Accept the Alert by clicking Ok
			if (boolAcceptAlert == true) {
				//Process the Accept
				try {
					objAlertPU.accept();
					strMethodDetails = strMethodDetails + 'Accepted the Alert Message.';
				}
				catch(error) {
					strMethodDetails = 'Exception occurred on Accepting the Alert!!! SEE ERROR STACK TRACE: ' + ExceptionUtils.getStackTrace(error);
				}
			}
			else if(boolDismissAlert == true) {
				//Dismiss the Alert by clicking Cancel
				try {
					objAlertPU.dismiss();
					strMethodDetails = strMethodDetails + 'Dismissed the Alert Message.';
				}
				catch(error) {
					strMethodDetails = 'Exception occurred on Dismissing the Alert!!! SEE ERROR STACK TRACE: ' + ExceptionUtils.getStackTrace(error);
				}
			}
			else {
				boolPassed = false;
				strMethodDetails = strMethodDetails + "FAILED!!! The user specifed FALSE for BOTH Accept and Dismiss Alert. We MUST perform ONE OF THE ACTIONS!!!";
			}
		}
	}
	else {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The ALERT was not DISPLAYED WHEN EXPECTED";
	}
	//Update the map
	mapResults.boolMethodPassed = boolPassed.toString(); // Using TS object assignment over `put`
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.boolAlertIsDisplayed = boolAlertIsDisplayed.toString();
	return mapResults;
}
/*
* -----------------------------  LibJaggaerWeb_CaptureImage ---------------------------
	* Capture
	* @param objAssgElement The element to highlight
	*
	* return NONE
	* Created 10/22/2020
	* Author PGKanaris
	* Last Edited:
	* Last Edited By:
	* Edit Comments: (Include email, date and details)
 */
function CommonWeb_CaptureImage (/**string*/strParentXpath, /**string*/strElemName, /**string*/strImgCaption, /**boolean*/boolOverRideHighLight) {
	let strStepAction = 'Capture Image'
	let strStepDescription = "Capture and image of the " + strElemName + " with a caption of: " + strImgCaption;
	let strStepExpectedResult = "The test will capture the image.";
	let strMethodDetails = '';
	let strSampleData = strImgCaption
	let /**boolean*/boolPassed = true;
	let /**boolean*/boolUseCustomCapture = false; //Trying to implement the step so the caption and all info is in the same step.
	var img = null;
	currentDate = new Date();
	let strCaptionOut =  "Captured " + strImgCaption + " image at: " + currentDate
	//Check if we are to capture snapshots
	//Get the object from the repo
	//Verify the object is present
	let strElemXpath = getORXPath(strElemName);
	let strElemFullPath = strParentXpath + strElemXpath
	var weTempElement = Navigator.SeSFind(strElemFullPath);	
	if (!weTempElement) {
		boolPassed = false;
		strMethodDetails = 'Capture Image FAILED!!! The objectID: ' + strElemName + ' WAS NOT found!!!';
		if (boolDoDebug === true){
			Tester.Message ( strMethodDetails);
		}
	}
	else{
		//Highlight if specified
		//TODO add a boolean highlight override for this. Some elements will change the screen layout if we highlight the object.
		if ( boolDoHighlight == true && boolOverRideHighLight == true) {
			CommonWeb.HighlightElement(weTempElement);
		}
		//TODO should we add date time to the caption?
		//Capture the object image
		if (boolUseCustomCapture == true){
			if (boolDoDebug === true) {
				strMethodDetails = 'Processing the image capture using GetBitmap and Img.sav';
				Tester.Message(strMethodDetails)
			}
			//Capture the image
			Tester.CaptureObjectImage('Sign In.png', Ses(weTempElement))
			Tester.Message('Image has been captured', 'Sign In.png')
			var img = File.Read('Sign In.png')
			if (!img) {
				strMethodDetails = "FAILED!!! No image was created!!!!!!";
				boolPassed = false;
			}
			else {
				strMethodDetails = "The image was created and will be saved to the step.";
				if (boolDoDebug === true) {
					Tester.Message("strMethodDetails: " + strMethodDetails);
				}
				//Get the date and time in the format specified
				var data = [
			    new SeSReportText("Captured " + strImgCaption + " image at: " + currentDate),
			    new SeSReportImage(img, strCaptionOut)
				];
				Tester.Assert(
					strStepDescription,
					boolStepPassed,
					data,
					strStepActualResults,
					 {"expectedResult":strStepExpectedResult, "sampleData":strSampleData, "action": strStepAction}
				);
			}
			
		}
		else{
			if (boolDoDebug === true) {
				Tester.Message('Processing the image capture using CaptureObjectImage')
			}
			Tester.CaptureObjectImage(strCaptionOut, SeS(weTempElement))
		}
	}
	if (boolPassed == false){
		//Update the TestStepResults
		TestStepResults.StepAction = strStepAction;
		TestStepResults.StepDesc = strStepDescription;
		TestStepResults.StepExpected = strStepExpectedResult;
		TestStepResults.StepActual = strMethodDetails;
		TestStepResults.StepData = strElemName + ": " + strImgCaption; 
		TestStepResults.StepPassed = boolPassed;
		TestExecReporting.ReportStepResults();
	}
}
/**
* -------------------------------------  takeWebElementScreenShot  -----------------------------------
* Take a screenshot of the specified WebElement
* @param objWE		  The element to capture a screencapture of.
* @param strWEName	  The name of the webelement.
*
* @author pkanaris
* @author Created: 04/15/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_takeWebElementScreenShot(/*WebElement*/objWE, /*String*/strWENam) { // objWE, return as `any`
	//Add the datetime to the file name
	strWEName = strWEName + " " + DateTime.timeCreateDateTimeValueForFileName();
	//Create outputs
	let mapResults = {}; // Mimic Groovy Map
	let boolPassed = true;
	let strMethodDetails; // Type inferred
	let strFileAbsPath;   // Type inferred
	let strScreenCptFileExt = GVars.GBLScreenCaptureExt("Value");
	let strTestOutputFolder = TCExecParams.getStrTestScreenShotsFolder();
	let strFilePath = strTestOutputFolder + "\\" + strWEName + strScreenCptFileExt;
	let fileScrnCapture; // Type inferred, will be a File object representation
	let weParent = TCObj.getWeParentObject(); // Type inferred
	//Add try catch
	try {
		fileScrnCapture = TCObj.getWeParentObject().getScreenshotAs(OutputType.FILE); // File object in Groovy
	}
	catch (eScrCapt1) {
		Tester.Message(ExceptionUtils.getStackTrace(eScrCapt1));
		strMethodDetails = 'Exception occurred!!! SEE ERROR STACK TRACE: ' + ExceptionUtils.getStackTrace(eScrCapt1) +
				" Will try to capture using object XPath";
		let weLocation = CWCore.returnWebElement(TCObj.getStrParentObjXPath());
		if (weLocation == null) {
			weLocation = CWCore.returnWebElement('//html[1]'); //Take the entire active page
		}
		if (weLocation != null) {
			try {
				fileScrnCapture = weLocation.getScreenshotAs(OutputType.FILE);
			}
			catch (e){
				strMethodDetails = strMethodDetails + "RETRY FAILED TO CAPTURE THE SCREEN!!! " + 'Exception occurred!!! SEE ERROR STACK TRACE: ' +
						ExceptionUtils.getStackTrace(e);
			}
		}
		else {
			boolPassed = false;
			strMethodDetails = 'SCREEN CAPTURE FAILED TO RETURN A WEBELEMENT USING ALL METHODS!!!';
		}

	}
	//if (!fileScrnCapture.exists()) { //Orignal Code PGK 05/12/2022. `exists` method on File
	if (fileScrnCapture == null) {
		//TODO update to show as failed step
		boolPassed = false;
		strMethodDetails = 'FAILED!!! Did not create a snapshot for: ' + strWEName + '!!! See Details: ' + strMethodDetails;
	}
	else {
		try{
			//TODO ADD the scaling code
			// FileUtils.copyFile and new File(strFilePath) should be handled by actual Node.js fs methods or similar
			FileUtils.copyFile(fileScrnCapture, { path: strFilePath }); // Simulate File object creation for destination
			FileUtils.forceDelete(fileScrnCapture);
			strMethodDetails = "Captured the screenshot:'" + strWEName + "'.";
			//Return the absolute path of new file
			// let flTmp = new File(strFilePath); // Groovy File object with getAbsolutePath()
			strFileAbsPath = strFilePath; // Assuming strFilePath is already absolute
			if (TCExecParams.getBoolDoDebug()== true) {
				Tester.Message("The Common Web takeWebElementScreenShot strFileAbsPath: " + strFileAbsPath);
			}
		} catch (e) {
			//logger.logWarning(e.getMessage())
			boolPassed = false;
			strMethodDetails = 'Exception occurred!!! SEE ERROR STACK TRACE: ' + ExceptionUtils.getStackTrace(e);
		}
	}
	//Update the map
	mapResults.boolMethodPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.strFilePath = strFileAbsPath;
	mapResults.strFileName = strWEName;
	mapResults.strFileExt = strScreenCptFileExt;
	return mapResults;
}

/**
* -------------------------------------  GetObjectText  -----------------------------------
* Return the object text value.
* @param strLocation		  The web page location value
* @param strElemFullPath	  The EditBox full XPath
* @param strElemName		  The meaningful name of the EditBox
* @return mapResults		  The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 01/11/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_GetObjectText (strLocation, strElemFullPath, strElemName) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'Element';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weElement;	 // Type inferred
	let strValue;	  // Type inferred
	let boolElemStale = false;
	//Return the element
	weElement = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weElement != null) {
		try {
			weElement.toString(); // arbitrary method call to check for stale reference
		}
		catch (e) { // Original: StaleElementReferenceException e. Now just generic error
			boolElemStale = true;
		}
		if (boolElemStale == false) {
			//Highlight
			if (TCExecParams.getBoolDoHighlight() == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weElement, 'Element', strElemName);
			}
			try {
				weElement.toString(); // arbitrary method call
			}
			catch (e) { // Original: StaleElementReferenceException e. Now just generic error
				boolElemStale = true;
				Tester.Message("The element is stale after highlighting!!!");
			}
			if (boolElemStale == false) {
				strValue = StringsAndNumbers.JComm_HandleNoData(weElement.getText());
			}
		}
		if (boolElemStale == true) {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemName + " IS CURRENTLY STALE!!!";
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.strValue = strValue;
	return mapResults;
}

/**
* -------------------------------------  GetObjectAttributeValue  -----------------------------------
* Return the attribute value for the assigned object.
* @param strLocation		  The web page location value
* @param strElemFullPath	  The EditBox full XPath
* @param strElemName		  The meaningful name of the EditBox
* @param strAttribute		 The assigned attribute for the element
* @return mapResults		  The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 01/11/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_GetObjectAttributeValue (strLocation, strElemFullPath, strElemName,
		strAttribute) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'Element';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weElement;	 // Type inferred
	let strAttributeValue; // Type inferred
	let boolElemStale = false;
	//Return the element
	weElement = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weElement != null) {
		try {
			weElement.toString(); // arbitrary method call
		}
		catch (e) { // Original: StaleElementReferenceException e. Now just generic error
			boolElemStale = true;
		}
		if (boolElemStale == false) {
			//Highlight
			if (TCExecParams.getBoolDoHighlight() == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weElement, 'Element', strElemName);
			}
			try {
				weElement.toString(); // arbitrary method call
			}
			catch (e) { // Original: StaleElementReferenceException e. Now just generic error
				boolElemStale = true;
				Tester.Message("The element is stale after highlighting!!!");
			}
			if (boolElemStale == false) {
				//Check if the attribute is present
				if (CWCore.isAttribtuePresent(weElement, strAttribute) == true) {
					//Return the element attribute value
					strAttributeValue = StringsAndNumbers.JComm_HandleNoData(weElement.getAttribute(strAttribute).toString());
				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
							strElemName + " DOES NOT HAVE THE ASSIGNED ATTRIBUTE of: " + strAttribute + "!!!";
				}
			}
		}
		if (boolElemStale == true) {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemName + " IS CURRENTLY STALE!!!";
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.strAttributeValue = strAttributeValue;
	return mapResults;
}
/***************************************** Popup Browser and Browser Tabs ***********************************************************/
/** Contains the various calls for working with popup browser
*
*/


/***************************************** Elements ***********************************************************/
/** COMMON_ELEMENT (such as verify element state) **/
/* COMMON */
/**
* -------------------------------------  VerifyElementState  -----------------------------------
* Set and verify the Element is in the correct visible and enabled state
* @param strLocation		  The web page location value
* @param strElemFullPath	  The Element's full XPath
* @param strTestObjType	   The type of object. (i.e. EditBox, Checkbox, ...)
* @param strElemName		  The meaningful name of the EditBox
* @param boolHighlightElement Highlight the element? True/False
* @param boolVisible		  The element is visible? true/false
* @param boolEnabled		  The element is enabled? true/false
*
* @return mapResults		  The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 04/30/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyElementState (strLocation, strElemFullPath, strTestObjType, strElemName,
		boolHighlightElement, boolVisible, boolEnabled) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weObject;	  // Type inferred
	let strTestObjText; // Declared in Groovy, but unused. Removed in TS to adhere to best practices although not strictly needed by the prompt.
	//Return the element
	weObject = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weObject != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true && boolHighlightElement == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weObject, strTestObjType, strElemName);
		}
		DateTime.WaitSecs(TCExecParams.getIntViewDelaySecs());
		//Check the element state (Enabled, Visible)
		let mapResultsObjectState = {}; // Mimic Groovy Map
		mapResultsObjectState = CWCore.objVerifyState(weObject, strElemName, boolVisible, boolEnabled);
		//Output results
		boolPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsObjectState.boolPassed);
		strMethodDetails = mapResultsObjectState.strMethodDetails;
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/** WebTable Results in Grid not panel (such as verify element state) **/
/* WebTable Results */
/**
* -------------------------------------  OutputResultHeader  -----------------------------------
* Output the header table results
* @param mapInputVariables The input variables that identify the header to include xpaths and properties
* @param Note: when strProperty* == gblNull we do not process the property.
* @param strLocName				   The pipe delimited location value
* @param strLocXpath				  The parent Xpath
* @param strTblName				   The name of the table
* @param strTblXPath				  The table Xpath
* @param strHdrRowXPath			   The Header row Xpath usually //thead//tr
* @param strHdrColXPath			   The header column Xpath usually //th
* @param strHdrColDragDropXPath	   The xpath for the drag and drop object. Enter gblNA if not applicable to the table.
* @param strHdrColCheckBoxXPath	   The xpath for the header checkbox object. Enter gblNA if not applicable to the table.
* @param strHdrColTextXPath		   The Xpath for the column text in the header column cell
* @param strHdrColSortXPath		   The Xpath for the column sort in the header column cell
* @param strHdrOutShtName			 The sheet name assigned for the output
* @param strHdrOutputColNames		 The delimited value containing the column names to be created in the output sheet
* @param btnActionXPath			   The Xpath for the action button
* @param lstActionXPath			   The Xpath for the list created when selecting the action button
* @param strHdrColSortToolTipXPath	The Xpath for the column tooltip
* @param strPropertyHdrColTextToolTip The property value of the tool tip object that contains the tooltip text
* @param strPropertyHdrColName		The cell property that will be used to find the assigned column name. Usually 'abbr'
* @param strPropertyHdrSortable	   The cell property that will distinguish is the column is sortable
* @param strPropertyHdrSortOrder	  The cell property that will distinguish what order the column is sorted in (ascending/descending).
* @param strHdrColMissingNames		For each column that does not have an identifiable name, add the name to the delimited value. i.e. colName1|colName2
* @param strHdrLabelContainerXpath	The Xpath for the Header Label Container used to determine the sort order.
* @param strPropertyHdrSortableValue  The element property that will distinguish is the column is sortable
* @param strPropertyHdrSortOrderValue The element property that will distinguish what order the column is sorted in (ascending/descending).
* @param strSortedOrderElementName	The name of the element which holds the sort order. Specifically HdrLabelContainer, other wise gblNull
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 04/30/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_OutputResultHeader (mapInputVariables) { // mapInputVariables and return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	//Return the method variables
	let strLocName = mapInputVariables.LocName;
	let strLocXpath = mapInputVariables.LocXPath;
	let strTblName = mapInputVariables.TblName;
	let strTblXPath = mapInputVariables.strTblXPath;
	let strHdrRowXPath = mapInputVariables.strHdrRowXPath;
	let strHdrColXPath = mapInputVariables.strHdrColXPath;
	let strHdrColDragDropXPath = mapInputVariables.strHdrColDragDropXPath;
	let strHdrColCheckBoxXPath = mapInputVariables.strHdrColCheckBoxXPath;
	let strHdrColTextXPath = mapInputVariables.strHdrColTextXPath;
	let strHdrColSortXPath = mapInputVariables.strHdrColSortXPath;
	let strHdrOutShtName = mapInputVariables.strHdrOutShtName;
	let strHdrOutputColNames = mapInputVariables.strHdrOutputColNames;
	let btnActionXPath = mapInputVariables.btnActionXPath;
	let lstActionXPath = mapInputVariables.lstActionXPath;
	let strHdrColSortToolTipXPath = mapInputVariables.strHdrColSortToolTipXPath;
	let strPropertyHdrColTextToolTip = mapInputVariables.strPropertyHdrColTextToolTip;
	let strPropertyHdrColName = mapInputVariables.strPropertyHdrColName;
	let strPropertyHdrSortable = mapInputVariables.strPropertyHdrSortable;
	let strPropertyHdrSortOrder = mapInputVariables.strPropertyHdrSortOrder;
	//Check for the alternate values
	let strHdrColMissingNames = gblNull;
	let strHdrLabelContainerXpath = gblNull;
	let strPropertyHdrSortableValue = gblNull;
	let strPropertyHdrSortOrderValue = gblNull;
	let strSortedOrderElementName = gblNull;
	if (mapInputVariables.hasOwnProperty('strHdrColMissingNames')){ // Use hasOwnProperty over containsKey
		strHdrColMissingNames = mapInputVariables.strHdrColMissingNames;
	}
	if (mapInputVariables.hasOwnProperty('strHdrLabelContainerXpath')){
		strHdrLabelContainerXpath = mapInputVariables.strHdrLabelContainerXpath;
	}
	if (mapInputVariables.hasOwnProperty('strPropertyHdrSortableValue')){
		strPropertyHdrSortableValue = mapInputVariables.strPropertyHdrSortableValue;
	}
	if (mapInputVariables.hasOwnProperty('strPropertyHdrSortOrderValue')) {
		strPropertyHdrSortOrderValue = mapInputVariables.strPropertyHdrSortOrderValue;
	}
	if (mapInputVariables.hasOwnProperty('strSortedOrderElementName')) {
		strSortedOrderElementName = mapInputVariables.strSortedOrderElementName;
	}
	//Check if the outputsheet name is valid
	let intLenOptHdrShName = strHdrOutShtName.length;
	if (intLenOptHdrShName < 32) {
		//Verify we have the table.
		let weTable = CWCore.returnWebElement(strLocXpath + strTblXPath);
		if (weTable == null) {
			boolPassed = false;
			strMethodDetails = "The location '" + strLocName + "' '" + strTblName + "' DID NOT RETURN an ELEMENT.";
		}
		else {
			//Highlight
			if (TCExecParams.getBoolDoHighlight() == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weTable, 'Element', strTblName);
			}
			//Check if the header row is present
			let weHeaderRow = CWCore.returnWebElement(strLocXpath + strTblXPath + strHdrRowXPath);
			if (weHeaderRow == null) {
				boolPassed = false;
				strMethodDetails = "The location '" + strLocName + "' '" + strTblName + "' DID NOT RETURN A HEADER ROW ELEMENT.";
			}
			else {
				//Move to the row to show the row we are porcessing in case of errors.
				let actions = new Actions(TCObj.getTcDriver());
				actions.moveToElement(weHeaderRow).perform();
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weHeaderRow, 'Element', strTblName + ' header row');
				}
				//Note, we open the file and do the work for each method and save/close the excel instance so save happens as we go.
				//Add the sheet to the input file
				let mapAddSheet = {}; // Mimic Groovy Map
				let strTCInputFilePath = TCObj.strTCInputFilePath;
				//TODO attempt to add the input stream to the test objects and use the input object to add the sheet in a new TObj WB and Sh
				mapAddSheet = ExcelData.excelAddSheetToWorkBook(strHdrOutShtName);
				let boolSheetAddedPassed = StringsAndNumbers.JComm_StringToBoolean (mapAddSheet.boolPassed);
				if (boolSheetAddedPassed == true) {
					//Add the column names
					let mapAddColNames = {}; // Mimic Groovy Map
					mapAddColNames = ExcelData.excelCreateHeaderCols(strHdrOutputColNames);
					let boolAddColNames = StringsAndNumbers.JComm_StringToBoolean (mapAddColNames.boolPassed);
					if (boolAddColNames == true) {
						//Return the header elements
						let lstColNames = mapAddColNames.lstColNames; // inferred List
						//Update the header column(s) to the assigned theme
						let strAssgTheme = 'OutputDataHdrStd';
						let mapUpdateHdrTheme = {}; // Mimic Groovy Map
						mapUpdateHdrTheme = ExcelData.excelSetRowCellFormatByTheme(0, strAssgTheme);
						let boolUPdHdrTheme = StringsAndNumbers.JComm_StringToBoolean (mapUpdateHdrTheme.boolPassed);
						if (boolUPdHdrTheme == false) {
							boolPassed = false;
							strMethodDetails = mapUpdateHdrTheme.strMethodDetails;
						}
						//Check for alternate name values
						let intAltColNameArrayCnt = -1;
						let intAltColNameUsed = -1;
						let arryAltColName; // inferred List
						//let strTempAltColName; // Declared but unused
						if (strHdrColMissingNames != gblNull) {
							//Split the value into an array
							//ADD a split for the value to return the date/time format
							let mapArray = {}; // Mimic Groovy Map
							mapArray = StringsAndNumbers.JComm_StringToArray(strHdrColMissingNames, gblDelimiter);
							intAltColNameArrayCnt = StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount);
							arryAltColName = mapArray.ArryOfValues;
						}
						//Return the data and update the output sheet
						//Get the column child elements
						let mapGetChildElements = CWCore.returnChildElements(weHeaderRow, strHdrColXPath);
						let lstHdrCols = mapGetChildElements.lstChildObjects; // inferred List<WebElement>
						let cntHdrCols = mapGetChildElements.cntChildObjs;
						if (cntHdrCols > 0) {
							let weHdrCol; // Type inferred
							let weTempHdrColValue; // Type inferred
							let weHdrColLabelContainer; // Type inferred
							let strTempValue; // Type inferred
							let strCellTheme = 'OutputDataCentered'; // Type inferred
							let mapSetCellValue = {}; // Mimic Groovy Map
							let intOutputCol; // Type inferred
							let intChildCnt; // Type inferred
							let boolAttribPresent; // Type inferred
							let intOutputExcelRow; // Type inferred
							let strChkBoxToolTip; // Type inferred
							for (let loopCols = 0; loopCols < cntHdrCols; loopCols++) {
								strChkBoxToolTip = gblNull;
								weHdrCol = lstHdrCols[loopCols];
								intOutputExcelRow = loopCols + 1; //Add one to output the correct row in excel
								//Highlight
								if (TCExecParams.getBoolDoHighlight() == true) {
									let mapHighlight = {}; // Mimic Groovy Map
									mapHighlight = CWCore.objHighlightElementJS(weHdrCol, 'Element', strTblName + ' header row column');
								}
								//Check if there is a child container
								if (strHdrLabelContainerXpath != gblNull) {
									weHdrColLabelContainer = null;
									intChildCnt = -1;
									intChildCnt = CWCore.returnWebElemChildElementCount(weHdrCol, strHdrLabelContainerXpath);
									if (intChildCnt == 1) {
										weHdrColLabelContainer = CWCore.returnChildElement(weHdrCol, strHdrLabelContainerXpath);
									}
								}
								if (weHdrColLabelContainer == null) {
									weTempHdrColValue = weHdrCol;
								}
								else {
									//Highlight
									if (TCExecParams.getBoolDoHighlight() == true) {
										let mapHighlight = {}; // Mimic Groovy Map
										mapHighlight = CWCore.objHighlightElementJS(weHdrColLabelContainer, 'Element', strTblName + ' header row column label container');
									}
									weTempHdrColValue = weHdrColLabelContainer;
								}
								intOutputCol = 0;
								//Note: when strProperty* == gblNull we do not process the property.
								//Check if drag and drop is present in the cell
								if (strHdrColDragDropXPath != gblNull) {
									//Check if the drag and drop is present in the cell
									let cntDragDrop = CWCore.returnWebElemChildElementCount(weHdrCol, strHdrColDragDropXPath);
									if (cntDragDrop == 1) {
										strTempValue = "elemDragDrop";
									}
									else if (cntDragDrop > 1){
										boolPassed = false;
										strMethodDetails = "FAILED!!!! The Header Column: " + loopCols + " contained the 'DragDrop' count of: " + cntDragDrop + " WHICH DOES NOT MATCH THE EXPECTED 1!!!";
									}
									else {
										let cntActionBtn = CWCore.returnWebElemChildElementCount(weHdrCol, btnActionXPath);
										if(cntActionBtn == 1) {
											strTempValue = "elemActionBtn";
										}
										else if (cntActionBtn == 0) {
											strTempValue = gblNull;
										}
										else {
											boolPassed = false;
											strMethodDetails = "FAILED!!!! The Header Column: " + loopCols + " contained the 'ActionBtn' count of: " + cntActionBtn + " WHICH DOES NOT MATCH THE EXPECTED 1!!!";
										}
									}

									if (boolPassed == true) {
										mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intOutputExcelRow, intOutputCol, strTempValue, strCellTheme);
										if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == true) {
											intOutputCol++;
										}
										else {
											boolPassed = false;
											strMethodDetails = mapSetCellValue.strMethodDetails;
										}
									}

								}
								//Check if Check Box is present in the cell
								if (strHdrColCheckBoxXPath != gblNull) {
									//Check if the drag and drop is present in the cell
									let cntChkBox = CWCore.returnWebElemChildElementCount(weHdrCol, strHdrColCheckBoxXPath);
									if (cntChkBox == 1) {
										strTempValue = "elemCheckbox";
										//Return the tooltip if present from the element
										if (strPropertyHdrColTextToolTip != gblNull) {
											let weChkbox = CWCore.returnChildElement(weHdrCol, strHdrColCheckBoxXPath);
											if (weChkbox != null) {
												boolAttribPresent = CWCore.isAttribtuePresent(weChkbox,strPropertyHdrColTextToolTip);
												if (boolAttribPresent == true) {
													strChkBoxToolTip = StringsAndNumbers.JComm_HandleNoData(weChkbox.getAttribute(strPropertyHdrColTextToolTip));
												}
												else {
													strChkBoxToolTip = "No ToolTip";
												}
											}
										}
									}
									else if (cntChkBox > 1){
										boolPassed = false;
										strMethodDetails = "FAILED!!!! The Header Column: " + loopCols + " contained the 'Checkbox' count of: " + cntChkBox + " WHICH DOES NOT MATCH THE EXPECTED 1!!!";
									}
									else {
										strTempValue = gblNull;
									}
									if (boolPassed == true) {
										mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intOutputExcelRow, intOutputCol, strTempValue, strCellTheme);
										if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == true) {
											intOutputCol++;
										}
										else {
											boolPassed = false;
											strMethodDetails = mapSetCellValue.strMethodDetails;
										}
									}
								}
								//Return the name
								if (strPropertyHdrColName != gblNull && boolPassed == true) {
									boolAttribPresent = CWCore.isAttribtuePresent(weHdrCol,strPropertyHdrColName);
									if (boolAttribPresent == true) {
										strTempValue = StringsAndNumbers.JComm_HandleNoData(weHdrCol.getAttribute(strPropertyHdrColName));
										if (strTempValue == gblNull) { //THIS SHOULD NOT HAPPEN
											intAltColNameUsed ++;
											strTempValue = "INVALIDNULLVALUEFORPROPERTY_" + strPropertyHdrColName.toUpperCase() + "_" + StringsAndNumbers.JComm_HandleNoData(arryAltColName[intAltColNameUsed]);
										}
									}
									//Check if we did not find a column name
									else if (intAltColNameArrayCnt >= 1 && intAltColNameArrayCnt > intAltColNameUsed ) {
										intAltColNameUsed ++;
										strTempValue = StringsAndNumbers.JComm_HandleNoData(arryAltColName[intAltColNameUsed]);
									}
									else {
										strTempValue = gblNull;
									}
									mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intOutputExcelRow, intOutputCol, strTempValue, strCellTheme);
									if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == true) {
										intOutputCol++;
									}
									else {
										boolPassed = false;
										strMethodDetails = mapSetCellValue.strMethodDetails;
									}
								}
								//Return the text
								if (boolPassed == true) {
									strTempValue = StringsAndNumbers.JComm_HandleNoData(weTempHdrColValue.getText());
									mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intOutputExcelRow, intOutputCol, strTempValue, strCellTheme);
									if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == true) {
										intOutputCol++;
									}
									else {
										boolPassed = false;
										strMethodDetails = mapSetCellValue.strMethodDetails;
									}
								}
								//Return isSortable
								if (strPropertyHdrSortable != gblNull && boolPassed == true) {
									boolAttribPresent = CWCore.isAttribtuePresent(weTempHdrColValue,strPropertyHdrSortable);
									if (boolAttribPresent == true) {
										strTempValue = StringsAndNumbers.JComm_HandleNoData(weTempHdrColValue.getAttribute(strPropertyHdrSortable));
										if (strTempValue != gblNull && strPropertyHdrSortableValue != gblNull) {
											//Split the value into an array
											let intArrayCnt; // Type inferred
											let arryValues; // Type inferred
											let strValueLoc; // Type inferred
											let strValue; // Type inferred
											//ADD a split for the value to return the date/time format
											let mapArray = {}; // Mimic Groovy Map
											mapArray = StringsAndNumbers.JComm_StringToArray(strPropertyHdrSortableValue, gblDelimiter);
											intArrayCnt = StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount);
											arryValues = mapArray.ArryOfValues;
											if (intArrayCnt == 2) {
												strValueLoc = arryValues[0];
												strValue = arryValues[1];
												let intInString = strTempValue.indexOf(strValue, 0);
												switch (strValueLoc){
													case '*STARTWITH*':
														if (intInString == 0) {
															strTempValue = 'true';
														}
														break;
													case '*CONTAINS*':
														if (intInString > 0) {
															strTempValue = 'true';
														}
														break;
													default:
														strTempValue = 'false';
														break;
												}
											}
											else {
												strTempValue = 'UNKNOWN INPUT Value:' + strPropertyHdrSortableValue;
											}
										}
									}
									else {
										strTempValue = 'false';
									}
									mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intOutputExcelRow, intOutputCol, strTempValue, strCellTheme);
									if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == true) {
										intOutputCol++;
									}
									else {
										boolPassed = false;
										strMethodDetails = mapSetCellValue.strMethodDetails;
									}
								}
								//Return sorted
								strTempValue = gblNull;
								if (strPropertyHdrSortOrder != gblNull && boolPassed == true) {
									let weTempSort; // Type inferred
									if (strSortedOrderElementName == 'HdrLabelContainer') {
										weTempSort = weTempHdrColValue;
									}
									else {
										weTempSort = weHdrCol;
									}
									boolAttribPresent = CWCore.isAttribtuePresent(weTempSort,strPropertyHdrSortOrder);
									if (boolAttribPresent == true) {
										strTempValue = StringsAndNumbers.JComm_HandleNoData(weTempSort.getAttribute(strPropertyHdrSortOrder));
										if (strPropertyHdrSortOrderValue != gblNull) {
											//Split the value into and array and check for each value
											let intArrayCnt; // Type inferred
											let arryValues; // Type inferred
											//ADD a split for the value
											let mapArray = {}; // Mimic Groovy Map
											mapArray = StringsAndNumbers.JComm_StringToArray(strPropertyHdrSortOrderValue, gblDelimiter);
											intArrayCnt = StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount);
											arryValues = mapArray.ArryOfValues;
											if (intArrayCnt == 3) {
												let strColSorted = arryValues[0];
												if (strColSorted == strTempValue || StringsAndNumbers.JComm_TextLocationInString(strTempValue, strColSorted) >= 0) {
													//Return the other values
													let strValueSortAsc = arryValues[1];
													let strValueSortDesc = arryValues[2];
													//Use the child object property to return the sort. Assumes the child object
													boolAttribPresent = CWCore.isAttribtuePresent(weTempHdrColValue,strPropertyHdrSortOrder);
													if (boolAttribPresent == true) {
														let strTempSortOrder = StringsAndNumbers.JComm_HandleNoData(weTempHdrColValue.getAttribute(strPropertyHdrSortOrder));
														// let intIndexSortAsc = strTempSortOrder.indexOf(strValueSortAsc); // Declared but unused
														// let intIndexSortDesc = strTempSortOrder.indexOf(strValueSortDesc); // Declared but unused
														if (strTempSortOrder.indexOf(strValueSortAsc) >= 0) {
															strTempValue = "Ascending";
														}
														else if(strTempSortOrder.indexOf(strValueSortDesc) >= 0) {
															strTempValue = "Descending";
														}
														else {
															strTempValue = "UNKNOWN SORT, Value of: " + strTempSortOrder + " DOES NOT CONTAIN ASC OR DESC DEFINED VALUES!!!";
														}
													}
													else {
														strTempValue = gblNull;
													}
												}
												else {
													strTempValue = gblNull;
												}
											}
											else {
												strTempValue = "UNKNOWN SORT NOT ENOUGH ITEMS IN NVALUE!!!";
											}
										}
									}
									else {
										strTempValue = gblNull;
									}
								}
								//always output the sortvalue
								mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intOutputExcelRow, intOutputCol, strTempValue, strCellTheme);
								if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == true) {
									intOutputCol++;
								}
								else {
									boolPassed = false;
									strMethodDetails = mapSetCellValue.strMethodDetails;
								}
								if (boolPassed == false) {
									break; //Exit the loop if we fail
								}
								//tooltip text
								strTempValue = gblNull;
								if (strHdrColSortToolTipXPath != gblNull && strChkBoxToolTip == gblNull) {
									//Get the column child elements
									let weToolTip; // Type inferred
									let mapGetToolTipElements = CWCore.returnChildElements(weTempHdrColValue, strHdrColSortToolTipXPath);
									let lstToolTip = mapGetToolTipElements.lstChildObjects; // inferred List<WebElement>
									let cntToolTip = mapGetToolTipElements.cntChildObjs;
									if (cntToolTip == 1) {
										weToolTip = lstToolTip[0];
										boolAttribPresent = CWCore.isAttribtuePresent(weToolTip,strPropertyHdrColTextToolTip);
										if (boolAttribPresent == true) {
											strTempValue = StringsAndNumbers.JComm_HandleNoData(weToolTip.getAttribute(strPropertyHdrColTextToolTip));
										}
										else {
											strTempValue = gblNull;
										}
									}
									else if (cntToolTip == 0) {
										strTempValue = "No ToolTip";
									}
									else {
										boolPassed = false;
										strMethodDetails = "FAILED!!!! The Header Column: " + loopCols + " contained the 'ToolTip' count of: " + cntToolTip + " WHICH DOES NOT MATCH THE EXPECTED 1!!!";
									}
								}
								//Check if there is a tooltip present
								else if (strPropertyHdrColTextToolTip != gblNull && strChkBoxToolTip == gblNull) {
									boolAttribPresent = CWCore.isAttribtuePresent(weTempHdrColValue,strPropertyHdrColTextToolTip);
									if (boolAttribPresent == true) {
										strTempValue = StringsAndNumbers.JComm_HandleNoData(weTempHdrColValue.getAttribute(strPropertyHdrColTextToolTip));
									}
									else {
										strTempValue = gblNull;
									}
								}
								else if (strChkBoxToolTip != gblNull) {
									//Update the tooltip to show the checkbox tooltip
									strTempValue = strChkBoxToolTip;
								}
								if (boolPassed == true) {
									mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intOutputExcelRow, intOutputCol, strTempValue, strCellTheme);
									if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == true) {
										intOutputCol++;
									}
									else {
										boolPassed = false;
										strMethodDetails = mapSetCellValue.strMethodDetails;
									}
								}
							}
							//Update the column width to fit
							if (boolPassed == true) {
								//TODO create a new function CommonWeb_to autofit the column
								let mapAutoFit = {}; // Mimic Groovy Map
								mapAutoFit = ExcelData.excelAutofitCols();
								boolPassed = StringsAndNumbers.JComm_StringToBoolean(mapAutoFit.boolPassed);
								strMethodDetails = mapAutoFit.strMethodDetails;
							}
						}
						else {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The TABLE HEADER contain NO COLUMNS!!!";
						}
					}
					else {
						boolPassed = false;
						strMethodDetails = mapAddColNames.strMethodDetails;
					}
				}
				else {
					boolPassed = false;
					strMethodDetails = mapAddSheet.strMethodDetails;
				}
			}
		}
	}
	else {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The assigned output sheet name '" + strHdrOutShtName + " is " + intLenOptHdrShName + " CHAR LENGTH WHICH EXCEEDS EXCEL TAB NAME LENGTH OF 31!!!!";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
* -------------------------------------  VerifyResultHeader  -----------------------------------
* Verify the header table results
* @param mapInputVariables The input variables that identify the header to include xpaths and properties
* @param Note: when strProperty* == gblNull we do not process the property.
* @param strLocName				   The pipe delimited location value
* @param strLocXpath				  The parent Xpath
* @param strTblName				   The name of the table
* @param strTblXPath				  The table Xpath
* @param strHdrRowXPath			   The Header row Xpath usually //thead//tr
* @param strHdrColXPath			   The header column Xpath usually //th
* @param strHdrColDragDropXPath	   The xpath for the drag and drop object. Enter gblNA if not applicable to the table.
* @param strHdrColCheckBoxXPath	   The xpath for the header checkbox object. Enter gblNA if not applicable to the table.
* @param strHdrColTextXPath		   The Xpath for the column text in the header column cell
* @param strHdrColSortXPath		   The Xpath for the column sort in the header column cell
* @param strHdrInputShtName		   The sheet name assigned for the input
* @param strHdrOutShtName			 The sheet name assigned for the output
* @param strHdrOutputColNames		 The delimited value containing the column names to be created in the output sheet
* @param btnActionXPath			   The Xpath for the action button
* @param lstActionXPath			   The Xpath for the list created when selecting the action button
* @param strHdrColSortToolTipXPath	The Xpath for the column tooltip
* @param strPropertyHdrColTextToolTip The property value of the tool tip object that contains the tooltip text
* @param strPropertyHdrColName		The cell property that will be used to find the assigned column name. Usually 'abbr'
* @param strPropertyHdrSortable	   The cell property that will distinguish is the column is sortable
* @param strPropertyHdrSortOrder	  The cell property that will distinguish what order the column is sorted in (ascending/descending).
* @param strHdrColMissingNames		For each column that does not have an identifiable name, add the name to the delimited value
* @param strHdrLabelContainerXpath	The Xpath for the Header Lable Container used to determine the sort order.
* @param strPropertyHdrSortableValue  The element property that will distinguish is the column is sortable
* @param strPropertyHdrSortOrderValue The element property that will distinguish what order the column is sorted in (ascending/descending).
* @param strSortedOrderElementName	The name of the element which holds the sort order. Specifically HdrLabelContainer, other wise gblNull
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 05/06/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyResultHeader (mapInputVariables) { // mapInputVariables and return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblSkip = GVars.GblSkip('Value');
	let gblIgnoreData = GVars.GblIgnoreData('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let cntInputCols = 6; // Type inferred
	//Return the method variables
	let strLocName = mapInputVariables.LocName;
	let strLocXpath = mapInputVariables.LocXPath;
	let strTblName = mapInputVariables.TblName;
	let strTblXPath = mapInputVariables.strTblXPath;
	let strHdrRowXPath = mapInputVariables.strHdrRowXPath;
	let strHdrColXPath = mapInputVariables.strHdrColXPath;
	let strHdrColDragDropXPath = mapInputVariables.strHdrColDragDropXPath;
	let strHdrColCheckBoxXPath = mapInputVariables.strHdrColCheckBoxXPath;
	let strHdrColTextXPath = mapInputVariables.strHdrColTextXPath;
	let strHdrColSortXPath = mapInputVariables.strHdrColSortXPath;
	let strHdrOutShtName = mapInputVariables.strHdrOutShtName;
	let strHdrOutputColNames = mapInputVariables.strHdrOutputColNames;
	let btnActionXPath = mapInputVariables.btnActionXPath;
	let lstActionXPath = mapInputVariables.lstActionXPath;
	let strHdrColSortToolTipXPath = mapInputVariables.strHdrColSortToolTipXPath;
	let strPropertyHdrColTextToolTip = mapInputVariables.strPropertyHdrColTextToolTip;
	let strPropertyHdrColName = mapInputVariables.strPropertyHdrColName;
	let strPropertyHdrSortable = mapInputVariables.strPropertyHdrSortable;
	let strPropertyHdrSortOrder = mapInputVariables.strPropertyHdrSortOrder;
	//Check for the alternate values
	let strHdrColMissingNames = gblNull;
	let strHdrLabelContainerXpath = gblNull;
	let strPropertyHdrSortableValue = gblNull;
	let strPropertyHdrSortOrderValue = gblNull;
	let strSortedOrderElementName = gblNull;
	if (mapInputVariables.hasOwnProperty('strHdrColMissingNames')){
		strHdrColMissingNames = mapInputVariables.strHdrColMissingNames;
	}
	if (mapInputVariables.hasOwnProperty('strHdrLabelContainerXpath')){
		strHdrLabelContainerXpath = mapInputVariables.strHdrLabelContainerXpath;
	}
	if (mapInputVariables.hasOwnProperty('strPropertyHdrSortableValue')){
		strPropertyHdrSortableValue = mapInputVariables.strPropertyHdrSortableValue;
	}
	if (mapInputVariables.hasOwnProperty('strPropertyHdrSortOrderValue')) {
		strPropertyHdrSortOrderValue = mapInputVariables.strPropertyHdrSortOrderValue;
	}
	if (mapInputVariables.hasOwnProperty('strSortedOrderElementName')) {
		strSortedOrderElementName = mapInputVariables.strSortedOrderElementName;
	}
	//Check if the outputsheet name is valid
	if (StringsAndNumbers.JComm_HandleNoData(strHdrOutShtName) == gblNull) {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The OUTPUT Sheet NAME is NULL!!!";
	}
	else {
		let intLenOptHdrShName = strHdrOutShtName.length;
		if (intLenOptHdrShName < 32) {
			//Verify we have the table.
			let weTable = CWCore.returnWebElement(strLocXpath + strTblXPath);
			if (weTable == null) {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location '" + strLocName + "' '" + strTblName + "' DID NOT RETURN an ELEMENT!!!";
			}
			else {
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weTable, 'Element', strTblName);
				}
				//Check if the header row is present
				let weHeaderRow = CWCore.returnWebElement(strLocXpath + strTblXPath + strHdrRowXPath);
				if (weHeaderRow == null) {
					boolPassed = false;
					strMethodDetails = "The location '" + strLocName + "' '" + strTblName + "' DID NOT RETURN A HEADER ROW ELEMENT.";
				}
				else {
					//Move to the row to show the row we are porcessing in case of errors.
					let actions = new Actions(TCObj.getTcDriver());
					actions.moveToElement(weHeaderRow).perform();
					//Highlight
					if (TCExecParams.getBoolDoHighlight() == true) {
						let mapHighlight = {}; // Mimic Groovy Map
						mapHighlight = CWCore.objHighlightElementJS(weHeaderRow, 'Element', strTblName + ' header row');
					}
					//Load the inputsheet
					let intInputRowCnt; // Type inferred
					let intInputColCnt; // Type inferred
					let intHdrColCnt; // Type inferred
					let lstHdrCols;	// Type inferred
					let shInput;	   // Type inferred
					let strHdrInputShtName = mapInputVariables.strHdrInputShtName;
					let mapOpenInputSheet = {}; // Mimic Groovy Map
					mapOpenInputSheet = ExcelData.excelGetSheetByName(TCObj.getObjWorkbook(), strHdrInputShtName);
					if (StringsAndNumbers.JComm_StringToBoolean(mapOpenInputSheet.boolPassed) == true) {
						shInput = mapOpenInputSheet.objWbSheet;
						//Return the row and column Count
						let mapSheetRowColCnt = {}; // Mimic Groovy Map
						mapSheetRowColCnt = ExcelData.excelGetRowAndColCount(shInput);
						if (StringsAndNumbers.JComm_StringToBoolean(mapSheetRowColCnt.boolPassed) == true) {
							intInputRowCnt = mapSheetRowColCnt.RowCount;
							intInputColCnt = mapSheetRowColCnt.ColCount;
						}
						else {
							boolPassed = false;
							strMethodDetails = mapSheetRowColCnt.strMethodDetails;
						}
						if (boolPassed == true) {
							//Return the header column count
							let mapHdrChildren = CWCore.returnChildElements(weHeaderRow, strHdrColXPath);
							intHdrColCnt = mapHdrChildren.cntChildObjs;
							lstHdrCols = mapHdrChildren.lstChildObjects;
						}
					}
					else {
						boolPassed = false;
						strMethodDetails = mapOpenInputSheet.strMethodDetails;
					}
					//Check if the number of columns displayed matches the number in the input sheet if not call output only and fail the step
					if (boolPassed == true) {
						if (intHdrColCnt != intInputRowCnt) {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The table header is displaying: " + intHdrColCnt + " BUT, EXPECTED : " + intInputRowCnt + " COLUMNS!!! OUTPUTTING HEADER ONLY see details: ";
						}
						else {
							//Return the input sheet column names and check if the match the output column names
							let strInputColValues; // Type inferred
							let mapGetInputColNames = {}; // Mimic Groovy Map
							mapGetInputColNames = ExcelData.excelGetHdrColNames(shInput);
							let boolInputColNames = StringsAndNumbers.JComm_StringToBoolean(mapGetInputColNames.boolPassed);
							if (boolInputColNames == true) {
								strInputColValues = mapGetInputColNames.ColValues;
								//Check if the match the output column names
								if (strInputColValues != strHdrOutputColNames) {
									boolPassed = false;
									strMethodDetails = "FAILED!!! The input column names '" + strInputColValues + "' DOES NOT MATCH the OUTPUT Columns '" + strHdrOutputColNames +"'!!!";
								}
							}
							else {
								boolPassed = false;
								strMethodDetails = mapGetInputColNames.strMethodDetails;
							}
						}
					}
					if (boolPassed == true) {
						//Create the output sheet
						let mapAddSheet = {}; // Mimic Groovy Map
						let strTCInputFilePath = TCObj.strTCInputFilePath;
						//TODO attempt to add the input stream to the test objects and use the input object to add the sheet in a new TObj WB and Sh
						mapAddSheet = ExcelData.excelAddSheetToWorkBook(strHdrOutShtName);
						let boolSheetAddedPassed = StringsAndNumbers.JComm_StringToBoolean (mapAddSheet.boolPassed);
						if (boolSheetAddedPassed == true) {
							//Add the column names
							let mapAddColNames = {}; // Mimic Groovy Map
							mapAddColNames= ExcelData.excelCreateHeaderCols(strHdrOutputColNames);
							let boolAddColNames = StringsAndNumbers.JComm_StringToBoolean (mapAddColNames.boolPassed);
							if (boolAddColNames == true) {
								//Return the header elements
								let lstColNames = mapAddColNames.lstColNames; // Type inferred
								//Update the header column(s) to the assigned theme
								let strAssgTheme = 'OutputDataHdrStd';
								let mapUpdateHdrTheme = {}; // Mimic Groovy Map
								mapUpdateHdrTheme = ExcelData.excelSetRowCellFormatByTheme(0, strAssgTheme);
								let boolUPdHdrTheme = StringsAndNumbers.JComm_StringToBoolean (mapUpdateHdrTheme.boolPassed);
								if (boolUPdHdrTheme == false) {
									boolPassed = false;
									strMethodDetails = mapUpdateHdrTheme.strMethodDetails;
								}
								//Process the input sheet and header
								let weHdrCol; // Type inferred
								let weTempHdrColValue; // Type inferred
								let weHdrColLabelContainer; // Type inferred
								let strInputTempColName; // Type inferred
								let strInputColError; // Type inferred
								let strInputTempValue; // Type inferred
								let strHdrTempValue; // Type inferred
								let strOutputCellTheme; // Type inferred
								let strCellTheme = 'OutputDataCentered'; // Type inferred
								let mapSetCellValue = {}; // Mimic Groovy Map
								let mapGetInputValue = {}; // Mimic Groovy Map
								let boolAttribPresent; // Type inferred
								let boolInputColFound; // Type inferred
								let intExcelRow;	   // Type inferred
								let intColIndex;	   // Type inferred
								let boolExitProcessing = false; // Type inferred
								//Cell counters
								let intCellPassed = 0; // Type inferred
								let intCellFailed = 0; // Type inferred
								let intCellSkipped = 0; // Type inferred
								let intCellIgnored = 0; // Type inferred
								let intChildCnt = -1; // Type inferred
								//Check for alternate name values
								let intAltColNameArrayCnt = -1; // Type inferred
								let intAltColNameUsed = -1; // Type inferred
								let arryAltColName; // Type inferred
								//let strTempAltColName; // Declared but unused
								if (strHdrColMissingNames != gblNull) {
									//Split the value into an array
									//ADD a split for the value to return the date/time format
									let mapArray = {}; // Mimic Groovy Map
									mapArray = StringsAndNumbers.JComm_StringToArray(strHdrColMissingNames, gblDelimiter);
									intAltColNameArrayCnt = StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount);
									arryAltColName = mapArray.ArryOfValues;
								}
								for (let loopHdrCols = 0; loopHdrCols < intHdrColCnt; loopHdrCols++) {
									weHdrCol = lstHdrCols[loopHdrCols];
									intExcelRow = loopHdrCols + 1; //Add one to output the correct row in excel
									//Highlight
									if (TCExecParams.getBoolDoHighlight() == true) {
										let mapHighlight = {}; // Mimic Groovy Map
										mapHighlight = CWCore.objHighlightElementJS(weHdrCol, 'Element', strTblName + ' header row column');
									}
									//Check if there is a child container
									if (strHdrLabelContainerXpath != gblNull) {
										weHdrColLabelContainer = null;
										intChildCnt = -1;
										intChildCnt = CWCore.returnWebElemChildElementCount(weHdrCol, strHdrLabelContainerXpath);
										if (intChildCnt == 1) {
											weHdrColLabelContainer = CWCore.returnChildElement(weHdrCol, strHdrLabelContainerXpath);
										}
									}
									if (weHdrColLabelContainer == null) {
										weTempHdrColValue = weHdrCol;
									}
									else {
										//Highlight
										if (TCExecParams.getBoolDoHighlight() == true) {
											let mapHighlight = {}; // Mimic Groovy Map
											mapHighlight = CWCore.objHighlightElementJS(weHdrColLabelContainer, 'Element', strTblName + ' header row column label container');
										}
										weTempHdrColValue = weHdrColLabelContainer;
									}
									boolInputColFound = false;
									strInputColError = gblNull;
									//Return the objects and values as assigned
									//Note: when strProperty* == gblNull we do not process the property or output any data
									//Case to process the hdr column against each of the columns in the input sheet
									for (let loopInputCol = 0; loopInputCol < intInputColCnt; loopInputCol++) {
										strInputTempColName = lstColNames[loopInputCol];
										strHdrTempValue = gblNull;
										switch (strInputTempColName) {
											case 'ColumnDragDrop':
												boolInputColFound = true;
												//Check if drag and drop is present in the cell
												if (strHdrColDragDropXPath != gblNull) {
													//Check if the drag and drop is present in the cell
													let cntDragDrop = CWCore.returnWebElemChildElementCount(weHdrCol, strHdrColDragDropXPath);
													if (cntDragDrop == 1) {
														strHdrTempValue = "elemDragDrop";
													}
													else if (cntDragDrop > 1){
														boolPassed = false;
														strMethodDetails = "FAILED!!!! The Header Column: " + loopHdrCols + " contained the 'DragDrop' count of: " + cntDragDrop + " WHICH DOES NOT MATCH THE EXPECTED 1!!!";
													}
													else {
														let cntActionBtn = CWCore.returnWebElemChildElementCount(weHdrCol, btnActionXPath);
														if(cntActionBtn == 1) {
															strHdrTempValue = "elemActionBtn";
														}
														else if (cntActionBtn == 0) {
															strHdrTempValue = gblNull;
														}
														else {
															boolPassed = false;
															strMethodDetails = "FAILED!!!! The Header Column: " + loopHdrCols + " contained the 'ActionBtn' count of: " + cntActionBtn + " WHICH DOES NOT MATCH THE EXPECTED 1!!!";
														}
													}
												}
												break;
											case 'ColumnCheckbox':
												boolInputColFound = true;
												//Check if checkbox is present in the cell
												if (strHdrColCheckBoxXPath != gblNull) {
													//Check if the drag and drop is present in the cell
													let cntChkBox = CWCore.returnWebElemChildElementCount(weHdrCol, strHdrColCheckBoxXPath);
													if (cntChkBox == 1) {
														strHdrTempValue = "elemCheckbox";
													}
													else if (cntChkBox > 1){
														boolPassed = false;
														strMethodDetails = "FAILED!!!! The Header Column: " + loopHdrCols + " contained the 'Checkbox' count of: " + cntChkBox + " WHICH DOES NOT MATCH THE EXPECTED 1!!!";
													}
													else {
														strHdrTempValue = gblNull;
													}
												}
												break;
											case 'ColumnName':
												boolInputColFound = true;
												//Return the column name
												if (strPropertyHdrColName != gblNull) {
													boolAttribPresent = CWCore.isAttribtuePresent(weHdrCol,strPropertyHdrColName);
													if (boolAttribPresent == true) {
														strHdrTempValue = StringsAndNumbers.JComm_HandleNoData(weHdrCol.getAttribute(strPropertyHdrColName));

													}
													//Check if we did not find a column name
													else if (intAltColNameArrayCnt >= 1 && intAltColNameArrayCnt > intAltColNameUsed ) {
														intAltColNameUsed ++;
														strHdrTempValue = StringsAndNumbers.JComm_HandleNoData(arryAltColName[intAltColNameUsed]);
													}
													else {
														strHdrTempValue = gblNull;
													}
													System.out.print(strHdrTempValue);
												}
												break;
											case 'ColumnText':
												boolInputColFound = true;
												strHdrTempValue = StringsAndNumbers.JComm_HandleNoData(weTempHdrColValue.getText());
												break;
											case 'ColumnIsSortable':
												boolInputColFound = true;
												//Return isSortable
												if (strPropertyHdrSortable != gblNull) {
													boolAttribPresent = CWCore.isAttribtuePresent(weTempHdrColValue,strPropertyHdrSortable);
													if (boolAttribPresent == true) {
														strHdrTempValue = StringsAndNumbers.JComm_HandleNoData(weTempHdrColValue.getAttribute(strPropertyHdrSortable));
														if (strHdrTempValue != gblNull && strPropertyHdrSortableValue != gblNull) {
															//Split the value into an array
															let intArrayCnt; // Type inferred
															let arryValues; // Type inferred
															let strValueLoc; // Type inferred
															let strValue; // Type inferred
															//ADD a split for the value to return the date/time format
															let mapArray = {}; // Mimic Groovy Map
															mapArray = StringsAndNumbers.JComm_StringToArray(strPropertyHdrSortableValue, gblDelimiter);
															intArrayCnt = StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount);
															arryValues = mapArray.ArryOfValues;
															if (intArrayCnt == 2) {
																strValueLoc = arryValues[0];
																strValue = arryValues[1];
																let intInString = strHdrTempValue.indexOf(strValue, 0);
																switch (strValueLoc){
																	case '*STARTWITH*':
																		if (intInString == 0) {
																			strHdrTempValue = 'true';
																		}
																		break;
																	case '*CONTAINS*':
																		if (intInString > 0) {
																			strHdrTempValue = 'true';
																		}
																		break;
																	default:
																		strHdrTempValue = 'false';
																		break;
																}
															}
															else {
																strHdrTempValue = 'UNKNOWN INPUT Value:' + strPropertyHdrSortableValue;
															}
														}
													}
													else {
														strHdrTempValue = 'false';
													}
												}
												break;
											case 'ColumnSort':
												boolInputColFound = true;
												//Return sorted
												if (strPropertyHdrSortOrder != gblNull) {
													let weTempSort; // Type inferred
													if (strSortedOrderElementName == 'HdrLabelContainer') {
														weTempSort = weTempHdrColValue;
													}
													else {
														weTempSort = weHdrCol;
													}
													boolAttribPresent = CWCore.isAttribtuePresent(weTempSort,strPropertyHdrSortOrder);
													if (boolAttribPresent == true) {
														strHdrTempValue = StringsAndNumbers.JComm_HandleNoData(weTempSort.getAttribute(strPropertyHdrSortOrder));
														if (strPropertyHdrSortOrderValue != gblNull) {
															//Split the value into and array and check for each value
															let intArrayCnt; // Type inferred
															let arryValues; // Type inferred
															//ADD a split for the value
															let mapArray = {}; // Mimic Groovy Map
															mapArray = StringsAndNumbers.JComm_StringToArray(strPropertyHdrSortOrderValue, gblDelimiter);
															intArrayCnt = StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount);
															arryValues = mapArray.ArryOfValues;
															if (intArrayCnt == 3) {
																let strColSorted = arryValues[0];
																if (strColSorted == strHdrTempValue || StringsAndNumbers.JComm_TextLocationInString(strHdrTempValue, strColSorted) >= 0) {
																	//Return the other values
																	let strValueSortAsc = arryValues[1];
																	let strValueSortDesc = arryValues[2];
																	//Use the child object property to return the sort. Assumes the child object
																	boolAttribPresent = CWCore.isAttribtuePresent(weTempHdrColValue,strPropertyHdrSortOrder);
																	if (boolAttribPresent == true) {
																		let strTempSortOrder = StringsAndNumbers.JComm_HandleNoData(weTempHdrColValue.getAttribute(strPropertyHdrSortOrder));
																		// let intIndexSortAsc = strTempSortOrder.indexOf(strValueSortAsc); // Declared but unused
																		// let intIndexSortDesc = strTempSortOrder.indexOf(strValueSortDesc); // Declared but unused
																		if (strTempSortOrder.indexOf(strValueSortAsc) >= 0) {
																			strHdrTempValue = "Ascending";
																		}
																		else if(strTempSortOrder.indexOf(strValueSortDesc) >= 0) {
																			strHdrTempValue = "Descending";
																		}
																		else {
																			strHdrTempValue = "UNKNOWN SORT, Value of: " + strTempSortOrder + " DOES NOT CONTAIN ASC OR DESC DEFINED VALUES!!!";
																		}
																	}
																	else {
																		strHdrTempValue = gblNull;
																	}
																}
																else {
																	strHdrTempValue = gblNull;
																}
															}
															else {
																strHdrTempValue = "UNKNOWN SORT NOT ENOUGH ITEMS IN VALUE!!!";
															}
														}
													}
													else {
														strHdrTempValue = gblNull;
													}
												}
												break;
											case 'ColumnTooltipText':
												boolInputColFound = true;
												//tooltip text
												if (strHdrColSortToolTipXPath != gblNull) {
													//Get the column child elements
													let weToolTip; // Type inferred
													let mapGetToolTipElements = CWCore.returnChildElements(weTempHdrColValue, strHdrColSortToolTipXPath);
													let lstToolTip = mapGetToolTipElements.lstChildObjects;
													let cntToolTip = mapGetToolTipElements.cntChildObjs;
													if (cntToolTip == 1) {
														weToolTip = lstToolTip[0];
														boolAttribPresent = CWCore.isAttribtuePresent(weToolTip,strPropertyHdrColTextToolTip);
														if (boolAttribPresent == true) {
															strHdrTempValue = StringsAndNumbers.JComm_HandleNoData(weToolTip.getAttribute(strPropertyHdrColTextToolTip));
														}
														else {
															strHdrTempValue = gblNull;
														}
													}
													else if (cntToolTip == 0) {
														strHdrTempValue = "No ToolTip";
													}
													else {
														strHdrTempValue = "'ToolTip' count of: " + cntToolTip;
													}
												}
												else {
													boolAttribPresent = CWCore.isAttribtuePresent(weTempHdrColValue,strPropertyHdrColTextToolTip);
													if (boolAttribPresent == true) {
														strHdrTempValue = StringsAndNumbers.JComm_HandleNoData(weTempHdrColValue.getAttribute(strPropertyHdrColTextToolTip));
													}
													else {
														strHdrTempValue = gblNull;
													}
												}
												break;
											default:
												strInputColError = "FAILED!!! The Excel Column Name: " + strInputTempColName + ' IS NOT DEFINED IN THE VERIFY HEADER METHOD!!!"';
										}
										if (boolInputColFound == true) {
											//Return the input value and check if we match
											let mapGetInputValue = {}; // Mimic Groovy Map
											mapGetInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, intExcelRow, strInputTempColName);
											let boolGetInpValue = StringsAndNumbers.JComm_StringToBoolean(mapGetInputValue.boolPassed);
											if (boolGetInpValue == true) {
												strInputTempValue = StringsAndNumbers.JComm_HandleNoData(mapGetInputValue.CellValue);
												intColIndex = mapGetInputValue.intColIndex;
											}
											else {
												boolPassed = false;
												strMethodDetails = mapGetInputValue.strMethodDetails;
											}
											if (strInputTempValue == gblSkip || strInputTempValue == gblIgnoreData) { //TODO should these be separated?
												strOutputCellTheme = 'TestRunDataSkipIgnore';
												if (strInputTempValue == gblSkip) {
													intCellSkipped++;
												}
												else {
													intCellIgnored++;
												}
											}
											if (strInputTempValue == strHdrTempValue) {
												strOutputCellTheme = 'TestRunPassStd';
												intCellPassed++;
											}
											else {
												strOutputCellTheme = 'TestRunFailStd';
												strHdrTempValue = "Expected: " + strInputTempValue + "HdrDisplayed: " + strHdrTempValue;
												boolPassed = false;
												intCellFailed++;
											}
											//Output the results to the cell
											mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intExcelRow, intColIndex, strHdrTempValue, strOutputCellTheme);
											if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
												boolPassed = false;
												strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
											}
										}
										else {
											boolPassed = false;
											strMethodDetails = strMethodDetails + strInputColError;
											boolExitProcessing = true;
											break;
										}
									}
									if (boolExitProcessing == true) {
										strMethodDetails = "FAILED!!! Processing STOPPED SEE DETAILS: " + strMethodDetails;
										break;
									}
									Tester.Message("End of Display Column: " + loopHdrCols);
								}
								//Set the column width and list how many cells passed, failed, skipped or ignored.
								let mapAutoFit = {}; // Mimic Groovy Map
								mapAutoFit = ExcelData.excelAutofitCols();
								if (StringsAndNumbers.JComm_StringToBoolean(mapAutoFit.boolPassed) == false) {
									boolPassed = false; // Fixed == false to = false
									strMethodDetails = strMethodDetails + mapAutoFit.strMethodDetails;
								}
								let intTotalCells = intCellPassed + intCellFailed + intCellSkipped + intCellIgnored;
								//For debug only
								strMethodDetails = "Processed " + intTotalCells + " total cells, passed: " + intCellPassed + ", failed: " +
								intCellFailed + ", skipped: " + intCellSkipped + ", and ignored: " + intCellIgnored + " cells. See OutputSheet: " + strHdrOutShtName + " for details.";
							}
							else {
								boolPassed = false;
								strMethodDetails = mapAddColNames.strMethodDetails;
							}
						}
						else {
							boolPassed = false;
							strMethodDetails = mapAddSheet.strMethodDetails;
						}
					}
					else {
						//Call the output header only
						let mapOutputHeader = {}; // Mimic Groovy Map
						mapOutputHeader = OutputResultHeader(mapInputVariables); // Call the function
						strMethodDetails = strMethodDetails + mapOutputHeader.strMethodDetails;
					}
				}
			}
		}
		else {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The assigned output sheet name '" + strHdrOutShtName + " is " + intLenOptHdrShName + " CHAR LENGTH WHICH EXCEEDS EXCEL TAB NAME LENGTH OF 31!!!!";
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------   OutputResultBody   -----------------------------------
* Output the body of the table results
* @param mapInputVariables The input variables that identify the header to include xpaths and properties and table results rows and columns
* @param NOTE: Each cell will be checked for the elements listed in the strResultsColObjNames
*
* @param strLocName						The pipe delimited location value
* @param strLocXpath						The parent Xpath
* @param strTblName						The name of the table
* @param strTblXPath						The table Xpath
* @param strHdrRowXPath					The Header row Xpath usually //thead//tr
* @param strHdrColXPath					The header column Xpath usually //th
* @param strPropertyHdrColName				The cell property that will be used to find the assigned column name. Usually 'abbr'
* @param strResultsOutShtName				The sheet name assigned for the output
* @param strResultsRowXPath				The Xpath for the rows within the results
* @param strResultsColXPath				The Xpath for the columns within the result rows
* @param arryElementNames					The array containing the list of element names that can be in the results cells.
* @param arryElemXPaths					The array containing the list of elemenT Xpaths that can be in the results cells.
* @param NOTE: the TWO Arrays must contain the same number of items
* @param strResultsCellStatElem			The name of the status element
* @param //Check for the alternate values
* @param strTblSepXpath					The XPath for the body section of the table. Should only be used if the body tag appears in a cell of the results.
* @param strHdrColMissingNames				For each column that does not have an identifiable name, add the name to the delimited value
* @param strMultElemColData				The delimited value with the name|count of elements where a cell contains more than one element. Currently only one column is supported
* @param //UL alternate values see: Sourcing|Templates and Libraries|Event Libraries for example
* @param strResultsCellULElemName			The user list element name
* @param strResultsCellULNoValueElemName	The user cell list item No Value element name
* @param arryULElementNames				The User List Array that contains the child element names
* @param arryULElemXPaths					The User List Array that contains the child element XPaths
*
* @param //Alternate Value for: Column contains a container with multiple elements
* @param strContainerMultElementInfo		The value will be the container object name|elementName i.e. tblSuppliersImageGroup|tblSuppliersGroupImg
*
* @param //Alternate values for body and header in one object i.e. thead see JI User Profile|User Roles and Access|Access which contains the header, data and subsection rows as data JTR-19321
* @param intTableHeaderRowCnt				The number of rows that contain header data
* @param strSectionRowInfo					The information about the section rows in a delimited format. rowProperty, rowValue, intCols. i.e. "class|SubSectionTitle|1" Must contain 3 values replace gblNotAblicable if value is not required or present.
*
* @param //Alternate values for image groups and elements in Elastic Search Results see JTR-20847-
* @param strGRIElemName					The Image group element name
* @param strGRIItemName					The name for the individual image element
* @param strGRIItemXPath					The Xpath for the image element within the group
* @param strGRIItemChildXPath				The Xpath for the child of the item which should contain the image. Multiple tags and class may be present.
* @param strGRIImgAttribute				The Image attribute that contains the value for the image for example the title'
*
* @param //Alternate value for the image elements attribute
* @param strImageAttribute					The attribute that will contain the value of the image
*
* @param //Alternate values for row is not displayed JTR-22460
* @param strTableRowHiddenValues			The table row is hidden when the attribute has the value specified. Contains the attribute | value pair. Enter gblNull if not present
*
* @param //Alternate value for custom Widget table on AP Home
* @param strElemContainsSVGNameXPath		The Xpath for the SVG Name if not in the SVG element. ./.. for the parent element or gblNull if not required
* @param strElemContainsSVGNameProperty 	The element property where the SVG name will be present i.e. class leave gblNull if not present

* @return mapResults 		The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 04/30/2022
* @author Last Edited: 06/27/2023
* @author Last Edited By: PGK
* @author Edit Comments: (Include email, date and details)
* @author updated to support table structure where the cell can contain a UL with li containers that may have 1 or more elements and values
* @author pkanaris added support for tables where the header and body are in one tag and the abiltiy for the table to contain a section row JTR-19321 05/28/2023
* @author pkanaris added support for tables which have image groups that contain 1 to many images with values. JTR-20847
*/
function CommonWeb_OutputResultBody (mapInputVariables) { // mapInputVariables and return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	//Return the method variables
	let strLocName = mapInputVariables.LocName;
	let strLocXpath = mapInputVariables.LocXPath;
	let strTblName = mapInputVariables.TblName;
	let strTblXPath = mapInputVariables.strTblXPath;
	let strHdrRowXPath = mapInputVariables.strHdrRowXPath;
	let strHdrColXPath = mapInputVariables.strHdrColXPath;
	let strPropertyHdrColName = mapInputVariables.strPropertyHdrColName;
	let strResultsOutShtName = mapInputVariables.strResultsOutShtName;
	let strResultsRowXPath = mapInputVariables.strResultsRowXPath;
	let strResultsColXPath = mapInputVariables.strResultsColXPath;
	//Cells can possibly have only one element
	let arryElementNames = mapInputVariables.arrayResultsColObjNames; // Type inferred
	let arryElemXPaths = mapInputVariables.arrayResultsColObjXPaths; // Type inferred
	let strResultsCellStatElem = mapInputVariables.strResultsCellStatElem;
	//Check for the alternate values
	let strTblSepXpath = gblNull;
	if (mapInputVariables.hasOwnProperty('strTblSepXpath')){ // Use hasOwnProperty
		strTblSepXpath = mapInputVariables.strTblSepXpath;
	}
	let strHdrColMissingNames = gblNull;
	let strMultElemColData = gblNull;
	if (mapInputVariables.hasOwnProperty('strHdrColMissingNames')){
		strHdrColMissingNames = mapInputVariables.strHdrColMissingNames;
	}
	if (mapInputVariables.hasOwnProperty('strMultElemColData')){
		strMultElemColData = mapInputVariables.strMultElemColData;
	}
	//Alternate values for row is not displayed JTR-22460
	let strTableRowHiddenAttribute = gblNull;
	let strTableRowHiddenValue = gblNull;
	if (mapInputVariables.hasOwnProperty('strTableRowHiddenValues')) { // Use hasOwnProperty
		let strTempHiddenRowValue = StringsAndNumbers.JComm_HandleNoData(mapInputVariables.strTableRowHiddenValues);
		if (strTempHiddenRowValue != gblNull) {
			//Left of pipe is the attribute
			strTableRowHiddenAttribute = StringsAndNumbers.JComm_GetLeftTextInString(strTempHiddenRowValue, gblDelimiter);
			//Right of pipe is the value
			strTableRowHiddenValue = StringsAndNumbers.JComm_GetRightTextInString(strTempHiddenRowValue, gblDelimiter);
		}
	}
	//UL alternate values see: Sourcing|Templates and Libraries|Event Libraries for example
	let strULElemName = gblNull;
	let strULNoValElemName = gblNull;
	let arryULElementNames; // Type inferred
	let arryULElemXPaths; // Type inferred
	if (mapInputVariables.hasOwnProperty('strResultsCellULElemName')){
		strULElemName = mapInputVariables.strResultsCellULElemName;
	}
	if (mapInputVariables.hasOwnProperty('strResultsCellULNoValueElemName')){
		strULNoValElemName = mapInputVariables.strResultsCellULNoValueElemName;
	}
	if (mapInputVariables.hasOwnProperty('arrayResultsColULObjNames')){
		arryULElementNames = mapInputVariables.arrayResultsColULObjNames;
	}
	if (mapInputVariables.hasOwnProperty('arrayResultsColULObjXPaths')){
		arryULElemXPaths = mapInputVariables.arrayResultsColULObjXPaths;
	}
	//Alternate values for Group-record-icons used in Supplier Search Results new UI Jun 2023 JTR-20847
	let strGRIElemName = gblNull;
	let strGRIItemName = gblNull;
	let strGRIItemXPath = gblNull;
	let strGRIItemChildXPath = gblNull;
	let strGRIImgAttribute = gblNull;
	if (mapInputVariables.hasOwnProperty('strGRIElemName')){
		strGRIElemName = mapInputVariables.strGRIElemName;
	}
	if (mapInputVariables.hasOwnProperty('strGRIItemName')){
		strGRIItemName = mapInputVariables.strGRIItemName;
	}
	if (mapInputVariables.hasOwnProperty('strGRIItemXPath')){
		strGRIItemXPath = mapInputVariables.strGRIItemXPath;
	}
	if (mapInputVariables.hasOwnProperty('strGRIItemChildXPath')){
		strGRIItemChildXPath = mapInputVariables.strGRIItemChildXPath;
	}
	if (mapInputVariables.hasOwnProperty('strGRIImgAttribute')){
		strGRIImgAttribute = mapInputVariables.strGRIImgAttribute;
	}
	//Alternate value for the image elements attribute
	let strImageAttribute = gblNull;
	if (mapInputVariables.hasOwnProperty('strImageAttribute')){
		strImageAttribute = mapInputVariables.strImageAttribute;
	}
	//Alternate value for custom Widget table on AP Home
	let strElemContainsSVGNameXPath = gblNull;
	let strElemContainsSVGNameProperty = gblNull;
	if (mapInputVariables.hasOwnProperty('strElemContainsSVGNameXPath')){
		strElemContainsSVGNameXPath = mapInputVariables.strElemContainsSVGNameXPath;
	}
	if (mapInputVariables.hasOwnProperty('strElemContainsSVGNameProperty')){
		strElemContainsSVGNameProperty = mapInputVariables.strElemContainsSVGNameProperty;
	}

	//Add for container in a cell with 1 to many items. See Supplier Search Results New UI for images
	let strContainerMultElementInfo = gblNull;
	let strContainerName = gblNull; // Type inferred
	let strContainElemName = gblNull; // Type inferred
	if (mapInputVariables.hasOwnProperty('strContainerMultElementInfo')){
		strContainerMultElementInfo = mapInputVariables.strContainerMultElementInfo;
		strContainerName = StringsAndNumbers.JComm_GetLeftTextInString(strContainerMultElementInfo, gblDelimiter);
		strContainElemName = StringsAndNumbers.JComm_GetRightTextInString(strContainerMultElementInfo, gblDelimiter);
	}
	//Add alternate for table where header and data is in one cell and a sub category row is present
	let intTableHeaderRowCnt = -1;
	let strSectionRowInfo = gblNull;
	let arrySectInfo; // Type inferred
	let intSectInfoArrySize; // Type inferred, not used in final code
	if (mapInputVariables.hasOwnProperty('intTableHeaderRowCnt')) {
		intTableHeaderRowCnt = mapInputVariables.intTableHeaderRowCnt;
	}
	if (mapInputVariables.hasOwnProperty('strSectionRowInfo')) {
		strSectionRowInfo = mapInputVariables.strSectionRowInfo;
		if (strSectionRowInfo == gblNull || strSectionRowInfo == gblNA) {
			strSectionRowInfo = gblNull; //Force to global null
		}
		else {
			let mapSplitSubSectInfoString = StringsAndNumbers.JComm_StringToArray(strSectionRowInfo, gblDelimiter);
			let intCntValues = StringsAndNumbers.JComm_StringToInteger(mapSplitSubSectInfoString.intItemCount);
			if (intCntValues == 4) { // Groovy original comment says 3, but expects 4 from its example usage.
				arrySectInfo = mapSplitSubSectInfoString.ArryOfValues;
				intSectInfoArrySize = arrySectInfo.length; // Use .length for JS Array
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The Results can contain a Sub Section row, BUT, THE info provided '" + strSectionRowInfo + "', DOES NOT CONTAIN 3 ITEMS!!!";
			}
		}
	}
	let cntDispCols; // Type inferred (number) //We need to capture the column count or expected count based on missing column names
	//Add variable for elements in the field for status? Check the arryElementNames and Xpaths.
	//Add variable for columnName|XPath semicolon separated values pair. This will be for multiple objects in a cell. Should be column specific
	//Check if the outputsheet name is valid
	let intLenOptResShName = strResultsOutShtName.length;
	if (intLenOptResShName < 32) {
		//Check for alternate name values
		let intAltColNameArrayCnt = -1;
		let intAltColNameUsed = -1;
		let arryAltColName; // Type inferred
		//let strTempAltColName; // Declared but not used
		if (strHdrColMissingNames != gblNull) {
			//Split the value into an array
			//ADD a split for the value to return the date/time format
			let mapArray = StringsAndNumbers.JComm_StringToArray(strHdrColMissingNames, gblDelimiter);
			intAltColNameArrayCnt = StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount);
			arryAltColName = mapArray.ArryOfValues;
		}
		//Check for columns with more than one element
		let strMultElemColName = gblNull;
		let intMultElemColCount = 1; //always at least 1
		if (strMultElemColData != gblNull) {
			//Split the value into an array
			//ADD a split for the value to return the date/time format
			let mapArray = StringsAndNumbers.JComm_StringToArray(strMultElemColData, gblDelimiter);
			if (StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount) == 2) {
				strMultElemColName = mapArray.ArryOfValues[0];
				intMultElemColCount = StringsAndNumbers.JComm_StringToInteger(mapArray.ArryOfValues[1]);
			}
		}
		//Verify we have the table.
		let weTable = CWCore.returnWebElement(strLocXpath + strTblXPath);
		if (weTable == null) {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location '" + strLocName + "' '" + strTblName + "' DID NOT RETURN an ELEMENT!!!";
		}
		else {
			//Highlight
			if (TCExecParams.getBoolDoHighlight() == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weTable, 'Element', strTblName);
			}
			//Check if the header row is present
			let weHeaderRow = CWCore.returnWebElement(strLocXpath + strTblXPath + strHdrRowXPath);
			if (weHeaderRow == null) {
				boolPassed = false;
				strMethodDetails = "FAILED!!!The location '" + strLocName + "' '" + strTblName + "' DID NOT RETURN A HEADER ROW ELEMENT!!!";
			}
			else {
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weHeaderRow, 'Element', strTblName + ' header row');
				}
				//Note, we open the file and do the work for each method and save/close the excel instance so save happens as we go.
				//Add the sheet to the input file
				let mapAddSheet = {}; // Mimic Groovy Map
				let strTCInputFilePath = TCObj.strTCInputFilePath;
				//Add the output sheet
				mapAddSheet = ExcelData.excelAddSheetToWorkBook(strResultsOutShtName);
				let boolSheetAddedPassed = StringsAndNumbers.JComm_StringToBoolean (mapAddSheet.boolPassed);
				if (boolSheetAddedPassed == true) {
					//Add the column names
					//Return all the column names and use to create the output sheet header columns
					let mapGetHdrChildElements = {}; // Mimic Groovy Map
					let strOutputColNames; // Type inferred
					mapGetHdrChildElements = CWCore.returnChildElements(weHeaderRow, strHdrColXPath);
					let lstHdrCols = mapGetHdrChildElements.lstChildObjects; // Type inferred (List<WebElement>)
					let arryColumnNames; // Type inferred, used later.
					let strHdrColNames; // Type inferred
					let cntHdrCols = mapGetHdrChildElements.cntChildObjs; // Type inferred
					if (cntHdrCols > 0) {
						cntDispCols = cntHdrCols; //TODO what happens if there is no header PGK 05/26/2023
						let weHdrCol; // Type inferred
						let strTempValue; // Type inferred
						let boolAttribPresent; // Type inferred
						//Return the column names and replace any spaces with '_'
						for (let loopHdrCols = 0; loopHdrCols < cntHdrCols; loopHdrCols++) {
							weHdrCol = lstHdrCols[loopHdrCols]; // Access array by index
							//Highlight
							if (TCExecParams.getBoolDoHighlight() == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weHdrCol, 'Element', strTblName + ' header row column');
							}
							boolAttribPresent = CWCore.isAttribtuePresent(weHdrCol,strPropertyHdrColName);
							strTempValue = gblNull;
							if (boolAttribPresent == true) {
								strTempValue = StringsAndNumbers.JComm_HandleNoData(weHdrCol.getAttribute(strPropertyHdrColName));
							}
							//Check if we did not find a column name
							else if (intAltColNameArrayCnt >= 1 && intAltColNameArrayCnt > intAltColNameUsed ) {
								intAltColNameUsed ++;
								strTempValue = StringsAndNumbers.JComm_HandleNoData(arryAltColName[intAltColNameUsed]);
							}
							if (strTempValue == gblNull) {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The header column '" + loopHdrCols +"' does not CONTAIN THE PROPERTY REQUIRED of:" + strPropertyHdrColName +"!!!";
								break; // Break loop
							}
							else {
								if (loopHdrCols == 0) {
									strOutputColNames = StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_");
									strHdrColNames = strTempValue;
								}
								else {
									strOutputColNames = strOutputColNames + gblDelimiter + StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_");
									strHdrColNames = strHdrColNames + gblDelimiter + strTempValue;
								}
								strOutputColNames = strOutputColNames + gblDelimiter + StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_") + "_Element"; //Add the column for the element
							}
						}
						if (boolPassed == true) {
							//Add column names to the array
							let mapStringToArry = StringsAndNumbers.JComm_StringToArray(strHdrColNames, gblDelimiter);
							arryColumnNames = mapStringToArry.ArryOfValues;
							let mapAddColNames = ExcelData.excelCreateHeaderCols(strOutputColNames);
							let boolAddColNames = StringsAndNumbers.JComm_StringToBoolean (mapAddColNames.boolPassed);
							if (boolAddColNames == true) {
								//Return the header elements
								//let lstColNames = mapAddColNames.lstColNames; // Declared but unused in Groovy
								//Update the header column(s) to the assigned theme
								let strAssgTheme = 'OutputDataHdrStd';
								let mapUpdateHdrTheme = ExcelData.excelSetRowCellFormatByTheme(0, strAssgTheme);
								let boolUPdHdrTheme = StringsAndNumbers.JComm_StringToBoolean (mapUpdateHdrTheme.boolPassed);
								if (boolUPdHdrTheme == false) {
									boolPassed = false;
									strMethodDetails = mapUpdateHdrTheme.strMethodDetails;
								}
							}
							else { // mapAddColNames failed
								boolPassed = false;
								strMethodDetails = mapAddColNames.strMethodDetails;
							}
						}
						//Return the data and update the output sheet
						if (boolPassed == true) { // If all initial checks passed
							let boolProcCellElements; // Type inferred
							if (arryElementNames == null || arryElemXPaths == null) {
								boolProcCellElements = false;
							}
							else if(arryElementNames.length != arryElemXPaths.length){ // Use .length for JS Array
								boolProcCellElements = false;
								Tester.Message("THE arryElementNames and arryElemXPaths SIZE DO NOT MATCH!!!" );
							}
							else {
								boolProcCellElements = true;
							}
							let intOutputExcelRow; // Type inferred
							let intBodyRowColCnt; // Type inferred
							let intOutputColumn; // Type inferred
							let intCellElemCnt; // Type inferred
							let strTempCellValue; // Type inferred
							let strTempCellElements; // Type inferred
							let weRow; // Type inferred (WebElement)
							let weCell; // Type inferred (WebElement)
							let mapGetRowCols; // Type inferred
							let lstBodyRowCols; // Type inferred (List<WebElement>)
							let mapSetCellValue; // Type inferred
							let strCellTheme = 'OutputDataLeftJustified'; // Type inferred
							//Check if the tbody is a separate XPath
							if (strTblSepXpath != gblNull) {
								//Return the body as a new table element
								weTable = CWCore.returnWebElement(strLocXpath + strTblXPath + strTblSepXpath);
								//Highlight
								if (TCExecParams.getBoolDoHighlight() == true) {
									let mapHighlight = {}; // Mimic Groovy Map
									mapHighlight = CWCore.objHighlightElementJS(weTable, 'Table Body', strTblName);
								}
							}
							//Return the row count and create the loop
							let mapGetRowElements = CWCore.returnChildElements(weTable, strResultsRowXPath);
							let lstBodyRows = mapGetRowElements.lstChildObjects;
							let intBodyRowCnt = mapGetRowElements.cntChildObjs;
							if (intBodyRowCnt == 0) {
								//Error
								boolPassed = false;
								strMethodDetails = "FAILED!!! The table contain 'ZERO'ROW(S)!!!!";
							}
							else {
								let intLoopStart = 0;
								if (intTableHeaderRowCnt > 0 && strTblSepXpath == gblNull) {
									intLoopStart = intTableHeaderRowCnt;
								}
								intOutputExcelRow = 0;
								for (let loopRow = intLoopStart; loopRow < intBodyRowCnt && boolPassed; loopRow++) { // Added boolPassed check
									//For each row count the cells and create the loop. May be different than the header columns.
									weRow = lstBodyRows[loopRow]; // Access array by index
									let strTempAttrValue = gblNull; // Type inferred
									//TODO Implement the hidden row for JTR-22460
									if (strTableRowHiddenAttribute != gblNull) {
										//Return the attribute value
										strTempAttrValue = weRow.getAttribute(strTableRowHiddenAttribute);
									}
									if (strTempAttrValue != strTableRowHiddenValue || strTempAttrValue == gblNull) {
										//Move to the row to show the row we are processing in case of errors.
										let actions = new Actions(TCObj.tcDriver);
										actions.moveToElement(weRow).perform();
										//Highlight
										if (TCExecParams.getBoolDoHighlight() == true) {
											let mapHighlight = {}; // Mimic Groovy Map
											mapHighlight = CWCore.objHighlightElementJS(weRow, 'Element', strTblName + ' body row');
										}
										//process the row otherwise ignore the row
										//Inject change for JTR-19321
										//Return the SubSection values if present
										let intSubSecColCnt = -1; // Type inferred
										let strColProp = gblNull; // Type inferred
										let strColPropValue = gblNull; // Type inferred
										let strColElemName = gblNull; // Type inferred
										if (strSectionRowInfo != gblNull) {
											intSubSecColCnt = StringsAndNumbers.JComm_StringToInteger(arrySectInfo[0]); // Access array by index
											strColProp = arrySectInfo[1]; // Access array by index
											strColPropValue = arrySectInfo[2]; // Access array by index
											strColElemName = arrySectInfo[3]; // Access array by index
										}
										//1|class|SubSectionTitle|elemSubSection
										//Check if displayed column count matches subsection and if so process
										//Return the row count and create the loop
										mapGetRowCols = CWCore.returnChildElements(weRow, strResultsColXPath);
										lstBodyRowCols = mapGetRowCols.lstChildObjects;
										intBodyRowColCnt = mapGetRowCols.cntChildObjs;
										if (intBodyRowColCnt == intSubSecColCnt) { // If it's a section row
											let boolIsSection = false; // Type inferred
											intOutputExcelRow ++;
											intOutputColumn = 0;
											//Process the subSection
											for (let loopRowCol = 0; loopRowCol < cntDispCols && boolPassed; loopRowCol++) { // Added boolPassed check
												strTempCellValue = gblNull;
												strTempCellElements = gblNull;
												if (loopRowCol < intSubSecColCnt) {
													//Return the cell object
													weCell = lstBodyRowCols[loopRowCol]; // Access array by index
													//Highlight
													if (TCExecParams.getBoolDoHighlight() == true) {
														let mapHighlight = {}; // Mimic Groovy Map
														mapHighlight = CWCore.objHighlightElementJS(weCell, 'Element', strTblName + ' body cell');
													}
													//Return the cell text
													strTempCellValue = StringsAndNumbers.JComm_HandleNoData(weCell.getText());
													//Confirm the cell belongs to the subsection
													boolAttribPresent = CWCore.isAttribtuePresent(weCell,strColProp);
													if (boolAttribPresent == true) {
														let strTempValue = StringsAndNumbers.JComm_HandleNoData(weCell.getAttribute(strColProp));
														if ( StringsAndNumbers.JComm_VerifyTextPresent(strTempValue, strColPropValue, "%Con%") == true) {
															boolIsSection = true;
															strTempCellElements = strColElemName;
														}
													}
												}
												//Output the values
												//Output the cell text after the elements.
												mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intOutputExcelRow, intOutputColumn, strTempCellValue, strCellTheme);
												if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == true) {
													intOutputColumn++;
													mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intOutputExcelRow, intOutputColumn, strTempCellElements, strCellTheme);
													if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == true) {
														intOutputColumn++;
													}
													else {
														boolPassed = false;
														strMethodDetails = mapSetCellValue.strMethodDetails;
													}
												}
												else {
													boolPassed = false;
													strMethodDetails = mapSetCellValue.strMethodDetails;
												}
											}
											if (boolIsSection == false) {
												boolPassed = false;
												strMethodDetails = `FAILED!!! Result row: ${loopRow + 1} CELL(s) PROPERTY DID NOT CONTAIN The VALUE '${strColPropValue}' !!!`; // Using Template Literal
											}
										}
										//Check if a spacer row is present with no columns
										else if (intBodyRowColCnt > 0) { // If it's a data row
											intOutputExcelRow ++;
											intOutputColumn = 0;
											//For each cell return the value and save to the value column based on the header columns
											for (let loopRowCol = 0; loopRowCol < intBodyRowColCnt && boolPassed; loopRowCol++) { // Added boolPassed check
												weCell = lstBodyRowCols[loopRowCol]; // Access array by index
												//Highlight
												if (TCExecParams.getBoolDoHighlight() == true) {
													let mapHighlight = {}; // Mimic Groovy Map
													mapHighlight = CWCore.objHighlightElementJS(weCell, 'Element', strTblName + ' body cell');
												}
												//Return the cell text
												strTempCellValue = StringsAndNumbers.JComm_HandleNoData(weCell.getText());
												if (boolProcCellElements == true) {
													//Set the wait to 1 second to speed up the processing of elements in the cells
													TCObj.tcDriver.manage().timeouts().implicitlyWait(100, TimeUnit.MILLISECONDS);
													let intMaxElem = 1; // Type inferred
													if (strMultElemColName == arryColumnNames[loopRowCol]) { // Access array by index
														intMaxElem = intMultElemColCount;
													}
													let cntCellElements = 0; // Type inferred
													strTempCellElements = gblNull; // Type inferred
													let boolCustomTheme = false; // Type inferred
													let strColor; // Type inferred
													let strBckgColor; // Type inferred
													let weCellElement; // Type inferred (WebElement)
													while ( cntCellElements < intMaxElem && boolPassed) { // Added boolPassed check
														//For each cell check to see if any of the expected elements are present and output
														//Process the array for elements
														for (let loopElements = 0; loopElements < arryElementNames.length && boolPassed; loopElements++) { // Added boolPassed check
															//TODO add in classes to get color such as status elements
															let strTempElemXpath = arryElemXPaths[loopElements]; // Access array by index
															if (strTempElemXpath == gblUndefined) {
																//Fail the method
																boolPassed = false;
																strMethodDetails = "FAILED!!! The " + arryElementNames[loopElements] + " RETURNED the VALUE : " + gblUndefined + "!!!";
																break; // Break loop
															}
															intCellElemCnt = CWCore.returnWebElemChildElementCount(weCell, arryElemXPaths[loopElements]); // Access array by index
															if (intCellElemCnt == 1) {
																cntCellElements++; //Increment the number of elements in the cell
																weCellElement = CWCore.returnChildElement(weCell, arryElemXPaths[loopElements]); // Access array by index
																//Highlight
																if (TCExecParams.getBoolDoHighlight() == true) {
																	let mapHighlight = {}; // Mimic Groovy Map
																	mapHighlight = CWCore.objHighlightElementJS(weCellElement, 'Element', strTblName + ' cell status element');
																}
																if (cntCellElements == 1) { // Only first element detected
																	//Check if the element is a UL
																	//TODO replace the ul with a new variable
																	if (arryElementNames[loopElements] == strULElemName) { //See Sourcing|Templates and Libraries|Event Libraries for UI example
																		//TODO we need to pass in the ul element names and xpaths
																		//let intItemCnt = 0; // Declared but not used
																		//let strItemValue; // Declared but not used
																		//let strItemElem; // Declared but not used
																		let weULItem; // Type inferred (WebElement)
																		//Process the UL to return the items
																		let mapULItems = CWCore.returnChildElements(weCellElement, arryULElemXPaths[0]); // Access array by index
																		let lstULItems = mapULItems.lstChildObjects; // Type inferred (List<WebElement>)
																		let intULItemCnt = mapULItems.cntChildObjs; // Type inferred
																		if (intULItemCnt == 0) {
																			//Error
																			boolPassed = false;
																			strMethodDetails = "FAILED!!! The Cell User List contains 'ZERO' ITEM(S)!!!!";
																		}
																		else {
																			let weULItemChild; // Type inferred (WebElement)
																			for (let loopULItemsInner = 0; loopULItemsInner < intULItemCnt && boolPassed; loopULItemsInner++) { // Added boolPassed check
																				//Return the listitem
																				weULItem = lstULItems[loopULItemsInner]; // Access array by index
																				//Add a ';' if loop is for greater than the first element to Cell and Element values
																				if (loopULItemsInner > 0) {
																					strTempCellValue = strTempCellValue + ';';
																					strTempCellElements = strTempCellElements + ';';
																				}
																				//Highlight
																				if (TCExecParams.getBoolDoHighlight() == true) {
																					let mapHighlight = {}; // Mimic Groovy Map
																					mapHighlight = CWCore.objHighlightElementJS(weULItem, 'Element', 'UserList item_' + loopULItemsInner);
																				}
																				//Check if the child items are present
																				//Loop through the UL items elements XPaths to check for a single item
																				//Start with the second element since 0 is the listitem
																				let intElemFoundCnt = 0; // Type inferred
																				for (let loopULItemElem = 1; loopULItemElem < arryULElementNames.length && boolPassed; loopULItemElem++) { // Access array by index, Added boolPassed check
																					weULItemChild = CWCore.returnChildElement(weULItem, arryULElemXPaths[loopULItemElem]); // Access array by index
																					if (weULItemChild != null) {
																						intElemFoundCnt++;
																						//Highlight
																						if (TCExecParams.getBoolDoHighlight() == true) {
																							let mapHighlight = {}; // Mimic Groovy Map
																							mapHighlight = CWCore.objHighlightElementJS(weULItemChild, 'Element', 'UserList item_child' + loopULItemsInner);
																						}
																						//Return the text and Add the element name and text to the variables
																						if (intElemFoundCnt == 1 && loopULItemsInner == 0) {
																							strTempCellValue = weULItemChild.getText();
																							strTempCellElements = arryULElementNames[loopULItemElem]; // Access array by index
																						}
																						else if (intElemFoundCnt == 1 && loopULItemsInner > 0) {
																							strTempCellValue = strTempCellValue + weULItemChild.getText();
																							strTempCellElements = strTempCellElements + arryULElementNames[loopULItemElem]; // Access array by index
																						}
																						else {
																							strTempCellValue = strTempCellValue + "|" + weULItemChild.getText();
																							strTempCellElements = strTempCellElements + "|" + arryULElementNames[loopULItemElem]; // Access array by index
																						}
																					}
																				}
																			}
																		}
																		if (arryElementNames[loopElements] == strGRIElemName) { //Check for the Group Record Image element
																			//let intItemCnt = 0; // Declared but unused
																			//let strItemValue; // Declared but unused
																			//let strItemElem; // Declared but not used
																			//let weItem; // Declared but not used
																			//Process the parent element to return the items
																			let mapGRIItems = CWCore.returnChildElements(weCellElement, strGRIItemXPath);
																			let lstGRIItems = mapGRIItems.lstChildObjects; // Type inferred (List<WebElement>)
																			let intGRIItemCnt = mapGRIItems.cntChildObjs; // Type inferred
																			if (intGRIItemCnt == 0) {
																				//Error
																				boolPassed = false;
																				strMethodDetails = "FAILED!!! The Cell " + strGRIElemName + " contains 'ZERO' ITEM(S)!!!!";
																			}
																			else {
																				let weGRIItem; // Type inferred (WebElement)
																				let weGRIItemChild; // Type inferred (WebElement)
																				let intElemFoundCnt = 0; // Type inferred
																				//Loop through the GRIItems each will contain a child that should have an image. Return the attribute value.
																				for (let intLoopItems = 0; intLoopItems < intGRIItemCnt && boolPassed; intLoopItems++) { // Added boolPassed check
																					weGRIItem = lstGRIItems[intLoopItems]; // Access array by index
																					intElemFoundCnt++;
																					let strItemValue = gblNull; // Type inferred
																					//Highlight
																					if (TCExecParams.getBoolDoHighlight() == true) {
																						let mapHighlight = {}; // Mimic Groovy Map
																						mapHighlight = CWCore.objHighlightElementJS(weGRIItem, 'Element', 'GRIElement Item' + intLoopItems);
																					}
																					//Return the item child by finding elements and verifying only 1 is present
																					let intGRIItemChildCnt = CWCore.returnWebElemChildElementCount(weGRIItem, strGRIItemChildXPath); // Assigns and uses intCellElemCnt in Groovy but that's local scope here.
																					//Check if only 1 item is present
																					if (intGRIItemChildCnt == 1) {
																						weGRIItemChild = CWCore.returnChildElement(weGRIItem, strGRIItemChildXPath);
																						//Return the value
																						boolAttribPresent = CWCore.isAttribtuePresent(weGRIItemChild,strGRIImgAttribute);
																						if (boolAttribPresent == true) {
																							strItemValue = StringsAndNumbers.JComm_HandleNoData(weGRIItemChild.getAttribute(strGRIImgAttribute));
																						}
																						//Add to the output value
																						if (intElemFoundCnt == 1) {
																							strTempCellValue = strItemValue;
																							strTempCellElements = 'Record Image';
																						}
																						else {
																							strTempCellValue = strTempCellValue + "|" + strItemValue;
																							strTempCellElements = strTempCellElements + "|" + 'Record Image';
																						}
																					}
																					else { // intGRIItemChildCnt is not 1
																						//Fail (original code handled empty block, adding explicit fail)
																						boolPassed = false;
																						strMethodDetails = "FAILED!!! The GRI Item contained " + intGRIItemChildCnt + " children, expected 1.";
																						break; // Break inner loop
																					}
																				}
																			}
																		}
																		else { // Not UL or GRI, general element
																			strTempCellElements = arryElementNames[loopElements]; // Access array by index
																			if (arryElementNames[loopElements] == 'checkbox') {
																				strTempCellValue = CWCore.objGetCheckboxChecked(weCellElement).toString();
																			}
																			else if (arryElementNames[loopElements] == 'image') {
																				if (CWCore.isAttribtuePresent(weCellElement, strImageAttribute) == true) {
																					strTempCellValue = StringsAndNumbers.JComm_HandleNoData(weCellElement.getAttribute(strImageAttribute));
																				}
																			}
																		}
																		if (StringsAndNumbers.JComm_TextLocationInString(arryElemXPaths[loopElements], 'svg') >= 0) { // Access array by index
																			//Check if the SVG image name is not in the svg element
																			if (strElemContainsSVGNameXPath == gblNull) {
																				let strSVGElem = StringsAndNumbers.JComm_HandleNoData(weCellElement.getText());
																				if (StringsAndNumbers.JComm_TextLocationInString(strTempCellValue, strSVGElem) == 0) {
																					strTempCellValue = StringsAndNumbers.JComm_GetRightTextInString(strTempCellValue, strSVGElem);
																				}
																				else {
																					strTempCellValue = StringsAndNumbers.JComm_GetLeftTextInString(strTempCellValue, strSVGElem);
																				}
																			}
																			else {
																				//Return the element based on the path
																				let weSVGNameElem = weCellElement.findElement(By.xpath(strElemContainsSVGNameXPath));
																				if (weSVGNameElem != null) {
																					boolAttribPresent = CWCore.isAttribtuePresent(weSVGNameElem,strElemContainsSVGNameProperty);
																					if (boolAttribPresent == true) {
																						//Return the value of the property
																						strTempCellValue = StringsAndNumbers.JComm_HandleNoData(weSVGNameElem.getAttribute(strElemContainsSVGNameProperty));
																					}
																					else {
																						strTempCellValue = 'NO IMAGE NAME!!!';
																					}
																				} else { // SVG name element was null
																					strTempCellValue = 'SVG NAME ELEMENT NOT FOUND!!!'; // Added more explicit error
																				}
																			}
																		}
																	}
																	else { // Element count not 1 for this xpath
																		strTempCellElements = strTempCellElements + gblDelimiter + arryElementNames[loopElements]; // Access array by index
																		if (strTempCellElements.includes('checkbox')) { // Check if checkbox
																			strTempCellValue = CWCore.objGetCheckboxChecked(weCellElement).toString();
																		}
																	}
																	//TODO add in classes to get color such as status elements
																	if (arryElementNames[loopElements] == strResultsCellStatElem) {
																		//Return the element
																		let weStatus = CWCore.returnChildElement(weCell, arryElemXPaths[loopElements]); // Access array by index
																		if (weStatus == null) {
																			boolPassed = false;
																			strMethodDetails = `FAILED!!! Result row: ${loopRow + 1} cell: ${loopRowCol + 1} DID NOT CONTAIN A STATUS ELEMENT AS EXEPCTED!!!`; // Using Template Literal
																		}
																		else {
																			//Highlight
																			if (TCExecParams.getBoolDoHighlight() == true) {
																				let mapHighlight = {}; // Mimic Groovy Map
																				mapHighlight = CWCore.objHighlightElementJS(weStatus, 'Element', strTblName + ' cell status element');
																			}
																			strColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("color"));
																			if (StringsAndNumbers.JComm_TextLocationInString(arryElemXPaths[loopElements], "svg") >= 0) { // Access array by index
																				strBckgColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("fill"));
																			}
																			else {
																				strBckgColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("background-color"));
																			}
																			//Check if the color is the same
																			//let strOrgColor = strColor; // Declared but unused
																			let strTempColor = StringsAndNumbers.JComm_GetRightTextInString(strColor, "(");
																			strTempColor = StringsAndNumbers.JComm_GetLeftTextInString(strTempColor, ")");
																			let strTempBckColor = StringsAndNumbers.JComm_GetRightTextInString(strBckgColor, "(");
																			strTempBckColor = StringsAndNumbers.JComm_GetLeftTextInString(strTempBckColor, ")");
																			//let intLocInColorToBck = StringsAndNumbers.JComm_TextLocationInString(strTempColor, strTempBckColor); // Declared but unused
																			//let intLocInBckToColor = StringsAndNumbers.JComm_TextLocationInString(strTempBckColor, strTempColor); // Declared but unused
																			let strTempStatusColor; // Type inferred
																			//Return the property and add to the strTempCellElements
																			if (StringsAndNumbers.JComm_TextLocationInString(strTempColor, strTempBckColor) >= 0 || StringsAndNumbers.JComm_TextLocationInString(strTempBckColor, strTempColor) >= 0) {
																				strColor = "RGB(255, 255, 255)"; //Change the color for the text to white. Will work for all but white backgrounds
																				strTempStatusColor = "Colors matched replaced original color:" + strOrgColor + " with Color:" + strColor + "Background Color:"+ strBckgColor;
																			}
																			else {
																				strTempStatusColor = "Color:" + strColor + "Background Color:"+ strBckgColor;
																			}
																			strTempCellElements = strTempCellElements + gblDelimiter + strTempStatusColor;
																			boolCustomTheme = true;
																		}
																	}
																} // End of inner loop (loopElements)
																if (cntCellElements == 0 )
																{
																	break; // Break from while loop if no elements found for current cell
																}
															} // End of while loop (cntCellElements)
															//Return the wait to maxwaittime
															TCObj.tcDriver.manage().timeouts().implicitlyWait(TCObj.getIntTempMaxWaitTime(), TimeUnit.SECONDS);
															//Output the cell text after the elements. Found some cells contain svg elements with text that must be removed. PGK 02/21/2023
															mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intOutputExcelRow, intOutputColumn, strTempCellValue, strCellTheme);
															if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == true) {
																intOutputColumn++;
															}
															else {
																boolPassed = false;
																strMethodDetails = mapSetCellValue.strMethodDetails;
															}
															if (boolCustomTheme == false) {
																mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intOutputExcelRow, intOutputColumn, strTempCellElements, strCellTheme);
															}
															else {
																//Process the output with a custom them
																let mapCustomTheme = {}; // Mimic Groovy Map
																mapCustomTheme.ExcelRow = intOutputExcelRow;
																mapCustomTheme.ExcelCol = intOutputColumn;
																mapCustomTheme.CellValue = strTempCellElements;
																mapCustomTheme.RGBAColor = strColor;
																mapCustomTheme.RGBABkColor = strBckgColor;
																mapCustomTheme.Alignment = 'Center'; // Assuming 'Center' is recognized
																//Call the set cell value with customtheme method
																mapSetCellValue = ExcelData.excelSetCellValueByRowColNumberCustTheme(mapCustomTheme);
															}
															if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == true) {
																intOutputColumn++;
															}
															else {
																boolPassed = false;
																strMethodDetails = mapSetCellValue.strMethodDetails;
															}
														}
														if (boolPassed == false) { //Col Break
															break; // Break from loopRowCol
														}
													} // End of inner loop (loopRowCol)
													if (boolPassed == false) { //Row break
														break; // Break from loopRow
													}
												}
											} // End of else if (intBodyRowColCnt > 0)
										} // End of if (strTempAttrValue != strTableRowHiddenValue || strTempAttrValue == gblNull)
									} // End of loopRow
									//Update the column width to fit
									if (boolPassed == true) {
										//TODO create a new function to autofit the column
										let mapAutoFit = {}; // Mimic Groovy Map
										mapAutoFit = ExcelData.excelAutofitCols();
										boolPassed = StringsAndNumbers.JComm_StringToBoolean(mapAutoFit.boolPassed);
										strMethodDetails = mapAutoFit.strMethodDetails;
									}
									if (boolPassed == true) {
										strMethodDetails = ` Successufully output ${intBodyRowCnt} row(s) consisting of ${intBodyRowColCnt} column(s).`; // Using template literal
									}
								} // End of else (intBodyRowCnt > 0)
							}
						} // End of if (boolPassed == true) for arryColumnNames
					}
					else { // cntHdrCols <= 0
						boolPassed = false;
						strMethodDetails = "FAILED!!! The TABLE HEADER contain NO COLUMNS!!!";
					}
				}
				else { // mapAddSheet failed
					boolPassed = false;
					strMethodDetails = mapAddSheet.strMethodDetails;
				}
			}
		}
	}
	else { // weHeaderRow was null
		// boolPassed = false and method details set earlier in this branch
	}
	// Fallthrough from original conditional tree might leave boolPassed unchanged.
	// Ensure strMethodDetails is always set at exit point for failure if no specific fail path was executed.
	if (!mapResults.hasOwnProperty('boolPassed')) { // If not set by earlier logic
		mapResults.boolPassed = boolPassed.toString();
		mapResults.strMethodDetails = strMethodDetails;
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  VerifyResultBody  -----------------------------------
* Verify the body of the table results
* @param mapInputVariables The input variables that identify the header to include xpaths and properties and table results rows and columns
* @param NOTE: Each cell will be checked for the elements listed in the strResultsColObjNames
*
* @param INPUT SHEET Variables
* @param strInputSheetName				The sheet name assigned for the input
* @param intInputDataStartRow			The input data row number to start using for verification
* @param intInputDataEndRow			The input data row number to end verification. Enter 999 for all rows
* @param boolMatchInputRowNumber		Must the order of displayed rows equal the input row order? true/false
*
* @param TABLE Variables
* @param strLocName						The pipe delimited location value
* @param strLocXpath						The parent Xpath
* @param strTblName						The name of the table
* @param strTblXPath						The table Xpath
* @param strHdrRowXPath					The Header row Xpath usually //thead//tr
* @param strHdrColXPath					The header column Xpath usually //th
* @param strPropertyHdrColName				The cell property that will be used to find the assigned column name. Usually 'abbr'
* @param strResultsOutShtName				The sheet name assigned for the output
* @param strResultsRowXPath				The Xpath for the rows within the results
* @param strResultsColXPath				The Xpath for the columns within the result rows
* @param arryElementNames					The array containing the list of element names that can be in the results cells.
* @param arryElemXPaths					The array containing the list of elemenT Xpaths that can be in the results cells.
* @param NOTE: the TWO Arrays must contain the same number of items
* @param strResultsCellStatElem			The name of the status element
* @param //Check for the alternate values
* @param strTblSepXpath					The XPath for the body section of the table. Should only be used if the body tag appears in a cell of the results.
* @param strHdrColMissingNames				For each column that does not have an identifiable name, add the name to the delimited value
* @param strMultElemColData				The delimited value with the name|count of elements where a cell contains more than one element. Currently only one column is supported
* @param //UL alternate values see: Sourcing|Templates and Libraries|Event Libraries for example
* @param strResultsCellULElemName			The user list element name
* @param strResultsCellULNoValueElemName	The user cell list item No Value element name
* @param arryULElementNames				The User List Array that contains the child element names
* @param arryULElemXPaths					The User List Array that contains the child element XPaths*
*
* @param //New alternate values for body and header in one object i.e. thead see JI User Profile|User Roles and Access|Access which contains the header, data and subsection rows as data JTR-19321
* @param intTableHeaderRowCnt				The number of rows that contain header data
* @param strSectionRowInfo					The information about the section rows in a delimited format. rowProperty, rowValue, intCols. i.e. "class|SubSectionTitle|1" Must contain 3 values replace gblNotAblicable if value is not required or present.
*
* @param //Alternate values for image groups and elements in Elastic Search Results see JTR-20847-
* @param strGRIElemName						The Image group element name
* @param strGRIItemName						The name for the individual image element
* @param strGRIItemXPath					The Xpath for the image element within the group
* @param strGRIItemChildXPath				The Xpath for the child of the item which should contain the image. Multiple tags and class may be present.
* @param strGRIImgAttribute					The Image attribute that contains the value for the image for example the title'
*
* @param //Alternate value for the image elements attribute
* @param strImageAttribute					The attribute that will contain the value of the image
*
* @param //Alternate value for custom Widget table on AP Home
* @param strElemContainsSVGNameXPath		The Xpath for the SVG Name if not in the SVG element. ./.. for the parent element or gblNull if not required
* @param strElemContainsSVGNameProperty 	The element property where the SVG name will be present i.e. class leave gblNull if not present
*
* @return mapResults 						The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 04/30/2022
* @author Last Edited:12/07/2023
* @author Last Edited By:kpluedde
* @author Edit Comments: (Include email, date and details)
* @author updated to support table structure where the cell can contain a UL with li containers that may have 1 or more elements and values
* @author pkanaris added support for tables where the header and body are in one tag and the abiltity for the table to contain a section row JTR-19321 05/28/2023
* @author pkanaris added support for tables which have image groups that contain 1 to many images with values. JTR-20847 06/27/2023
* @author kpluedde added support for tables which have image attributes that contain 1 to many images with values. JJTR-22649 12/07/2023
*/
function CommonWeb_VerifyResultBody(mapInputVariables) { // mapInputVariables and return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	let gblSkip = GVars.GblSkip('Value');
	let gblIgnoreData = GVars.GblIgnoreData('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails; // Type inferred from assignment
	let boolPassed = true;
	//Return the method variables
	//Input Sheet variables
	let strInputSheetName = mapInputVariables.InputDataSheetName;
	let intInputDataStartRow = mapInputVariables.InputDataRowStart;
	let intInputDataEndRow = mapInputVariables.InputDataRowEnd;
	let boolMatchInputRowNumber = mapInputVariables.boolMatchInputRowNumber;
	//Application variables
	let strLocName = mapInputVariables.LocName;
	let strLocXpath = mapInputVariables.LocXPath;
	let strTblName = mapInputVariables.TblName;
	let strTblXPath = mapInputVariables.strTblXPath;
	let strHdrRowXPath = mapInputVariables.strHdrRowXPath;
	let strHdrColXPath = mapInputVariables.strHdrColXPath;
	let strPropertyHdrColName = mapInputVariables.strPropertyHdrColName;
	let strResultsOutShtName = mapInputVariables.strResultsOutShtName;
	let strResultsRowXPath = mapInputVariables.strResultsRowXPath;
	let strResultsColXPath = mapInputVariables.strResultsColXPath;
	//Cells can possibly have only one element
	let arryElementNames = mapInputVariables.arrayResultsColObjNames;
	let arryElemXPaths = mapInputVariables.arrayResultsColObjXPaths;
	let strResultsCellStatElem = mapInputVariables.strResultsCellStatElem;
	//Alternate properties
	let strTblSepXpath = gblNull;
	if (mapInputVariables.hasOwnProperty('strTblSepXpath')) { // Use hasOwnProperty for existence check
		strTblSepXpath = mapInputVariables.strTblSepXpath;
	}
	let strHdrColMissingNames = gblNull;
	let strMultElemColData = gblNull;
	if (mapInputVariables.hasOwnProperty('strHdrColMissingNames')) {
		strHdrColMissingNames = mapInputVariables.strHdrColMissingNames;
	}
	if (mapInputVariables.hasOwnProperty('strMultElemColData')) {
		strMultElemColData = mapInputVariables.strMultElemColData;
	}
	//UL alternate values see: Sourcing|Templates and Libraries|Event Libraries for example
	let strULElemName = gblNull;
	let strULNoValElemName = gblNull;
	let arryULElementNames;
	let arryULElemXPaths;
	if (mapInputVariables.hasOwnProperty('strResultsCellULElemName')) {
		strULElemName = mapInputVariables.strResultsCellULElemName;
	}
	if (mapInputVariables.hasOwnProperty('strResultsCellULNoValueElemName')) {
		strULNoValElemName = mapInputVariables.strResultsCellULNoValueElemName;
	}
	if (mapInputVariables.hasOwnProperty('arrayResultsColULObjNames')) {
		arryULElementNames = mapInputVariables.arrayResultsColULObjNames;
	}
	if (mapInputVariables.hasOwnProperty('arrayResultsColULObjXPaths')) {
		arryULElemXPaths = mapInputVariables.arrayResultsColULObjXPaths;
	}
	//Alternate values for Group-record-icons used in Supplier Search Results new UI Jun 2023 JTR-20847
	let strGRIElemName = gblNull;
	let strGRIItemName = gblNull;
	let strGRIItemXPath = gblNull;
	let strGRIItemChildXPath = gblNull;
	let strGRIImgAttribute = gblNull;
	if (mapInputVariables.hasOwnProperty('strGRIElemName')) {
		strGRIElemName = mapInputVariables.strGRIElemName;
	}
	if (mapInputVariables.hasOwnProperty('strGRIItemName')) {
		strGRIItemName = mapInputVariables.strGRIItemName;
	}
	if (mapInputVariables.hasOwnProperty('strGRIItemXPath')) {
		strGRIItemXPath = mapInputVariables.strGRIItemXPath;
	}
	if (mapInputVariables.hasOwnProperty('strGRIItemChildXPath')) {
		strGRIItemChildXPath = mapInputVariables.strGRIItemChildXPath;
	}
	if (mapInputVariables.hasOwnProperty('strGRIImgAttribute')) {
		strGRIImgAttribute = mapInputVariables.strGRIImgAttribute;
	}
	//Alternate value for the image elements attribute
	let strImageAttribute = gblNull;
	if (mapInputVariables.hasOwnProperty('strImageAttribute')) {
		strImageAttribute = mapInputVariables.strImageAttribute;
	}
	//Alternate value for custom Widget table on AP Home
	let strElemContainsSVGNameXPath = gblNull;
	let strElemContainsSVGNameProperty = gblNull;
	if (mapInputVariables.hasOwnProperty('strElemContainsSVGNameXPath')) {
		strElemContainsSVGNameXPath = mapInputVariables.strElemContainsSVGNameXPath;
	}
	if (mapInputVariables.hasOwnProperty('strElemContainsSVGNameProperty')) {
		strElemContainsSVGNameProperty = mapInputVariables.strElemContainsSVGNameProperty;
	}
	//Add alternate for table where header and data is in one cell and a sub category row is present
	let intTableHeaderRowCnt = -1;
	let strSectionRowInfo = gblNull;
	let arrySectInfo;
	let intSectInfoArrySize;
	if (mapInputVariables.hasOwnProperty('intTableHeaderRowCnt')) {
		intTableHeaderRowCnt = mapInputVariables.intTableHeaderRowCnt;
	}
	let intLoopHdrCount; //Holds the number of rows in the header to be added to the loopResultStart
	if (intTableHeaderRowCnt > 0) {
		intLoopHdrCount = intTableHeaderRowCnt;
	} else {
		intLoopHdrCount = 0; // Initialize if not set
	}

	if (mapInputVariables.hasOwnProperty('strSectionRowInfo')) {
		strSectionRowInfo = mapInputVariables.strSectionRowInfo;
		if (strSectionRowInfo == gblNull || strSectionRowInfo == gblNA) {
			strSectionRowInfo = gblNull; //Force to global null
		}
		else {
			let mapSplitSubSectInfoString = StringsAndNumbers.JComm_StringToArray(strSectionRowInfo, gblDelimiter);
			let intCntValues = StringsAndNumbers.JComm_StringToInteger(mapSplitSubSectInfoString.intItemCount);
			if (intCntValues == 4) {
				arrySectInfo = mapSplitSubSectInfoString.ArryOfValues;
				intSectInfoArrySize = arrySectInfo.length; // Use .length for JS arrays
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The Results can contain a Sub Section row, BUT, THE info provided '" + strSectionRowInfo + "', DOES NOT CONTAIN 3 ITEMS!!!";
			}
		}
	}
	let cntDispCols; //We need to capture the column count or expected count based on missing column names
	//Check for columns with more than one element
	let strMultElemColName = gblNull;
	let intMultElemColCount = 1; //always at least 1
	if (strMultElemColData != gblNull) {
		//Split the value into an array
		//ADD a split for the value to return the date/time format
		let mapArray = {}; // Mimic Groovy Map
		mapArray = StringsAndNumbers.JComm_StringToArray(strMultElemColData, gblDelimiter);
		if (StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount) == 2) {
			strMultElemColName = mapArray.ArryOfValues[0];
			intMultElemColCount = StringsAndNumbers.JComm_StringToInteger(mapArray.ArryOfValues[1]);
		}
	}
	let intResultsRowCnt; // Type inferred
	let boolProcCellElements; // Type inferred
	let arryOutputColNames; // Type inferred
	//Check if the outputsheet name is valid
	let intLenOptResShName = strResultsOutShtName.length;
	if (intLenOptResShName < 32) {
		//Verify we have the table.
		let weTable = CWCore.returnWebElement(strLocXpath + strTblXPath);
		if (weTable == null) {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location '" + strLocName + "' '" + strTblName + "' DID NOT RETURN an ELEMENT!!!";
		}
		else {
			//Highlight
			if (TCExecParams.getBoolDoHighlight() == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weTable, 'Element', strTblName);
			}
			//Check if the header row is present
			let weHeaderRow = CWCore.returnWebElement(strLocXpath + strTblXPath + strHdrRowXPath);
			if (weHeaderRow == null) {
				boolPassed = false;
				strMethodDetails = "FAILED!!!The location '" + strLocName + "' '" + strTblName + "' DID NOT RETURN A HEADER ROW ELEMENT!!!";
			}
			else {
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weHeaderRow, 'Element', strTblName + ' header row');
				}
				//Check for alternate name values
				let intAltColNameArrayCnt = -1;
				let intAltColNameUsed = -1;
				let arryAltColName;
				//let strTempAltColName; // Declared but unused
				if (strHdrColMissingNames != gblNull) {
					//Split the value into an array
					//ADD a split for the value to return the date/time format
					let mapArray = {}; // Mimic Groovy Map
					mapArray = StringsAndNumbers.JComm_StringToArray(strHdrColMissingNames, gblDelimiter);
					intAltColNameArrayCnt = StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount);
					arryAltColName = mapArray.ArryOfValues;
				}
				//Note, we open the file and do the work for each method and save/close the excel instance so save happens as we go.
				//Add the sheet to the input file
				let mapAddSheet = {}; // Mimic Groovy Map
				let strTCInputFilePath = TCObj.strTCInputFilePath;
				//TODO attempt to add the input stream to the test objects and use the input object to add the sheet in a new TObj WB and Sh
				mapAddSheet = ExcelData.excelAddSheetToWorkBook(strResultsOutShtName);
				let boolSheetAddedPassed = StringsAndNumbers.JComm_StringToBoolean(mapAddSheet.boolPassed);
				if (boolSheetAddedPassed == true) {
					//Add the column names
					//Return all the column names and use to create the output sheet header columns
					let mapGetHdrChildElements = {}; // Mimic Groovy Map
					let strOutputColNames;
					mapGetHdrChildElements = CWCore.returnChildElements(weHeaderRow, strHdrColXPath);

					let lstHdrCols = mapGetHdrChildElements.lstChildObjects;
					let arryColumnNames;
					let strHdrColNames;
					let cntHdrCols = mapGetHdrChildElements.cntChildObjs;
					if (cntHdrCols > 0) {
						cntDispCols = cntHdrCols; //TODO what happens if there is no header PGK 05/30/2023
						let weHdrCol;
						let strTempValue;
						let boolAttribPresent;
						//Return the column names and replace any spaces with '_'
						for (let loopHdrCols = 0; loopHdrCols < cntHdrCols; loopHdrCols++) {
							weHdrCol = lstHdrCols[loopHdrCols];
							//Highlight
							if (TCExecParams.getBoolDoHighlight() == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weHdrCol, 'Element', strTblName + ' header row column');
							}
							boolAttribPresent = CWCore.isAttribtuePresent(weHdrCol, strPropertyHdrColName);
							if (boolAttribPresent == true) {
								strTempValue = StringsAndNumbers.JComm_HandleNoData(weHdrCol.getAttribute(strPropertyHdrColName));
							}
							else if (intAltColNameArrayCnt >= 1 && intAltColNameArrayCnt > intAltColNameUsed) {
								intAltColNameUsed++;
								strTempValue = StringsAndNumbers.JComm_HandleNoData(arryAltColName[intAltColNameUsed]);
							}
							else {
								strTempValue = gblNull;
							}
							if (strTempValue != gblNull) {
								if (loopHdrCols == 0) {
									strOutputColNames = StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_");
									strHdrColNames = strTempValue;
								}
								else {
									strOutputColNames = strOutputColNames + gblDelimiter + StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_");
									strHdrColNames = strHdrColNames + gblDelimiter + strTempValue;
								}
								strOutputColNames = strOutputColNames + gblDelimiter + StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_") + "_Element"; //Add the column for the element
							}
							else {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The header column '" + loopHdrCols + "' does not CONTAIN THE PROPERTY REQUIRED of:" + strPropertyHdrColName + "!!!";
								break;
							}
						}
						if (boolPassed == true) {
							//Add column names to the array
							let mapStringToArry = {}; // Mimic Groovy Map
							mapStringToArry = StringsAndNumbers.JComm_StringToArray(strHdrColNames, gblDelimiter);
							arryColumnNames = mapStringToArry.ArryOfValues;
							let mapAddColNames = {}; // Mimic Groovy Map
							mapAddColNames = ExcelData.excelCreateHeaderCols(strOutputColNames);
							let boolAddColNames = StringsAndNumbers.JComm_StringToBoolean(mapAddColNames.boolPassed);
							if (boolAddColNames == true) {
								//Return the header elements
								// let lstColNames = mapAddColNames.lstColNames; // Declared but unused in Groovy func
								//Update the header column(s) to the assigned theme
								let strAssgTheme = 'OutputDataHdrStd';
								let mapUpdateHdrTheme = {}; // Mimic Groovy Map
								mapUpdateHdrTheme = ExcelData.excelSetRowCellFormatByTheme(0, strAssgTheme);
								let boolUPdHdrTheme = StringsAndNumbers.JComm_StringToBoolean(mapUpdateHdrTheme.boolPassed);
								if (boolUPdHdrTheme == false) {
									boolPassed = false;
									strMethodDetails = mapUpdateHdrTheme.strMethodDetails;
								}
							}
							else {
								boolPassed = false;
								strMethodDetails = mapAddColNames.strMethodDetails;
							}
						}
						//Load the inputsheet
						let intInputRowCnt;
						let intInputColCnt;
						// let intHdrColCnt; // Already declared above
						let shInput; // Mimic XSSFSheet
						let strHdrInputShtName = mapInputVariables.strHdrInputShtName; // Used later
						let mapOpenInputSheet = {}; // Mimic Groovy Map
						mapOpenInputSheet = ExcelData.excelGetSheetByName(TCObj.getObjWorkbook(), strInputSheetName);
						if (StringsAndNumbers.JComm_StringToBoolean(mapOpenInputSheet.boolPassed) == true) {
							shInput = mapOpenInputSheet.objWbSheet;
							//Return the row and column Count
							let mapSheetRowColCnt = {}; // Mimic Groovy Map
							mapSheetRowColCnt = ExcelData.excelGetRowAndColCount(shInput);
							if (StringsAndNumbers.JComm_StringToBoolean(mapSheetRowColCnt.boolPassed) == true) {
								intInputRowCnt = mapSheetRowColCnt.RowCount;
								intInputColCnt = mapSheetRowColCnt.ColCount;
							}
							else {
								boolPassed = false;
								strMethodDetails = mapSheetRowColCnt.strMethodDetails;
							}
							// intHdrColCnt is reassigned here, but it was already derived from the displayed header.
							// In TS, we'd stick to the one value derived from the actual display.
							// The Groovy code probably means to re-check counts from input, but the variable name clashes.
							if (boolPassed == true) {
								let mapHdrChildren = CWCore.returnChildElements(weHeaderRow, strHdrColXPath);
								// intHdrColCnt = mapHdrChildren.cntChildObjs; // Not re-assigning, using the one from displayed header.
								// lstHdrCols = mapHdrChildren.lstChildObjects; // Already defined.
							}
						}
						else {
							boolPassed = false;
							strMethodDetails = mapOpenInputSheet.strMethodDetails;
						}
						//Check if the number of columns displayed matches the number in the input sheet if not call output only and fail the step
						if (boolPassed == true) {
							if (cntHdrCols != intInputColCnt / 2) { //Each displayed column consist of the value and element column
								boolPassed = false;
								strMethodDetails = "FAILED!!! The table header is displaying: " + cntHdrCols + " BUT, EXPECTED: " + intInputColCnt / 2 + " COLUMNS!!! OUTPUTTING HEADER ONLY see details: ";
							}
							else {
								//Return the input sheet column names and check if the match the output column names
								let strInputColValues;
								let mapGetInputColNames = {}; // Mimic Groovy Map
								mapGetInputColNames = ExcelData.excelGetHdrColNames(shInput);
								let boolInputColNames = StringsAndNumbers.JComm_StringToBoolean(mapGetInputColNames.boolPassed);
								if (boolInputColNames == true) {
									strInputColValues = mapGetInputColNames.ColValues;
									//Check if the match the output column names
									if (strInputColValues != strOutputColNames) {
										boolPassed = false;
										strMethodDetails = "FAILED!!! The input column names '" + strInputColValues + "' DOES NOT MATCH the OUTPUT Columns '" + strOutputColNames + "'!!!";
									}
									else {
										strMethodDetails = "The column names matched and the input file contains: " + intInputRowCnt + " rows, and " + intInputColCnt + " columns.";
										let mapSplitElementString = StringsAndNumbers.JComm_StringToArray(strOutputColNames, gblDelimiter);
										// let intCntElemNames= StringsAndNumbers.JComm_StringToInteger(mapSplitElementString.intItemCount); // Declared but unused
										arryOutputColNames = mapSplitElementString.ArryOfValues;
									}
								}
								else {
									boolPassed = false;
									strMethodDetails = mapGetInputColNames.strMethodDetails;
								}
							}
						}
						if (boolPassed == true) {
							//Check if the cell elements will be verified
							if (arryElementNames == null || arryElemXPaths == null) {
								boolProcCellElements = false;
							}
							else if (arryElementNames.length != arryElemXPaths.length) { // Use .length for JS arrays
								boolPassed = false;
								strMethodDetails = "FAILED!!! THE arryElementNames and arryElemXPaths SIZE DO NOT MATCH!!!";
							}
							else {
								boolProcCellElements = true;
							}
						}
						//Return the data and verify the results outputting
						if (boolPassed == true) {
							//Set the start and end rows
							let loopInputRowStart;
							let loopInputRowEnd;
							let loopResultStart;
							let loopResultEnd;
							let intInputCol;
							let intDataElem;
							// let intOutputRow; // Declared but unused
							let boolInputRowMatch = true; //reset after each loop processed for input data.
							let mapSetLoopStartEnd = {}; // Mimic Groovy Map
							mapSetLoopStartEnd = ExcelData.excelReturnLoopStartEndRow(intInputDataStartRow, intInputDataEndRow, intInputRowCnt);
							loopInputRowStart = mapSetLoopStartEnd.intLoopStart;
							loopInputRowEnd = mapSetLoopStartEnd.intLoopEnd;
							let intInputRowsToProcess = loopInputRowEnd - loopInputRowStart + 1;
							let strTempInputValue;
							let strTempInputElem;
							let strTempDispValue;
							let strTempDispElem;
							let strColor;
							let strBckgColor;
							let intRowsMatched = 0; // Initialize to 0, not type inferred from any assignment
							let intOutputExcelRow;
							let intBodyRowColCnt;
							let intOutputColumn;
							let intCellElemCnt;
							let intOutputValColIndex;
							let intOutputElemColIndex;
							//Status Counters
							let intCellSkipped;
							let intCellIgnored;
							let intCellPassed;
							let intCellFailed;

							let strTempCellValue;
							let strTempCellElements;
							let weRow;
							let weCell;
							let mapGetRowCols;
							let lstBodyRowCols;
							let mapGetInputValue;
							let mapSetCellValue;
							let boolGetInpValue;
							//Check if the tbody is a separate XPath
							if (strTblSepXpath != gblNull) {
								//Return the body as a new table element
								weTable = CWCore.returnWebElement(strLocXpath + strTblXPath + strTblSepXpath);
								//Highlight
								if (TCExecParams.getBoolDoHighlight() == true) {
									let mapHighlight = {}; // Mimic Groovy Map
									mapHighlight = CWCore.objHighlightElementJS(weTable, 'Table Body', strTblName);
								}
								intLoopHdrCount = 0; // Reset header count for body, as it's separate
							}
							//Return the row count and create the loop
							let mapGetRowElements = CWCore.returnChildElements(weTable, strResultsRowXPath);
							let lstBodyRows = mapGetRowElements.lstChildObjects;
							let intBodyRowCnt = mapGetRowElements.cntChildObjs;
							intOutputExcelRow = 1; // Increment before loop starts
							//Check the results against the rows to process
							if (intInputRowsToProcess <= intBodyRowCnt) {
								let intSpacerCnt = 0; //Used to change the input row if the spacer is present.
								for (let loopInputRow = loopInputRowStart; loopInputRow <= loopInputRowEnd; loopInputRow++) {
									boolInputRowMatch = true;

									//Loop for the results
									if (boolMatchInputRowNumber == true) { //Input row number must match the displayed row
										loopResultStart = loopInputRow - 1 + intLoopHdrCount; //Zero based index for displayed rows
										loopResultEnd = loopInputRow + intLoopHdrCount;
									}
									else { //Input row data must be displayed within the results any row match is good
										loopResultStart = 0 + intLoopHdrCount; //Zero based index for displayed rows
										loopResultEnd = intBodyRowCnt + intLoopHdrCount;
									}
									for (let loopDisplayRow = loopResultStart; loopDisplayRow < loopResultEnd; loopDisplayRow++) {
										boolInputRowMatch = true;
										intInputCol = 0;
										intDataElem = 1; //Set to verify data first
										weRow = lstBodyRows[loopDisplayRow]; // Access array by index
										//Move to the row to show the row we are processing in case of errors.
										let actions = new Actions(TCObj.getTcDriver());
										actions.moveToElement(weRow).perform();
										//Highlight
										if (TCExecParams.getBoolDoHighlight() == true) {
											let mapHighlight = {}; // Mimic Groovy Map
											mapHighlight = CWCore.objHighlightElementJS(weRow, 'Element', strTblName + ' body row');
										}
										//TODO inject change for JTR-19321
										//Return the SubSection values if present
										let intSubSecColCnt = -1;
										let strColProp = gblNull;
										let strColPropValue = gblNull;
										let strColElemName = gblNull;
										if (strSectionRowInfo != gblNull) {
											intSubSecColCnt = StringsAndNumbers.JComm_StringToInteger(arrySectInfo[0]);
											strColProp = arrySectInfo[1];
											strColPropValue = arrySectInfo[2];
											strColElemName = arrySectInfo[3];
										}
										//1|class|SubSectionTitle|elemSubSection
										//Return the column elements and count
										mapGetRowCols = CWCore.returnChildElements(weRow, strResultsColXPath);
										lstBodyRowCols = mapGetRowCols.lstChildObjects;
										intBodyRowColCnt = mapGetRowCols.cntChildObjs;
										if (intBodyRowColCnt == intSubSecColCnt) {
											let boolIsSection = false;
											intOutputColumn = 0;
											//Create the temp counters
											intCellSkipped = 0; // Initialize for inner loop counts
											intCellIgnored = 0;
											intCellPassed = 0;
											intCellFailed = 0;
											let boolComparePass = true;
											//Process the subSection
											for (let loopRowCol = 0; loopRowCol < cntDispCols; loopRowCol++) {
												strTempCellValue = gblNull;
												strTempCellElements = gblNull;
												if (loopRowCol < intSubSecColCnt) {
													//Return the cell object
													weCell = lstBodyRowCols[loopRowCol];
													//Highlight
													if (TCExecParams.getBoolDoHighlight() == true) {
														let mapHighlight = {}; // Mimic Groovy Map
														mapHighlight = CWCore.objHighlightElementJS(weCell, 'Element', strTblName + ' body cell');
													}
													//Return the cell text
													strTempCellValue = StringsAndNumbers.JComm_HandleNoData(weCell.getText());
													//Confirm the cell belongs to the subsection
													let boolAttribPresent = CWCore.isAttribtuePresent(weCell, strColProp);
													if (boolAttribPresent == true) {
														let strTempValue = StringsAndNumbers.JComm_HandleNoData(weCell.getAttribute(strColProp));
														if (StringsAndNumbers.JComm_VerifyTextPresent(strTempValue, strColPropValue, "%Con%") == true) {
															boolIsSection = true;
															strTempCellElements = strColElemName;
														}
													}
												}
												//Get the input values
												strTempInputValue = gblNull;
												strTempInputElem = gblNull;
												//Return the expected values
												mapGetInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow - intSpacerCnt, arryOutputColNames[intInputCol]);
												boolGetInpValue = StringsAndNumbers.JComm_StringToBoolean(mapGetInputValue.boolPassed);
												if (boolGetInpValue == true) {
													strTempInputValue = StringsAndNumbers.JComm_HandleNoData(mapGetInputValue.CellValue);
													intOutputValColIndex = mapGetInputValue.intColIndex;
													//Check if the value is a dynamic date and/or time
													if (strTempInputValue.indexOf('D[') > 0 || strTempInputValue.indexOf('T[') > 0) {
														let myTemp = DateTime.datetimeReturnDynamicFormatedDateTime(strTempInputValue);
														strTempInputValue = myTemp;
													}
												}
												intInputCol++; //Increment the input column number
												mapGetInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow - intSpacerCnt, arryOutputColNames[intInputCol]);
												boolGetInpValue = StringsAndNumbers.JComm_StringToBoolean(mapGetInputValue.boolPassed);
												if (boolGetInpValue == true) {
													strTempInputElem = StringsAndNumbers.JComm_HandleNoData(mapGetInputValue.CellValue);
													intOutputElemColIndex = mapGetInputValue.intColIndex;
												}
												intInputCol++; //Increment the input column number
												//Compare each value and output the results
												let strCompareDispValue;
												let strCompareInputValue;
												let strOutputCellTheme; // Type inferred
												let intOutputColForCell; // Renamed to avoid clash with outer intOutputColumn
												for (let loopCompareInner = 1; loopCompareInner <= 2; loopCompareInner++) {
													if (loopCompareInner == 1) {
														strCompareDispValue = strTempCellValue;
														strCompareInputValue = strTempInputValue;
														intOutputColForCell = intInputCol - 2;
													}
													else {
														strCompareDispValue = strTempCellElements;
														strCompareInputValue = strTempInputElem;
														intOutputColForCell = intInputCol - 1;
													}
													//Check if they match
													if (strCompareInputValue == gblSkip || strCompareInputValue == gblIgnoreData) { //TODO should these be separated?
														strOutputCellTheme = 'TestRunDataSkipIgnore';
														if (strCompareInputValue == gblSkip) {
															intCellSkipped++;
														}
														else {
															intCellIgnored++;
														}
													}
													else if (strCompareInputValue == strCompareDispValue) {
														strOutputCellTheme = 'TestRunPassStd';
														intCellPassed++;
													}
													else {
														strOutputCellTheme = 'TestRunFailStd';
														strCompareDispValue = `Expected: ${strCompareInputValue}${gblLineFeed} Cell Displayed: ${strCompareDispValue}`; // Using template literals
														boolComparePass = false;
														boolInputRowMatch = false;
														intCellFailed++;
														//Check if we should break if row number does not need to match
													}
													//Output the results to the cell
													if (boolComparePass == true || boolMatchInputRowNumber == true) {
														mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intOutputExcelRow, intOutputColForCell, strCompareDispValue, strOutputCellTheme);
														if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
															boolPassed = false;
															strMethodDetails = mapSetCellValue.strMethodDetails;
														}
													}
													//Check if we should continue the compare
													if (boolComparePass == false && boolMatchInputRowNumber == true) {
														if (TCExecParams.getBoolDoDebug() == true) {
															Tester.Message("Compare failed and matchinputRowNumber == true so continue compare");
														}
													}
													else if (boolComparePass == false) {
														if (TCExecParams.getBoolDoDebug() == true) {
															Tester.Message("Compare failed and matchinputRowNumber == false so break for compare");
														}
														//Break so we move to the next row in the display
														break;
													}
												}
												//Check if we should continue the compare
												if (boolComparePass == false && boolMatchInputRowNumber == true) {
													boolPassed = false; // Fixed == false to = false
													strMethodDetails = strMethodDetails + "FAILED!!! Display Row: " + loopDisplayRow + " and cell in column: " +
														loopRowCol + " DID NOT MATCH SEE OUTPUT SHEET!!!";
												}
												else if (boolComparePass == false) {
													if (TCExecParams.getBoolDoDebug() == true) {
														Tester.Message("Compare failed and matchinputRowNumber == false so break for columns");
													}
													//Break so we move to the next row in the display
													break;
												}
											}
											if (boolIsSection == false) {
												boolPassed = false;
												strMethodDetails = `FAILED!!! Result row: ${loopDisplayRow + 1} CELL(s) PROPERTY DID NOT CONTAIN The VALUE '${strColPropValue}' !!!`;
											}
											if (boolComparePass == true) {
												intRowsMatched++;
											}
											//Check if row passed only if row number is to match as well
											if (boolComparePass == false && boolMatchInputRowNumber == true) {
												boolPassed = false;
											}
											if (boolMatchInputRowNumber == true) {
												//Increment to output row only if row number match is true
												intOutputExcelRow++;
												//Update the steps processed counts
												// These inner loop counters should probably be reset per row/col or handled centrally.
												// Copying behavior from Groovy where they are just incremented from the last value.
												// intCellSkipped = intCellSkipped + intCellSkippedTemp; // The Groovy original copies here, but 'intCellSkippedTemp' is not declared above this loop.
												// intCellIgnored = intCellIgnored + intCellIgnoredTemp;
												// intCellPassed = intCellPassed + intCellPassedTemp;
												// intCellFailed = intCellFailed + intCellFailedTemp;
											}
										}
										//Check if a spacer row is present with no columns
										else if (intBodyRowColCnt > 0) { //Skips process the row.
											if (intBodyRowColCnt == cntHdrCols) { // Assuming intHdrColCnt from outer scope.
												//Create the temp counters (these should probably be local to this block)
												intCellSkipped = 0; // Initialize for inner loop counts
												intCellIgnored = 0;
												intCellPassed = 0;
												intCellFailed = 0;
												let boolComparePass = true;
												for (let loopDisplayCol = 0; loopDisplayCol < cntHdrCols; loopDisplayCol++) { // Use cntHdrCols from outer scope
													strTempInputValue = gblNull;
													strTempInputElem = gblNull;
													//Return the expected values
													mapGetInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow - intSpacerCnt, arryOutputColNames[intInputCol]);
													boolGetInpValue = StringsAndNumbers.JComm_StringToBoolean(mapGetInputValue.boolPassed);
													if (boolGetInpValue == true) {
														strTempInputValue = StringsAndNumbers.JComm_HandleNoData(mapGetInputValue.CellValue);
														intOutputValColIndex = mapGetInputValue.intColIndex;
														//Check if the value is a dynamic date and/or time
														if (strTempInputValue.indexOf('D[') > 0 || strTempInputValue.indexOf('T[') > 0) {
															let myTemp = DateTime.datetimeReturnDynamicFormatedDateTime(strTempInputValue);
															strTempInputValue = myTemp;
														}
													}
													intInputCol++; //Increment the input column number
													mapGetInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow - intSpacerCnt, arryOutputColNames[intInputCol]);
													boolGetInpValue = StringsAndNumbers.JComm_StringToBoolean(mapGetInputValue.boolPassed);
													if (boolGetInpValue == true) {
														strTempInputElem = StringsAndNumbers.JComm_HandleNoData(mapGetInputValue.CellValue);
														intOutputElemColIndex = mapGetInputValue.intColIndex;
													}
													intInputCol++; //Increment the input column number
													weCell = lstBodyRowCols[loopDisplayCol];
													//Highlight
													if (TCExecParams.getBoolDoHighlight() == true) {
														let mapHighlight = {}; // Mimic Groovy Map
														mapHighlight = CWCore.objHighlightElementJS(weCell, 'Element', strTblName + ' body cell');
													}
													//Return the cell value
													strTempDispValue = StringsAndNumbers.JComm_HandleNoData(weCell.getText());
													//Return the cell element
													strTempDispElem = gblNull;
													//Return the cell element
													// strTempDispElem = gblNull // Already set
													let weCellElement; // Type inferred
													if (boolProcCellElements == true) {
														//Set the wait to 1 second to speed up the processing of elements in the cells
														TCObj.tcDriver.manage().timeouts().implicitlyWait(100, TimeUnit.MILLISECONDS);
														//Process the array for elements
														let intMaxElem = 1;
														if (strMultElemColName == arryColumnNames[loopDisplayCol]) {
															intMaxElem = intMultElemColCount;
														}
														// strTempCellElements = gblNull; // Already set in outer loop
														//Process the array for elements
														let cntCellElements = 0;
														while (cntCellElements < intMaxElem) {
															for (let loopElements = 0; loopElements < arryElementNames.length; loopElements++) {
																//TODO add in classes to get color such as status elements
																intCellElemCnt = CWCore.returnWebElemChildElementCount(weCell, arryElemXPaths[loopElements]);
																if (intCellElemCnt == 1) {
																	cntCellElements++; //Increment the number of elements in the cell
																	weCellElement = CWCore.returnChildElement(weCell, arryElemXPaths[loopElements]);
																	//Highlight
																	if (TCExecParams.getBoolDoHighlight() == true) {
																		let mapHighlight = {}; // Mimic Groovy Map
																		mapHighlight = CWCore.objHighlightElementJS(weCellElement, 'Element', strTblName + ' cell status element');
																	}
																	if (cntCellElements == 1) {
																		//Check if the element is a UL
																		if (arryElementNames[loopElements] == strULElemName) { //See Sourcing|Templates and Libraries|Event Libraries for UI example
																			//TODO we need to pass in the ul element names and xpaths
																			//let intItemCnt = 0 // Declared but unused
																			//let strItemValue; // Declared but unused
																			//let strItemElem; // Declared but unused
																			let weULItem;
																			//Process the UL to return the items
																			let mapULItems = CWCore.returnChildElements(weCellElement, arryULElemXPaths[0]);
																			let lstULItems = mapULItems.lstChildObjects;
																			let intULItemCnt = mapULItems.cntChildObjs;
																			if (intULItemCnt == 0) {
																				//Error
																				boolPassed = false;
																				strMethodDetails = "FAILED!!! The Cell User List contain 'ZERO' ITEM(S)!!!!";
																			}
																			else {
																				let weULItemChild;
																				for (let loopULItemsInner = 0; loopULItemsInner < intULItemCnt; loopULItemsInner++) { //Start with the second element since 0 is the listitem
																					//Return the listitem
																					weULItem = lstULItems[loopULItemsInner];
																					//Add a ';' if loop is for greater than the first element to Cell and Element values
																					if (loopULItemsInner > 0) {
																						strTempCellValue = strTempCellValue + ';';
																						strTempCellElements = strTempCellElements + ';';
																					}
																					//Highlight
																					if (TCExecParams.getBoolDoHighlight() == true) {
																						let mapHighlight = {}; // Mimic Groovy Map
																						mapHighlight = CWCore.objHighlightElementJS(weULItem, 'Element', 'UserList item_' + loopULItemsInner);
																					}
																					//Check if the child items are present
																					//Loop through the UL items elements XPaths to check for a single item
																					//Start with the second element since 0 is the listitem
																					let intElemFoundCnt = 0;
																					for (let loopULItemElem = 1; loopULItemElem < arryULElementNames.length; loopULItemElem++) {
																						weULItemChild = CWCore.returnChildElement(weULItem, arryULElemXPaths[loopULItemElem]);
																						if (weULItemChild != null) {
																							intElemFoundCnt++;
																							//Highlight
																							if (TCExecParams.getBoolDoHighlight() == true) {
																								let mapHighlight = {}; // Mimic Groovy Map
																								mapHighlight = CWCore.objHighlightElementJS(weULItemChild, 'Element', 'UserList item_child' + loopULItemsInner);
																							}
																							//Return the text and Add the element name and text to the variables
																							if (intElemFoundCnt == 1 && loopULItemsInner == 0) {
																								strTempCellValue = weULItemChild.getText();
																								strTempCellElements = arryULElementNames[loopULItemElem];
																							}
																							else if (intElemFoundCnt == 1 && loopULItemsInner > 0) {
																								strTempCellValue = strTempCellValue + weULItemChild.getText();
																								strTempCellElements = strTempCellElements + arryULElementNames[loopULItemElem];
																							}
																							else {
																								strTempCellValue = strTempCellValue + "|" + weULItemChild.getText();
																								strTempCellElements = strTempCellElements + "|" + arryULElementNames[loopULItemElem];
																							}
																						}
																					}
																				}
																			}
																			strTempDispElem = strTempCellElements;
																			strTempDispValue = strTempCellValue;
																		}
																		else if (arryElementNames[loopElements] == strGRIElemName) { //Check for the Group Record Image element
																			//let intItemCnt = 0 // Declared but unused
																			//let strItemValue; // Declared but unused
																			//let strItemElem; // Declared but unused
																			//let weItem; // Declared but unused
																			//Process the parent element to return the items
																			let mapGRIItems = CWCore.returnChildElements(weCellElement, strGRIItemXPath);
																			let lstGRIItems = mapGRIItems.lstChildObjects;
																			let intGRIItemCnt = mapGRIItems.cntChildObjs;
																			if (intGRIItemCnt == 0) {
																				//Error
																				boolPassed = false;
																				strMethodDetails = "FAILED!!! The Cell " + strGRIElemName + " contains 'ZERO' ITEM(S)!!!!";
																			}
																			else {
																				let weGRIItem;
																				let weGRIItemChild;
																				let intElemFoundCnt = 0;
																				//Loop through the GRIItems each will contain a child that should have an image. Return the attribute value.
																				for (let intLoopItems = 0; intLoopItems < intGRIItemCnt; intLoopItems++) {
																					weGRIItem = lstGRIItems[intLoopItems];
																					intElemFoundCnt++;
																					let strItemValue = gblNull;
																					//Highlight
																					if (TCExecParams.getBoolDoHighlight() == true) {
																						let mapHighlight = {}; // Mimic Groovy Map
																						mapHighlight = CWCore.objHighlightElementJS(weGRIItem, 'Element', 'GRIElement Item' + intLoopItems);
																					}
																					//Return the item child by finding elements and verifying only 1 is present
																					let intGRIItemChildCnt = CWCore.returnWebElemChildElementCount(weGRIItem, strGRIItemChildXPath); // Original also set intCellElemCnt =
																					//Check if only 1 item is present
																					if (intGRIItemChildCnt == 1) {
																						weGRIItemChild = CWCore.returnChildElement(weGRIItem, strGRIItemChildXPath);
																						//Return the value
																						let boolAttribPresent = CWCore.isAttribtuePresent(weGRIItemChild, strGRIImgAttribute);
																						if (boolAttribPresent == true) {
																							strItemValue = StringsAndNumbers.JComm_HandleNoData(weGRIItemChild.getAttribute(strGRIImgAttribute));
																						}
																						//Add to the output value
																						if (intElemFoundCnt == 1) {
																							strTempDispValue = strItemValue;
																							strTempDispElem = 'Record Image';
																						}
																						else {
																							strTempDispValue = strTempDispValue + "|" + strItemValue;
																							strTempDispElem = strTempDispElem + "|" + 'Record Image';
																						}
																					}
																					else {
																						//Fail (original code handled empty block, adding explicit fail)
																						boolPassed = false;
																						strMethodDetails = "FAILED!!! The GRI Item contains " + intGRIItemChildCnt + " children, expected 1.";
																					}
																				}
																			}
																		}
																		else {
																			strTempDispElem = arryElementNames[loopElements];
																			let intInString = StringsAndNumbers.JComm_TextLocationInString(strTempDispElem.toUpperCase(), 'CHECKBOX');
																			if (StringsAndNumbers.JComm_TextLocationInString(strTempDispElem.toUpperCase(), 'CHECKBOX') >= 0) {
																				strTempDispValue = CWCore.objGetCheckboxChecked(weCellElement).toString();
																				if (strTempDispValue == gblNull) {
																					strTempDispValue = 'false';
																				}
																			}
																			else if (arryElementNames[loopElements] == 'image') {
																				if (CWCore.isAttribtuePresent(weCellElement, strImageAttribute) == true) {
																					strTempDispValue = StringsAndNumbers.JComm_HandleNoData(weCellElement.getAttribute(strImageAttribute));
																				}
																				else {
																					strTempDispValue = gblNull; // Ensure it's explicitly null if attribute not present
																				}
																			}
																		}
																		if (StringsAndNumbers.JComm_TextLocationInString(arryElemXPaths[loopElements], 'svg') >= 0) {
																			//Check if the SVG image name is not in the svg element
																			if (strElemContainsSVGNameXPath == gblNull) {
																				let strSVGElem = StringsAndNumbers.JComm_HandleNoData(weCellElement.getText());
																				// The original Groovy logic used JComm_TextLocationInString(strTempCellValue, strSVGElem) == 0 etc.
																				// StringsAndNumbers.JComm_GetRightTextInString/GetLeftTextInString is for delimited strings.
																				// Assuming intent is to remove prefix/suffix if it exists.
																				if (strTempDispValue.startsWith(strSVGElem)) {
																					strTempDispValue = strTempDispValue.substring(strSVGElem.length);
																				} else if (strTempDispValue.endsWith(strSVGElem)) {
																					strTempDispValue = strTempDispValue.substring(0, strTempDispValue.length - strSVGElem.length);
																				}
																			}
																			else {
																				//Return the element based on the path
																				// findElement on WebElement assumes it's supported (it often is in WebDriver bindings)
																				let weSVGNameElem = weCellElement.findElement(By.xpath(strElemContainsSVGNameXPath));
																				if (weSVGNameElem != null) {
																					//Check for the element property
																					let boolAttribPresent = CWCore.isAttribtuePresent(weSVGNameElem, strElemContainsSVGNameProperty);
																					if (boolAttribPresent == true) {
																						//Return the value of the property
																						strTempDispValue = StringsAndNumbers.JComm_HandleNoData(weSVGNameElem.getAttribute(strElemContainsSVGNameProperty));
																					}
																					else {
																						strTempDispValue = 'NO IMAGE NAME!!!';
																					}
																				}
																				else {
																					strTempDispValue = 'SVG NAME ELEMENT NOT FOUND!!!'; // Added more explicit error
																				}
																			}
																		}
																	}
																	else {
																		strTempDispElem = strTempDispElem + gblDelimiter + arryElementNames[loopElements];
																	}
																	//TODO add in classes to get color such as status elements
																	if (arryElementNames[loopElements] == strResultsCellStatElem) {
																		//Return the element
																		let weStatus = CWCore.returnChildElement(weCell, arryElemXPaths[loopElements]);
																		if (weStatus == null) {
																			boolPassed = false;
																			strMethodDetails = "FAILED!!! Result row: " + loopDisplayRow + " cell: " + (loopDisplayCol + 1) + " DID NOT CONTAIN A STATUS ELEMENT AS EXEPCTED!!!";
																		}
																		else {
																			//Highlight
																			if (TCExecParams.getBoolDoHighlight() == true) {
																				let mapHighlight = {}; // Mimic Groovy Map
																				mapHighlight = CWCore.objHighlightElementJS(weStatus, 'Element', strTblName + ' cell status element');
																			}
																			strColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("color"));
																			if (StringsAndNumbers.JComm_TextLocationInString(arryElemXPaths[loopElements], "svg") >= 0) {
																				strBckgColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("fill"));
																			}
																			else {
																				strBckgColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("background-color"));
																			}
																			//Check if the color is the same
																			let strOrgColor = strColor;
																			let strTempColor = StringsAndNumbers.JComm_GetRightTextInString(strColor, "(");
																			strTempColor = StringsAndNumbers.JComm_GetLeftTextInString(strTempColor, ")");
																			let strTempBckColor = StringsAndNumbers.JComm_GetRightTextInString(strBckgColor, "(");
																			strTempBckColor = StringsAndNumbers.JComm_GetLeftTextInString(strTempBckColor, ")");
																			//let intLocInColorToBck = StringsAndNumbers.JComm_TextLocationInString(strTempColor, strTempBckColor); // Declared but unused
																			//let intLocInBckToColor = StringsAndNumbers.JComm_TextLocationInString(strTempBckColor, strTempColor); // Declared but unused
																			let strTempStatusColor;
																			//Return the property and add to the strTempCellElements
																			if (StringsAndNumbers.JComm_TextLocationInString(strTempColor, strTempBckColor) >= 0 || StringsAndNumbers.JComm_TextLocationInString(strTempBckColor, strTempColor) >= 0) {
																				strColor = "RGB(255, 255, 255)"; //Change the color for the text to white. Will work for all but white backgrounds
																				strTempStatusColor = "Colors matched replaced original color:" + strOrgColor + " with Color:" + strColor + "Background Color:" + strBckgColor;
																			}
																			else {
																				strTempStatusColor = "Color:" + strColor + "Background Color:" + strBckgColor;
																			}
																			strTempDispElem = strTempDispElem + gblDelimiter + strTempStatusColor;
																		}
																	}
																}
															}
															if (cntCellElements == 0) {
																break;
															}
														}
														//Return the wait to maxwaittime
														TCObj.tcDriver.manage().timeouts().implicitlyWait(TCObj.getIntTempMaxWaitTime(), TimeUnit.SECONDS);
													}
													//Compare each value and output the results
													let strCompareDispValueFinal; // Renamed to avoid shadowing
													let strCompareInputValueFinal; // Renamed to avoid shadowing
													let strOutputCellThemeFinal; // Renamed to avoid shadowing
													let intOutputColForFinal; // Renamed to avoid shadowing
													for (let loopCompareFinal = 1; loopCompareFinal <= 2; loopCompareFinal++) {
														if (loopCompareFinal == 1) {
															strCompareDispValueFinal = strTempDispValue;
															strCompareInputValueFinal = strTempInputValue;
															intOutputColForFinal = intInputCol - 2;
														}
														else {
															strCompareDispValueFinal = strTempDispElem;
															strCompareInputValueFinal = strTempInputElem;
															intOutputColForFinal = intInputCol - 1;
														}
														//Check if they match
														if (strCompareInputValueFinal == gblSkip || strCompareInputValueFinal == gblIgnoreData) { //TODO should these be separated?
															strOutputCellThemeFinal = 'TestRunDataSkipIgnore';
															if (strCompareInputValueFinal == gblSkip) {
																intCellSkipped++; // Use outer counters
															}
															else {
																intCellIgnored++; // Use outer counters
															}
														}
														else if (strCompareInputValueFinal == strCompareDispValueFinal) {
															strOutputCellThemeFinal = 'TestRunPassStd';
															intCellPassed++; // Use outer counters
														}
														else {
															strOutputCellThemeFinal = 'TestRunFailStd';
															strCompareDispValueFinal = `Expected: ${strCompareInputValueFinal}${gblLineFeed} Cell Displayed: ${strCompareDispValueFinal}`;
															boolComparePass = false;
															boolInputRowMatch = false;
															intCellFailed++; // Use outer counters
															//Check if we should break if row number does not need to match
														}
														//Output the results to the cell
														if (boolComparePass == true || boolMatchInputRowNumber == true) {
															mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intOutputExcelRow, intOutputColForFinal, strCompareDispValueFinal, strOutputCellThemeFinal);
															if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
																boolPassed = false;
																strMethodDetails = mapSetCellValue.strMethodDetails;
															}
														}
														//Check if we should continue the compare
														if (boolComparePass == false && boolMatchInputRowNumber == true) {
															if (TCExecParams.getBoolDoDebug() == true) {
																Tester.Message("Compare failed and matchinputRowNumber == true so continue compare");
															}
														}
														else if (boolComparePass == false) {
															if (TCExecParams.getBoolDoDebug() == true) {
																Tester.Message("Compare failed and matchinputRowNumber == false so break for compare");
															}
															//Break so we move to the next row in the display
															break;
														}
													}
													//Check if we should continue the compare
													if (boolComparePass == false && boolMatchInputRowNumber == true) {
														// Note: `boolPassed == false` is a comparison, needs to be `boolPassed = false` for assignment
														boolPassed = false;
														strMethodDetails = strMethodDetails + "FAILED!!! Display Row: " + loopDisplayRow + " and cell in column: " +
															loopDisplayCol + " DID NOT MATCH SEE OUTPUT SHEET!!!";
													}
													else if (boolComparePass == false) {
														if (TCExecParams.getBoolDoDebug() == true) {
															Tester.Message("Compare failed and matchinputRowNumber == false so break for columns");
														}
														//Break so we move to the next row in the display
														break;
													}
												}
												if (boolComparePass == true) {
													intRowsMatched++;
												}
												//Check if row passed only if row number is to match as well
												if (boolComparePass == false && boolMatchInputRowNumber == true) {
													boolPassed = false;
												}
												if (boolMatchInputRowNumber == true) {
													//Increment to output row only if row number match is true
													intOutputExcelRow++;
													//Update the steps processed counts
													// Note: Groovy original uses intCellSkippedTemp etc here, but those are re-declared locally.
													// Assuming these are cumulative counters.
													// intCellSkipped = intCellSkipped + intCellSkippedTemp;
													// intCellIgnored = intCellIgnored + intCellIgnoredTemp;
													// intCellPassed = intCellPassed + intCellPassedTemp;
													// intCellFailed = intCellFailed + intCellFailedTemp;
												}
											}
											else {
												boolPassed = false;
												strMethodDetails = "FAILED!!! Result row: " + loopDisplayRow + " is displaying " + intBodyRowColCnt +
													" column(s), HOWEVER EXPECTED COLUMNS BASED ON THE HEADER COUNT IS: " + cntHdrCols + "!!!";
											}
											if (boolInputRowMatch == true) {
												strMethodDetails = strMethodDetails + "Results in row: " + loopDisplayRow + " matched the results assigned.";
												break;
											}
										}
										else {
											intSpacerCnt++;
											loopInputRowEnd++; //Also add 1 to the inputrowend since the input sheet does not account for spacer rows
										}
									}
								}
								//Update the column width to fit
								let mapAutoFit = {}; // Mimic Groovy Map
								mapAutoFit = ExcelData.excelAutofitCols();
								if (StringsAndNumbers.JComm_StringToBoolean(mapAutoFit.boolPassed) == false) {
									strMethodDetails = strMethodDetails + mapAutoFit.strMethodDetails;
								}
								//Check if we matched all the assigned rows.
								if (intRowsMatched == intInputRowsToProcess) {
									strMethodDetails = "Successfully processed the assigned " + intInputRowsToProcess + " row(s). See output sheet for details." +
										gblLineFeed + strMethodDetails;
								}
								else {
									boolPassed = false;
									strMethodDetails = "FAILED!!! ONLY" + intRowsMatched + " row(s) MATCHED OF THE ASSIGNED " + intInputRowsToProcess + " row(s) to process. See output sheet for details." +
										gblLineFeed + strMethodDetails;
								}
							}
							else {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The user specified to process " + intInputRowsToProcess +
									" row(s) WHICH IS GREATER THAN THE DISPLAYED " + intBodyRowCnt + " ROWS!!!";
							}
						}
					}
				}
				else {
					boolPassed = false;
					strMethodDetails = mapAddSheet.strMethodDetails;
				}
			}
		}
	}
	else {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The assigned output sheet name '" + strResultsOutShtName + " is " + intLenOptResShName + " CHAR LENGTH WHICH EXCEEDS EXCEL TAB NAME LENGTH OF 31!!!!";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

//TODO update method comments when completed
/**
* -------------------------------------  SetSelectRowFromResultBody  -----------------------------------
* Verify the body of the table results
* @param mapInputVariables The input variables that identify the header to include xpaths and properties and table results rows and columns
* @param NOTE: Each cell will be checked for the elements listed in the strResultsColObjNames
* @param NOTE: Elements in input sheets and variables passed in must contain the exact value in this keyword.
* @param examples lstAddlColElemTypes = ['Editbox','Checkbox','Graphic Checkbox','JavaScript Checkbox','listbox', 'Single Value Filter List Box']
* @param INPUT SHEET Variables
* @param strInputSheetName				The sheet name assigned for the input
* @param intInputDataStartRow			The input data row number to start using for verification
* @param intInputDataEndRow			The input data row number to end verification. Enter 999 for all rows
* @param boolMatchInputRowNumber		Must the order of displayed rows equal the input row order? true/false
* @param strSelectColumnName 			The column containing the element to select/edit/set i.e. gblUseInputColName + gblDelimiter + "Select_Column"
* @param strSelectElementName			The column containing the element i.e. gblUseInputColName + gblDelimiter + "Select_ElemName"
* @param strSelectElementChecked		The column containing the element checked state true/false i.e gblUseInputColName + gblDelimiter + "Select_Checked"
* @param strSelectElementType 			The element type that will be in the column name i.e gblUseInputColName + gblDelimiter + "Select_ElementType
* @param strSelectedItemValueSelected 	The value to set/select/ or click as applicable. i.e. gblUseInputColName + gblDelimiter + "Select_Data
* @param TABLE Variables
* @param strLocName					The pipe delimited location value
* @param strLocXpath					The parent Xpath
* @param strTblName					The name of the table
* @param strTblXPath					The table Xpath
* @param strHdrRowXPath				The Header row Xpath usually //thead//tr
* @param strHdrColXPath				The header column Xpath usually //th
* @param strPropertyHdrColName			The cell property that will be used to find the assigned column name. Usually 'abbr'
* @param strResultsOutShtName			The sheet name assigned for the output
* @param strResultsRowXPath			The Xpath for the rows within the results
* @param strResultsColXPath			The Xpath for the columns within the result rows
* @param arryElementNames				The array containing the list of element names that can be in the results cells.
* @param arryElemXPaths				The array containing the list of elemenT Xpaths that can be in the results cells.
* @param NOTE: the TWO Arrays must contain the same number of items
* @param strResultsCellStatElem		The name of the status element
* @param //Check for the alternate values
* @param strTblSepXpath				The XPath for the body section of the table. Should only be used if the body tag appears in a cell of the results.
* @param strHdrColMissingNames			For each column that does not have an identifiable name, add the name to the delimited value
* @param strInputSelSetColNames
* @param strMultElemColData			The delimited value with the name|count of elements where a cell contains more than one element. Currently only one column is supported
* @param //UL alternate values see: Sourcing|Templates and Libraries|Event Libraries for example
* @param strResultsCellULElemName		The user list element name
* @param strResultsCellULNoValueElemNameThe user cell list item No Value element name
* @param arryULElementNames			The User List Array that contains the child element names
* @param arryULElemXPaths				The User List Array that contains the child element XPaths
*
* @param //Alternate values for image groups and elements in Elastic Search Results see JTR-20847-
* @param strGRIElemName				The Image group element name
* @param strGRIItemName				The name for the individual image element
* @param strGRIItemXPath				The Xpath for the image element within the group
* @param strGRIItemChildXPath			The Xpath for the child of the item which should contain the image. Multiple tags and class may be present.
* @param strGRIImgAttribute			The Image attribute that contains the value for the image for example the title'
* @return mapResults 					The results showing Passed and method details. Includes the number of the row matched.
*
* @param //Alternate value for custom Widget table on AP Home
* @param strElemContainsSVGNameXPath		The Xpath for the SVG Name if not in the SVG element. ./.. for the parent element or gblNull if not required
* @param strElemContainsSVGNameProperty 	The element property where the SVG name will be present i.e. class leave gblNull if not present
*
*
* @author pkanaris
* @author Created: 04/30/2022
* @author Last Edited: 06/27/2023
* @author Last Edited By: PGK
* @author Edit Comments: (Include email, date and details)
* pkanaris@jaggear.com added support to select a specific items from the cell containing the userlist.
* Variable must include the value pair and will be in the excel spreadsheet.
* The following are how the three variables must be passed in from the module to select items. The values for each row are used in place of the input variable
* strSelectColumnName = gblUseInputColName + gblDelimiter + "Select_Column"
* strSelectElementName = gblUseInputColName + gblDelimiter + "Select_Element"
* strSelectedItemValueSelected = gblUseInputColName + gblDelimiter + "Select_Data"
* Select_Data value for selecting a list item must include the value pair and will be in the excel spreadsheet.
* @author updated to support table structure where the cell can contain a UL with li containers that may have 1 or more elements and values
* @author pkanaris added support for tables which have image groups that contain 1 to many images with values. JTR-20847
*/
function CommonWeb_SetSelectRowFromResultBody(mapInputVariables) { // mapInputVariables and return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	let gblSkip = GVars.GblSkip('Value');
	let gblIgnoreData = GVars.GblIgnoreData('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	let gblUseInputColName = GVars.GBLUseInputColName('Value');
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	//Return the method variables
	//Input Sheet variables
	let strInputSheetName = mapInputVariables.InputDataSheetName;
	let intInputDataStartRow = mapInputVariables.InputDataRowStart;
	let intInputDataEndRow = mapInputVariables.InputDataRowEnd;
	// Note: boolMatchInputRowNumber not taken from mapInputVariables here as per Groovy source in this function
	// But it is present in the Javadoc, so might be intended. Omitting for now as per code.
	let strSelectColumnName = mapInputVariables.strSelectColumnName;
	let strSelectElementName = mapInputVariables.strSelectElementName;
	let strSelectElementValueColName = mapInputVariables.strSelectElementValueColName;
	let strSelectElementType = mapInputVariables.strSelectElementType; //What type of listbox, checkbox
	let strSelectedItemValueSelected = mapInputVariables.strSelectedItemValueSelected;
	//Application variables
	let strLocName = mapInputVariables.LocName;
	let strLocXpath = mapInputVariables.LocXPath;
	let strTblName = mapInputVariables.TblName;
	let strTblXPath = mapInputVariables.strTblXPath;
	let strHdrRowXPath = mapInputVariables.strHdrRowXPath;
	let strHdrColXPath = mapInputVariables.strHdrColXPath;
	let strPropertyHdrColName = mapInputVariables.strPropertyHdrColName;
	let strResultsOutShtName = mapInputVariables.strResultsOutShtName;
	let strResultsRowXPath = mapInputVariables.strResultsRowXPath;
	let strResultsColXPath = mapInputVariables.strResultsColXPath;
	//Cells can possibly have only one element
	let arryElementNames = mapInputVariables.arrayResultsColObjNames;
	let arryElemXPaths = mapInputVariables.arrayResultsColObjXPaths;
	let strResultsCellStatElem = mapInputVariables.strResultsCellStatElem;
	//Check for the alternate values
	let strTblSepXpath = gblNull;
	if (mapInputVariables.hasOwnProperty('strTblSepXpath')) {
		strTblSepXpath = mapInputVariables.strTblSepXpath;
	}
	let strHdrColMissingNames = gblNull;
	let strInputSelSetColNames = gblNull;
	let strMultElemColData = gblNull;
	if (mapInputVariables.hasOwnProperty('strHdrColMissingNames')) {
		strHdrColMissingNames = mapInputVariables.strHdrColMissingNames;
	}
	if (mapInputVariables.hasOwnProperty('strInputSelSetColNames')) {
		strInputSelSetColNames = mapInputVariables.strInputSelSetColNames;
	}
	if (mapInputVariables.hasOwnProperty('strMultElemColData')) {
		strMultElemColData = mapInputVariables.strMultElemColData;
	}
	//UL alternate values see: Sourcing|Templates and Libraries|Event Libraries for example
	let strULElemName = gblNull;
	let strULNoValElemName = gblNull;
	let arryULElementNames;
	let arryULElemXPaths;
	if (mapInputVariables.hasOwnProperty('strResultsCellULElemName')) {
		strULElemName = mapInputVariables.strResultsCellULElemName;
	}
	if (mapInputVariables.hasOwnProperty('strResultsCellULNoValueElemName')) {
		strULNoValElemName = mapInputVariables.strResultsCellULNoValueElemName;
	}
	if (mapInputVariables.hasOwnProperty('arrayResultsColULObjNames')) {
		arryULElementNames = mapInputVariables.arrayResultsColULObjNames;
	}
	if (mapInputVariables.hasOwnProperty('arrayResultsColULObjXPaths')) {
		arryULElemXPaths = mapInputVariables.arrayResultsColULObjXPaths;
	}
	//Alternate values for Group-record-icons used in Supplier Search Results new UI Jun 2023 JTR-20847
	let strGRIElemName = gblNull;
	let strGRIItemName = gblNull;
	let strGRIItemXPath = gblNull;
	let strGRIItemChildXPath = gblNull;
	let strGRIImgAttribute = gblNull;
	if (mapInputVariables.hasOwnProperty('strGRIElemName')) {
		strGRIElemName = mapInputVariables.strGRIElemName;
	}
	if (mapInputVariables.hasOwnProperty('strGRIItemName')) {
		strGRIItemName = mapInputVariables.strGRIItemName;
	}
	if (mapInputVariables.hasOwnProperty('strGRIItemXPath')) {
		strGRIItemXPath = mapInputVariables.strGRIItemXPath;
	}
	if (mapInputVariables.hasOwnProperty('strGRIItemChildXPath')) {
		strGRIItemChildXPath = mapInputVariables.strGRIItemChildXPath;
	}
	if (mapInputVariables.hasOwnProperty('strGRIImgAttribute')) {
		strGRIImgAttribute = mapInputVariables.strGRIImgAttribute;
	}
	//Alternate value for custom Widget table on AP Home
	let strElemContainsSVGNameXPath = gblNull;
	let strElemContainsSVGNameProperty = gblNull;
	if (mapInputVariables.hasOwnProperty('strElemContainsSVGNameXPath')) {
		strElemContainsSVGNameXPath = mapInputVariables.strElemContainsSVGNameXPath;
	}
	if (mapInputVariables.hasOwnProperty('strElemContainsSVGNameProperty')) {
		strElemContainsSVGNameProperty = mapInputVariables.strElemContainsSVGNameProperty;
	}
	//Alternate value for list box
	let strElemSingleValListFullPath = gblNull;
	let strElemSingleValListName = gblNull;
	if (mapInputVariables.hasOwnProperty('strElemSingleValListFullPath')) {
		strElemSingleValListFullPath = mapInputVariables.strElemSingleValListFullPath;
	}
	if (mapInputVariables.hasOwnProperty('strElemSingleValListName')) {
		strElemSingleValListName = mapInputVariables.strElemSingleValListName;
	}
	//Alternate value for the filter list boxes
	let mapSetVerfiyFilterEditBoxValues; // Type inferred
	if (mapInputVariables.hasOwnProperty('mapSetVerfiyFilterEditBoxValues')) {
		mapSetVerfiyFilterEditBoxValues = mapInputVariables.mapSetVerfiyFilterEditBoxValues;
	}
	//Check for columns with more than one element
	let strMultElemColName = gblNull;
	let intMultElemColCount = 1; //always at least 1
	if (strMultElemColData != gblNull) {
		//Split the value into an array
		//ADD a split for the value to return the date/time format
		let mapArray = {}; // Mimic Groovy Map
		mapArray = StringsAndNumbers.JComm_StringToArray(strMultElemColData, gblDelimiter);
		if (StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount) == 2) {
			strMultElemColName = mapArray.ArryOfValues[0];
			intMultElemColCount = StringsAndNumbers.JComm_StringToInteger(mapArray.ArryOfValues[1]);
		}
	}
	let intResultsRowCnt; // Type inferred
	let intRowMatched;	// Type inferred
	let boolProcCellElements; // Type inferred
	let arryOutputColNames; // Type inferred
	let strDelLstColNames; // Type inferred
	let intAddlColumnCnt = 0; // Type inferred
	let strSelectItemValue = gblNull; // Type inferred
	//Check for alternate name values
	let intAltColNameArrayCnt = -1; // Type inferred
	let intAltColNameUsed = -1; // Type inferred
	let arryAltColName; // Type inferred
	//let strTempAltColName; // Declared but unused
	if (strHdrColMissingNames != gblNull) {
		//Split the value into an array
		//ADD a split for the value to return the date/time format
		let mapArray = {}; // Mimic Groovy Map
		mapArray = StringsAndNumbers.JComm_StringToArray(strHdrColMissingNames, gblDelimiter);
		intAltColNameArrayCnt = StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount);
		arryAltColName = mapArray.ArryOfValues;
	}
	let lstCheckboxes = ['Checkbox', 'Graphic Checkbox', 'JavaScript Checkbox']; // Mimic Groovy def list
	let lstListBoxes = ['listbox', 'Single Value Filter List Box']; // Mimic Groovy def list
	let lstAddlColElemTypes = ['Editbox', 'Checkbox', 'Graphic Checkbox', 'JavaScript Checkbox', 'listbox', 'Single Value Filter List Box']; // Mimic Groovy def list
	if (lstAddlColElemTypes.includes(strSelectElementType) == true) { // Use .includes for JS array
		intAddlColumnCnt++;
	}
	//Check for using select data from the input sheet
	if (StringsAndNumbers.JComm_TextLocationInString(strSelectColumnName, gblUseInputColName) >= 0) {
		intAddlColumnCnt++;
	}
	if (StringsAndNumbers.JComm_TextLocationInString(strSelectElementName, gblUseInputColName) >= 0) {
		intAddlColumnCnt++;
	}
	// strSelectElementChecked is in Javadoc but not used to compute intAddlColumnCnt
	if (StringsAndNumbers.JComm_TextLocationInString(strSelectedItemValueSelected, gblUseInputColName) >= 0) {
		intAddlColumnCnt++;
	}
	//Verify we have the table.
	let weTable = CWCore.returnWebElement(strLocXpath + strTblXPath);
	if (weTable == null) {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The location '" + strLocName + "' '" + strTblName + "' DID NOT RETURN an ELEMENT!!!";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weTable, 'Element', strTblName);
		}
		//Check if the header row is present
		let weHeaderRow = CWCore.returnWebElement(strLocXpath + strTblXPath + strHdrRowXPath);
		if (weHeaderRow == null) {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location '" + strLocName + "' '" + strTblName + "' DID NOT RETURN A HEADER ROW ELEMENT!!!";
		}
		else {
			//Highlight
			if (TCExecParams.getBoolDoHighlight() == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weHeaderRow, 'Element', strTblName + ' header row');
			}
			//Add the column names
			//Return all the column names and use to create the output sheet header columns
			let mapGetHdrChildElements = {}; // Mimic Groovy Map
			let strColNames;
			mapGetHdrChildElements = CWCore.returnChildElements(weHeaderRow, strHdrColXPath);
			let lstHdrCols = mapGetHdrChildElements.lstChildObjects;
			let arryColumnNames;
			let cntHdrCols = mapGetHdrChildElements.cntChildObjs;
			if (cntHdrCols > 0) {
				let weHdrCol;
				let strTempValue;
				let boolAttribPresent;
				//Return the column names and replace any spaces with '_'
				for (let loopHdrCols = 0; loopHdrCols < cntHdrCols; loopHdrCols++) {
					strTempValue = gblNull;
					weHdrCol = lstHdrCols[loopHdrCols];
					//Highlight
					if (TCExecParams.getBoolDoHighlight() == true) {
						let mapHighlight = {}; // Mimic Groovy Map
						mapHighlight = CWCore.objHighlightElementJS(weHdrCol, 'Element', strTblName + ' header row column');
					}
					boolAttribPresent = CWCore.isAttribtuePresent(weHdrCol, strPropertyHdrColName);
					if (boolAttribPresent == true) {
						strTempValue = StringsAndNumbers.JComm_HandleNoData(weHdrCol.getAttribute(strPropertyHdrColName));
					}
					else if (intAltColNameArrayCnt >= 1 && intAltColNameArrayCnt > intAltColNameUsed) {
						intAltColNameUsed++;
						strTempValue = StringsAndNumbers.JComm_HandleNoData(arryAltColName[intAltColNameUsed]);
					}
					else {
						strTempValue = gblNull;
					}
					if (strTempValue != gblNull) {
						if (loopHdrCols == 0) {
							strColNames = StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_");
							strDelLstColNames = strTempValue;
						}
						else {
							strColNames = strColNames + gblDelimiter + StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_");
							strDelLstColNames = strDelLstColNames + "|" + strTempValue;
						}
						strColNames = strColNames + gblDelimiter + StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_") + "_Element"; //Add the column for the element

					}
					else {
						boolPassed = false;
						strMethodDetails = "FAILED!!! The header column '" + loopHdrCols + "' does not CONTAIN THE PROPERTY REQUIRED of:" + strPropertyHdrColName + "!!!";
						break;
					}
				}
				if (boolPassed == true) {
					//Add column names to the array
					let mapStringToArry = {}; // Mimic Groovy Map
					mapStringToArry = StringsAndNumbers.JComm_StringToArray(strDelLstColNames, gblDelimiter);
					arryColumnNames = mapStringToArry.ArryOfValues;
				}
				//Load the inputsheet
				let intInputRowCnt;
				let intInputColCnt;
				// let intHdrColCnt; // Already defined from displayed header
				let shInput; // Mimic XSSFSheet object
				// let strHdrInputShtName = mapInputVariables.strHdrInputShtName; // Not used in this block, declared but unused.
				let mapOpenInputSheet = {}; // Mimic Groovy Map
				mapOpenInputSheet = ExcelData.excelGetSheetByName(TCObj.getObjWorkbook(), strInputSheetName);
				if (StringsAndNumbers.JComm_StringToBoolean(mapOpenInputSheet.boolPassed) == true) {
					shInput = mapOpenInputSheet.objWbSheet;
					//Return the row and column Count
					let mapSheetRowColCnt = {}; // Mimic Groovy Map
					mapSheetRowColCnt = ExcelData.excelGetRowAndColCount(shInput);
					if (StringsAndNumbers.JComm_StringToBoolean(mapSheetRowColCnt.boolPassed) == true) {
						intInputRowCnt = mapSheetRowColCnt.RowCount;
						intInputColCnt = mapSheetRowColCnt.ColCount;
					}
					else {
						boolPassed = false;
						strMethodDetails = mapSheetRowColCnt.strMethodDetails;
					}
					// This section of Groovy code is redundant regarding `intHdrColCnt` and `lstHdrCols`
					// as they were already correctly derived from `weHeaderRow` above.
					// Keep the original control flow but ignore the re-assignment if it's meant to be the displayed count.
					if (boolPassed == true) {
						let mapHdrChildren = CWCore.returnChildElements(weHeaderRow, strHdrColXPath);
						// intHdrColCnt = mapHdrChildren.cntChildObjs; // Do not re-assign display specific count
						// lstHdrCols = mapHdrChildren.lstChildObjects; // Do not re-assign
					}
				}
				else {
					boolPassed = false;
					strMethodDetails = mapOpenInputSheet.strMethodDetails;
				}
				//Check if the number of columns displayed matches the number in the input sheet if not call output only and fail the step
				if (boolPassed == true) {
					let intInputTempColCnt = intInputColCnt - intAddlColumnCnt;
					if (cntHdrCols != intInputTempColCnt / 2) { //Each displayed column consist of the value and element column
						boolPassed = false;
						strMethodDetails = "FAILED!!! The table header is displaying: " + cntHdrCols + " BUT, EXPECTED: " + intInputTempColCnt / 2 + " COLUMNS!!! OUTPUTTING HEADER ONLY see details: ";
					}
					else {
						//Return the input sheet column names and check if the match the output column names
						let strInputColValues;
						let mapGetInputColNames = {}; // Mimic Groovy Map
						mapGetInputColNames = ExcelData.excelGetHdrColNames(shInput);
						let boolInputColNames = StringsAndNumbers.JComm_StringToBoolean(mapGetInputColNames.boolPassed);
						if (boolInputColNames == true) {
							let strTempInputColNames;
							strInputColValues = mapGetInputColNames.ColValues;
							if (intAddlColumnCnt > 0) {
								//Trim off the end values by count
								let strTmpValue = strInputColValues;
								let strTempRight;
								for (let loopTrim = 0; loopTrim < intAddlColumnCnt; loopTrim++) {
									if (loopTrim == 0) {
										strTempRight = StringsAndNumbers.JComm_GetRightTextInString(strTmpValue, gblDelimiter);
									}
									else {
										strTempRight = StringsAndNumbers.JComm_GetRightTextInString(strTmpValue, gblDelimiter) + gblDelimiter + strTempRight;
									}
									// This line needs careful translation
									strTmpValue = StringsAndNumbers.JComm_GetLeftTextInString(strTmpValue, gblDelimiter + StringsAndNumbers.JComm_GetRightTextInString(strTmpValue, gblDelimiter));
								}
								strTempInputColNames = StringsAndNumbers.JComm_GetLeftTextInString(strInputColValues, gblDelimiter + strTempRight);
							}
							else {
								strTempInputColNames = strInputColValues;
							}
							//Check if they match the output column names
							if (strTempInputColNames != strColNames) {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The input column names '" + strTempInputColNames + "' DOES NOT MATCH the OUTPUT Columns '" + strColNames + "'!!!";
							}
							else {
								strMethodDetails = "The column names matched and the input file contains: " + strTempInputColNames + " rows, and " + intInputColCnt + " columns. Does Not Include Edit or Set input column(s).";
								let mapSplitElementString = StringsAndNumbers.JComm_StringToArray(strColNames, gblDelimiter);
								// let intCntElemNames= StringsAndNumbers.JComm_StringToInteger(mapSplitElementString.intItemCount); // Declared but unused
								arryOutputColNames = mapSplitElementString.ArryOfValues;
							}
						}
						else {
							boolPassed = false;
							strMethodDetails = mapGetInputColNames.strMethodDetails;
						}
					}
				}
				if (boolPassed == true) {
					//Check if the cell elements will be verified
					if (arryElementNames == null || arryElemXPaths == null) {
						boolProcCellElements = false;
					}
					else if (arryElementNames.length != arryElemXPaths.length) { // Use .length for JS Array
						boolPassed = false;
						strMethodDetails = "FAILED!!! THE arryElementNames and arryElemXPaths SIZE DO NOT MATCH!!!";
					}
					else {
						boolProcCellElements = true;
					}
				}
				//Return the data and verify the results
				if (boolPassed == true) {
					//Set the start and end rows
					let loopInputRowStart;
					let loopInputRowEnd;
					let loopResultStart;
					let loopResultEnd;
					let intInputCol;
					let intDataElem; // Renamed to avoid shadowing
					// let intOutputRow; // Declared but unused
					let boolInputRowMatch = true; //reset after each loop processed for input data.
					let mapSetLoopStartEnd = {}; // Mimic Groovy Map
					mapSetLoopStartEnd = ExcelData.excelReturnLoopStartEndRow(intInputDataStartRow, intInputDataEndRow, intInputRowCnt);
					loopInputRowStart = mapSetLoopStartEnd.intLoopStart;
					loopInputRowEnd = mapSetLoopStartEnd.intLoopEnd;
					let intInputRowsToProcess = loopInputRowEnd - loopInputRowStart + 1;
					let strTempInputValue;
					let strTempInputElem;
					let strTempDispValue;
					let strTempDispElem;
					let strColor;
					let strBckgColor;
					let intRowsMatched = 0; // Initialize for cumulative count
					let intOutputExcelRow;
					let intBodyRowColCnt;
					let intOutputColumn;
					let intCellElemCnt;
					let intOutputValColIndex;
					let intOutputElemColIndex;
					//Status Counters
					let intCellSkipped = 0; // Initialize for cumulative count
					let intCellIgnored = 0;
					let intCellPassed = 0;
					let intCellFailed = 0;

					let strTempCellValue;
					let strTempCellElements;
					let weRow;
					let weCell;
					let mapGetRowCols;
					let lstBodyRowCols;
					let mapGetInputValue;
					let mapSetCellValue;
					let boolGetInpValue;
					//Check if the tbody is a separate XPath
					if (strTblSepXpath != gblNull) {
						//Return the body as a new table element
						weTable = CWCore.returnWebElement(strLocXpath + strTblXPath + strTblSepXpath);
						//Highlight
						if (TCExecParams.getBoolDoHighlight() == true) {
							let mapHighlight = {}; // Mimic Groovy Map
							mapHighlight = CWCore.objHighlightElementJS(weTable, 'Table Body', strTblName);
						}
					}
					let mapGetRowElements = CWCore.returnChildElements(weTable, strResultsRowXPath);
					let lstBodyRows = mapGetRowElements.lstChildObjects;
					let intBodyRowCnt = mapGetRowElements.cntChildObjs;
					intRowMatched = -1; //Set to show no rows are matching
					intOutputExcelRow = 1; //Increment the row out to set for first output
					//Check the results against the rows to process
					if (intInputRowsToProcess <= intBodyRowCnt) {
						for (let loopInputRow = loopInputRowStart; loopInputRow <= loopInputRowEnd; loopInputRow++) {
							boolInputRowMatch = true;
							// Note: loopResultStart/End are reset per input row due to boolMatchInputRowNumber logic
							loopResultStart = 0; //Zero based index for displayed rows
							loopResultEnd = intBodyRowCnt;
							for (let loopDisplayRow = loopResultStart; loopDisplayRow < loopResultEnd; loopDisplayRow++) {
								boolInputRowMatch = true;
								intInputCol = 0;
								intDataElem = 1; //Set to verify data first
								weRow = lstBodyRows[loopDisplayRow]; // Access array by index
								//Move to the row to show the row we are processing in case of errors.
								let actions = new Actions(TCObj.getTcDriver());
								actions.moveToElement(weRow).perform();
								//Highlight
								if (TCExecParams.getBoolDoHighlight() == true) {
									let mapHighlight = {}; // Mimic Groovy Map
									mapHighlight = CWCore.objHighlightElementJS(weRow, 'Element', strTblName + ' body row');
								}
								//TODO inject change for JTR-19321 (This logic is for section rows)
								//Return the SubSection values if present
								let intSubSecColCnt = -1;
								let strColProp = gblNull;
								let strColPropValue = gblNull;
								let strColElemName = gblNull;
								if (strSectionRowInfo != gblNull) {
									intSubSecColCnt = StringsAndNumbers.JComm_StringToInteger(arrySectInfo[0]);
									strColProp = arrySectInfo[1];
									strColPropValue = arrySectInfo[2];
									strColElemName = arrySectInfo[3];
								}
								//1|class|SubSectionTitle|elemSubSection
								//Return the column elements and count
								mapGetRowCols = CWCore.returnChildElements(weRow, strResultsColXPath);
								lstBodyRowCols = mapGetRowCols.lstChildObjects;
								intBodyRowColCnt = mapGetRowCols.cntChildObjs;
								//Check for a spacer row (Original logic in Groovy mixes spacer and section row checks in `VerifyResultBody`)
								if (intBodyRowColCnt == intSubSecColCnt && strSectionRowInfo !== gblNull) { // If section row, process it
									let boolIsSection = false;
									intOutputColumn = 0;
									//Create the temp counters for the current row
									let intCellSkippedTemp = 0;
									let intCellIgnoredTemp = 0;
									let intCellPassedTemp = 0;
									let intCellFailedTemp = 0;
									let boolComparePass = true;
									//Process the subSection
									for (let loopRowCol = 0; loopRowCol < cntDispCols; loopRowCol++) { // Iterate through displayed columns
										strTempCellValue = gblNull;
										strTempCellElements = gblNull;
										if (loopRowCol < intSubSecColCnt) {
											//Return the cell object
											weCell = lstBodyRowCols[loopRowCol];
											//Highlight
											if (TCExecParams.getBoolDoHighlight() == true) {
												let mapHighlight = {}; // Mimic Groovy Map
												mapHighlight = CWCore.objHighlightElementJS(weCell, 'Element', strTblName + ' body cell');
											}
											//Return the cell text
											strTempCellValue = StringsAndNumbers.JComm_HandleNoData(weCell.getText());
											//Confirm the cell belongs to the subsection
											let boolAttribPresent = CWCore.isAttribtuePresent(weCell, strColProp);
											if (boolAttribPresent == true) {
												let strTempValue = StringsAndNumbers.JComm_HandleNoData(weCell.getAttribute(strColProp));
												if (StringsAndNumbers.JComm_VerifyTextPresent(strTempValue, strColPropValue, "%Con%") == true) {
													boolIsSection = true;
													strTempCellElements = strColElemName;
												}
											}
										}
										//Get the input values FOR THE CURRENT INPUT ROW
										strTempInputValue = gblNull;
										strTempInputElem = gblNull;
										//Return the expected values for the "current" input row.
										mapGetInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow - intSpacerCnt, arryOutputColNames[intInputCol]);
										boolGetInpValue = StringsAndNumbers.JComm_StringToBoolean(mapGetInputValue.boolPassed);
										if (boolGetInpValue == true) {
											strTempInputValue = StringsAndNumbers.JComm_HandleNoData(mapGetInputValue.CellValue);
											intOutputValColIndex = mapGetInputValue.intColIndex;
											//Check if the value is a dynamic date and/or time
											if (strTempInputValue.indexOf('D[') > 0 || strTempInputValue.indexOf('T[') > 0) {
												let myTemp = DateTime.datetimeReturnDynamicFormatedDateTime(strTempInputValue);
												strTempInputValue = myTemp;
											}
										}
										intInputCol++; //Increment the input column number for the element part
										mapGetInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow - intSpacerCnt, arryOutputColNames[intInputCol]);
										boolGetInpValue = StringsAndNumbers.JComm_StringToBoolean(mapGetInputValue.boolPassed);
										if (boolGetInpValue == true) {
											strTempInputElem = StringsAndNumbers.JComm_HandleNoData(mapGetInputValue.CellValue);
											intOutputElemColIndex = mapGetInputValue.intColIndex;
										}
										intInputCol++; //Increment the input column number for the next pair

										//Compare each value and output the results
										let strCompareDispValue;
										let strCompareInputValue;
										let strOutputCellTheme;
										let intOutputColForCell; // Renamed to avoid outer intOutputColumn
										for (let loopCompareInner = 1; loopCompareInner <= 2; loopCompareInner++) {
											if (loopCompareInner == 1) {
												strCompareDispValue = strTempCellValue;
												strCompareInputValue = strTempInputValue;
												intOutputColForCell = intInputCol - 2;
											}
											else {
												strCompareDispValue = strTempCellElements;
												strCompareInputValue = strTempInputElem;
												intOutputColForCell = intInputCol - 1;
											}
											//Check if they match
											if (strCompareInputValue == gblSkip || strCompareInputValue == gblIgnoreData) {
												strOutputCellTheme = 'TestRunDataSkipIgnore';
												if (strCompareInputValue == gblSkip) {
													intCellSkippedTemp++;
												}
												else {
													intCellIgnoredTemp++;
												}
											}
											else if (strCompareInputValue == strCompareDispValue) {
												strOutputCellTheme = 'TestRunPassStd';
												intCellPassedTemp++;
											}
											else {
												strOutputCellTheme = 'TestRunFailStd';
												strCompareDispValue = `Expected: ${strCompareInputValue}${gblLineFeed} Cell Displayed: ${strCompareDispValue}`;
												boolComparePass = false;
												boolInputRowMatch = false;
												intCellFailedTemp++;
												//Check if we should break if row number does not need to match
											}
											//Output the results to the cell
											if (boolComparePass == true || boolMatchInputRowNumber == true) {
												mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intOutputExcelRow, intOutputColForCell, strCompareDispValue, strOutputCellTheme);
												if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
													boolPassed = false;
													strMethodDetails = mapSetCellValue.strMethodDetails;
												}
											}
											//Check if we should continue the compare
											if (boolComparePass == false && boolMatchInputRowNumber == true) {
												if (TCExecParams.getBoolDoDebug() == true) {
													Tester.Message("Compare failed and matchinputRowNumber == true so continue compare");
												}
											}
											else if (boolComparePass == false) {
												if (TCExecParams.getBoolDoDebug() == true) {
													Tester.Message("Compare failed and matchinputRowNumber == false so break for compare");
												}
												//Break so we move to the next row in the display
												break;
											}
										}
										//Check if we should continue the compare
										if (boolComparePass == false && boolMatchInputRowNumber == true) {
											boolPassed = false; // Assignment
											strMethodDetails = strMethodDetails + "FAILED!!! Display Row: " + loopDisplayRow + " and cell in column: " +
												loopRowCol + " DID NOT MATCH SEE OUTPUT SHEET!!!";
										}
										else if (boolComparePass == false) {
											if (TCExecParams.getBoolDoDebug() == true) {
												Tester.Message("Compare failed and matchinputRowNumber == false so break for columns");
											}
											//Break so we move to the next row in the display
											break;
										}
									}
									if (boolIsSection == false) {
										boolPassed = false;
										strMethodDetails = `FAILED!!! Result row: ${loopDisplayRow + 1} CELL(s) PROPERTY DID NOT CONTAIN The VALUE '${strColPropValue}' !!!`;
									}
									if (boolComparePass == true) {
										intRowsMatched++;
									}
									//Check if row passed only if row number is to match as well
									if (boolComparePass == false && boolMatchInputRowNumber == true) {
										boolPassed = false;
									}
									if (boolMatchInputRowNumber == true) { // Cumulative counters for the entire function CommonWeb_scope
										//Increment to output row only if row number match is true
										intOutputExcelRow++;
										//Update the steps processed counts
										intCellSkipped = intCellSkipped + intCellSkippedTemp; // The 'Temp' variables scope needs careful handling.
										intCellIgnored = intCellIgnored + intCellIgnoredTemp;
										intCellPassed = intCellPassed + intCellPassedTemp;
										intCellFailed = intCellFailed + intCellFailedTemp;
									}
								}
								//Check if a spacer row is present with no columns
								else if (intBodyRowColCnt > 0) { //Skips process the row.
									if (intBodyRowColCnt == cntHdrCols) { // Assuming intHdrColCnt from outer scope.
										//Create the temp counters (these should probably be local to this block, if "Temp" suffix is correct)
										let intCellSkippedTemp = 0;
										let intCellIgnoredTemp = 0;
										let intCellPassedTemp = 0;
										let intCellFailedTemp = 0;
										let boolComparePass = true;
										for (let loopDisplayCol = 0; loopDisplayCol < cntHdrCols; loopDisplayCol++) { // Use cntHdrCols from outer scope
											strTempInputValue = gblNull;
											strTempInputElem = gblNull;
											//Return the expected values
											mapGetInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow - intSpacerCnt, arryOutputColNames[intInputCol]);
											boolGetInpValue = StringsAndNumbers.JComm_StringToBoolean(mapGetInputValue.boolPassed);
											if (boolGetInpValue == true) {
												strTempInputValue = StringsAndNumbers.JComm_HandleNoData(mapGetInputValue.CellValue);
												intOutputValColIndex = mapGetInputValue.intColIndex;
												//Check if the value is a dynamic date and/or time
												if (strTempInputValue.indexOf('D[') > 0 || strTempInputValue.indexOf('T[') > 0) {
													let myTemp = DateTime.datetimeReturnDynamicFormatedDateTime(strTempInputValue);
													strTempInputValue = myTemp;
												}
											}
											intInputCol++; //Increment the input column number
											mapGetInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow - intSpacerCnt, arryOutputColNames[intInputCol]);
											boolGetInpValue = StringsAndNumbers.JComm_StringToBoolean(mapGetInputValue.boolPassed);
											if (boolGetInpValue == true) {
												strTempInputElem = StringsAndNumbers.JComm_HandleNoData(mapGetInputValue.CellValue);
												intOutputElemColIndex = mapGetInputValue.intColIndex;
											}
											intInputCol++; //Increment the input column number
											weCell = lstBodyRowCols[loopDisplayCol];
											//Highlight
											if (TCExecParams.getBoolDoHighlight() == true) {
												let mapHighlight = {}; // Mimic Groovy Map
												mapHighlight = CWCore.objHighlightElementJS(weCell, 'Element', strTblName + ' body cell');
											}
											//Return the cell value
											strTempDispValue = StringsAndNumbers.JComm_HandleNoData(weCell.getText());
											//Return the cell element
											strTempDispElem = gblNull; // Reset for this cell
											let weCellElement; // Type inferred
											if (boolProcCellElements == true) {
												//Set the wait to 1 second to speed up the processing of elements in the cells
												TCObj.tcDriver.manage().timeouts().implicitlyWait(100, TimeUnit.MILLISECONDS);
												//Process the array for elements
												let intMaxElem = 1;
												if (strMultElemColName == arryColumnNames[loopDisplayCol]) {
													intMaxElem = intMultElemColCount;
												}
												strTempCellElements = gblNull; // Reset for this cell
												//Process the array for elements
												let cntCellElements = 0;
												while (cntCellElements < intMaxElem) {
													for (let loopElements = 0; loopElements < arryElementNames.length; loopElements++) {
														//TODO add in classes to get color such as status elements
														intCellElemCnt = CWCore.returnWebElemChildElementCount(weCell, arryElemXPaths[loopElements]);
														if (intCellElemCnt == 1) {
															cntCellElements++; //Increment the number of elements in the cell
															weCellElement = CWCore.returnChildElement(weCell, arryElemXPaths[loopElements]);
															//Highlight
															if (TCExecParams.getBoolDoHighlight() == true) {
																let mapHighlight = {}; // Mimic Groovy Map
																mapHighlight = CWCore.objHighlightElementJS(weCellElement, 'Element', strTblName + ' cell status element');
															}
															if (cntCellElements == 1) { // Only first element detected
																//Check if the element is a UL
																if (arryElementNames[loopElements] == strULElemName) { //See Sourcing|Templates and Libraries|Event Libraries for UI example
																	//TODO we need to pass in the ul element names and xpaths
																	//let intItemCnt = 0; // Declared but unused
																	//let strItemValue; // Declared but unused
																	//let strItemElem; // Declared but unused
																	let weULItem;
																	//Process the UL to return the items
																	let mapULItems = CWCore.returnChildElements(weCellElement, arryULElemXPaths[0]);
																	let lstULItems = mapULItems.lstChildObjects;
																	let intULItemCnt = mapULItems.cntChildObjs;
																	if (intULItemCnt == 0) {
																		//Error
																		boolPassed = false;
																		strMethodDetails = "FAILED!!! The Cell User List contain 'ZERO' ITEM(S)!!!!";
																	}
																	else {
																		let weULItemChild;
																		for (let loopULItemsInner = 0; loopULItemsInner < intULItemCnt; loopULItemsInner++) {
																			//Return the listitem
																			weULItem = lstULItems[loopULItemsInner];
																			//Add a ';' if loop is for greater than the first element to Cell and Element values
																			if (loopULItemsInner > 0) {
																				strTempCellValue = strTempCellValue + ';';
																				strTempCellElements = strTempCellElements + ';';
																			}
																			//Highlight
																			if (TCExecParams.getBoolDoHighlight() == true) {
																				let mapHighlight = {}; // Mimic Groovy Map
																				mapHighlight = CWCore.objHighlightElementJS(weULItem, 'Element', 'UserList item_' + loopULItemsInner);
																			}
																			//Check if the child items are present
																			//Loop through the UL items elements XPaths to check for a single item
																			//Start with the second element since 0 is the listitem
																			let intElemFoundCnt = 0;
																			for (let loopULItemElem = 1; loopULItemElem < arryULElementNames.length; loopULItemElem++) {
																				weULItemChild = CWCore.returnChildElement(weULItem, arryULElemXPaths[loopULItemElem]);
																				if (weULItemChild != null) {
																					intElemFoundCnt++;
																					//Highlight
																					if (TCExecParams.getBoolDoHighlight() == true) {
																						let mapHighlight = {}; // Mimic Groovy Map
																						mapHighlight = CWCore.objHighlightElementJS(weULItemChild, 'Element', 'UserList item_child' + loopULItemsInner);
																					}
																					//Return the text and Add the element name and text to the variables
																					if (intElemFoundCnt == 1 && loopULItemsInner == 0) {
																						strTempCellValue = weULItemChild.getText();
																						strTempCellElements = arryULElementNames[loopULItemElem];
																					}
																					else if (intElemFoundCnt == 1 && loopULItemsInner > 0) {
																						strTempCellValue = strTempCellValue + weULItemChild.getText();
																						strTempCellElements = strTempCellElements + arryULElementNames[loopULItemElem];
																					}
																					else {
																						strTempCellValue = strTempCellValue + "|" + weULItemChild.getText();
																						strTempCellElements = strTempCellElements + "|" + arryULElementNames[loopULItemElem];
																					}
																				}
																			}
																		}
																	}
																	strTempDispElem = strTempCellElements;
																	strTempDispValue = strTempCellValue;
																}
																else if (arryElementNames[loopElements] == strGRIElemName) { //Check for the Group Record Image element
																	//let intItemCnt = 0; // Declared but unused
																	//let strItemValue; // Declared but unused
																	//let strItemElem; // Declared but unused
																	//let weItem; // Declared but unused
																	//Process the parent element to return the items
																	let mapGRIItems = CWCore.returnChildElements(weCellElement, strGRIItemXPath);
																	let lstGRIItems = mapGRIItems.lstChildObjects;
																	let intGRIItemCnt = mapGRIItems.cntChildObjs;
																	if (intGRIItemCnt == 0) {
																		//Error
																		boolPassed = false;
																		strMethodDetails = "FAILED!!! The Cell " + strGRIElemName + " contains 'ZERO' ITEM(S)!!!!";
																	}
																	else {
																		let weGRIItem;
																		let weGRIItemChild;
																		let intElemFoundCnt = 0;
																		//Loop through the GRIItems each will contain a child that should have an image. Return the attribute value.
																		for (let intLoopItemsInner = 0; intLoopItemsInner < intGRIItemCnt; intLoopItemsInner++) {
																			weGRIItem = lstGRIItems[intLoopItemsInner];
																			intElemFoundCnt++;
																			let strItemValue = gblNull;
																			//Highlight
																			if (TCExecParams.getBoolDoHighlight() == true) {
																				let mapHighlight = {}; // Mimic Groovy Map
																				mapHighlight = CWCore.objHighlightElementJS(weGRIItem, 'Element', 'GRIElement Item' + intLoopItemsInner);
																			}
																			//Return the item child by finding elements and verifying only 1 is present
																			let intGRIItemChildCnt = CWCore.returnWebElemChildElementCount(weGRIItem, strGRIItemChildXPath); // Original also set intCellElemCnt =
																			//Check if only 1 item is present
																			if (intGRIItemChildCnt == 1) {
																				weGRIItemChild = CWCore.returnChildElement(weGRIItem, strGRIItemChildXPath);
																				//Return the value
																				let boolAttribPresent = CWCore.isAttribtuePresent(weGRIItemChild, strGRIImgAttribute);
																				if (boolAttribPresent == true) {
																					strItemValue = StringsAndNumbers.JComm_HandleNoData(weGRIItemChild.getAttribute(strGRIImgAttribute));
																				}
																				//Add to the output value
																				if (intElemFoundCnt == 1) {
																					strTempDispValue = strItemValue;
																					strTempDispElem = 'Record Image';
																				}
																				else {
																					strTempDispValue = strTempDispValue + "|" + strItemValue;
																					strTempDispElem = strTempDispElem + "|" + 'Record Image';
																				}
																			}
																			else {
																				//Fail (original code handled empty block, adding explicit fail)
																				boolPassed = false;
																				strMethodDetails = "FAILED!!! The GRI Item contains " + intGRIItemChildCnt + " children, expected 1.";
																			}
																		}
																	}
																}
																else {
																	strTempDispElem = arryElementNames[loopElements];
																	//let intInString = StringsAndNumbers.JComm_TextLocationInString(strTempDispElem.toUpperCase(),'CHECKBOX'); // Declared but unused
																	if (StringsAndNumbers.JComm_TextLocationInString(strTempDispElem.toUpperCase(), 'CHECKBOX') >= 0) {
																		strTempDispValue = CWCore.objGetCheckboxChecked(weCellElement).toString();
																		if (strTempDispValue == gblNull) {
																			strTempDispValue = 'false';
																		}
																	}
																	else if (arryElementNames[loopElements] == 'image') {
																		if (CWCore.isAttribtuePresent(weCellElement, strImageAttribute) == true) {
																			strTempDispValue = StringsAndNumbers.JComm_HandleNoData(weCellElement.getAttribute(strImageAttribute));
																		}
																		else {
																			strTempDispValue = gblNull; // Explicitly set if attribute not found from original
																		}
																	}
																}
																if (StringsAndNumbers.JComm_TextLocationInString(arryElemXPaths[loopElements], 'svg') >= 0) {
																	//Check if the SVG image name is not in the svg element
																	if (strElemContainsSVGNameXPath == gblNull) {
																		let strSVGElem = StringsAndNumbers.JComm_HandleNoData(weCellElement.getText());
																		// Groovy used `JComm_TextLocationInString`. In TS, use `startsWith` and `endsWith`.
																		if (strTempDispValue.startsWith(strSVGElem)) {
																			strTempDispValue = strTempDispValue.substring(strSVGElem.length);
																		}
																		else if (strTempDispValue.endsWith(strSVGElem)) {
																			strTempDispValue = strTempDispValue.substring(0, strTempDispValue.length - strSVGElem.length);
																		}
																	}
																	else {
																		//Return the element based on the path
																		let weSVGNameElem = weCellElement.findElement(By.xpath(strElemContainsSVGNameXPath));
																		if (weSVGNameElem != null) {
																			//Check for the element property
																			let boolAttribPresent = CWCore.isAttribtuePresent(weSVGNameElem, strElemContainsSVGNameProperty);
																			if (boolAttribPresent == true) {
																				//Return the value of the property
																				strTempDispValue = StringsAndNumbers.JComm_HandleNoData(weSVGNameElem.getAttribute(strElemContainsSVGNameProperty));
																			}
																			else {
																				strTempDispValue = 'NO IMAGE NAME!!!';
																			}
																		}
																		else {
																			strTempDispValue = 'SVG NAME ELEMENT NOT FOUND!!!'; // Added this explicit error
																		}
																	}
																}
															}
															else { // This means the element was NOT UNIQUE (cntCellElements > 1 for this element) or was not one of the first elements.
																strTempDispElem = strTempDispElem + gblDelimiter + arryElementNames[loopElements];
															}
															//TODO add in classes to get color such as status elements
															if (arryElementNames[loopElements] == strResultsCellStatElem) {
																//Return the element
																let weStatus = CWCore.returnChildElement(weCell, arryElemXPaths[loopElements]);
																if (weStatus == null) {
																	boolPassed = false;
																	strMethodDetails = "FAILED!!! Result row: " + loopDisplayRow + " cell: " + (loopDisplayCol + 1) + " DID NOT CONTAIN A STATUS ELEMENT AS EXEPCTED!!!";
																}
																else {
																	//Highlight
																	if (TCExecParams.getBoolDoHighlight() == true) {
																		let mapHighlight = {}; // Mimic Groovy Map
																		mapHighlight = CWCore.objHighlightElementJS(weStatus, 'Element', strTblName + ' cell status element');
																	}
																	strColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("color"));
																	if (StringsAndNumbers.JComm_TextLocationInString(arryElemXPaths[loopElements], "svg") >= 0) {
																		strBckgColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("fill"));
																	}
																	else {
																		strBckgColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("background-color"));
																	}
																	//Check if the color is the same
																	let strOrgColor = strColor;
																	let strTempColor = StringsAndNumbers.JComm_GetRightTextInString(strColor, "(");
																	strTempColor = StringsAndNumbers.JComm_GetLeftTextInString(strTempColor, ")");
																	let strTempBckColor = StringsAndNumbers.JComm_GetRightTextInString(strBckgColor, "(");
																	strTempBckColor = StringsAndNumbers.JComm_GetLeftTextInString(strTempBckColor, ")");
																	//let intLocInColorToBck = StringsAndNumbers.JComm_TextLocationInString(strTempColor, strTempBckColor); // Declared but unused
																	//let intLocInBckToColor = StringsAndNumbers.JComm_TextLocationInString(strTempBckColor, strTempColor); // Declared but unused
																	let strTempStatusColor;
																	//Return the property and add to the strTempCellElements
																	if (StringsAndNumbers.JComm_TextLocationInString(strTempColor, strTempBckColor) >= 0 || StringsAndNumbers.JComm_TextLocationInString(strTempBckColor, strTempColor) >= 0) {
																		strColor = "RGB(255, 255, 255)"; //Change the color for the text to white. Will work for all but white backgrounds
																		strTempStatusColor = "Colors matched replaced original color:" + strOrgColor + " with Color:" + strColor + "Background Color:" + strBckgColor;
																	}
																	else {
																		strTempStatusColor = "Color:" + strColor + "Background Color:" + strBckgColor;
																	}
																	strTempDispElem = strTempDispElem + gblDelimiter + strTempStatusColor;
																}
															}
														}
													}
													if (cntCellElements == 0) {
														break;
													}
												}
												//Return the wait to maxwaittime
												TCObj.tcDriver.manage().timeouts().implicitlyWait(TCObj.getIntTempMaxWaitTime(), TimeUnit.SECONDS);
											}
											//Compare each value
											let strCompareDispValueForWrite; // Renamed to avoid shadowing
											let strCompareInputValueForWrite; // Renamed to avoid shadowing
											let strOutputCellThemeForWrite; // Renamed to avoid shadowing
											let intOutputColForWrite; // Renamed to avoid shadowing
											for (let loopCompareWrite = 1; loopCompareWrite <= 2; loopCompareWrite++) {
												if (loopCompareWrite == 1) {
													strCompareDispValueForWrite = strTempDispValue;
													strCompareInputValueForWrite = strTempInputValue;
													intOutputColForWrite = intInputCol - 2;
												}
												else {
													strCompareDispValueForWrite = strTempDispElem;
													strCompareInputValueForWrite = strTempInputElem;
													intOutputColForWrite = intInputCol - 1;
												}
												//Check if they match
												if (strCompareInputValueForWrite == gblSkip || strCompareInputValueForWrite == gblIgnoreData) {
													strOutputCellThemeForWrite = 'TestRunDataSkipIgnore';
													if (strCompareInputValueForWrite == gblSkip) {
														intCellSkippedTemp++;
													}
													else {
														intCellIgnoredTemp++;
													}
												}
												else if (strCompareInputValueForWrite != strCompareDispValueForWrite) {
													boolComparePass = false;
													boolInputRowMatch = false;
													break; // Break inner loop if content doesn't match
												} else { // It matches
													boolComparePass = true;
													// No increment on passed, as it sums up all matched in this loop
												}
											}
											if (boolComparePass == false){ // If comparison failed in this column pair
												if (TCExecParams.getBoolDoDebug()== true) {
													Tester.Message("Compare failed and matchinputRowNumber == false so break for columns");
												}
												//Break so we move to the next row in the display
												break;
											}
										}

										if (boolComparePass == true) { // If this display row fully matched the input row
											strTempCellValue = null; // Reset for next iteration (though not explicitly used)
											intRowMatched = loopDisplayRow; //Return the row that matched this input row (0-indexed)
											weCell = null; // Reset for next (though not explicitly used)
											//Check if we need to set or select anything
											//Identify the column that will be used for the selection value
											//Check if the strSelectColumnName contains use input column name
											let strAssgColName = gblNull;
											if (StringsAndNumbers.JComm_TextLocationInString(strSelectColumnName, gblUseInputColName) >= 0) {
												let strTempColName = StringsAndNumbers.JComm_GetRightTextInString(strSelectColumnName, gblDelimiter);
												//Return the value based on the matched input row number and the assigned column name
												// Accessing .CellValue from a map returned by excelGetCellValueByRowNumColName
												strAssgColName = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow, strTempColName).CellValue;
											}
											else {
												strAssgColName = strSelectColumnName;
											}
											let mapSplitElementString = StringsAndNumbers.JComm_StringToArray(strDelLstColNames, gblDelimiter);
											// let intCntElemNames= StringsAndNumbers.JComm_StringToInteger(mapSplitElementString.intItemCount); // Declared but unused
											let arryColName = mapSplitElementString.ArryOfValues; // Inferred to string[]
											let intColIndex = arryColName.indexOf(strAssgColName);
											if (intColIndex >= 0) {
												//Return the cell from the row
												weCell = lstBodyRowCols[intColIndex];
												if (weCell != null) {
													//Highlight
													if (TCExecParams.getBoolDoHighlight() == true) {
														let mapHighlight = {}; // Mimic Groovy Map
														mapHighlight = CWCore.objHighlightElementJS(weCell, 'Cell', strTblName + ' cell in row');
													}
													//Get the element xpath that matches the assigned
													let boolFoundElement = false;
													//Check if the strSelectColumnName contains use input column name
													let strAssgSelElemName = gblNull;
													if (StringsAndNumbers.JComm_TextLocationInString(strSelectElementName, gblUseInputColName) >= 0) {
														// strSelectElementName has "Select_Element" appended to gblUseInputColName.
														// Assuming JComm_GetRightTextInString gets the part after the delimiter, then .CellValue is needed.
														strAssgSelElemName = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow, StringsAndNumbers.JComm_GetRightTextInString(strSelectElementName, gblUseInputColName + gblDelimiter)).CellValue;
													}
													else {
														strAssgSelElemName = strSelectElementName;
													}
													//Check if the strSelectColumnName contains use input column name
													let strAssgSetSelData = gblNull;
													if (StringsAndNumbers.JComm_TextLocationInString(strSelectedItemValueSelected, gblUseInputColName) >= 0) {
														strAssgSetSelData = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow, StringsAndNumbers.JComm_GetRightTextInString(strSelectedItemValueSelected, gblUseInputColName + gblDelimiter)).CellValue;
													}
													else {
														strAssgSetSelData = strSelectItemValue;
													}
													let weCellElement; // Type inferred
													//Check if the cell has ul in it as the ul may have multiple items in the li
													if (arryElementNames.includes(strULElemName) >= 0) { // Check existence using .includes for JS Array
														weCellElement = CWCore.returnChildElement(weCell, arryElemXPaths[arryElementNames.indexOf(strULElemName)]);
														//Highlight
														if (TCExecParams.getBoolDoHighlight() == true) {
															let mapHighlight = {}; // Mimic Groovy Map
															mapHighlight = CWCore.objHighlightElementJS(weCellElement, 'Element', 'UserList');
														}
														if (weCellElement != null) {
															boolFoundElement = true; //Set to true since we found the ul. If the element is not in the li separate fail message is presented.
															//TODO we need to pass in the ul element names and xpaths
															//let intItemCnt = 0; // Declared but unused
															//let strItemValue; // Declared but unused
															//let strItemElem; // Declared but unused
															let weULItem;
															//Process the UL to return the items
															let mapULItems = CWCore.returnChildElements(weCellElement, arryULElemXPaths[0]); // Assumes arryULElemXPaths[0] is UL's direct child xpath
															let lstULItems = mapULItems.lstChildObjects;
															let intULItemCnt = mapULItems.cntChildObjs;
															if (intULItemCnt == 0) {
																//Error
																boolPassed = false;
																strMethodDetails = "FAILED!!! The Cell User List contain 'ZERO' ITEM(S)!!!!";
															}
															else {
																//Loop through the li to find the matching value pair
																let boolValueFound = false;
																let weULItemChild;
																for (let loopULItemsInner = 0; loopULItemsInner < intULItemCnt; loopULItemsInner++) { // Loop starts at 0 to iterate.
																	//Return the listitem
																	weULItem = lstULItems[loopULItemsInner];
																	//Highlight (Note: Groovy original had weULItemChild here, fixed to weULItem)
																	if (TCExecParams.getBoolDoHighlight() == true) {
																		let mapHighlight = {}; // Mimic Groovy Map
																		mapHighlight = CWCore.objHighlightElementJS(weULItem, 'Element', 'UserList item_' + loopULItemsInner);
																	}
																	//Check if the child items are present
																	//Loop through the UL items elements XPaths to check for a single item
																	//Start with the second element since 0 is the listitem (conceptually in Groovy, here loop from 1 for child elems)
																	//let intElemFoundCnt = 0; // local counter reset
																	let intElemIndex = -1;
																	let strULItemValue = null;
																	for (let loopULItemElem = 1; loopULItemElem < arryULElementNames.length; loopULItemElem++) { // Loop starts at 1
																		weULItemChild = CWCore.returnChildElement(weULItem, arryULElemXPaths[loopULItemElem]);
																		if (weULItemChild != null) {
																			//intElemFoundCnt++; // Accumulate for `intElemFoundCnt == 1` condition logic
																			if (arryULElementNames[loopULItemElem] == strAssgSelElemName) {
																				//Capture the element index
																				intElemIndex = loopULItemElem;
																				boolFoundElement = true; // Found the specified element
																			}
																			//Highlight
																			if (TCExecParams.getBoolDoHighlight() == true) {
																				let mapHighlight = {}; // Mimic Groovy Map
																				mapHighlight = CWCore.objHighlightElementJS(weULItemChild, 'Element', 'UserList item_child' + loopULItemsInner);
																			}
																			//Return the text
																			if (strULItemValue == null) {
																				strULItemValue = weULItemChild.getText();
																			}
																			else {
																				strULItemValue = strULItemValue + gblDelimiter + weULItemChild.getText();
																			}
																			// Note: The `strTempCellValue` and `strTempCellElements` assignment logic here gets complex.
																			// It seems to be trying to build a delimited string of all child elements for output,
																			// but also uses `intElemFoundCnt == 1` and `loopULItems == 0/ >0` from `OutputResultBody`.
																			// For simplification in `SetSelectRowFromResultBody`, focusing on `strULItemValue` for match.
																		}
																	}
																	//Check if the value is a match
																	if (strULItemValue == strAssgSetSelData) {
																		//boolFoundElement = true; // Already set by above logic if `strAssgSelElemName` found
																		boolValueFound = true;
																		//Check for the assigned element
																		//Return the Element
																		weULItemChild = CWCore.returnChildElement(weULItem, arryULElemXPaths[intElemIndex]); // Get the specific element to click on
																		//Highlight
																		if (TCExecParams.getBoolDoHighlight() == true) {
																			let mapHighlight = {}; // Mimic Groovy Map
																			mapHighlight = CWCore.objHighlightElementJS(weULItemChild, 'Element', 'UserList item_child' + loopULItemsInner);
																		}
																		strMethodDetails = strMethodDetails + `Results in row: ${loopDisplayRow} matched the results assigned. Clicked the element '${strAssgSelElemName}' in the column '${strAssgColName}' that matched the value of: ${weULItemChild.getText()}.`;
																		//Perform the click action currently only setup to click links
																		weULItemChild.click();
																		boolPassed = true; // Success for this operation
																		intRowsMatched++; // Increment overall match count
																		break; // Break from loopULItemsInner, matched so done with this li
																	}
																}
																if (boolValueFound == false) {
																	boolPassed = false;
																	strMethodDetails = strMethodDetails + `FAILED!!! Results in row: ${loopDisplayRow} matched the results assigned. HOWEVER, did not find the ASSIGNED VALUES of: ${strAssgSetSelData} in the RESULTS DISPLAYED of: ${strTempCellValue}!!!`;
																}
															}
														}
													}
													if (boolFoundElement == false && boolPassed == true) { //No element was found so far try the element name
														//Continue process the element
														let intElementIndex = arryElementNames.indexOf(strAssgSelElemName);
														if (intElementIndex >= 0) {
															// let strElemXPath = arryElemXPaths[intElementIndex]; // Declared but unused
															intCellElemCnt = CWCore.returnWebElemChildElementCount(weCell, arryElemXPaths[intElementIndex]);
															let weElementLocated = CWCore.returnChildElement(weCell, arryElemXPaths[intElementIndex]);
															if (weElementLocated != null) {
																//Highlight
																if (TCExecParams.getBoolDoHighlight() == true) {
																	let mapHighlight = {}; // Mimic Groovy Map
																	mapHighlight = CWCore.objHighlightElementJS(weElementLocated, 'Element', strTblName + ' cell status element');
																}
																boolFoundElement = true;
																//TODO add code for UL and list items. Requires SetSelect Column to contain value|elementName|instance
																//Else process using the original method statements
																let boolActionProcessed = false;
																//Editbox
																if (strSelectElementType == 'editbox' && strAssgSetSelData != gblNA && strAssgSetSelData != gblSkip) {
																	// this.SetVerifyEditBoxValue - Assuming this is a globally available function
																	// Cannot call 'this.SetVerifyEditBoxValue' directly, assuming external helper function.
																	SetVerifyEditBoxValue(strLocName, weElementLocated, strAssgSelElemName, strAssgSetSelData, false, false, false);
																	boolActionProcessed = true; // Action was handled by custom function
																}
																//Return the Input setedit column information from strInputSelSetColNames
																//Convert to array of values
																//Check for alternate name values
																let intSetEditValuseArrayCnt = -1; //Used for the SetEdit Values count which contains the input colum name, object type, and value
																let intSetEditValueArrayCnt = -1; //used to count the number of items in the setedit value
																//let intSetEditValueUsed = -1; // Declared but unused
																let arrySetEditValues;
																let arrySetEditValue;
																let strSetEditTempValues;
																let strSetEditInputColName;
																let strSetEditTempValue;
																if (intAddlColumnCnt > 0) {
																	let mapSetEditArray = {}; // Mimic Groovy Map
																	mapSetEditArray = StringsAndNumbers.JComm_StringToArray(strInputSelSetColNames, ';');
																	intSetEditValuseArrayCnt = StringsAndNumbers.JComm_StringToInteger(mapSetEditArray.intItemCount);
																	arrySetEditValues = mapSetEditArray.ArryOfValues;
																}
																//Checkboxes
																if (lstCheckboxes.includes(strSelectElementType) == true && strAssgSetSelData != gblNA && strAssgSetSelData != gblSkip) {
																	let boolIsChecked; // Type inferred
																	boolActionProcessed = true;
																	if (intSetEditValuseArrayCnt == 1) { // Only expecting one set of values for single element type.
																		//For each column after the element column name split to determine input column, element type, variable type
																		//Return the SetEdit Column Name
																		strSetEditTempValues = arrySetEditValues[0]; //Always index zero as only one column is required
																		//Split the returned value into three items using gblDelimiter
																		let mapSetEditValueArray = {}; // Mimic Groovy Map
																		mapSetEditValueArray = StringsAndNumbers.JComm_StringToArray(strSetEditTempValues, gblDelimiter);
																		intSetEditValueArrayCnt = StringsAndNumbers.JComm_StringToInteger(mapSetEditValueArray.intItemCount);
																		if (intSetEditValueArrayCnt == 3) {
																			arrySetEditValue = mapSetEditValueArray.ArryOfValues;
																			//Return the input column name
																			strSetEditInputColName = arrySetEditValue[0];
																			//Return the value for the assigned column and row
																			strSetEditTempValue = (ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow, strSetEditInputColName)).CellValue;
																			//Check if the value is boolean and covert to boolean value
																			if (StringsAndNumbers.JComm_StringIsBoolean(strSetEditTempValue) == true) {
																				boolIsChecked = StringsAndNumbers.JComm_StringToBoolean(strSetEditTempValue);
																				//strSelectElementType
																				switch (strSelectElementType) {
																					case 'Checkbox':
																						//Call the standard Check box method
																						// Assuming SetVerifyCheckBoxWebElement is a global function
																						SetVerifyCheckBoxWebElement(strLocName, weElementLocated, strAssgSelElemName, boolIsChecked, true);
																						break;
																					case 'Graphic Checkbox':
																						//Call the Graphic Group Check box method
																						//CommonWeb.SetVerifyGraphicChkBoxGroupItems(mapSetVerChkBoxGroupValues) // Assuming mapSetVerChkBoxGroupValues is provided
																						strMethodDetails = "Graphic Checkbox not implemented in template.";
																						break;
																					case 'JavaScript Checkbox':
																						//Call the Graphic JavaScript Check box method
																						//CommonWeb.SetVerifyCheckBoxJavaScript (strLocName, strChkboxXpath, strChkBox, boolCkBoxIsChecked, boolVerChecked)
																						strMethodDetails = "JavaScript Checkbox not implemented in template.";
																						break;
																					default:
																						boolPassed = false;
																						strMethodDetails = "FAILED!!! Matching row was found, found cell by column name of: ' " + strSelectColumnName + "' found element by Name/Xpath'" + strSelectElementName + "' BUT, CHECKBOX TYPE '" + strSelectElementType + "' WAS NOT FOUND IN SWITCH!!!";
																						break;
																				}
																			}
																			else {
																				boolPassed = false;
																				strMethodDetails = "FAILED!!! Matching row was found, and object type: " + strSelectElementType + " BUT, input value: " + strSetEditTempValue + " IS NOT BOOLEAN!!!";
																			}
																		}
																		else {
																			boolPassed = false;
																			strMethodDetails = "FAILED!!! Matching row was found, and object type: " + strSelectElementType + " BUT, the SetEdit Value: " + strSetEditTempValues + " DOES NOT CONTAIN THREE VALUES!!! ";
																		}
																	}
																	else {
																		boolPassed = false;
																		strMethodDetails = "FAILED!!! Matching row was found, and object type: " + strSelectElementType + " BUT, the SetEdit Value: " + strSetEditTempValues + " DOES NOT CONTAIN ONLY 1 SET of VALUES!!! ";
																	}
																}
																//ListBoxes
																else if (lstListBoxes.includes(strSelectElementType) == true && strSelectItemValue != gblNA && strSelectItemValue != gblSkip) { // Use .includes for JS Array
																	boolActionProcessed = true;
																	switch (strSelectElementType) {
																		case 'listbox': // Original had 'Listbox', JS is case-sensitive
																			//TODO check for alternate values for strElemSingleValListFullPath, strElemSingleValListName.
																			if (strElemSingleValListFullPath != gblNull && strElemSingleValListName != gblNull) {
																				// Assuming objSelectVerifyStandardSingleValueList is a global function
																				objSelectVerifyStandardSingleValueList(strLocName, strElemSingleValListFullPath, strElemSingleValListFullPath, strSetEditTempValue, true, false, true);
																			}
																			break;
																		case 'Single Value Filter List Box':
																			if (mapSetVerfiyFilterEditBoxValues != null) {
																				//TODO Check the values passed and replace the search and select values with strSetEditTempValue
																				//CommonWeb.objSetVerifyEditSingleValueFilterListBox(mapSetVerfiyFilterEditBoxValues) // Assuming mapSetVerfiyFilterEditBoxValues is configured
																				strMethodDetails = "Single Value Filter List Box not implemented in template.";
																			}
																			break;
																		default:
																			boolPassed = false;
																			strMethodDetails = "FAILED!!! Matching row was found, found cell by column name of: ' " + strSelectColumnName + "' found element by Name/Xpath'" + strSelectElementName + "' BUT, LISTBOX TYPE '" + strSelectElementType + "' WAS NOT FOUND IN SWITCH!!!";
																			break;
																	}
																}
																if (boolActionProcessed == false) {
																	//Return the element link
																	//Highlight
																	if (TCExecParams.getBoolDoHighlight() == true) {
																		let mapHighlight = {}; // Mimic Groovy Map
																		mapHighlight = CWCore.objHighlightElementJS(weElementLocated, 'Element', strTblName + ' cell status element');
																	}
																	// Note: weElement.getAttribute('href') only works for anchor tags.
																	// TCObj.getTcDriver().navigate().to(strHRef) - Assuming navigate().to is a valid method.
																	weElementLocated.click(); // Always attempt click if no specific action taken.
																}
															}
														}
													}
													if (boolFoundElement == false) {
														boolPassed = false;
														strMethodDetails = "FAILED!!! Matching row was found, found cell by column name of: ' " + strSelectColumnName + "' HOWEVER, ELEMENT NAME '" + strSelectElementName + "' WAS NOT FOUND!!!";
													}
												}
												else {
													boolPassed = false;
													strMethodDetails = "FAILED!!! Matching row was found, Attempted to get the cell by column name of: ' " + strSelectColumnName + "' NO ELEMENT WAS RETURNED!!!";
												}
											}
											else {
												boolPassed = false;
												strMethodDetails = "FAILED!!! Matching row was found, HOWEVER the SPECIFIED SELECT COLUMN NAME OF: ' " + strSelectColumnName + "' IS INVALID!!!";
											}
											break; // Break from loopDisplayRow, found a match and processed.
										}
									}
									else {
										boolPassed = false;
										strMethodDetails = "FAILED!!! Result row: " + loopDisplayRow + " is displaying " + intBodyRowColCnt +
											" column(s), HOWEVER EXPECTED COLUMNS BASED ON THE HEADER COUNT IS: " + cntHdrCols + "!!!";
									}
									if (boolInputRowMatch == true) { // If _this_ input row matched _a_ displayed row
										strMethodDetails = strMethodDetails + "Results in row: " + loopDisplayRow + " matched the results assigned.";
										// A break should happen here to move to the next input row
										break; // Break from loopDisplayRow. Inner loop, done with input row.
									}
								}
								//Update the column width to fit
								let mapAutoFit = {}; // Mimic Groovy Map
								mapAutoFit = ExcelData.excelAutofitCols();
								if (StringsAndNumbers.JComm_StringToBoolean(mapAutoFit.boolPassed) == false) {
									strMethodDetails = strMethodDetails + mapAutoFit.strMethodDetails;
								}
								//Check if we matched all the assigned rows.
								if (intRowMatched != -1 && boolPassed == true) { // Based on zero based indexing, -1 means no match found
									strMethodDetails = strMethodDetails + "Successfully processed the assigned " + intInputRowsToProcess + " row. Returning the row match of: " +
										intRowMatched + ".";
								}
								else {
									boolPassed = false;
									strMethodDetails = strMethodDetails + "FAILED!!! Proccesed the displayed row count of: " + intBodyRowCnt + " row)s), suggest RUN VERIFY TO CHECK DATA!!!";
								}
							}
						}
					}
					else {
						boolPassed = false;
						strMethodDetails = "FAILED!!! The user specified to process " + intInputRowsToProcess +
							" row(s) WHICH IS GREATER THAN THE DISPLAYED " + intBodyRowCnt + " ROWS!!!";
					}
				}
				else {
					boolPassed = false;
					strMethodDetails = mapAddSheet.strMethodDetails;
				}
			}
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.intRowMatched = intRowMatched;
	return mapResults;
}

// These functions are assumed to be globally available or imported in the TS environment.
// Their implementations would also follow the "no interfaces, no explicit types" rule.
/*
function CommonWeb_SetVerifyEditBoxValue(locName, weElement, elemName, setValue, verifyValue, highlight, screenCapture) {
	// ... implementation without interfaces or explicit types
}

function CommonWeb_SetVerifyCheckBoxWebElement(locName, weElement, chkBoxName, checkValue, verifyChecked) {
	// ... implementation without interfaces or explicit types
}

function CommonWeb_objSelectVerifyStandardSingleValueList(locName, listXPath, listName, valueToSelect, verifySelected, highlight, screenCapture) {
	// ... implementation without interfaces or explicit types
}
*/

/**
* -------------------------------------  SetSelectMultiElemInRowsOfResultBody  -----------------------------------
* Set or Select Elements in more then one column of a row in the result body
* @param mapInputVariables The input variables that identify the header to include xpaths and properties and table results rows and columns
* @param NOTE: Each cell will be checked for the elements listed in the strResultsColObjNames
* @param		Rows in the input should only be present if changes to values or selection is required and can be in any order.
* @param		Provide sufficent data to find the unique row.
*
* @param INPUT SHEET Variables
* @param strInputSheetName				The sheet name assigned for the input
* @param intInputDataStartRow			The input data row number to start using for verification
* @param intInputDataEndRow			The input data row number to end verification. Enter 999 for all rows
*
* @param TABLE Variables
* @param strLocName						The pipe delimited location value
* @param strLocXpath						The parent Xpath
* @param strTblName						The name of the table
* @param strTblXPath						The table Xpath
* @param strHdrRowXPath					The Header row Xpath usually //thead//tr
* @param strHdrColXPath					The header column Xpath usually //th
* @param strPropertyHdrColName				The cell property that will be used to find the assigned column name. Usually 'abbr'
* @param strResultsOutShtName				The sheet name assigned for the output
* @param strResultsRowXPath				The Xpath for the rows within the results
* @param strResultsColXPath				The Xpath for the columns within the result rows
* @param arryElementNames					The array containing the list of element names that can be in the results cells.
* @param arryElemXPaths					The array containing the list of elemenT Xpaths that can be in the results cells.
* @param NOTE: the TWO Arrays must contain the same number of items
* @param strResultsCellStatElem			The name of the status element
* @param strHdrColMissingNames				For each column that does not have an identifiable name, add the name to the delimited value
* @param strMultElemColData				The delimited value with the name|count of elements where a cell contains more than one element. Currently only one column is supported
* @param strSelectElemColumns				The Select columns in order that they must be set as a value pair (strSelectColumnName|strSelectElementType|strSelectedItemValueSelected) seperated by a semicolon for each input column impacted.
* @param									i.e. New_Action|New_Action_Element;New_DocumentSearch|New_DocumentSearch_Element;New_PersonaShopping|New_PersonaShopping_Element
* @param strSelectColumnName 				The column containing the element to select/edit/set
* @param strSelectElementType 				The element type that will be in the column name
* @param strSelectedItemValueSelected 		The value to set/select/ or click as applicable.
*
* @param mapSetVerfiyFilterEditBoxValues	The map containing the values for the filter listbox child items see below:
* @param	strElemEditBoxName				The name of the editbox child
* @param 	strElemEditBoxXPath				The Xpath for the editbox used as a filter of items displayed in the list box
* @param	strElemListBoxName				The name of the listbox that opens after setting the value
* @param	strElemListBoxXPath				The Xpath for the listbox that opens after setting the value.
* @param 	strElemListBoxItemName			The name to associate with the item(s)
* @param	strElemListBoxItemXPath			The Xpath for the item(s) within the listbox
*
*
* @return mapResults 		The results showing Passed and method details. Includes the number of the row matched.
*
* @author pkanaris
* @author Created: 06/05/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetSelectMultiElemInRowsOfResultBody (mapInputVariables) { // mapInputVariables and return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	let gblSkip = GVars.GblSkip('Value');
	let gblIgnoreData = GVars.GblIgnoreData('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails;
	let boolPassed = true;
	//Return the method variables
	//Input Sheet variables
	let strInputSheetName = mapInputVariables.InputDataSheetName;
	let intInputDataStartRow = mapInputVariables.InputDataRowStart;
	let intInputDataEndRow = mapInputVariables.InputDataRowEnd;
	//Application variables
	let strLocName = mapInputVariables.LocName;
	let strLocXpath = mapInputVariables.LocXPath;
	let strTblName = mapInputVariables.TblName;
	let strTblXPath = mapInputVariables.strTblXPath;
	let strHdrRowXPath = mapInputVariables.strHdrRowXPath;
	let strHdrColXPath = mapInputVariables.strHdrColXPath;
	let strPropertyHdrColName = mapInputVariables.strPropertyHdrColName;
	let strResultsOutShtName = mapInputVariables.strResultsOutShtName;
	let strResultsRowXPath = mapInputVariables.strResultsRowXPath;
	let strResultsColXPath = mapInputVariables.strResultsColXPath;
	//Cells can possibly have only one element
	let arryElementNames = mapInputVariables.arrayResultsColObjNames;
	let arryElemXPaths = mapInputVariables.arrayResultsColObjXPaths;
	let strResultsCellStatElem = mapInputVariables.strResultsCellStatElem;
	//Alternate properties
	let strTblSepXpath = gblNull;
	if (mapInputVariables.hasOwnProperty('strTblSepXpath')){ // Use hasOwnProperty
		strTblSepXpath = mapInputVariables.strTblSepXpath;
	}
	let strHdrColMissingNames = gblNull;
	let strMultElemColData = gblNull;
	if (mapInputVariables.hasOwnProperty('strHdrColMissingNames')){
		strHdrColMissingNames = mapInputVariables.strHdrColMissingNames;
	}
	if (mapInputVariables.hasOwnProperty('strMultElemColData')){
		strMultElemColData = mapInputVariables.strMultElemColData;
	}
	//UL alternate values see: Sourcing|Templates and Libraries|Event Libraries for example
	let strULElemName = gblNull;
	let strULNoValElemName = gblNull;
	let arryULElementNames;
	let arryULElemXPaths;
	let arryMultiColSelItems;
	let intMultiColSelItemsCnt = 0;
	let intMultiColSelValueCnt = 0;
	if (mapInputVariables.hasOwnProperty('strResultsCellULElemName')){
		strULElemName = mapInputVariables.strResultsCellULElemName;
	}
	if (mapInputVariables.hasOwnProperty('strResultsCellULNoValueElemName')){
		strULNoValElemName = mapInputVariables.strResultsCellULNoValueElemName;
	}
	if (mapInputVariables.hasOwnProperty('arrayResultsColULObjNames')){
		arryULElementNames = mapInputVariables.arrayResultsColULObjNames;
	}
	if (mapInputVariables.hasOwnProperty('arrayResultsColULObjXPaths')){
		arryULElemXPaths = mapInputVariables.arrayResultsColULObjXPaths;
	}
	//Check for columns with more than one element
	let strMultElemColName = gblNull;
	let intMultElemColCount = 1; //always at least 1
	if (strMultElemColData != gblNull) {
		//Split the value into an array
		//ADD a split for the value to return the date/time format
		let mapArray = {}; // Mimic Groovy Map
		mapArray = StringsAndNumbers.JComm_StringToArray(strMultElemColData, gblDelimiter);
		if (StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount) == 2) {
			strMultElemColName = mapArray.ArryOfValues[0];
			intMultElemColCount = StringsAndNumbers.JComm_StringToInteger(mapArray.ArryOfValues[1]);
		}
	}
	//Add alternate for table where header and data is in one cell and a sub category row is present
	let intTableHeaderRowCnt = -1;
	let strSectionRowInfo = gblNull;
	let arrySectInfo;
	let intSectInfoArrySize;
	if (mapInputVariables.hasOwnProperty('intTableHeaderRowCnt')) {
		intTableHeaderRowCnt = mapInputVariables.intTableHeaderRowCnt;
	}
	let intLoopHdrCount; //Holds the number of rows in the header to be added to the loopResultStart
	if (intTableHeaderRowCnt > 0) {
		intLoopHdrCount = intTableHeaderRowCnt;
	} else {
		intLoopHdrCount = 0; // Initialize if not set
	}

	if (mapInputVariables.hasOwnProperty('strSectionRowInfo')) {
		strSectionRowInfo = mapInputVariables.strSectionRowInfo;
		let mapSplitSubSectInfoString = StringsAndNumbers.JComm_StringToArray(strSectionRowInfo, gblDelimiter);
		let intCntValues = StringsAndNumbers.JComm_StringToInteger(mapSplitSubSectInfoString.intItemCount);
		if (intCntValues == 4) { // Groovy original comment says 3, but expects 4 from its example usage.
			arrySectInfo = mapSplitSubSectInfoString.ArryOfValues;
			intSectInfoArrySize = arrySectInfo.length; // Use .length for JS Array
		}
		else {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The Results can contain a Sub Section row, BUT, THE info provided '" + strSectionRowInfo + "', DOES NOT CONTAIN 3 ITEMS!!!";
		}

	}
	//Return the singlevalue filter listbox element map
	let mapSetVerfiyFilterEditBoxValues = {}; // Mimic Groovy Map
	let boolDoSingleFilterLstBox = false;
	if (mapInputVariables.hasOwnProperty('mapSetVerfiyFilterEditBoxValues')) {
		boolDoSingleFilterLstBox = true;
		mapSetVerfiyFilterEditBoxValues = mapInputVariables.mapSetVerfiyFilterEditBoxValues;
	}
	//Add variables for working with multi-setselect columns
	let strMultSetSelColInfo;
	let arryMultSetSelColInfo;
	//let intMultSetSelColInfo; // Declared but unused directly
	if (mapInputVariables.hasOwnProperty('strSetSelectColumnInfo')) {
		strMultSetSelColInfo = mapInputVariables.strSetSelectColumnInfo;
		//TODO if not null split values into array
		if (strMultSetSelColInfo != gblNull) {
			let mapArray = {}; // Mimic Groovy Map
			mapArray = StringsAndNumbers.JComm_StringToArray(strMultSetSelColInfo, ";"); //Use item splitter
			intMultiColSelItemsCnt = StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount);
			if ( intMultiColSelItemsCnt > 0) {
				arryMultiColSelItems = mapArray.ArryOfValues;
				intMultiColSelValueCnt = intMultiColSelItemsCnt * 3; //3 values per item colName|inputColValue|inputColElem
			}
		}
	}
	let intResultsRowCnt;
	let intRowMatched;
	let boolProcCellElements;
	let arryOutputColNames;
	let strDelLstColNames;
	let intAddlColumnCnt = 0;
	let strSelectItemValue = gblNull; // Note: strSelectItemValue is unused after this initial assignment
	//Check for alternate name values
	let intAltColNameArrayCnt = -1;
	let intAltColNameUsed = -1;
	let arryAltColName;
	//let strTempAltColName; // Declared but unused
	if (strHdrColMissingNames != gblNull) {
		//Split the value into an array
		//ADD a split for the value to return the date/time format
		let mapArray = {}; // Mimic Groovy Map
		mapArray = StringsAndNumbers.JComm_StringToArray(strHdrColMissingNames, gblDelimiter);
		intAltColNameArrayCnt = StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount);
		arryAltColName = mapArray.ArryOfValues;
	}
	//Define the list of possible elements that can be in the cells
	let lstCheckboxes = ['Checkbox','Graphic Checkbox','JavaScript Checkbox']; // Mimic Groovy def list
	let lstListBoxes = ['listbox', 'Single Value Filter List Box']; // Mimic Groovy def list
	let lstAddlColElemTypes = ['Editbox','Checkbox','Graphic Checkbox','JavaScript Checkbox','listbox', 'Single Value Filter List Box']; // Mimic Groovy def list
	//Verify we have the table.
	let weTable = CWCore.returnWebElement(strLocXpath + strTblXPath);
	if (weTable == null) {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The location '" + strLocName + "' '" + strTblName + "' DID NOT RETURN an ELEMENT!!!";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weTable, 'Element', strTblName);
		}
	}
	//Check if the header row is present
	let weHeaderRow = CWCore.returnWebElement(strLocXpath + strTblXPath + strHdrRowXPath);
	if (weHeaderRow == null) {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The location '" + strLocName + "' '" + strTblName + "' DID NOT RETURN A HEADER ROW ELEMENT!!!";
	}
	else {
		//Return all the column names
		let mapGetHdrChildElements = {}; // Mimic Groovy Map
		let strColNames;
		mapGetHdrChildElements = CWCore.returnChildElements(weHeaderRow, strHdrColXPath);
		let lstHdrCols = mapGetHdrChildElements.lstChildObjects;
		//let arryColumnNames; // Already declared above
		let cntHdrCols = mapGetHdrChildElements.cntChildObjs;
		if (cntHdrCols > 0) {
			let weHdrCol;
			let strTempValue;
			let boolAttribPresent;
			//Return the column names and replace any spaces with '_'
			for (let loopHdrCols = 0; loopHdrCols < cntHdrCols; loopHdrCols++) {
				strTempValue = gblNull;
				weHdrCol = lstHdrCols[loopHdrCols];
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weHdrCol, 'Element', strTblName + ' header row column');
				}
				boolAttribPresent = CWCore.isAttribtuePresent(weHdrCol, strPropertyHdrColName);
				if (boolAttribPresent == true) {
					strTempValue = StringsAndNumbers.JComm_HandleNoData(weHdrCol.getAttribute(strPropertyHdrColName));
				}
				//Check if we did not find a column name
				else if (intAltColNameArrayCnt >= 1 && intAltColNameArrayCnt > intAltColNameUsed) {
					intAltColNameUsed++;
					strTempValue = StringsAndNumbers.JComm_HandleNoData(arryAltColName[intAltColNameUsed]);
				}
				else {
					strTempValue = gblNull;
				}
				if (strTempValue != gblNull) {
					if (loopHdrCols == 0) {
						strColNames = StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_");
						strDelLstColNames = strTempValue;
					}
					else {
						strColNames = strColNames + gblDelimiter + StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_");
						strDelLstColNames = strDelLstColNames + "|" + strTempValue;
					}
					strColNames = strColNames + gblDelimiter + StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_") + "_Element"; //Add the column for the element

				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The header column '" + loopHdrCols + "' does not CONTAIN THE PROPERTY REQUIRED of:" + strPropertyHdrColName + "!!!";
					break;
				}
			}
			if (boolPassed == true) {
				//Add column names to the array
				let mapStringToArry = {}; // Mimic Groovy Map
				mapStringToArry = StringsAndNumbers.JComm_StringToArray(strDelLstColNames, gblDelimiter);
				arryColumnNames = mapStringToArry.ArryOfValues;
			}
			//Load the inputsheet
			let intInputRowCnt;
			let intInputColCnt;
			let shInput; // Mimic XSSFSheet
			// let strHdrInputShtName = mapInputVariables.strHdrInputShtName; // Not used in this function, despite being in Javadoc
			let mapOpenInputSheet = {}; // Mimic Groovy Map
			mapOpenInputSheet = ExcelData.excelGetSheetByName(TCObj.getObjWorkbook(), strInputSheetName);
			if (StringsAndNumbers.JComm_StringToBoolean(mapOpenInputSheet.boolPassed) == true) {
				shInput = mapOpenInputSheet.objWbSheet;
				//Return the row and column Count
				let mapSheetRowColCnt = {}; // Mimic Groovy Map
				mapSheetRowColCnt = ExcelData.excelGetRowAndColCount(shInput);
				if (StringsAndNumbers.JComm_StringToBoolean(mapSheetRowColCnt.boolPassed) == true) {
					intInputRowCnt = mapSheetRowColCnt.RowCount;
					intInputColCnt = mapSheetRowColCnt.ColCount;
				}
				else {
					boolPassed = false;
					strMethodDetails = mapSheetRowColCnt.strMethodDetails;
				}
				// Original Groovy re-gets header columns here, but `cntHdrCols` and `lstHdrCols` are already populated.
				// This might be a redundancy where the original thought might have been to use `intHdrColCnt` from Excel.
				if (boolPassed == true) {
					let mapHdrChildren = CWCore.returnChildElements(weHeaderRow, strHdrColXPath);
					// intHdrColCnt = mapHdrChildren.cntChildObjs; // Do not reassign, as `cntHdrCols` is from actual display
					// lstHdrCols = mapHdrChildren.lstChildObjects; // Do not reassign
				}
			}
			else {
				boolPassed = false;
				strMethodDetails = mapOpenInputSheet.strMethodDetails;
			}
			//Check if the number of columns displayed matches the number in the input sheet if not call output only and fail the step
			if (boolPassed == true) {
				let intInputTempColCnt = intInputColCnt - intAddlColumnCnt - (intMultiColSelItemsCnt * 2); //Select Columns are a value and the element thus *2
				if (cntHdrCols != intInputTempColCnt / 2) { //Each displayed column consist of the value and element column
					boolPassed = false;
					strMethodDetails = "FAILED!!! The table header is displaying: " + cntHdrCols + " BUT, EXPECTED: " + intInputTempColCnt / 2 + " COLUMNS!!! OUTPUTTING HEADER ONLY see details: " + gblLineFeed + strColNames;
				}
				else {
					//Return the input sheet column names and check if the match the output column names
					let strInputColValues;
					let mapGetInputColNames = {}; // Mimic Groovy Map
					mapGetInputColNames = ExcelData.excelGetHdrColNames(shInput);
					let boolInputColNames = StringsAndNumbers.JComm_StringToBoolean(mapGetInputColNames.boolPassed);
					if (boolInputColNames == true) {
						let strTempInputColNames;
						strInputColValues = mapGetInputColNames.ColValues;
						let intTotalAddlColumns = intAddlColumnCnt + (intMultiColSelItemsCnt * 2);
						if (intTotalAddlColumns > 0) {
							//Trim off the end values by count
							let strTmpValue = strInputColValues;
							let strTempRight;
							for (let loopTrim = 0; loopTrim < intTotalAddlColumns; loopTrim++) {
								if (loopTrim == 0) {
									strTempRight = StringsAndNumbers.JComm_GetRightTextInString(strTmpValue, gblDelimiter);
								}
								else {
									strTempRight = StringsAndNumbers.JComm_GetRightTextInString(strTmpValue, gblDelimiter) + gblDelimiter + strTempRight;
								}
								strTmpValue = StringsAndNumbers.JComm_GetLeftTextInString(strTmpValue, gblDelimiter + StringsAndNumbers.JComm_GetRightTextInString(strTmpValue, gblDelimiter));
							}
							strTempInputColNames = StringsAndNumbers.JComm_GetLeftTextInString(strInputColValues, gblDelimiter + strTempRight);
						}
						else {
							strTempInputColNames = strInputColValues;
						}
						//Check if they match the output column names
						if (strTempInputColNames != strColNames) {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The input column names '" + strTempInputColNames + "' DOES NOT MATCH the OUTPUT Columns '" + strColNames + "'!!!";
						}
						else {
							strMethodDetails = "The column names matched and the input file contains: " + strTempInputColNames + " rows, and " + intInputColCnt + " columns. Does Not Include Edit or Set input column(s).";
							let mapSplitElementString = StringsAndNumbers.JComm_StringToArray(strColNames, gblDelimiter);
							//let intCntElemNames= StringsAndNumbers.JComm_StringToInteger(mapSplitElementString.intItemCount); // Declared but unused
							arryOutputColNames = mapSplitElementString.ArryOfValues;
						}
					}
					else {
						boolPassed = false;
						strMethodDetails = mapGetInputColNames.strMethodDetails;
					}
				}
			}
			if (boolPassed == true) {
				//Check if the cell elements will be verified
				if (arryElementNames == null || arryElemXPaths == null) {
					boolProcCellElements = false;
				}
				else if (arryElementNames.length != arryElemXPaths.length) { // Use .length for JS Array
					boolPassed = false;
					strMethodDetails = "FAILED!!! THE arryElementNames and arryElemXPaths SIZE DO NOT MATCH!!!";
				}
				else {
					boolProcCellElements = true;
				}
			}
			//Return the data and verify the results
			if (boolPassed == true) {
				//Set the start and end rows
				let loopInputRowStart;
				let loopInputRowEnd;
				let loopResultStart;
				let loopResultEnd;
				let intInputCol;
				let intDataElem;
				//let intOutputRow; // Declared but unused
				let boolInputRowMatch = true; //reset after each loop processed for input data.
				let mapSetLoopStartEnd = {}; // Mimic Groovy Map
				mapSetLoopStartEnd = ExcelData.excelReturnLoopStartEndRow(intInputDataStartRow, intInputDataEndRow, intInputRowCnt);
				loopInputRowStart = mapSetLoopStartEnd.intLoopStart;
				loopInputRowEnd = mapSetLoopStartEnd.intLoopEnd;
				let intInputRowsToProcess = loopInputRowEnd - loopInputRowStart + 1;
				let strTempInputValue;
				let strTempInputElem;
				let strTempDispValue;
				let strTempDispElem;
				let strColor;
				let strBckgColor;
				let strTempRowDetails;
				let intRowsMatched = 0; // Initialize for cumulative counts
				let intOutputExcelRow;
				let intBodyRowColCnt;
				let intOutputColumn;
				let intCellElemCnt;
				let intOutputValColIndex;
				let intOutputElemColIndex;
				//Status Counters
				// These are for cumulative counts across all loops if successful, or local to current row/col
				// Based on usage, these seem cumulative.
				let intCellSkipped = 0;
				let intCellIgnored = 0;
				let intCellPassed = 0;
				let intCellFailed = 0;

				let strTempCellValue;
				let strTempCellElements;
				let weRow;
				let weCell;
				let mapGetRowCols;
				let lstBodyRowCols;
				let mapGetInputValue;
				let mapSetCellValue;
				let boolGetInpValue;
				//Check if the tbody is a separate XPath
				if (strTblSepXpath != gblNull) {
					//Return the body as a new table element
					weTable = CWCore.returnWebElement(strLocXpath + strTblXPath + strTblSepXpath);
					//Highlight
					if (TCExecParams.getBoolDoHighlight() == true) {
						let mapHighlight = {}; // Mimic Groovy Map
						mapHighlight = CWCore.objHighlightElementJS(weTable, 'Table Body', strTblName);
					}
					intLoopHdrCount = 0; // Reset header count for body, as it's separate
				}
				let mapGetRowElements = CWCore.returnChildElements(weTable, strResultsRowXPath);
				let lstBodyRows = mapGetRowElements.lstChildObjects;
				let intBodyRowCnt = mapGetRowElements.cntChildObjs;
				intRowMatched = -1; //Set to show no rows are matching
				intOutputExcelRow = 1; //Increment the row out to set for first output
				//Check the results against the rows to process
				if (intInputRowsToProcess <= intBodyRowCnt) {
					let intSubSecColCnt = StringsAndNumbers.JComm_StringToInteger(arrySectInfo[0]); // Sub-section logic from verify method
					for (let loopInputRow = loopInputRowStart; loopInputRow <= loopInputRowEnd; loopInputRow++) {
						if (boolPassed == false) {
							break; // Exit the loop on failure
						}
						boolInputRowMatch = true;
						loopResultStart = 0; //Zero based index for displayed rows
						loopResultEnd = intBodyRowCnt;
						for (let loopDisplayRow = loopResultStart; loopDisplayRow < loopResultEnd; loopDisplayRow++) {
							intInputCol = 0;
							intDataElem = 1; //Set to verify data first
							weRow = lstBodyRows[loopDisplayRow]; // Access array by index
							//Move to the row to show the row we are processing in case of errors.
							let actions = new Actions(TCObj.tcDriver);
							actions.moveToElement(weRow).perform();
							//Highlight
							if (TCExecParams.getBoolDoHighlight() == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weRow, 'Element', strTblName + ' body row');
							}
							//Return the column elements and count
							mapGetRowCols = CWCore.returnChildElements(weRow, strResultsColXPath);
							lstBodyRowCols = mapGetRowCols.lstChildObjects;
							intBodyRowColCnt = mapGetRowCols.cntChildObjs;
							//Check for a spacer row and make sure not SubSection row
							if (intBodyRowColCnt > 0 && intBodyRowColCnt != intSubSecColCnt) { //Skips processing the row.
								if (intBodyRowColCnt == cntHdrCols) { // Corrected `intHdrColCnt` to `cntHdrCols` from outer scope.
									//Create the temp counters
									let intCellSkippedTemp = 0; // Local counters for this iteration
									let intCellIgnoredTemp = 0;
									let intCellPassedTemp = 0;
									let intCellFailedTemp = 0;
									let boolComparePass = true;
									for (let loopDisplayCol = 0; loopDisplayCol < cntHdrCols; loopDisplayCol++) {
										strTempInputValue = gblNull;
										strTempInputElem = gblNull;
										//Return the expected values
										mapGetInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow, arryOutputColNames[intInputCol]);
										boolGetInpValue = StringsAndNumbers.JComm_StringToBoolean(mapGetInputValue.boolPassed);
										if (boolGetInpValue == true) {
											strTempInputValue = StringsAndNumbers.JComm_HandleNoData(mapGetInputValue.CellValue);
											intOutputValColIndex = mapGetInputValue.intColIndex;
											//Check if the value is a dynamic date and/or time
											if (strTempInputValue.indexOf('D[') > 0 || strTempInputValue.indexOf('T[') > 0) {
												let myTemp = DateTime.datetimeReturnDynamicFormatedDateTime(strTempInputValue);
												strTempInputValue = myTemp;
											}
										}
										intInputCol++; //Increment the input column number
										mapGetInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow, arryOutputColNames[intInputCol]);
										boolGetInpValue = StringsAndNumbers.JComm_StringToBoolean(mapGetInputValue.boolPassed);
										if (boolGetInpValue == true) {
											strTempInputElem = StringsAndNumbers.JComm_HandleNoData(mapGetInputValue.CellValue);
											intOutputElemColIndex = mapGetInputValue.intColIndex;
										}
										intInputCol++; //Increment the input column number
										weCell = lstBodyRowCols[loopDisplayCol];
										//Highlight
										if (TCExecParams.getBoolDoHighlight() == true) {
											let mapHighlight = {}; // Mimic Groovy Map
											mapHighlight = CWCore.objHighlightElementJS(weCell, 'Element', strTblName + ' body cell');
										}
										//Return the cell value
										strTempDispValue = StringsAndNumbers.JComm_HandleNoData(weCell.getText());
										//Return the cell element
										strTempDispElem = gblNull;
										let weCellElement;
										if (boolProcCellElements == true) { // boolProcCellElements indicates whether to process elements
											//Set the wait to 1 second to speed up the processing of elements in the cells
											TCObj.tcDriver.manage().timeouts().implicitlyWait(100, TimeUnit.MILLISECONDS);
											//Process the array for elements
											let intMaxElem = 1;
											if (strMultElemColName == arryColumnNames[loopDisplayCol]) {
												intMaxElem = intMultElemColCount;
											}
											strTempCellElements = gblNull; // Reset for this cell
											//Process the array for elements
											let cntCellElements = 0;
											while (cntCellElements < intMaxElem && boolPassed) { // Added boolPassed check
												for (let loopElements = 0; loopElements < arryElementNames.length && boolPassed; loopElements++) { // Added boolPassed check
													//TODO add in classes to get color such as status elements
													intCellElemCnt = CWCore.returnWebElemChildElementCount(weCell, arryElemXPaths[loopElements]);
													if (intCellElemCnt == 1) {
														cntCellElements++; //Increment the number of elements in the cell
														weCellElement = CWCore.returnChildElement(weCell, arryElemXPaths[loopElements]);
														//Highlight
														if (TCExecParams.getBoolDoHighlight() == true) {
															let mapHighlight = {}; // Mimic Groovy Map
															mapHighlight = CWCore.objHighlightElementJS(weCellElement, 'Element', strTblName + ' cell status element');
														}
														if (cntCellElements == 1) { // Only first element detected
															//Check if the element is a UL
															if (arryElementNames[loopElements] == strULElemName) { //See Sourcing|Templates and Libraries|Event Libraries for UI example
																//TODO we need to pass in the ul element names and xpaths
																//let intItemCnt = 0; // Declared but unused
																//let strItemValue; // Declared but unused
																//let strItemElem; // Declared but unused
																let weULItem;
																//Process the UL to return the items
																let mapULItems = CWCore.returnChildElements(weCellElement, arryULElemXPaths[0]);
																let lstULItems = mapULItems.lstChildObjects;
																let intULItemCnt = mapULItems.cntChildObjs;
																if (intULItemCnt == 0) {
																	//Error
																	boolPassed = false;
																	strMethodDetails = "FAILED!!! The Cell User List contain 'ZERO' ITEM(S)!!!!";
																}
																else {
																	let weULItemChild;
																	for (let loopULItemsInner = 0; loopULItemsInner < intULItemCnt && boolPassed; loopULItemsInner++) { // Added boolPassed check
																		//Return the listitem
																		weULItem = lstULItems[loopULItemsInner];
																		//Add a ';' if loop is for greater than the first element to Cell and Element values
																		if (loopULItemsInner > 0) {
																			strTempCellValue = strTempCellValue + ';';
																			strTempCellElements = strTempCellElements + ';';
																		}
																		//Highlight
																		if (TCExecParams.getBoolDoHighlight() == true) {
																			let mapHighlight = {}; // Mimic Groovy Map
																			mapHighlight = CWCore.objHighlightElementJS(weULItem, 'Element', 'UserList item_' + loopULItemsInner);
																		}
																		//Check if the child items are present
																		//Loop through the UL items elements XPaths to check for a single item
																		//Start with the second element since 0 is the listitem
																		let intElemFoundCnt = 0;
																		for (let loopULItemElem = 1; loopULItemElem < arryULElementNames.length && boolPassed; loopULItemElem++) { // Added boolPassed check
																			weULItemChild = CWCore.returnChildElement(weULItem, arryULElemXPaths[loopULItemElem]);
																			if (weULItemChild != null) {
																				intElemFoundCnt++;
																				//Highlight
																				if (TCExecParams.getBoolDoHighlight() == true) {
																					let mapHighlight = {}; // Mimic Groovy Map
																					mapHighlight = CWCore.objHighlightElementJS(weULItemChild, 'Element', 'UserList item_child' + loopULItemsInner);
																				}
																				//Return the text and Add the element name and text to the variables
																				if (intElemFoundCnt == 1 && loopULItemsInner == 0) {
																					strTempCellValue = weULItemChild.getText();
																					strTempCellElements = arryULElementNames[loopULItemElem];
																				}
																				else if (intElemFoundCnt == 1 && loopULItemsInner > 0) {
																					strTempCellValue = strTempCellValue + weULItemChild.getText();
																					strTempCellElements = strTempCellElements + arryULElementNames[loopULItemElem];
																				}
																				else {
																					strTempCellValue = strTempCellValue + "|" + weULItemChild.getText();
																					strTempCellElements = strTempCellElements + "|" + arryULElementNames[loopULItemElem];
																				}
																			}
																		}
																	}
																}
																strTempDispElem = strTempCellElements;
																strTempDispValue = strTempCellValue;
															}
															else if (arryElementNames[loopElements] == 'image') { // Added to match previous
																if (CWCore.isAttribtuePresent(weCellElement, strImageAttribute) == true) {
																	strTempDispValue = StringsAndNumbers.JComm_HandleNoData(weCellElement.getAttribute(strImageAttribute));
																}
																else {
																	strTempDispValue = gblNull;
																}
																strTempDispElem = arryElementNames[loopElements];
															}
															else {
																strTempDispElem = arryElementNames[loopElements];
																if (arryElementNames[loopElements] == 'checkbox') { // Check if checkbox/graphic checkbox
																	strTempDispValue = CWCore.objGetCheckboxChecked(weCellElement).toString();
																}
															}
															// For SVG
															if (StringsAndNumbers.JComm_TextLocationInString(arryElemXPaths[loopElements], 'svg') >= 0) {
																let strSVGElem = StringsAndNumbers.JComm_HandleNoData(weCellElement.getText());
																if (strElemContainsSVGNameXPath == gblNull) {
																	// Simplified translation of the original complex JComm_TextLocationInString logic
																	if (strTempDispValue && strTempDispValue.startsWith(strSVGElem)) {
																		strTempDispValue = strTempDispValue.substring(strSVGElem.length);
																	} else if (strTempDispValue && strTempDispValue.endsWith(strSVGElem)) {
																		strTempDispValue = strTempDispValue.substring(0, strTempDispValue.length - strSVGElem.length);
																	}
																}
																else {
																	let weSVGNameElem = weCellElement.findElement(By.xpath(strElemContainsSVGNameXPath));
																	if (weSVGNameElem != null) {
																		let boolAttribPresentInner = CWCore.isAttribtuePresent(weSVGNameElem, strElemContainsSVGNameProperty);
																		if (boolAttribPresentInner == true) {
																			strTempDispValue = StringsAndNumbers.JComm_HandleNoData(weSVGNameElem.getAttribute(strElemContainsSVGNameProperty));
																		}
																		else {
																			strTempDispValue = 'NO IMAGE NAME!!!';
																		}
																	} else {
																		strTempDispValue = 'SVG NAME ELEMENT NOT FOUND!!!';
																	}
																}
															}
														}
														else { // This means the element was NOT UNIQUE (cntCellElements > 1 for this element) or was not one of the first elements.
															strTempDispElem = strTempDispElem + gblDelimiter + arryElementNames[loopElements];
														}
														//TODO add in classes to get color such as status elements
														if (arryElementNames[loopElements] == strResultsCellStatElem) {
															//Return the element
															let weStatus = CWCore.returnChildElement(weCell, arryElemXPaths[loopElements]);
															if (weStatus == null) {
																boolPassed = false;
																strMethodDetails = "FAILED!!! Result row: " + loopDisplayRow + " cell: " + (loopDisplayCol + 1) + " DID NOT CONTAIN A STATUS ELEMENT AS EXEPCTED!!!";
															}
															else {
																//Highlight
																if (TCExecParams.getBoolDoHighlight() == true) {
																	let mapHighlight = {}; // Mimic Groovy Map
																	mapHighlight = CWCore.objHighlightElementJS(weStatus, 'Element', strTblName + ' cell status element');
																}
																strColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("color"));
																if (StringsAndNumbers.JComm_TextLocationInString(arryElemXPaths[loopElements], "svg") >= 0) {
																	strBckgColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("fill"));
																}
																else {
																	strBckgColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("background-color"));
																}
																//Check if the color is the same
																//let strOrgColor = strColor; // Declared but unused
																let strTempColor = StringsAndNumbers.JComm_GetRightTextInString(strColor, "(");
																strTempColor = StringsAndNumbers.JComm_GetLeftTextInString(strTempColor, ")");
																let strTempBckColor = StringsAndNumbers.JComm_GetRightTextInString(strBckgColor, "(");
																strTempBckColor = StringsAndNumbers.JComm_GetLeftTextInString(strTempBckColor, ")");
																//let intLocInColorToBck = StringsAndNumbers.JComm_TextLocationInString(strTempColor, strTempBckColor); // Declared but unused
																//let intLocInBckToColor = StringsAndNumbers.JComm_TextLocationInString(strTempBckColor, strTempColor); // Declared but unused
																let strTempStatusColor;
																//Return the property and add to the strTempCellElements
																if (StringsAndNumbers.JComm_TextLocationInString(strTempColor, strTempBckColor) >= 0 || StringsAndNumbers.JComm_TextLocationInString(strTempBckColor, strTempColor) >= 0) {
																	strColor = "RGB(255, 255, 255)"; //Change the color for the text to white. Will work for all but white backgrounds
																	strTempStatusColor = "Colors matched replaced original color:" + strOrgColor + " with Color:" + strColor + "Background Color:" + strBckgColor;
																}
																else {
																	strTempStatusColor = "Color:" + strColor + "Background Color:" + strBckgColor;
																}
																strTempDispElem = strTempDispElem + gblDelimiter + strTempStatusColor;
															}
														}
													}
													if (cntCellElements == 0) {
														break;
													}
												}
												//Return the wait to maxwaittime
												TCObj.tcDriver.manage().timeouts().implicitlyWait(TCObj.getIntTempMaxWaitTime(), TimeUnit.SECONDS);
											}
											//Compare each value
											let strCompareDispValue; // Re-declared for this smaller scope
											let strCompareInputValue; // Re-declared for this smaller scope
											let strOutputCellTheme; // Re-declared for this smaller scope
											let intOutputColForCell; // Re-declared for this smaller scope
											for (let loopCompareInner = 1; loopCompareInner <= 2 && boolPassed; loopCompareInner++) { // Added boolPassed check
												if (loopCompareInner == 1) {
													strCompareDispValue = strTempDispValue;
													strCompareInputValue = strTempInputValue;
													intOutputColForCell = intInputCol - 2;
												}
												else {
													strCompareDispValue = strTempDispElem;
													strCompareInputValue = strTempInputElem;
													intOutputColForCell = intInputCol - 1;
												}
												//Check if they match
												if (strCompareInputValue == gblSkip || strCompareInputValue == gblIgnoreData) {
													strOutputCellTheme = 'TestRunDataSkipIgnore';
													if (strCompareInputValue == gblSkip) {
														intCellSkippedTemp++;
													}
													else {
														intCellIgnoredTemp++;
													}
												}
												else if (strCompareInputValue != strCompareDispValue) {
													boolComparePass = false;
													boolInputRowMatch = false;
													break;
												}
												else { // It matches
													boolComparePass = true;
												}
											}
											if (boolComparePass == false){
												if (TCExecParams.getBoolDoDebug()== true) {
													Tester.Message("Compare failed and matchinputRowNumber == false so break for columns");
												}
												//Break so we move to the next row in the display
												break;
											}
										}
										let boolRowSet = false; // Initialize for this path
										//Set the selection if the verify for the row passed
										if (boolComparePass == true) {
											//Process the displayed row using the assigned values for the columns
											let strTempSelValue;
											let strTempSelColName;
											let strTempSelColValue;
											let strTempSelColElem;
											let strTempSetSelValue;
											let strTempSetSelElem;
											let strTempSetSelXpath;
											let strTempElemDetails;
											let arrySelColValues;
											let weSetSelCell;
											let weSetSelElem;
											let intSetSelColNumber;
											let intSetSelElemIndex;
											//let boolSetSel = false; // Declared but unused
											boolRowSet = true;
											strTempRowDetails = "Displayed Row: " + loopDisplayRow + " ";
											//Input row loopInputRow and display row loopDisplayRow are required along with intMultiColSelItemsCnt and array of elements
											for (let intLoopSelValues = 0; intLoopSelValues < intMultiColSelItemsCnt && boolPassed; intLoopSelValues++) { // Added boolPassed check
												boolRowSet = false; // Reset boolRowSet for each multi-column element
												strTempSelValue = arryMultiColSelItems[intLoopSelValues];
												let mapSplitElementString = StringsAndNumbers.JComm_StringToArray(strTempSelValue, gblDelimiter);
												let intCntSelValueItems = StringsAndNumbers.JComm_StringToInteger(mapSplitElementString.intItemCount);
												if (intCntSelValueItems == 3) { //Must always be 3 columnName, columnValue, columnElement
													//Return the values
													arrySelColValues = mapSplitElementString.ArryOfValues;
													strTempSelColName = arrySelColValues[0]; // Select Column Name
													strTempSelColValue = arrySelColValues[1]; // Input Sheet Column Name for Value
													strTempSelColElem = arrySelColValues[2]; // Input Sheet Column Name for Element Type
													//Return the value and element from the input sheet.
													let mapGetSetSelValue = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow, strTempSelColValue);
													strTempSetSelValue = mapGetSetSelValue.CellValue;
													let mapGetSetSelElem = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow, strTempSelColElem);
													strTempSetSelElem = mapGetSetSelElem.CellValue;
													if ((strTempSetSelValue != gblNull && strTempSetSelValue != gblNA) && (strTempSetSelElem != gblNull && strTempSetSelElem != gblNA)) {
														intSetSelElemIndex = StringsAndNumbers.JComm_ReturnStringIndexFromArray(arryElementNames, strTempSetSelElem);
														//TODO add error handling at each level to avoid crashes
														if (intSetSelElemIndex != -1) {
															strTempSetSelXpath = arryElemXPaths[intSetSelElemIndex];
															//Get the cell by display row and column name
															intSetSelColNumber = StringsAndNumbers.JComm_ReturnStringIndexFromArray(arryColumnNames, strTempSelColName);
															weSetSelCell = lstBodyRowCols[intSetSelColNumber];
															if (weSetSelCell != null) {
																//Highlight
																if (TCExecParams.getBoolDoHighlight() == true && weSetSelCell != null) { // weSetSelCell null check is redundant here
																	let mapHighlight = {}; // Mimic Groovy Map
																	mapHighlight = CWCore.objHighlightElementJS(weSetSelCell, 'Display Cell', strTblName);
																}
																//Return the web element
																//Map mapGetCellElement = CWCore.returnChildElements(null, ""); // Commented out in original
																weSetSelElem = CWCore.returnChildElement(weSetSelCell, strTempSetSelXpath);
																//Highlight
																if (TCExecParams.getBoolDoHighlight() == true && weSetSelElem != null) {
																	let mapHighlight = {}; // Mimic Groovy Map
																	mapHighlight = CWCore.objHighlightElementJS(weSetSelElem, 'Cell Clement', strTblName);
																}
																//TODO if graphicsCheckbox or filterListbox are used always place longer text at the top of the if else
																if (StringsAndNumbers.JComm_TextLocationInString(strTempSetSelElem.toLowerCase(), "checkbox") >= 0) {
																	//Convert value to boolean
																	if (StringsAndNumbers.JComm_StringIsBoolean(strTempSetSelValue) == true) {
																		//Call setverify checkbox element with always verify
																		let mapSetVerifyCkbox = SetVerifyCheckBoxWebElement("Display Results", weSetSelElem, "Checkbox", StringsAndNumbers.JComm_StringToBoolean(strTempSetSelValue), true);
																		//Check if it passed
																		if (StringsAndNumbers.JComm_StringToBoolean(mapSetVerifyCkbox.boolPassed) == true) {
																			strTempRowDetails = strTempRowDetails + ", set the checkbox ' " + strTempSetSelElem + " in column " + strTempSelColName + " to " + strTempSetSelValue;
																			boolRowSet = true; // Action was performed
																		}
																		else {
																			boolPassed = false;
																			strMethodDetails = mapSetVerifyCkbox.strMethodDetails;
																			break;
																		}
																	}
																	else {
																		//TODO Error handling
																		boolPassed = false;
																		strMethodDetails = "FAILED!!! the value for setting the CHECKBOX is NOT BOOLEAN!!!";
																		break;
																	}
																}
																else if (StringsAndNumbers.JComm_TextLocationInString(strTempSetSelElem.toLowerCase(), "link") >= 0) {
																	//Convert value to boolean
																	if (StringsAndNumbers.JComm_StringIsBoolean(strTempSetSelValue) == true) {
																		if (StringsAndNumbers.JComm_StringToBoolean(strTempSetSelValue) == true) {
																			//click the link
																			weSetSelElem.click();
																			strTempRowDetails = strTempRowDetails + ", selected the link ' " + strTempSetSelElem + " in column " + strTempSelColName;
																			boolRowSet = true; // Action was performed
																		}
																	}
																}
																else if (StringsAndNumbers.JComm_TextLocationInString(strTempSetSelElem.toLowerCase(), "lstbox") >= 0) { // Original 'lstbox'
																	let mapSetVerifyListBox = {}; // Mimic Groovy Map
																	//Call setverify Listbox element with always verify
																	mapSetVerifyListBox = objPickItemFromListBox("Display Results", weSetSelElem, "ListBox", strTempSetSelValue, true);
																	//Check if it passed
																	if (StringsAndNumbers.JComm_StringToBoolean(mapSetVerifyListBox.boolPassed) == true) {
																		strTempRowDetails = strTempRowDetails + ", set the Listbox ' " + strTempSetSelElem + " in column " + strTempSelColName + " to " + strTempSetSelValue;
																		boolRowSet = true; // Action was performed
																	}
																	else {
																		boolPassed = false;
																		strMethodDetails = mapSetVerifyListBox.strMethodDetails;
																		break;
																	}
																}
																else if (StringsAndNumbers.JComm_TextLocationInString(strTempSetSelElem.toLowerCase(), "singlefilterlistbox") >= 0) {
																	if (boolDoSingleFilterLstBox == false) {
																		boolPassed = false;
																		strMethodDetails = "FAILED!!! The map of values DID NOT CONTAIN a FILTERLISTBOX Map of VALUES!!!";
																		break;
																	}
																	else {
																		let mapSetVerifySingleFilterListBox = {}; // Mimic Groovy Map
																		//Call setverify Listbox element with always verify
																		//TODO create a new filer listbox keyword that accepts the element
																		// add the values passed in for the filter and select. NOTE select value must be the complete item value
																		// mapSetVerfiyFilterEditBoxValues map that contains the values passed in from the module
																		// Assuming objSetVerifyEditSingleValueFilterListBox is a global function
																		mapSetVerifySingleFilterListBox = objSetVerifyEditSingleValueFilterListBox(null); // Passing null for mapSetVerfiyFilterEditBoxValues, as per Groovy
																		//Check if it passed
																		if (StringsAndNumbers.JComm_StringToBoolean(mapSetVerifySingleFilterListBox.boolPassed) == true) {
																			strTempRowDetails = strTempRowDetails + ", set the Listbox ' " + strTempSetSelElem + " in column " + strTempSelColName + " to " + strTempSetSelValue;
																			boolRowSet = true; // Action was performed
																		}
																		else {
																			boolPassed = false;
																			strMethodDetails = mapSetVerifySingleFilterListBox.strMethodDetails;
																			break;
																		}
																	}
																}
															}
															else {
																boolRowSet = false;
																boolPassed = false;
																strMethodDetails = "FAILED!!! A CELL WAS NOT RETURNED!!!";
																break;
															}
														}
														else {
															boolRowSet = false;
															boolPassed = false;
															strMethodDetails = `FAILED!!! The specifed element name '${intSetSelElemIndex}' was NOT FOUND in the arrary of elements ${arryElementNames.toString()}!!!`;
															break;
														}
													}
													else {
														boolRowSet = false;
														boolPassed = false;
														strMethodDetails = `FAILED!!! the select item value '${strTempSelValue}' DOES NOT have THREE ITEMS!!!!`;
														break;
													}
												}
												if (boolPassed == true && boolRowSet == true) { // Row set successfully for this input row, break from inner (display row) loop
													break; //Exit as we have found and set the row
												}
											}
											// if (boolRowSet == false) { // This block commented out in original Groovy, no change needed.

											// }
											//Check all the braces below and update as needed with the correct values
											//TODO if compare fails and match rows is expected, fail
										}
										if (boolPassed == false) { // If boolPassed became false, break from display row loop
											boolInputRowMatch = false; // Ensure input row match is false if inner operation failed.
											break;
										}
									} // End loopDisplayRow
									// At this point, boolInputRowMatch indicates if the current input row found a matching display row
									if (boolInputRowMatch == true) { // Current Input Row had a match
										strMethodDetails = strMethodDetails + gblLineFeed + `Results in row: ${loopDisplayRow} matched the results assigned and elements set/selected.`;
										intRowMatched = loopDisplayRow; // Update the matched row index
										break; // Break from loopInputRow since a match was found and action taken for this input row
									}
								} // End loopInputRow
								if (boolPassed != false) { // Overall success check
									strMethodDetails = strMethodDetails + gblLineFeed + strTempRowDetails;
								}
							}
							else {
								boolPassed = false; // If intBodyRowCnt 0 or not expected size, or intSubSecColCnt mismatch
								strMethodDetails = strMethodDetails + "FAILED the INPUTROW(S) TO PROCESS of: " + intInputRowsToProcess +
								" is GREATER THAN THE DISPLAYED ROW(s) of: " + intBodyRowCnt + "!!!";
							}
						}
					}
				}
				else {
					boolPassed = false; // Initial checks for header, input sheet failed
					strMethodDetails = strMethodDetails + "FAILED: Initial table or input sheet setup checks.";
				}
			}
		}
	}

	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure final detail message is captured
	mapResults.intRowMatched = intRowMatched;
	return mapResults;
}

/**
* -------------------------------------  GetIFrameIndex  -----------------------------------
* Return the IFrame Index based on property name and value specified
* @param strLocXPath		The web page XPath
* @param strProperty		The property that we should check
* @param strPropertyValue   The value of the property we must match
*
* @return mapMethodResults
*
* @author pkanaris
* @author Created: 03/28/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_GetIFrameIndex (strLocXPath, strProperty, strPropertyValue) { // Return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	let wePage;
	let strIFrameXPath = '//iframe';
	//let strTestObjText; // Declared but unused
	let intIFrameIndex = null; // Initialize to null to mimic Groovy's nullable int
	//let intIframeMatchInstance = -1; //Set to negative one so the instance count will match zero based. Declared but unused.
	//Return the element
	wePage = CWCore.returnWebElement(strLocXPath);
	if (wePage == null) {
		boolPassed = false;
		strMethodDetails = "The location Xpath'" + strLocXPath + "' DID NOT RETURN THE WEBPAGE!!!.";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(wePage, 'Page', 'MainPage');
		}
		//Return the webpage child IFrames and count
		let lstIFrames;
		let weFrame;
		let intIFrameCnt;
		let strIFramePropValue;
		let mapGetWEChildre = {}; // Mimic Groovy Map
		mapGetWEChildre = CWCore.returnChildElements(wePage, strIFrameXPath);
		lstIFrames = mapGetWEChildre.lstChildObjects;
		intIFrameCnt = mapGetWEChildre.cntChildObjs;
		if (intIFrameCnt > 0) {
			for (let loopIframes = 0; loopIframes < intIFrameCnt; loopIframes++) {
				weFrame = lstIFrames[loopIframes];
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weFrame, 'IFRame', 'IFrame Index: ' + loopIframes);
				}
				strIFramePropValue = weFrame.getAttribute(strProperty);
				if (strIFramePropValue == strPropertyValue) {
					intIFrameIndex = loopIframes;
					break;
				}
			}
			if (intIFrameIndex != null) {
				strMethodDetails = "The IFrame matching the property: " + strProperty + " and the value of: " +
				strPropertyValue + " was found and is IFrame index: " + intIFrameIndex + ".";
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! NO IFRAME FOUND matching the property: " + strProperty + " and the value of: " +
				strPropertyValue + ". Checked " + intIFrameCnt + " IFrame objects!!!!";
			}
		}
		else {
			boolPassed = false;
			strMethodDetails = "FAILED!!! DID NOT FIND ANY IFRAMES in the ASSIGNED WEBPAGE!!!";
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString(); // Using TS object assignment over `put`
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.intFrameIndex = intIFrameIndex;
	return mapResults;
}
/**
* -------------------------------------  GetChildIFrameWebElement  -----------------------------------
* Return the IFrame Index based on property name and value specified
* @param strParentXPath	 The parent Xpath
* @param strProperty		The property that we should check
* @param strPropertyValue   The value of the property we must match
*
* @return mapMethodResults
*
* @author pkanaris
* @author Created: 03/28/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_GetChildIFrameWebElement (strParentXPath, strProperty, strPropertyValue) { // Return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	let wePage;
	let weIFrame = null; // Initialize to null
	let strIFrameXPath = '//iframe';
	//let strTestObjText; // Declared but unused
	let intIFrameIndex = null; // Initialize to null
	//Return the element
	wePage = CWCore.returnWebElement(strParentXPath);
	if (wePage == null) {
		boolPassed = false;
		strMethodDetails = "The location Xpath'" + strParentXPath + "' DID NOT RETURN THE WEBPAGE!!!.";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(wePage, 'Page', 'MainPage');
		}
		//Return the webpage child IFrames and count
		let lstIFrames;
		//let weFrame; // Declared but unused in current loop context
		let intIFrameCnt;
		let strIFramePropValue;
		let mapGetWEChildre = {}; // Mimic Groovy Map
		mapGetWEChildre = CWCore.returnChildElements(wePage, strIFrameXPath);
		lstIFrames = mapGetWEChildre.lstChildObjects;
		intIFrameCnt = mapGetWEChildre.cntChildObjs;
		if (intIFrameCnt > 0) {
			for (let loopIframes = 0; loopIframes < intIFrameCnt; loopIframes++) {
				weIFrame = lstIFrames[loopIframes]; // Assign the current WebElement
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weIFrame, 'IFRame', 'IFrame Index: ' + loopIframes);
				}
				strIFramePropValue = weIFrame.getAttribute(strProperty);
				if (strIFramePropValue == strPropertyValue) {
					intIFrameIndex = loopIframes;
					break;
				}
			}
			if (intIFrameIndex != null) {
				strMethodDetails = "The IFrame matching the property: " + strProperty + " and the value of: " +
				strPropertyValue + " was found and is IFrame index: " + intIFrameIndex + ".";
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! NO IFRAME FOUND matching the property: " + strProperty + " and the value of: " +
				strPropertyValue + ". Checked " + intIFrameCnt + " IFrame objects!!!!";
			}
		}
		else {
			boolPassed = false;
			strMethodDetails = "FAILED!!! DID NOT FIND ANY IFRAMES in the ASSIGNED WEBPAGE!!!";
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.intFrameIndex = intIFrameIndex;
	mapResults.weIFrame = weIFrame;
	return mapResults;
}
/**
* -------------------------------------  MoveToElement  -----------------------------------
* Move to the specified element
* @param strLocation		The web page location value
* @param strElemFullPath	The Element's full XPath
* @param strElemName		The meaningful name of the EditBox
*
* @author pkanaris
* @author Created: 03/28/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_MoveToElement (strLocation, strElemFullPath, strElemName) { // Return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	let we;
	//let strTestObjText; // Declared but unused
	//Return the element
	we = CWCore.returnWebElement(strElemFullPath);
	if (we == null) {
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(we, 'Element', strElemName);
		}
		//Move to the element
		//Create actions to allow moving to an object in case it is not displayed.
		let actions = new Actions(TCObj.getTcDriver());
		actions.moveToElement(we).perform();
		strMethodDetails = "Moved to the elememnt '" + strElemName + "' in the location '" + strLocation + "'.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  ScrollToWebElement  -----------------------------------
* Scroll until the element is displayed.
* NOTE: THE PAGE AND SECTION MUST BE SCROLLABLE
* @param strLocation		The web page location value
* @param we				 The webelement object
* @param strElemName		The meaningful name of the EditBox
*
* @author pkanaris
* @author Created: 12/12/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_ScrollToWebElement (strLocation, we, strElemName) { // we is `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	//let strTestObjText; // Declared but unused
	if (we == null) {
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	else {
		//Scroll to the element
		//Use Javascript
		let driver = TCObj.getTcDriver();
		// Casting to any to allow executeScript without explicit JavascriptExecutor interface
		driver.executeScript("arguments[0].scrollIntoView(true);", we);
		sleep(100); // Mimic Thread.sleep
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(we, 'Element', strElemName);
		}
		strMethodDetails = "Scrolled to the elememnt '" + strElemName + "' in the location '" + strLocation + "'.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  MoveToWebElement (Overload) -----------------------------------
* Move to the specified element
* @param strLocation		The web page location value
* @param we				 The webelement object
* @param strElemName		The meaningful name of the EditBox
*
* @author pkanaris
* @author Created: 05/05/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_MoveToWebElement (strLocation, we, strElemName) { // Renamed from original duplicate, return type `any`
	//GlobalVars
	//GlobalVars
	let gblNull = GblNull;
	let gblUndefined = GblUndefined;
	let gblLineFeed = GblLineFeed;
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	//let strTestObjText; // Declared but unused
	if (we == null) {
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	else {
		//Highlight
		if (boolDoHighlight == true) {
			CommonWeb.HighlightElement(we);
		}
		//Move to the element
		Tester.SuppressReport(true);  
		we.DoMouseMove()
		Tester.SuppressReport(false)
		strMethodDetails = "Moved to the elememnt '" + strElemName + "' in the location '" + strLocation + "'.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  MoveToWebElementAndClick  -----------------------------------
* Move to the specified element and click the element
* @param strLocation		The web page location value
* @param we				 The webelement object
* @param strElemName		The meaningful name of the EditBox
*
* @author pkanaris
* @author Created: 01/06/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_MoveToWebElementAndClick (strLocation, we, strElemName) { // we is `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	//let strTestObjText; // Declared but unused
	if (we == null) {
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(we, 'Element', strElemName);
		}
		//Move to the element
		//Create actions to allow moving to an object in case it is not displayed.
		let actions = new Actions(TCObj.getTcDriver());
		actions.moveToElement(we).click().perform();
		strMethodDetails = "Moved to the elememnt '" + strElemName + "' in the location '" + strLocation + "'.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  MoveToAndClickElement  -----------------------------------
* Move to the specified element and click on the element
* NOTE: No Wait or Highlighting is present in this method
* @param strLocation		The web page location value
* @param strElemFullPath	The Element's full XPath
* @param strElemName		The meaningful name of the EditBox
*
* @author pkanaris
* @author Created: 04/04/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_MoveToAndClickElement (strLocation, strElemFullPath, strElemName) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	let we;
	//let strTestObjText; // Declared but unused
	//Return the element
	we = CWCore.returnWebElement(strElemFullPath);
	if (we == null) {
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(we, 'Element', strElemName);
		}
		//Move to the element
		//Create actions to allow moving to an object in case it is not displayed.
		let actions = new Actions(TCObj.getTcDriver());
		actions.moveToElement(we);
		we.click(); // Standard element click
		strMethodDetails = "Moved to and clicked the elememnt '" + strElemName + "' in the location '" + strLocation + "'.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  MoveToAndClickElementCustomTiming  -----------------------------------
* Move to the specified element and click on the element using the custom timings provided
* NOTE: No Wait or Highlighting is present in this method
* @param strLocation		The web page location value
* @param strElemFullPath	The Element's full XPath
* @param strElemName		The meaningful name of the EditBox
* @param intMaxWaitSecs   The number of seconds to continue trying to find the element
* @param intPollingMS		 The number of milliseconds between tries to find the element
*
* @author pkanaris
* @author Created: 04/19/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_MoveToAndClickElementCustomTiming (strLocation, strElemFullPath, strElemName, intMaxWaitSecs, intPollingMS) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	let we;
	//let strTestObjText; // Declared but unused
	//Return the element
	we = CWCore.returnWebElementCustomTiming(strElemFullPath, intMaxWaitSecs, intPollingMS);
	if (we == null) {
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(we, 'Element', strElemName);
		}
		//Move to the element
		//Create actions to allow moving to an object in case it is not displayed.
		let actions = new Actions(TCObj.getTcDriver());
		actions.moveToElement(we, 1, 1).perform(); // Move to element with offset 1,1
		sleep(1000); // Mimic Thread.sleep
		//we.click() // Commented in original Groovy, so not added.
		strMethodDetails = "Moved to and clicked the elememnt '" + strElemName + "' in the location '" + strLocation + "'.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/* MESSAGE BOX */
/**
* -------------------------------------  VerifyElementTextAndColor  -----------------------------------
* Verify the element text, color, and background color
* @param strLocation		The web page location value
* @param strElemFullPath	The full XPath
* @param strElemName		The meaningful name of the Element
* @param strAssignValue   The assigned value to be verified. The data must have status|Color:rgba(xxx, xxx, xxx, x); Background Color:rgba(xxx, xxx, xxx, x)
*
* @return mapResults		 The results showing Passed and method details.
*
* @author kpluedde
* @author Created: 09/08/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyElementTextAndColor (strLocation, strElemFullPath, strElemName,
	strAssignValue) { // Return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	let boolDoHighlight = TCExecParams.getBoolDoHighlight();
	let intViewDelay = TCExecParams.getIntViewDelaySecs();
	//Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	let weImage;
	//Item Status object variables
	let strColor;
	let strBckgColor;
	let strDisplayedColor;
	let strDisplayedText;
	let strExpectedText;
	let strExpectedColorText;
	//let strStartLocation = strLocation; // Declared but unused
	weImage = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weImage != null) {
		if (boolDoHighlight == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weImage, 'Element', "Image");
		}
		// split array
		let ArryOfItems;
		let mapSplitTextString = StringsAndNumbers.JComm_StringToArray(strAssignValue, gblDelimiter);
		let intCntValues = StringsAndNumbers.JComm_StringToInteger(mapSplitTextString.intItemCount);
		if (intCntValues == 2) {
			ArryOfItems = mapSplitTextString.ArryOfValues;
			strExpectedText = ArryOfItems[0];
			strExpectedColorText = ArryOfItems[1];
			//Return the Text/name
			strDisplayedText = StringsAndNumbers.JComm_HandleNoData(weImage.getText());
			// Compare values
			if (strDisplayedText == strExpectedText) {
				boolPassed = true;
				strMethodDetails = "The Text of: '" + strExpectedText + "' matched. ";

			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The EXPECTED Text of '" + strExpectedText + "' DOES NOT MATCH THE DISPLAYED TEXT of '" + strDisplayedText + "'!!! ";
			}
			//Return the colors
			strColor = StringsAndNumbers.JComm_HandleNoData(weImage.getCssValue("color"));
			strBckgColor = StringsAndNumbers.JComm_HandleNoData(weImage.getCssValue("background-color"));
			strDisplayedColor = "Color:" + strColor + "; Background Color:" + strBckgColor;
			//Compare color
			if (strExpectedColorText == strDisplayedColor) {
				boolPassed = true;
				strMethodDetails = strMethodDetails + "The color matched of '" + strExpectedColorText + "'";

			}
			else {
				boolPassed = false;
				strMethodDetails = strMethodDetails + "FAILED!!! The EXPECTED Text of '" + strExpectedColorText + "' DOES NOT MATCH THE DISPLAYED TEXT of '" + strDisplayedColor + "'!!!";
			}

		}
		else {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The Results DO NOT CONTAIN 2 ITEMS!!!";
		}
	}
	else {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure final detail message
	return mapResults;
}
/* EDITBOX */
/**
* -------------------------------------  SetVerifyEditBoxValue  -----------------------------------
* Set and verify the assigned value for the EditBox
* @param strLocation		The web page location value
* @param strLocXPath		The location XPath
* @param strElemName		The meaningful name of the EditBox
* @param strAssignValue   The assigned value to Set/Enter and Verify
* @param boolClearValue   Clear the value in the EditBox before Set/Enter? true/false
* @param boolVerifyValue  Verify the value in the EditBox after Set/Enter? true/false
* @param boolTabOutOfField  Tab out of the field after entry? True/False
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 04/20/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyEditBoxValue (/**string*/strActionName,/**string*/strLocation, /**string*/strLocXPath, /**string*/strElemName,
		/**string*/strAssignValue, /**boolean*/boolClearValue, /**boolean*/boolVerifyValue, /**boolean*/boolTabOutOfField) { // Return type any
	if (strAssignValue !== GblSkip || strAssignValue !== GblNotApplicable){ //Process the step only if either is not true
		//GlobalVars
		let gblNull = GblNull;
		let gblUndefined = GblUndefined;
		let gblLineFeed = GblLineFeed;
		//Create the output values
		// Declare the variables
		let strTestObjType = 'EditBox';
		let strStepAction =  strActionName;
		let strStepDescription = strActionName + ": " + strAssignValue;
		let strStepExpectedResult = "The assigned value is entered and displayed in the field.";
		let strMethodDetails;
		let boolPassed = true;
		var /**WebElementWrapper*/ weEditBox ;
		//let strTestObjText; // Declared but unused
		let boolElemStale = false;
		//Return the Element Xpath and concantenate to the strLocXPath
		let strElemXpath = getORXPath(strElemName);
		if (boolDoDebug === true) {
			Tester.Message('The SetVerifyEditBoxValue.strLocXPath: ' + strLocXPath)
		}
		if (boolDoDebug === true) {
			Tester.Message('The SetVerifyEditBoxValue.Element ' + strElemName + ' XPath is: ' + strElemXpath)
		}
		let strElemFullPath = strLocXPath + strElemXpath
		if (boolDoDebug === true) {
			Tester.Message('The SetVerifyEditBoxValue.FullXPath is: ' + strElemFullPath)
		}
		//Return the element
		weEditBox = CommonWebCore.returnWebElement(strElemFullPath);
		//Process the element
		if (weEditBox != null) {
			//Highlight
			if (boolDoHighlight == true) {
				CommonWeb.HighlightElement(weEditBox);
				//CommonWeb.HighlightElement(weEditBox);
			}
			//Check the element state (Enabled, Visible)
			let mapResultsEditState = {}; // Mimic Groovy Map
			mapResultsEditState = CommonWebCore.objVerifyState(weEditBox, strElemName, true, true);
			//Output results
			let boolEditStatePassed = StringsAndNumbers.JComm_StringToBoolean(mapResultsEditState.boolPassed);
			let strVerEditStateResults = mapResultsEditState.strMethodDetails;
			if (boolDoDebug === true) {
				Tester.Message('The SetVerifyEditBoxValue.EditStatePassed is: ' + boolEditStatePassed);
				Tester.Message('The SetVerifyEditBoxValue.EditStateResults is: ' + strVerEditStateResults);
			}
			//SetVerify the values
			if (boolEditStatePassed == true) {
				let mapResultsEditSet = {}; // Mimic Groovy Map
				mapResultsEditSet = CommonWebCore.objSetEditBoxValue(weEditBox, strElemName, strAssignValue, boolClearValue, false, boolTabOutOfField);
				let boolEditSetVerPassed = StringsAndNumbers.JComm_StringToBoolean(mapResultsEditSet.boolPassed);
				let strEditSetVerResults = mapResultsEditSet.strMethodDetails;
				try {
					weEditBox.toString(); // arbitrary method call
				}
				catch (e) {
					boolElemStale = true;
					Tester.Message("The element is stale after setting the value!!!");
				}
				//Update step details
				if (boolEditSetVerPassed == true) {
					if (boolVerifyValue == true) {
						let mapResultsVerify = {}; // Mimic Groovy Map
						mapResultsVerify = CommonWebCore.objVerifyEditBoxValue(weEditBox, strElemName, strAssignValue);
						let boolVerPassed = StringsAndNumbers.JComm_StringToBoolean(mapResultsVerify.boolPassed);
						let strVerResults = mapResultsVerify.strMethodDetails;
						try {
							weEditBox.toString(); // arbitrary method call
						}
						catch (e) {
							boolElemStale = true;
							Tester.Message("The element is stale after verifying the value!!!");
						}
						if (boolVerPassed == true) {
							strMethodDetails = "The location name: " + strLocation + " and element name '" +
							strElemName + "' is set and verified as matching the specified value of: '" + strAssignValue + "'.";
						}
						else {
							strMethodDetails = "FAILED value was set but verifed failed see details : " + strVerResults;
							boolPassed = false;
						}
					}
					else {
						strMethodDetails = "The location name: " + strLocation + " and element name '" +
							strElemName + "' is set to the specified value of: '" + strAssignValue + "'.";
					}
				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
							strElemName + "' is NOT SET TO THE ASSIGNED VALUE!!! see details below:" + gblLineFeed + strEditSetVerResults;
				}
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strVerEditStateResults;
			}
		}
		else {
			boolPassed = false;
			strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
		}
		//Update the TestStepResults
		TestStepResults.StepAction = strStepAction;
		TestStepResults.StepDesc = strStepDescription;
		TestStepResults.StepExpected = strStepExpectedResult;
		TestStepResults.StepActual = strMethodDetails;
		TestStepResults.StepData = strElemName + ": " + strAssignValue; 
		TestStepResults.StepPassed = boolPassed;
		TestExecReporting.ReportStepResults();
	}
}

// Helper functions (placeholders) that would need to be implemented or imported
// and also follow the "no interfaces, no declared types" rule.
// These are called from within the template.
/*
function CommonWeb_SetVerifyCheckBoxWebElement(locName, weElement, chkBoxName, checkValue, verifyChecked) {
	// ... implementation
}

function CommonWeb_objPickItemFromListBox(locName, weElement, listName, valueToSelect, verifySelected) {
	// ... implementation
}

function CommonWeb_objSetVerifyEditSingleValueFilterListBox(mapValues) {
	// ... implementation
}
*/
/**
* -------------------------------------  SetSelectMultiElemInRowsOfResultBody  -----------------------------------
* Set or Select Elements in more then one column of a row in the result body
* @param mapInputVariables The input variables that identify the header to include xpaths and properties and table results rows and columns
* @param NOTE: Each cell will be checked for the elements listed in the strResultsColObjNames
* @param		Rows in the input should only be present if changes to values or selection is required and can be in any order.
* @param		Provide sufficent data to find the unique row.
*
* @param INPUT SHEET Variables
* @param strInputSheetName				The sheet name assigned for the input
* @param intInputDataStartRow			The input data row number to start using for verification
* @param intInputDataEndRow			The input data row number to end verification. Enter 999 for all rows

* @param TABLE Variables
* @param strLocName						The pipe delimited location value
* @param strLocXpath						The parent Xpath
* @param strTblName						The name of the table
* @param strTblXPath						The table Xpath
* @param strHdrRowXPath					The Header row Xpath usually //thead//tr
* @param strHdrColXPath					The header column Xpath usually //th
* @param strPropertyHdrColName				The cell property that will be used to find the assigned column name. Usually 'abbr'
* @param strResultsOutShtName				The sheet name assigned for the output
* @param strResultsRowXPath				The Xpath for the rows within the results
* @param strResultsColXPath				The Xpath for the columns within the result rows
* @param arryElementNames					The array containing the list of element names that can be in the results cells.
* @param arryElemXPaths					The array containing the list of elemenT Xpaths that can be in the results cells.
* @param NOTE: the TWO Arrays must contain the same number of items
* @param strResultsCellStatElem			The name of the status element
* @param strHdrColMissingNames				For each column that does not have an identifiable name, add the name to the delimited value
* @param strMultElemColData				The delimited value with the name|count of elements where a cell contains more than one element. Currently only one column is supported
* @param strSelectElemColumns				The Select columns in order that they must be set as a value pair (strSelectColumnName|strSelectElementType|strSelectedItemValueSelected) seperated by a semicolon for each input column impacted.
* @param									i.e. New_Action|New_Action_Element;New_DocumentSearch|New_DocumentSearch_Element;New_PersonaShopping|New_PersonaShopping_Element
* @param strSelectColumnName 				The column containing the element to select/edit/set
* @param strSelectElementType 				The element type that will be in the column name
* @param strSelectedItemValueSelected 		The value to set/select/ or click as applicable.
*
* @param mapSetVerfiyFilterEditBoxValues	The map containing the values for the filter listbox child items see below:
* @param	strElemEditBoxName				The name of the editbox child
* @param 	strElemEditBoxXPath				The Xpath for the editbox used as a filter of items displayed in the list box
* @param	strElemListBoxName				The name of the listbox that opens after setting the value
* @param	strElemListBoxXPath				The Xpath for the listbox that opens after setting the value.
* @param 	strElemListBoxItemName			The name to associate with the item(s)
* @param	strElemListBoxItemXPath			The Xpath for the item(s) within the listbox
*
*
* @return mapResults 		The results showing Passed and method details. Includes the number of the row matched.
*
* @author pkanaris
* @author Created: 06/05/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetSelectMultiElemInRowsOfResultBody (mapInputVariables) { // mapInputVariables and return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	let gblSkip = GVars.GblSkip('Value');
	let gblIgnoreData = GVars.GblIgnoreData('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails;
	let boolPassed = true;
	//Return the method variables
	//Input Sheet variables
	let strInputSheetName = mapInputVariables.InputDataSheetName;
	let intInputDataStartRow = mapInputVariables.InputDataRowStart;
	let intInputDataEndRow = mapInputVariables.InputDataRowEnd;
	//Application variables
	let strLocName = mapInputVariables.LocName;
	let strLocXpath = mapInputVariables.LocXPath;
	let strTblName = mapInputVariables.TblName;
	let strTblXPath = mapInputVariables.strTblXPath;
	let strHdrRowXPath = mapInputVariables.strHdrRowXPath;
	let strHdrColXPath = mapInputVariables.strHdrColXPath;
	let strPropertyHdrColName = mapInputVariables.strPropertyHdrColName;
	let strResultsOutShtName = mapInputVariables.strResultsOutShtName;
	let strResultsRowXPath = mapInputVariables.strResultsRowXPath;
	let strResultsColXPath = mapInputVariables.strResultsColXPath;
	//Cells can possibly have only one element
	let arryElementNames = mapInputVariables.arrayResultsColObjNames;
	let arryElemXPaths = mapInputVariables.arrayResultsColObjXPaths;
	let strResultsCellStatElem = mapInputVariables.strResultsCellStatElem;
	//Alternate properties
	let strTblSepXpath = gblNull;
	if (mapInputVariables.hasOwnProperty('strTblSepXpath')){ // Use hasOwnProperty
		strTblSepXpath = mapInputVariables.strTblSepXpath;
	}
	let strHdrColMissingNames = gblNull;
	let strMultElemColData = gblNull;
	if (mapInputVariables.hasOwnProperty('strHdrColMissingNames')){
		strHdrColMissingNames = mapInputVariables.strHdrColMissingNames;
	}
	if (mapInputVariables.hasOwnProperty('strMultElemColData')){
		strMultElemColData = mapInputVariables.strMultElemColData;
	}
	//UL alternate values see: Sourcing|Templates and Libraries|Event Libraries for example
	let strULElemName = gblNull;
	let strULNoValElemName = gblNull;
	let arryULElementNames;
	let arryULElemXPaths;
	let arryMultiColSelItems;
	let intMultiColSelItemsCnt = 0;
	let intMultiColSelValueCnt = 0;
	if (mapInputVariables.hasOwnProperty('strResultsCellULElemName')){
		strULElemName = mapInputVariables.strResultsCellULElemName;
	}
	if (mapInputVariables.hasOwnProperty('strResultsCellULNoValueElemName')){
		strULNoValElemName = mapInputVariables.strResultsCellULNoValueElemName;
	}
	if (mapInputVariables.hasOwnProperty('arrayResultsColULObjNames')){
		arryULElementNames = mapInputVariables.arrayResultsColULObjNames;
	}
	if (mapInputVariables.hasOwnProperty('arrayResultsColULObjXPaths')){
		arryULElemXPaths = mapInputVariables.arrayResultsColULObjXPaths;
	}
	//Check for columns with more than one element
	let strMultElemColName = gblNull;
	let intMultElemColCount = 1; //always at least 1
	if (strMultElemColData != gblNull) {
		//Split the value into an array
		//ADD a split for the value to return the date/time format
		let mapArray = {}; // Mimic Groovy Map
		mapArray = StringsAndNumbers.JComm_StringToArray(strMultElemColData, gblDelimiter);
		if (StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount) == 2) {
			strMultElemColName = mapArray.ArryOfValues[0];
			intMultElemColCount = StringsAndNumbers.JComm_StringToInteger(mapArray.ArryOfValues[1]);
		}
	}
	//Add alternate for table where header and data is in one cell and a sub category row is present
	let intTableHeaderRowCnt = -1;
	let strSectionRowInfo = gblNull;
	let arrySectInfo;
	let intSectInfoArrySize; // Not used later in this function
	if (mapInputVariables.hasOwnProperty('intTableHeaderRowCnt')) {
		intTableHeaderRowCnt = mapInputVariables.intTableHeaderRowCnt;
	}
	let intLoopHdrCount; //Holds the number of rows in the header to be added to the loopResultStart
	if (intTableHeaderRowCnt > 0) {
		intLoopHdrCount = intTableHeaderRowCnt;
	} else {
		intLoopHdrCount = 0; // Initialize if not set
	}

	if (mapInputVariables.hasOwnProperty('strSectionRowInfo')) {
		strSectionRowInfo = mapInputVariables.strSectionRowInfo;
		if (strSectionRowInfo == gblNull || strSectionRowInfo == gblNA) {
			strSectionRowInfo = gblNull;
		}
		else {
			let mapSplitSubSectInfoString = StringsAndNumbers.JComm_StringToArray(strSectionRowInfo, gblDelimiter);
			let intCntValues = StringsAndNumbers.JComm_StringToInteger(mapSplitSubSectInfoString.intItemCount);
			if (intCntValues == 4) { // Groovy original comment says 3, but expects 4 from its example usage.
				arrySectInfo = mapSplitSubSectInfoString.ArryOfValues;
				intSectInfoArrySize = arrySectInfo.length; // Use .length for JS Array
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The Results can contain a Sub Section row, BUT, THE info provided '" + strSectionRowInfo + "', DOES NOT CONTAIN 3 ITEMS!!!";
			}
		}
	}
	//Return the singlevalue filter listbox element map
	let mapSetVerfiyFilterEditBoxValues = {}; // Mimic Groovy Map
	let boolDoSingleFilterLstBox = false;
	if (mapInputVariables.hasOwnProperty('mapSetVerfiyFilterEditBoxValues')) {
		boolDoSingleFilterLstBox = true;
		mapSetVerfiyFilterEditBoxValues = mapInputVariables.mapSetVerfiyFilterEditBoxValues;
	}
	//Add variables for working with multi-setselect columns
	let strMultSetSelColInfo;
	//let arryMultSetSelColInfo; // Declared below
	//let intMultSetSelColInfo; // Declared but unused directly
	if (mapInputVariables.hasOwnProperty('strSetSelectColumnInfo')) {
		strMultSetSelColInfo = mapInputVariables.strSetSelectColumnInfo;
		//TODO if not null split values into array
		if (strMultSetSelColInfo != gblNull) {
			let mapArray = {}; // Mimic Groovy Map
			mapArray = StringsAndNumbers.JComm_StringToArray(strMultSetSelColInfo, ";"); //Use item splitter
			intMultiColSelItemsCnt = StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount);
			if ( intMultiColSelItemsCnt > 0) {
				arryMultiColSelItems = mapArray.ArryOfValues;
				intMultiColSelValueCnt = intMultiColSelItemsCnt * 3; //3 values per item colName|inputColValue|inputColElem
			}
		}
	}
	let intResultsRowCnt;
	let intRowMatched; // Initialized below
	let boolProcCellElements;
	let arryOutputColNames;
	let strDelLstColNames;
	let intAddlColumnCnt = 0; // Note: This doesn't seem to be calculated in the current setup for this function.
	let strSelectItemValue = gblNull; // Note: strSelectItemValue is unused after this initial assignment or default (not in javadoc)
	//Check for alternate name values
	let intAltColNameArrayCnt = -1;
	let intAltColNameUsed = -1;
	let arryAltColName;
	//let strTempAltColName; // Declared but unused
	if (strHdrColMissingNames != gblNull) {
		//Split the value into an array
		//ADD a split for the value to return the date/time format
		let mapArray = {}; // Mimic Groovy Map
		mapArray = StringsAndNumbers.JComm_StringToArray(strHdrColMissingNames, gblDelimiter);
		intAltColNameArrayCnt = StringsAndNumbers.JComm_StringToInteger(mapArray.intItemCount);
		arryAltColName = mapArray.ArryOfValues;
	}
	//Define the list of possible elements that can be in the cells
	let lstCheckboxes = ['Checkbox', 'Graphic Checkbox', 'JavaScript Checkbox']; // Mimic Groovy def list
	let lstListBoxes = ['listbox', 'Single Value Filter List Box']; // Mimic Groovy def list
	//let lstAddlColElemTypes = ['Editbox','Checkbox','Graphic Checkbox','JavaScript Checkbox','listbox', 'Single Value Filter List Box']; // Declared but unused

	//Verify we have the table.
	let weTable = CWCore.returnWebElement(strLocXpath + strTblXPath);
	if (weTable == null) {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The location '" + strLocName + "' '" + strTblName + "' DID NOT RETURN an ELEMENT!!!";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weTable, 'Element', strTblName);
		}
	} // No else block for boolPassed if weTable is not null, so strMethodDetails is not set here if weTable is null
	let cntHdrCols = 0; // Default initialization

	//Check if the header row is present
	let weHeaderRow = CWCore.returnWebElement(strLocXpath + strTblXPath + strHdrRowXPath);
	if (weHeaderRow == null) {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The location '" + strLocName + "' '" + strTblName + "' DID NOT RETURN A HEADER ROW ELEMENT!!!";
	}
	else {
		//Return all the column names
		let mapGetHdrChildElements = {}; // Mimic Groovy Map
		let strColNames;
		mapGetHdrChildElements = CWCore.returnChildElements(weHeaderRow, strHdrColXPath);
		let lstHdrCols = mapGetHdrChildElements.lstChildObjects;
		let arryColumnNames; // Declared, used in a later scope
		cntHdrCols = mapGetHdrChildElements.cntChildObjs; // Now initialized

		if (cntHdrCols > 0) {
			let weHdrCol;
			let strTempValue;
			let boolAttribPresent;
			//Return the column names and replace any spaces with '_'
			for (let loopHdrCols = 0; loopHdrCols < cntHdrCols; loopHdrCols++) {
				strTempValue = gblNull;
				weHdrCol = lstHdrCols[loopHdrCols];
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weHdrCol, 'Element', strTblName + ' header row column');
				}
				boolAttribPresent = CWCore.isAttribtuePresent(weHdrCol, strPropertyHdrColName);
				if (boolAttribPresent == true) {
					strTempValue = StringsAndNumbers.JComm_HandleNoData(weHdrCol.getAttribute(strPropertyHdrColName));
				}
				//Check if we did not find a column name
				else if (intAltColNameArrayCnt >= 1 && intAltColNameArrayCnt > intAltColNameUsed) {
					intAltColNameUsed++;
					strTempValue = StringsAndNumbers.JComm_HandleNoData(arryAltColName[intAltColNameUsed]);
				}
				else {
					strTempValue = gblNull;
				}
				if (strTempValue != gblNull) {
					if (loopHdrCols == 0) {
						strColNames = StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_");
						strDelLstColNames = strTempValue;
					}
					else {
						strColNames = strColNames + gblDelimiter + StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_");
						strDelLstColNames = strDelLstColNames + "|" + strTempValue;
					}
					strColNames = strColNames + gblDelimiter + StringsAndNumbers.JComm_ReplaceValue(strTempValue, " ", "_") + "_Element"; //Add the column for the element

				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The header column '" + loopHdrCols + "' does not CONTAIN THE PROPERTY REQUIRED of:" + strPropertyHdrColName + "!!!";
					break;
				}
			}
			if (boolPassed == true) {
				//Add column names to the array
				let mapStringToArry = {}; // Mimic Groovy Map
				mapStringToArry = StringsAndNumbers.JComm_StringToArray(strDelLstColNames, gblDelimiter);
				arryColumnNames = mapStringToArry.ArryOfValues;
			}
			//Load the inputsheet
			let intInputRowCnt;
			let intInputColCnt;
			// let intHdrColCnt; // Already defined from displayed header
			let shInput; // Mimic XSSFSheet
			// let strHdrInputShtName = mapInputVariables.strHdrInputShtName; // Used later in VerifyResultHeader

			let mapOpenInputSheet = {}; // Mimic Groovy Map
			mapOpenInputSheet = ExcelData.excelGetSheetByName(TCObj.getObjWorkbook(), strInputSheetName);
			if (StringsAndNumbers.JComm_StringToBoolean(mapOpenInputSheet.boolPassed) == true) {
				shInput = mapOpenInputSheet.objWbSheet;
				//Return the row and column Count
				let mapSheetRowColCnt = {}; // Mimic Groovy Map
				mapSheetRowColCnt = ExcelData.excelGetRowAndColCount(shInput);
				if (StringsAndNumbers.JComm_StringToBoolean(mapSheetRowColCnt.boolPassed) == true) {
					intInputRowCnt = mapSheetRowColCnt.RowCount;
					intInputColCnt = mapSheetRowColCnt.ColCount;
				}
				else {
					boolPassed = false;
					strMethodDetails = mapSheetRowColCnt.strMethodDetails;
				}
				// The original Groovy `intHdrColCnt` was often reassigned here from Excel,
				// but in TS we're using the one derived from the displayed header for consistency.
				// This section is now just to potentially populate `lstHdrCols` and related header info IF it wasn't done before.
				if (boolPassed == true) {
					let mapHdrChildren = CWCore.returnChildElements(weHeaderRow, strHdrColXPath);
					// cntHdrCols = mapHdrChildren.cntChildObjs; // Do not reassign
					// lstHdrCols = mapHdrChildren.lstChildObjects; // Already defined from outer scope
				}
			}
			else {
				boolPassed = false;
				strMethodDetails = mapOpenInputSheet.strMethodDetails;
			}
			//Check if the number of columns displayed matches the number in the input sheet if not call output only and fail the step
			if (boolPassed == true) {
				let intInputTempColCnt = intInputColCnt - intAddlColumnCnt - (intMultiColSelItemsCnt * 2); //Select Columns are a value and the element thus *2
				if (cntHdrCols != intInputTempColCnt / 2) { //Each displayed column consist of the value and element column
					boolPassed = false;
					strMethodDetails = "FAILED!!! The table header is displaying: " + cntHdrCols + " BUT, EXPECTED: " + intInputTempColCnt / 2 + " COLUMNS!!! OUTPUTTING HEADER ONLY see details: " + gblLineFeed + strColNames;
				}
				else {
					//Return the input sheet column names and check if the match the output column names
					let strInputColValues;
					let mapGetInputColNames = {}; // Mimic Groovy Map
					mapGetInputColNames = ExcelData.excelGetHdrColNames(shInput);
					let boolInputColNames = StringsAndNumbers.JComm_StringToBoolean(mapGetInputColNames.boolPassed);
					if (boolInputColNames == true) {
						let strTempInputColNames;
						strInputColValues = mapGetInputColNames.ColValues;
						let intTotalAddlColumns = intAddlColumnCnt + (intMultiColSelItemsCnt * 2);
						if (intTotalAddlColumns > 0) {
							//Trim off the end values by count
							let strTmpValue = strInputColValues;
							let strTempRight;
							for (let loopTrim = 0; loopTrim < intTotalAddlColumns; loopTrim++) {
								if (loopTrim == 0) {
									strTempRight = StringsAndNumbers.JComm_GetRightTextInString(strTmpValue, gblDelimiter);
								}
								else {
									strTempRight = StringsAndNumbers.JComm_GetRightTextInString(strTmpValue, gblDelimiter) + gblDelimiter + strTempRight;
								}
								strTmpValue = StringsAndNumbers.JComm_GetLeftTextInString(strTmpValue, gblDelimiter + StringsAndNumbers.JComm_GetRightTextInString(strTmpValue, gblDelimiter));
							}
							strTempInputColNames = StringsAndNumbers.JComm_GetLeftTextInString(strInputColValues, gblDelimiter + strTempRight);
						}
						else {
							strTempInputColNames = strInputColValues;
						}
						//Check if they match the output column names
						if (strTempInputColNames != strColNames) {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The input column names '" + strTempInputColNames + "' DOES NOT MATCH the OUTPUT Columns '" + strColNames + "'!!!";
						}
						else {
							strMethodDetails = "The column names matched and the input file contains: " + strTempInputColNames + " rows, and " + intInputColCnt + " columns. Does Not Include Edit or Set input column(s).";
							let mapSplitElementString = StringsAndNumbers.JComm_StringToArray(strColNames, gblDelimiter);
							//let intCntElemNames= StringsAndNumbers.JComm_StringToInteger(mapSplitElementString.intItemCount); // Declared but unused
							arryOutputColNames = mapSplitElementString.ArryOfValues;
						}
					}
					else {
						boolPassed = false;
						strMethodDetails = mapGetInputColNames.strMethodDetails;
					}
				}
			}
			if (boolPassed == true) {
				//Check if the cell elements will be verified
				if (arryElementNames == null || arryElemXPaths == null) {
					boolProcCellElements = false;
				}
				else if (arryElementNames.length != arryElemXPaths.length) { // Use .length for JS Array
					boolPassed = false;
					strMethodDetails = "FAILED!!! THE arryElementNames and arryElemXPaths SIZE DO NOT MATCH!!!";
				}
				else {
					boolProcCellElements = true;
				}
			}
			//Return the data and verify the results
			if (boolPassed == true) {
				//Set the start and end rows
				let loopInputRowStart;
				let loopInputRowEnd;
				let loopResultStart;
				let loopResultEnd;
				let intInputCol;
				let intDataElem;
				//let intOutputRow; // Declared but unused

				let mapSetLoopStartEnd = {}; // Mimic Groovy Map
				mapSetLoopStartEnd = ExcelData.excelReturnLoopStartEndRow(intInputDataStartRow, intInputDataEndRow, intInputRowCnt);
				loopInputRowStart = mapSetLoopStartEnd.intLoopStart;
				loopInputRowEnd = mapSetLoopStartEnd.intLoopEnd;
				let intInputRowsToProcess = loopInputRowEnd - loopInputRowStart + 1;
				let strTempInputValue;
				let strTempInputElem;
				let strTempDispValue;
				let strTempDispElem;
				let strColor;
				let strBckgColor;
				strTempRowDetails = ""; // Initialize this for safe concatenation
				intRowsMatched = 0; // Initialize for cumulative count
				let intOutputExcelRow;
				let intBodyRowColCnt;
				let intOutputColumn;
				let intCellElemCnt;
				let intOutputValColIndex;
				let intOutputElemColIndex;
				//Status Counters (local to this function's scope)
				let intCellSkipped = 0;
				let intCellIgnored = 0;
				let intCellPassed = 0;
				let intCellFailed = 0;

				let strTempCellValue;
				let strTempCellElements;
				let weRow;
				let weCell;
				let mapGetRowCols;
				let lstBodyRowCols;
				let mapGetInputValue;
				let mapSetCellValue;
				let boolGetInpValue;
				//Check if the tbody is a separate XPath
				if (strTblSepXpath != gblNull) {
					//Return the body as a new table element
					weTable = CWCore.returnWebElement(strLocXpath + strTblXPath + strTblSepXpath);
					//Highlight
					if (TCExecParams.getBoolDoHighlight() == true) {
						let mapHighlight = {}; // Mimic Groovy Map
						mapHighlight = CWCore.objHighlightElementJS(weTable, 'Table Body', strTblName);
					}
					intLoopHdrCount = 0; // Reset header count for body, as it's separate
				}
				let mapGetRowElements = CWCore.returnChildElements(weTable, strResultsRowXPath);
				let lstBodyRows = mapGetRowElements.lstChildObjects;
				let intBodyRowCnt = mapGetRowElements.cntChildObjs;
				intRowMatched = -1; //Set to show no rows are matching
				intOutputExcelRow = 1; //Increment the row out to set for first output

				let intSubSecColCnt = (arrySectInfo && arrySectInfo.length >= 4) ? StringsAndNumbers.JComm_StringToInteger(arrySectInfo[0]) : -1; // Sub-section logic from verify method
				for (let loopInputRow = loopInputRowStart; loopInputRow <= loopInputRowEnd; loopInputRow++) {
					if (boolPassed == false) { // Assuming boolPassed can change to false inside the loop
						break; // Exit the loop on failure
					}
					let boolInputRowMatch = true; // reset for each input row processing iteration
					loopResultStart = 0; //Zero based index for displayed rows
					loopResultEnd = intBodyRowCnt;
					let boolCurrentInputRowProcessed = false; // Flag to indicate if current input row found a match and was processed.

					for (let loopDisplayRow = loopResultStart; loopDisplayRow < loopResultEnd; loopDisplayRow++) {
						// Reset temporary comparison flags for each display row if it means a new comparison context.
						// Based on Groovy logic, `boolComparePass` for a displayed row is crucial to continue processing that specific row.
						let boolComparePass = true; // Reset per display row attempt.


						intInputCol = 0;
						intDataElem = 1; //Set to verify data first
						weRow = lstBodyRows[loopDisplayRow]; // Access array by index
						//Move to the row to show the row we are processing in case of errors.
						let actions = new Actions(TCObj.tcDriver);
						actions.moveToElement(weRow).perform();
						//Highlight
						if (TCExecParams.getBoolDoHighlight() == true) {
							let mapHighlight = {}; // Mimic Groovy Map
							mapHighlight = CWCore.objHighlightElementJS(weRow, 'Element', strTblName + ' body row');
						}
						//Return the column elements and count
						mapGetRowCols = CWCore.returnChildElements(weRow, strResultsColXPath);
						lstBodyRowCols = mapGetRowCols.lstChildObjects;
						intBodyRowColCnt = mapGetRowCols.cntChildObjs;
						//Check for a spacer row and make sure not SubSection row
						if (intBodyRowColCnt > 0 && intBodyRowColCnt != intSubSecColCnt) { //Skips processing the row.
							if (intBodyRowColCnt == cntHdrCols) { // Corrected `intHdrColCnt` to `cntHdrCols` from outer scope.
								//Create the temp counters (these are for current display row processing only)
								let intCellSkippedTemp = 0;
								let intCellIgnoredTemp = 0;
								let intCellPassedTemp = 0;
								let intCellFailedTemp = 0;
								// boolComparePass = true; // Re-initialized inside this specific loop scope

								for (let loopDisplayCol = 0; loopDisplayCol < cntHdrCols; loopDisplayCol++) {
									strTempInputValue = gblNull;
									strTempInputElem = gblNull;
									//Return the expected values
									mapGetInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow, arryOutputColNames[intInputCol]);
									boolGetInpValue = StringsAndNumbers.JComm_StringToBoolean(mapGetInputValue.boolPassed);
									if (boolGetInpValue == true) {
										strTempInputValue = StringsAndNumbers.JComm_HandleNoData(mapGetInputValue.CellValue);
										intOutputValColIndex = mapGetInputValue.intColIndex;
										//Check if the value is a dynamic date and/or time
										if (strTempInputValue.indexOf('D[') > 0 || strTempInputValue.indexOf('T[') > 0) {
											let myTemp = DateTime.datetimeReturnDynamicFormatedDateTime(strTempInputValue);
											strTempInputValue = myTemp;
										}
									}
									intInputCol++; //Increment the input column number
									mapGetInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow, arryOutputColNames[intInputCol]);
									boolGetInpValue = StringsAndNumbers.JComm_StringToBoolean(mapGetInputValue.boolPassed);
									if (boolGetInpValue == true) {
										strTempInputElem = StringsAndNumbers.JComm_HandleNoData(mapGetInputValue.CellValue);
										intOutputElemColIndex = mapGetInputValue.intColIndex;
									}
									intInputCol++; //Increment the input column number
									weCell = lstBodyRowCols[loopDisplayCol];
									//Highlight
									if (TCExecParams.getBoolDoHighlight() == true) {
										let mapHighlight = {}; // Mimic Groovy Map
										mapHighlight = CWCore.objHighlightElementJS(weCell, 'Element', strTblName + ' body cell');
									}
									//Return the cell value
									strTempDispValue = StringsAndNumbers.JComm_HandleNoData(weCell.getText());
									//Return the cell element
									strTempDispElem = gblNull;
									let weCellElement;
									if (boolProcCellElements == true) { // boolProcCellElements indicates whether to process elements
										//Set the wait to 1 second to speed up the processing of elements in the cells
										TCObj.tcDriver.manage().timeouts().implicitlyWait(100, TimeUnit.MILLISECONDS);
										//Process the array for elements
										let intMaxElem = 1;
										if (strMultElemColName == arryColumnNames[loopDisplayCol]) {
											intMaxElem = intMultElemColCount;
										}
										strTempCellElements = gblNull; // Reset for this cell
										//Process the array for elements
										let cntCellElements = 0;
										while (cntCellElements < intMaxElem && boolPassed) { // Added boolPassed check
											for (let loopElements = 0; loopElements < arryElementNames.length && boolPassed; loopElements++) { // Added boolPassed check
												//TODO add in classes to get color such as status elements
												intCellElemCnt = CWCore.returnWebElemChildElementCount(weCell, arryElemXPaths[loopElements]);
												if (intCellElemCnt == 1) {
													cntCellElements++; //Increment the number of elements in the cell
													weCellElement = CWCore.returnChildElement(weCell, arryElemXPaths[loopElements]);
													//Highlight
													if (TCExecParams.getBoolDoHighlight() == true) {
														let mapHighlight = {}; // Mimic Groovy Map
														mapHighlight = CWCore.objHighlightElementJS(weCellElement, 'Element', strTblName + ' cell status element');
													}
													if (cntCellElements == 1) { // Only first element detected
														//Check if the element is a UL
														if (arryElementNames[loopElements] == strULElemName) { //See Sourcing|Templates and Libraries|Event Libraries for UI example
															//TODO we need to pass in the ul element names and xpaths
															//let intItemCnt = 0; // Declared but unused
															//let strItemValue; // Declared but unused
															//let strItemElem; // Declared but unused
															let weULItem;
															//Process the UL to return the items
															let mapULItems = CWCore.returnChildElements(weCellElement, arryULElemXPaths[0]);
															let lstULItems = mapULItems.lstChildObjects;
															let intULItemCnt = mapULItems.cntChildObjs;
															if (intULItemCnt == 0) {
																//Error
																boolPassed = false;
																strMethodDetails = "FAILED!!! The Cell User List contain 'ZERO' ITEM(S)!!!!";
															}
															else {
																let weULItemChild;
																for (let loopULItemsInner = 0; loopULItemsInner < intULItemCnt && boolPassed; loopULItemsInner++) { // Added boolPassed check
																	//Return the listitem
																	weULItem = lstULItems[loopULItemsInner];
																	//Add a ';' if loop is for greater than the first element to Cell and Element values
																	if (loopULItemsInner > 0) {
																		strTempCellValue = strTempCellValue + ';';
																		strTempCellElements = strTempCellElements + ';';
																	}
																	//Highlight
																	if (TCExecParams.getBoolDoHighlight() == true) {
																		let mapHighlight = {}; // Mimic Groovy Map
																		mapHighlight = CWCore.objHighlightElementJS(weULItem, 'Element', 'UserList item_' + loopULItemsInner);
																	}
																	//Check if the child items are present
																	//Loop through the UL items elements XPaths to check for a single item
																	//Start with the second element since 0 is the listitem
																	let intElemFoundCnt = 0;
																	for (let loopULItemElem = 1; loopULItemElem < arryULElementNames.length && boolPassed; loopULItemElem++) { // Added boolPassed check
																		weULItemChild = CWCore.returnChildElement(weULItem, arryULElemXPaths[loopULItemElem]);
																		if (weULItemChild != null) {
																			intElemFoundCnt++;
																			//Highlight
																			if (TCExecParams.getBoolDoHighlight() == true) {
																				let mapHighlight = {}; // Mimic Groovy Map
																				mapHighlight = CWCore.objHighlightElementJS(weULItemChild, 'Element', 'UserList item_child' + loopULItemsInner);
																			}
																			//Return the text and Add the element name and text to the variables
																			if (intElemFoundCnt == 1 && loopULItemsInner == 0) {
																				strTempCellValue = weULItemChild.getText();
																				strTempCellElements = arryULElementNames[loopULItemElem];
																			}
																			else if (intElemFoundCnt == 1 && loopULItemsInner > 0) {
																				strTempCellValue = strTempCellValue + weULItemChild.getText();
																				strTempCellElements = strTempCellElements + arryULElementNames[loopULItemElem];
																			}
																			else {
																				strTempCellValue = strTempCellValue + "|" + weULItemChild.getText();
																				strTempCellElements = strTempCellElements + "|" + arryULElementNames[loopULItemElem];
																			}
																		}
																	}
																}
															}
															strTempDispElem = strTempCellElements;
															strTempDispValue = strTempCellValue;
														}
														else if (arryElementNames[loopElements] == 'image') { // Added to match previous
															if (CWCore.isAttribtuePresent(weCellElement, 'src') == true) { // Assuming 'src' for image attribute
																strTempDispValue = StringsAndNumbers.JComm_HandleNoData(weCellElement.getAttribute('src'));
															}
															else {
																strTempDispValue = gblNull;
															}
															strTempDispElem = arryElementNames[loopElements];
														}
														else {
															strTempDispElem = arryElementNames[loopElements];
															if (arryElementNames[loopElements] == 'checkbox') { // Check if checkbox/graphic checkbox
																strTempDispValue = CWCore.objGetCheckboxChecked(weCellElement).toString();
															}
														}
														if (StringsAndNumbers.JComm_TextLocationInString(arryElemXPaths[loopElements], 'svg') >= 0) {
															let strSVGElem = StringsAndNumbers.JComm_HandleNoData(weCellElement.getText());
															if (strTempDispValue && strTempDispValue.startsWith(strSVGElem)) {
																strTempDispValue = strTempDispValue.substring(strSVGElem.length);
															} else if (strTempDispValue && strTempDispValue.endsWith(strSVGElem)) {
																strTempDispValue = strTempDispValue.substring(0, strTempDispValue.length - strSVGElem.length);
															}
														}
													}
													else { // This means the element was NOT UNIQUE (cntCellElements > 1 for this element) or was not one of the first elements.
														strTempDispElem = strTempDispElem + gblDelimiter + arryElementNames[loopElements];
													}
													//TODO add in classes to get color such as status elements
													if (arryElementNames[loopElements] == strResultsCellStatElem) {
														//Return the element
														let weStatus = CWCore.returnChildElement(weCell, arryElemXPaths[loopElements]);
														if (weStatus == null) {
															boolPassed = false;
															strMethodDetails = "FAILED!!! Result row: " + loopDisplayRow + " cell: " + (loopDisplayCol + 1) + " DID NOT CONTAIN A STATUS ELEMENT AS EXEPCTED!!!";
														}
														else {
															//Highlight
															if (TCExecParams.getBoolDoHighlight() == true) {
																let mapHighlight = {}; // Mimic Groovy Map
																mapHighlight = CWCore.objHighlightElementJS(weStatus, 'Element', strTblName + ' cell status element');
															}
															strColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("color"));
															if (StringsAndNumbers.JComm_TextLocationInString(arryElemXPaths[loopElements], "svg") >= 0) {
																strBckgColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("fill"));
															}
															else {
																strBckgColor = StringsAndNumbers.JComm_HandleNoData(weStatus.getCssValue("background-color"));
															}
															//Check if the color is the same
															//let strOrgColor = strColor; // Declared but unused
															let strTempColor = StringsAndNumbers.JComm_GetRightTextInString(strColor, "(");
															strTempColor = StringsAndNumbers.JComm_GetLeftTextInString(strTempColor, ")");
															let strTempBckColor = StringsAndNumbers.JComm_GetRightTextInString(strBckgColor, "(");
															strTempBckColor = StringsAndNumbers.JComm_GetLeftTextInString(strTempBckColor, ")");
															//let intLocInColorToBck = StringsAndNumbers.JComm_TextLocationInString(strTempColor, strTempBckColor); // Declared but unused
															//let intLocInBckToColor = StringsAndNumbers.JComm_TextLocationInString(strTempBckColor, strTempColor); // Declared but unused
															let strTempStatusColor;
															//Return the property and add to the strTempCellElements
															if (StringsAndNumbers.JComm_TextLocationInString(strTempColor, strTempBckColor) >= 0 || StringsAndNumbers.JComm_TextLocationInString(strTempBckColor, strTempColor) >= 0) {
																strColor = "RGB(255, 255, 255)"; //Change the color for the text to white. Will work for all but white backgrounds
																strTempStatusColor = "Colors matched replaced original color:" + strOrgColor + " with Color:" + strColor + "Background Color:" + strBckgColor;
															}
															else {
																strTempStatusColor = "Color:" + strColor + "Background Color:" + strBckgColor;
															}
															strTempDispElem = strTempDispElem + gblDelimiter + strTempStatusColor;
														}
													}
												}
												if (cntCellElements == 0) {
													break;
												}
											}
											//Return the wait to maxwaittime
											TCObj.tcDriver.manage().timeouts().implicitlyWait(TCObj.getIntTempMaxWaitTime(), TimeUnit.SECONDS);
										}
										//Compare each value
										let strCompareDispValueFinal; // Re-declared for this smaller scope
										let strCompareInputValueFinal; // Re-declared for this smaller scope
										let strOutputCellThemeFinal; // Re-declared for this smaller scope
										let intOutputColForCellFinal; // Re-declared for this smaller scope
										for (let loopCompareInner = 1; loopCompareInner <= 2; loopCompareInner++) { // Correct inner loop condition
											if (loopCompareInner == 1) {
												strCompareDispValueFinal = strTempDispValue;
												strCompareInputValueFinal = strTempInputValue;
												intOutputColForCellFinal = intInputCol - 2;
											}
											else {
												strCompareDispValueFinal = strTempDispElem;
												strCompareInputValueFinal = strTempInputElem;
												intOutputColForCellFinal = intInputCol - 1;
											}
											//Check if they match
											if (strCompareInputValueFinal == gblSkip || strCompareInputValueFinal == gblIgnoreData) {
												strOutputCellThemeFinal = 'TestRunDataSkipIgnore';
												if (strCompareInputValueFinal == gblSkip) {
													intCellSkippedTemp++;
												}
												else {
													intCellIgnoredTemp++;
												}
											}
											else if (strCompareInputValueFinal != strCompareDispValueFinal) {
												boolComparePass = false;
												boolInputRowMatch = false; // Set to false if any column in this row doesn't match
												break; // Break inner loop if content doesn't match
											} else { // It matches
												boolComparePass = true;
												// No increment on passed, as it sums up all matched in this loop
											}
										}
										if (boolComparePass == false){ // If comparison failed in this column pair
											if (TCExecParams.getBoolDoDebug()== true) {
												Tester.Message("Compare failed and matchinputRowNumber == false so break for columns");
											}
											//Break so we move to the next column processing for the current display row
											break;
										}
									}
									let boolRowSet = false; // Initialize for this path
									//Set the selection if the verify for the row passed
									if (boolComparePass == true) { // boolComparePass indicates this *displayed* row matches the *input* row
										//Process the displayed row using the assigned values for the columns
										let strTempSelValue;
										let strTempSelColName;
										let strTempSelColValue;
										let strTempSelColElem;
										let strTempSetSelValue;
										let strTempSetSelElem;
										let strTempSetSelXpath;
										//let strTempElemDetails; // Declared but unused
										let arrySelColValues;
										let weSetSelCell;
										let weSetSelElem;
										let intSetSelColNumber;
										let intSetSelElemIndex;
										boolRowSet = true; // Assume row will be set unless an action fails
										strTempRowDetails = "Displayed Row: " + loopDisplayRow + " ";
										//Input row loopInputRow and display row loopDisplayRow are required along with intMultiColSelItemsCnt and array of elements
										for (let intLoopSelValues = 0; intLoopSelValues < intMultiColSelItemsCnt && boolPassed; intLoopSelValues++) { // Added boolPassed check
											// This `boolRowSet = false` here means if any individual SELECT action fails, the *entire* current display row operation is considered failed for all `intMultiColSelItemsCnt` actions.
											// Original Groovy has `boolRowSet=false` *before* the inner loop's item processing.
											// If this is set to false here, the outer `boolPassed` needs to capture it.
											boolRowSet = true; // This specific action group is okay so far.
											strTempSelValue = arryMultiColSelItems[intLoopSelValues];
											let mapSplitElementString = StringsAndNumbers.JComm_StringToArray(strTempSelValue, gblDelimiter);
											let intCntSelValueItems = StringsAndNumbers.JComm_StringToInteger(mapSplitElementString.intItemCount);
											if (intCntSelValueItems == 3) { //Must always be 3 columnName, columnValue, columnElement
												//Return the values
												arrySelColValues = mapSplitElementString.ArryOfValues;
												strTempSelColName = arrySelColValues[0]; // Select Column Name for this action
												strTempSelColValue = arrySelColValues[1]; // Input Sheet Column Name for Value
												strTempSelColElem = arrySelColValues[2]; // Input Sheet Column Name for Element Type
												//Return the value and element from the input sheet.
												let mapGetSetSelValue = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow, strTempSelColValue);
												strTempSetSelValue = mapGetSetSelValue.CellValue;
												let mapGetSetSelElem = ExcelData.excelGetCellValueByRowNumColName(shInput, loopInputRow, strTempSelColElem);
												strTempSetSelElem = mapGetSetSelElem.CellValue;
												if ((strTempSetSelValue != gblNull && strTempSetSelValue != gblNA) && (strTempSetSelElem != gblNull && strTempSetSelElem != gblNA)) {
													intSetSelElemIndex = StringsAndNumbers.JComm_ReturnStringIndexFromArray(arryElementNames, strTempSetSelElem);
													if (intSetSelElemIndex != -1) {
														strTempSetSelXpath = arryElemXPaths[intSetSelElemIndex];
														//Get the cell by display row and column name
														intSetSelColNumber = StringsAndNumbers.JComm_ReturnStringIndexFromArray(arryColumnNames, strTempSelColName);
														weSetSelCell = lstBodyRowCols[intSetSelColNumber];
														if (weSetSelCell != null) {
															//Highlight
															if (TCExecParams.getBoolDoHighlight() == true) {
																let mapHighlight = {}; // Mimic Groovy Map
																mapHighlight = CWCore.objHighlightElementJS(weSetSelCell, 'Display Cell', strTblName);
															}
															//Return the web element
															weSetSelElem = CWCore.returnChildElement(weSetSelCell, strTempSetSelXpath);
															//Highlight
															if (TCExecParams.getBoolDoHighlight() == true && weSetSelElem != null) {
																let mapHighlight = {}; // Mimic Groovy Map
																mapHighlight = CWCore.objHighlightElementJS(weSetSelElem, 'Cell Clement', strTblName);
															}
															//TODO if graphicsCheckbox or filterListbox are used always place longer text at the top of the if else
															if (StringsAndNumbers.JComm_TextLocationInString(strTempSetSelElem.toLowerCase(), "checkbox") >= 0) {
																//Convert value to boolean
																if (StringsAndNumbers.JComm_StringIsBoolean(strTempSetSelValue) == true) {
																	let mapSetVerifyCkbox = SetVerifyCheckBoxWebElement("Display Results", weSetSelElem, "Checkbox", StringsAndNumbers.JComm_StringToBoolean(strTempSetSelValue), true);
																	if (StringsAndNumbers.JComm_StringToBoolean(mapSetVerifyCkbox.boolPassed) == true) {
																		strTempRowDetails = strTempRowDetails + ", set the checkbox ' " + strTempSetSelElem + " in column " + strTempSelColName + " to " + strTempSetSelValue;
																		//boolRowSet = true; // Action was performed (this implies current multisel item passed)
																	}
																	else {
																		boolPassed = false;
																		strMethodDetails = mapSetVerifyCkbox.strMethodDetails;
																		break; // Break from intLoopSelValues
																	}
																}
																else {
																	boolPassed = false;
																	strMethodDetails = "FAILED!!! the value for setting the CHECKBOX is NOT BOOLEAN!!!";
																	break; // Break from intLoopSelValues
																}
															}
															else if (StringsAndNumbers.JComm_TextLocationInString(strTempSetSelElem.toLowerCase(), "link") >= 0) {
																//Convert value to boolean
																if (StringsAndNumbers.JComm_StringIsBoolean(strTempSetSelValue) == true) {
																	if (StringsAndNumbers.JComm_StringToBoolean(strTempSetSelValue) == true) {
																		weSetSelElem.click();
																		strTempRowDetails = strTempRowDetails + ", selected the link ' " + strTempSetSelElem + " in column " + strTempSelColName;
																		//boolRowSet = true; // Action was performed
																	} // No else for false, so it just skips action
																}
															}
															else if (StringsAndNumbers.JComm_TextLocationInString(strTempSetSelElem.toLowerCase(), "lstbox") >= 0) { // Original 'lstbox'
																let mapSetVerifyListBox = {}; // Mimic Groovy Map
																mapSetVerifyListBox = objPickItemFromListBox("Display Results", weSetSelElem, "ListBox", strTempSetSelValue, true);
																if (StringsAndNumbers.JComm_StringToBoolean(mapSetVerifyListBox.boolPassed) == true) {
																	strTempRowDetails = strTempRowDetails + ", set the Listbox ' " + strTempSetSelElem + " in column " + strTempSelColName + " to " + strTempSetSelValue;
																	//boolRowSet = true; // Action was performed
																}
																else {
																	boolPassed = false;
																	strMethodDetails = mapSetVerifyListBox.strMethodDetails;
																	break; // Break from intLoopSelValues
																}
															}
															else if (StringsAndNumbers.JComm_TextLocationInString(strTempSetSelElem.toLowerCase(), "singlefilterlistbox") >= 0) {
																if (boolDoSingleFilterLstBox == false) {
																	boolPassed = false;
																	strMethodDetails = "FAILED!!! The map of values DID NOT CONTAIN a FILTERLISTBOX Map of VALUES!!!";
																	break; // Break from intLoopSelValues
																}
																else {
																	let mapSetVerifySingleFilterListBox = {}; // Mimic Groovy Map
																	mapSetVerifySingleFilterListBox = objSetVerifyEditSingleValueFilterListBox(null); // Passing null for mapSetVerfiyFilterEditBoxValues as per Groovy
																	if (StringsAndNumbers.JComm_StringToBoolean(mapSetVerifySingleFilterListBox.boolPassed) == true) {
																		strTempRowDetails = strTempRowDetails + ", set the Listbox ' " + strTempSetSelElem + " in column " + strTempSelColName + " to " + strTempSetSelValue;
																		//boolRowSet = true; // Action was performed
																	}
																	else {
																		boolPassed = false;
																		strMethodDetails = mapSetVerifySingleFilterListBox.strMethodDetails;
																		break; // Break from intLoopSelValues
																	}
																}
															} else { // No specific type matched, so it might be a general click or text/value set
															   if (strTempSetSelValue != gblNull && strTempSetSelValue != gblNA) {
																   // Assuming direct click if no specific type matched for boolIsChecked == true
																	// For generic editbox, assuming SetVerifyEditBoxValue, but need to check if it's already defined
																	// For this template, a direct click or sendKeys if applicable
																	try {
																		weSetSelElem.click(); // Try a generic click
																		strTempRowDetails += `, clicked generic element '${strTempSetSelElem}' in column '${strTempSelColName}'`;
																		//boolRowSet = true;
																	} catch (clickError) {
																		// Fallback to sendKeys if it's an input field
																		try {
																			weSetSelElem.clear();
																			weSetSelElem.sendKeys(strTempSetSelValue);
																			strTempRowDetails += `, set generic element '${strTempSetSelElem}' in column '${strTempSelColName}' to '${strTempSetSelValue}'`;
																			//boolRowSet = true;
																		} catch (sendKeysError) {
																			boolPassed = false;
																			strMethodDetails = `FAILED!!! Could not click or set generic element '${strTempSetSelElem}' in column '${strTempSelColName}'.`;
																			break; // Break from intLoopSelValues
																		}
																	}
															   }
															}
														}
														else { // weSetSelCell was null
															boolRowSet = false;
															boolPassed = false;
															strMethodDetails = "FAILED!!! A CELL WAS NOT RETURNED!!!";
															break; // Break from intLoopSelValues
														}
													}
													else { // intSetSelElemIndex == -1
														boolRowSet = false;
														boolPassed = false;
														strMethodDetails = `FAILED!!! The specifed element name '${strTempSetSelElem}' was NOT FOUND in the array of elements ${arryElementNames.toString()}!!!`;
														break; // Break from intLoopSelValues
													}
												}
												// boolPassed can change either due to this inner loop failing, or earlier in the function
												if (boolPassed == false) { // If an action failed, break this entire loop for current display row
													break;
												}
											}
											if (boolPassed == true) { // If all actions in intMultiColSelItemsCnt succeeded for this display row
												boolRowSet = true; // Confirm this display row was completely set.
												break; // Exit loopDisplayCol loop, as a match was found and processed.
											}
										}
										// Original Groovy has this commented out block, keeping it out.
										// if (boolRowSet == false) { /* ... */ }

									} // End if (boolComparePass == true) (This indicates a row matched)

									if (boolPassed == false) { // If boolPassed became false, break from display row loop
										// boolInputRowMatch should have been set to false too if comparePass became false.
										break;
									}
								} // End loopDisplayRow (iterating through displayed rows to find a match for current input row)

								if (boolInputRowMatch == true) { // Current Input Row had a match for its data pattern in a displayed row
									strMethodDetails = strMethodDetails + gblLineFeed + `Results in row: ${loopDisplayRow} matched the results assigned and elements set/selected.`;
									intRowMatched = loopDisplayRow; // Update the matched row index (0-indexed)
									break; // Break from loopInputRow since a match was found and action taken for this input row, don't need to check other input rows
								}
							} // End if (intBodyRowColCnt == cntHdrCols)
							else { // if intBodyRowColCnt is not equal to cntHdrCols (e.g., spacer row)
								// This would be where you might handle spacer rows or other cases, if needed.
								// Original code does nothing here, so just continues to next display row.
							}
						} // End loopInputRow
						// If intRowMatched is still -1 after all loopDisplayRow iterations, this input row found no match.
						if (intRowMatched == -1) {
							boolPassed = false; // Input row did not match any displayed row
							strMethodDetails = strMethodDetails + gblLineFeed + `FAILED: Input row ${loopInputRow} did not find a matching displayed row.`;
						}
					}
				}
			} // End if (boolPassed == true) for header column logic
		} // End if (cntHdrCols > 0)
	} // End else for weHeaderRow null check

	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure final detail message is captured
	mapResults.intRowMatched = intRowMatched;
	return mapResults;
}

/**
* -------------------------------------  GetIFrameIndex  -----------------------------------
* Return the IFrame Index based on property name and value specified
* @param strLocXPath		The web page XPath
* @param strProperty		The property that we should check
* @param strPropertyValue   The value of the property we must match
*
* @return mapMethodResults
*
* @author pkanaris
* @author Created: 03/28/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_GetIFrameIndex (strLocXPath, strProperty, strPropertyValue) { // Return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	let wePage;
	let strIFrameXPath = '//iframe';
	//let strTestObjText; // Declared but unused
	let intIFrameIndex = null; // Initialize to null to mimic Groovy's nullable int
	//let intIframeMatchInstance = -1; //Set to negative one so the instance count will match zero based. Declared but unused.
	//Return the element
	wePage = CWCore.returnWebElement(strLocXPath);
	if (wePage == null) {
		boolPassed = false;
		strMethodDetails = "The location Xpath'" + strLocXPath + "' DID NOT RETURN THE WEBPAGE!!!.";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(wePage, 'Page', 'MainPage');
		}
		//Return the webpage child IFrames and count
		let lstIFrames;
		let weFrame;
		let intIFrameCnt;
		let strIFramePropValue;
		let mapGetWEChildre = {}; // Mimic Groovy Map
		mapGetWEChildre = CWCore.returnChildElements(wePage, strIFrameXPath);
		lstIFrames = mapGetWEChildre.lstChildObjects;
		intIFrameCnt = mapGetWEChildre.cntChildObjs;
		if (intIFrameCnt > 0) {
			for (let loopIframes = 0; loopIframes < intIFrameCnt; loopIframes++) {
				weFrame = lstIFrames[loopIframes];
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weFrame, 'IFRame', 'IFrame Index: ' + loopIframes);
				}
				strIFramePropValue = weFrame.getAttribute(strProperty);
				if (strIFramePropValue == strPropertyValue) {
					intIFrameIndex = loopIframes;
					break;
				}
			}
			if (intIFrameIndex != null) {
				strMethodDetails = "The IFrame matching the property: " + strProperty + " and the value of: " +
				strPropertyValue + " was found and is IFrame index: " + intIFrameIndex + ".";
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! NO IFRAME FOUND matching the property: " + strProperty + " and the value of: " +
				strPropertyValue + ". Checked " + intIFrameCnt + " IFrame objects!!!!";
			}
		}
		else {
			boolPassed = false;
			strMethodDetails = "FAILED!!! DID NOT FIND ANY IFRAMES in the ASSIGNED WEBPAGE!!!";
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString(); // Using TS object assignment over `put`
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.intFrameIndex = intIFrameIndex;
	return mapResults;
}
/**
* -------------------------------------  GetChildIFrameWebElement  -----------------------------------
* Return the IFrame Index based on property name and value specified
* @param strParentXPath	 The parent Xpath
* @param strProperty		The property that we should check
* @param strPropertyValue   The value of the property we must match
*
* @return mapMethodResults
*
* @author pkanaris
* @author Created: 03/28/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_GetChildIFrameWebElement (strParentXPath, strProperty, strPropertyValue) { // Return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	let wePage;
	let weIFrame = null; // Initialize to null
	let strIFrameXPath = '//iframe';
	//let strTestObjText; // Declared but unused
	let intIFrameIndex = null; // Initialize to null
	//Return the element
	wePage = CWCore.returnWebElement(strParentXPath);
	if (wePage == null) {
		boolPassed = false;
		strMethodDetails = "The location Xpath'" + strParentXPath + "' DID NOT RETURN THE WEBPAGE!!!.";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(wePage, 'Page', 'MainPage');
		}
		//Return the webpage child IFrames and count
		let lstIFrames;
		//let weFrame; // Declared but unused in current loop context
		let intIFrameCnt;
		let strIFramePropValue;
		let mapGetWEChildre = {}; // Mimic Groovy Map
		mapGetWEChildre = CWCore.returnChildElements(wePage, strIFrameXPath);
		lstIFrames = mapGetWEChildre.lstChildObjects;
		intIFrameCnt = mapGetWEChildre.cntChildObjs;
		if (intIFrameCnt > 0) {
			for (let loopIframes = 0; loopIframes < intIFrameCnt; loopIframes++) {
				weIFrame = lstIFrames[loopIframes]; // Assign the current WebElement
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weIFrame, 'IFRame', 'IFrame Index: ' + loopIframes);
				}
				strIFramePropValue = weIFrame.getAttribute(strProperty);
				if (strIFramePropValue == strPropertyValue) {
					intIFrameIndex = loopIframes;
					break;
				}
			}
			if (intIFrameIndex != null) {
				strMethodDetails = "The IFrame matching the property: " + strProperty + " and the value of: " +
				strPropertyValue + " was found and is IFrame index: " + intIFrameIndex + ".";
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! NO IFRAME FOUND matching the property: " + strProperty + " and the value of: " +
				strPropertyValue + ". Checked " + intIFrameCnt + " IFrame objects!!!!";
			}
		}
		else {
			boolPassed = false;
			strMethodDetails = "FAILED!!! DID NOT FIND ANY IFRAMES in the ASSIGNED WEBPAGE!!!";
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.intFrameIndex = intIFrameIndex;
	mapResults.weIFrame = weIFrame;
	return mapResults;
}
/**
* -------------------------------------  MoveToElement  -----------------------------------
* Move to the specified element
* @param strLocation		The web page location value
* @param strElemFullPath	The Element's full XPath
* @param strElemName		The meaningful name of the EditBox
*
* @author pkanaris
* @author Created: 03/28/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_MoveToElement (strLocation, strElemFullPath, strElemName) { // Return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	let we;
	//let strTestObjText; // Declared but unused
	//Return the element
	we = CWCore.returnWebElement(strElemFullPath);
	if (we == null) {
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(we, 'Element', strElemName);
		}
		//Move to the element
		//Create actions to allow moving to an object in case it is not displayed.
		let actions = new Actions(TCObj.getTcDriver());
		actions.moveToElement(we).perform();
		strMethodDetails = "Moved to the elememnt '" + strElemName + "' in the location '" + strLocation + "'.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  ScrollToWebElement  -----------------------------------
* Scroll until the element is displayed.
* NOTE: THE PAGE AND SECTION MUST BE SCROLLABLE
* @param strLocation		The web page location value
* @param we				 The webelement object
* @param strElemName		The meaningful name of the EditBox
*
* @author pkanaris
* @author Created: 12/12/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_ScrollToWebElement (strLocation, we, strElemName) { // we is `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	//let strTestObjText; // Declared but unused
	if (we == null) {
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	else {
		//Scroll to the element
		//Use Javascript
		let driver = TCObj.getTcDriver();
		driver.executeScript("arguments[0].scrollIntoView(true);", we); // Casting to any to allow executeScript call
		sleep(100); // Mimic Thread.sleep
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(we, 'Element', strElemName);
		}
		strMethodDetails = "Scrolled to the elememnt '" + strElemName + "' in the location '" + strLocation + "'.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  MoveToWebElement (Overload) -----------------------------------
* Move to the specified element
* @param strLocation		The web page location value
* @param we				 The webelement object
* @param strElemName		The meaningful name of the EditBox
*
* @author pkanaris
* @author Created: 05/05/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
// The parameter signature of this function CommonWeb_is identical to ScrollToWebElement and overlaps with the global `MoveToWebElement`
// from the previous section. In Groovy, method overloading works by parameter types.
// In TypeScript/JavaScript, this would be a re-declaration unless parameters are different or exported differently.
// For template, assuming unique functions, renaming may be necessary, or choosing the first one.
// Renaming this one since the one with string elemFullPath is declared first.
function CommonWeb_MoveToWebElement_From_WebElement_Param (strLocation, we, strElemName) { // we is `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	//let strTestObjText; // Declared but unused
	if (we == null) {
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(we, 'Element', strElemName);
		}
		//Move to the element
		//Create actions to allow moving to an object in case it is not displayed.
		let actions = new Actions(TCObj.getTcDriver());
		actions.moveToElement(we).perform();
		strMethodDetails = "Moved to the elememnt '" + strElemName + "' in the location '" + strLocation + "'.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  MoveToWebElementAndClick  -----------------------------------
* Move to the specified element and click the element
* @param strLocation		The web page location value
* @param we				 The webelement object
* @param strElemName		The meaningful name of the EditBox
*
* @author pkanaris
* @author Created: 01/06/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_MoveToWebElementAndClick (strLocation, we, strElemName) { // we is `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	//let strTestObjText; // Declared but unused
	if (we == null) {
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(we, 'Element', strElemName);
		}
		//Move to the element
		//Create actions to allow moving to an object in case it is not displayed.
		let actions = new Actions(TCObj.getTcDriver());
		actions.moveToElement(we).click().perform(); // Perform click as part of the action sequence
		strMethodDetails = "Moved to the elememnt '" + strElemName + "' in the location '" + strLocation + "'.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  MoveToAndClickElement  -----------------------------------
* Move to the specified element and click on the element
* NOTE: No Wait or Highlighting is present in this method
* @param strLocation		The web page location value
* @param strElemFullPath	The Element's full XPath
* @param strElemName		The meaningful name of the EditBox
*
* @author pkanaris
* @author Created: 04/04/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_MoveToAndClickElement (strLocation, strElemFullPath, strElemName) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	let we;
	//let strTestObjText; // Declared but unused
	//Return the element
	we = CWCore.returnWebElement(strElemFullPath);
	if (we == null) {
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(we, 'Element', strElemName);
		}
		//Move to the element
		//Create actions to allow moving to an object in case it is not displayed.
		let actions = new Actions(TCObj.getTcDriver());
		actions.moveToElement(we); // Just move, no perform() yet for Groovy's actions.moveToElement(we)
		we.click(); // Then perform click
		strMethodDetails = "Moved to and clicked the elememnt '" + strElemName + "' in the location '" + strLocation + "'.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  MoveToAndClickElementCustomTiming  -----------------------------------
* Move to the specified element and click on the element using the custom timings provided
* NOTE: No Wait or Highlighting is present in this method
* @param strLocation		The web page location value
* @param strElemFullPath	The Element's full XPath
* @param strElemName		The meaningful name of the EditBox
* @param intMaxWaitSecs   The number of seconds to continue trying to find the element
* @param intPollingMS		 The number of milliseconds between tries to find the element
*
* @author pkanaris
* @author Created: 04/19/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_MoveToAndClickElementCustomTiming (strLocation, strElemFullPath, strElemName, intMaxWaitSecs, intPollingMS) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	let we;
	//let strTestObjText; // Declared but unused
	//Return the element
	we = CWCore.returnWebElementCustomTiming(strElemFullPath, intMaxWaitSecs, intPollingMS);
	if (we == null) {
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	else {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(we, 'Element', strElemName);
		}
		//Move to the element
		//Create actions to allow moving to an object in case it is not displayed.
		let actions = new Actions(TCObj.getTcDriver());
		actions.moveToElement(we, 1, 1).perform(); // Move to element with offset 1,1
		sleep(1000); // Mimic Thread.sleep
		//we.click() // Commented in original Groovy, so not added.
		strMethodDetails = "Moved to and clicked the elememnt '" + strElemName + "' in the location '" + strLocation + "'.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/* MESSAGE BOX */
/**
* -------------------------------------  HighlightElement  -----------------------------------
* Highlight the specified element.
* @param strLocation		The web page location value
* @param strElemFullPath	The Message Box full XPath
* @param strElemName		The meaningful name of the Message Box
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 03/30/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_HighlightElementGroovyVersion (strLocation, strElemFullPath, strElemName) { // Return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'MsgBox';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	let weElem;
	//let strTestObjText; // Declared but unused
	//Return the element
	weElem = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weElem != null) {
		//Highlight
		let mapHighlight = {}; // Mimic Groovy Map
		mapHighlight = CWCore.objHighlightElementJS(weElem, strTestObjType, strElemName);
		boolPassed = StringsAndNumbers.JComm_StringToBoolean(mapHighlight.boolPassed);
		strMethodDetails = mapHighlight.strMethodDetails;
	}
	else {
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  VerifyElementTextAndColor  -----------------------------------
* Verify the element text, color, and background color
* @param strLocation		The web page location value
* @param strElemFullPath	The full XPath
* @param strElemName		The meaningful name of the Element
* @param strAssignValue   The assigned value to be verified. The data must have status|Color:rgba(xxx, xxx, xxx, x); Background Color:rgba(xxx, xxx, xxx, x)
*
* @return mapResults		 The results showing Passed and method details.
*
* @author kpluedde
* @author Created: 09/08/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyElementTextAndColor (strLocation, strElemFullPath, strElemName,
	strAssignValue) { // Return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	let boolDoHighlight = TCExecParams.getBoolDoHighlight();
	//let intViewDelay = TCExecParams.getIntViewDelaySecs(); // Declared but unused
	//Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails;
	let boolPassed = true;
	let weImage;
	//Item Status object variables
	let strColor;
	let strBckgColor;
	let strDisplayedColor;
	let strDisplayedText;
	let strExpectedText;
	let strExpectedColorText;
	//let strStartLocation = strLocation; // Declared but unused
	weImage = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weImage != null) {
		if (boolDoHighlight == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weImage, 'Element', "Image");
		}
		// split array
		let ArryOfItems;
		let mapSplitTextString = StringsAndNumbers.JComm_StringToArray(strAssignValue, gblDelimiter);
		let intCntValues = StringsAndNumbers.JComm_StringToInteger(mapSplitTextString.intItemCount);
		if (intCntValues == 2) {
			ArryOfItems = mapSplitTextString.ArryOfValues;
			strExpectedText = ArryOfItems[0];
			strExpectedColorText = ArryOfItems[1];
			//Return the Text/name
			strDisplayedText = StringsAndNumbers.JComm_HandleNoData(weImage.getText());
			// Compare values
			if (strDisplayedText == strExpectedText) {
				boolPassed = true;
				strMethodDetails = "The Text of: '" + strExpectedText + "' matched. ";

			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The EXPECTED Text of '" + strExpectedText + "' DOES NOT MATCH THE DISPLAYED TEXT of '" + strDisplayedText + "'!!! ";
			}
			//Return the colors
			strColor = StringsAndNumbers.JComm_HandleNoData(weImage.getCssValue("color"));
			strBckgColor = StringsAndNumbers.JComm_HandleNoData(weImage.getCssValue("background-color"));
			strDisplayedColor = "Color:" + strColor + "; Background Color:" + strBckgColor;
			//Compare color
			if (strExpectedColorText == strDisplayedColor) {
				boolPassed = true;
				strMethodDetails = strMethodDetails + "The color matched of '" + strExpectedColorText + "'";

			}
			else {
				boolPassed = false;
				strMethodDetails = strMethodDetails + "FAILED!!! The EXPECTED Text of '" + strExpectedColorText + "' DOES NOT MATCH THE DISPLAYED TEXT of '" + strDisplayedColor + "'!!!";
			}

		}
		else {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The Results DO NOT CONTAIN 2 ITEMS!!!";
		}
	}
	else {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure final detail message
	return mapResults;
}
/* EDITBOX */
/**
* -------------------------------------  SetVerifyEditBoxSelectEnter  -----------------------------------
* Set and verify the assigned value for the EditBox
* @param strLocation		The web page location value
* @param strElemFullXPath   The EditBox full XPath
* @param strElemName		The meaningful name of the EditBox
* @param strAssignValue   The assigned value to Set/Enter and Verify
* @param boolClearValue   Clear the value in the EditBox before Set/Enter? true/false
* @param boolVerifyValue  Verify the value in the EditBox after Set/Enter? true/false
* @param boolClickEnter	 Select Enter after we set the value? True/False
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 04/20/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyEditBoxSelectEnter (strLocation, strElemFullXPath, strElemName,
		strAssignValue, boolClearValue, boolVerifyValue, boolClickEnter ) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'EditBox';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weEditBox;	 // Type inferred
	//let strTestObjText; // Declared but unused

	//Return the element
	weEditBox = CWCore.returnWebElement(strElemFullXPath);
	//Process the element
	if (weEditBox != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weEditBox, 'Edit', strElemName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsEditState = {}; // Mimic Groovy Map
		mapResultsEditState = CWCore.objVerifyState(weEditBox, strElemName, true, true);
		//Output results
		let boolEditStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsEditState.boolPassed);
		let strVerEditStateResults = mapResultsEditState.strMethodDetails;
		//SetVerify the values
		if (boolEditStatePassed == true) {
			let mapResultsEditSet = {}; // Mimic Groovy Map
			mapResultsEditSet = CWCore.objSetVerEditBoxValueClickEnter(weEditBox, strElemName, strElemFullXPath, strAssignValue, boolClearValue, boolVerifyValue, boolClickEnter);
			let boolEditSetVerPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsEditSet.boolPassed);
			let strEditSetVerResults = mapResultsEditSet.strMethodDetails;
			//Update step details
			if (boolEditSetVerPassed == true) {
				strMethodDetails = "The location name: " + strLocation + " and element name '" +
				strElemName + "' is set to the value of: " + strAssignValue + ".";
			}
			else
			{
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strElemName + "' is NOT SET TO THE ASSIGNED VALUE!!! see details below:" + gblLineFeed + strEditSetVerResults;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strVerEditStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString(); // Using TS object assignment over `put`
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  SetVerifyHiddenFilePathValue  -----------------------------------
* Set and verify the assigned value for the hidden file path input
* NOTE: the input file may be displayed as a button but tag is input.
* @param strLocation		The web page location value
* @param strElemFullPath	The EditBox full XPath
* @param strElemName		The meaningful name of the EditBox
* @param strAssignValue   The assigned value to Set/Enter and Verify
* @param boolClearValue   Clear the value in the EditBox before Set/Enter? true/false
* @param boolVerifyValue  Verify the value in the EditBox after Set/Enter? true/false
* @param boolTabOutOfField  Tab out of the field after entry? True/False
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 04/20/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyHiddenFilePathValue (strLocation, strElemFullPath, strElemName,
		strAssignValue, boolClearValue, boolVerifyValue, boolTabOutOfField) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'EditBox';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weEditBox;	 // Type inferred
	//let strTestObjText; // Declared but unused

	//Return the element
	weEditBox = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weEditBox != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weEditBox, 'Edit', strElemName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsEditState = {}; // Mimic Groovy Map
		mapResultsEditState = CWCore.objVerifyState(weEditBox, strElemName, true, true);
		//Output results
		let boolEditStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsEditState.boolPassed);
		let strVerEditStateResults = mapResultsEditState.strMethodDetails;
		//SetVerify the values
		if (boolEditStatePassed == true) {
			//let mapResultsEditSet = {}; // Declared but unused
			//Verify is set to false and a separate verify is used
			//Should we do a separate set since an error is thrown if the file is not found.
			try{
				weEditBox.sendKeys(strAssignValue);
			}
			catch (e) { // Catch any exception. Checking specific types inside.
				if (e instanceof FileNotFoundException) { // Simulate check for Java exception type
					boolPassed = false;
					strMethodDetails = "FAILED!!! The FILE PATH: " + strAssignValue + " was NOT FOUND!!!";
				} else if (e.name === 'InvalidArgumentError' || e.name === 'InvalidArgumentException') { // Simulate WebDriver.InvalidArgumentException
					//Check if the file exist (Simulated fs.existsSync)
					// In a real Node.js environment, `fs.existsSync(strAssignValue)` would be used.
					if (fs.existsSync(strAssignValue) == false) {
						strMethodDetails = "FAILED!!! NO FILE EXIST FOR PATH: " + strAssignValue + "!!!";
					}
					else{
						// Note: ExceptionUtils.getRootCauseMessage(iex) requires specific library
						strMethodDetails = 'Exception occurred!!! SEE ERROR: ' + ExceptionUtils.getStackTrace(e); // Simplified
					}
					boolPassed = false;
				}
				else { // Generic exception
					Tester.Message(ExceptionUtils.getStackTrace(e));
					strMethodDetails = 'Exception occurred!!! SEE ERROR STACK TRACE: ' + ExceptionUtils.getStackTrace(e);
					boolPassed = false;
				}
			}
			if (boolPassed == true && boolVerifyValue == true) {
				//Verify is separate and may not match the path. Get right of the last "\" to capture the file name only.
				let strFileSep = "\\";
				let strFileName = StringsAndNumbers.JComm_GetRightTextInString(strAssignValue, strFileSep);
				let strTempDisplayValue = weEditBox.getAttribute('value');
				//Check if the file name is present and pass it.
				//File name must be at the end '%End%'
				let boolInString = StringsAndNumbers.JComm_VerifyTextPresent(strTempDisplayValue, strFileName, '%End%');
				if (boolInString == false) {
					boolPassed = false;
					strMethodDetails = "FAILED!!! Attempted to set the file path to: " + strAssignValue + " but failed!!!" + gblLineFeed +
					" DISPLAYING path: " + strTempDisplayValue;
				}
			}
			//Update step details
			if (boolPassed == true) {
				strMethodDetails = "The location name: " + strLocation + " and element name '" +
						strElemName + "' is set to the specified value of: '" + strAssignValue + "'.";
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strVerEditStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  SetVerifyMultiLineTextBoxValue  -----------------------------------
* Set and verify the assigned value(s) for the EditBox. Supports multiline text input and verification
* @param strLocation		The web page location value
* @param strElemFullPath	The EditBox full XPath
* @param strElemName		The meaningful name of the EditBox
* @param strAssignValue   The assigned value to Set/Enter and Verify
* @param boolClearValue   Clear the value in the EditBox before Set/Enter? true/false
* @param boolVerifyValue  Verify the value in the EditBox after Set/Enter? true/false
* @param boolTabOutOfField  Tab out of the field after entry? True/False
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 05/21/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyMultiLineTextBoxValue (strLocation, strElemFullPath, strElemName,
		strAssignValue, boolClearValue, boolVerifyValue, boolTabOutOfField ) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'EditBox';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weEditBox;	 // Type inferred
	//let strTestObjText; // Declared but unused

	//Return the element
	weEditBox = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weEditBox != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weEditBox, 'Edit', strElemName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsEditState = {}; // Mimic Groovy Map
		mapResultsEditState = CWCore.objVerifyState(weEditBox, strElemName, true, true);
		//Output results
		let boolEditStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsEditState.boolPassed);
		let strVerEditStateResults = mapResultsEditState.strMethodDetails;
		//SetVerify the values
		if (boolEditStatePassed == true) {
			let mapResultsEditSet = {}; // Mimic Groovy Map
			mapResultsEditSet = CWCore.objSetMultiLineTextBoxValue(weEditBox, strElemName, strAssignValue, boolClearValue, boolVerifyValue, boolTabOutOfField);
			let boolEditSetVerPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsEditSet.boolPassed);
			let strEditSetVerResults = mapResultsEditSet.strMethodDetails;
			//Update step details
			if (boolEditSetVerPassed == true) {
				strMethodDetails = "The location name: " + strLocation + " and element name '" +
						strElemName + "' is set to the specified value of: '" + strAssignValue + "'.";
			}
			else
			{
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strElemName + "' is NOT SET TO THE ASSIGNED VALUE!!! see details below:" + gblLineFeed + strEditSetVerResults;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strVerEditStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  SetVerifyEditBoxValueAttribute  -----------------------------------
* Set and verify the assigned value for the EditBox 'Value' attribute using Javascript
* @param strLocation		The web page location value
* @param strEditFullPath	The EditBox full XPath
* @param strEditName		The meaningful name of the EditBox
* @param strValueEditFullPath   The full Xpath for the value edit field
* @param strAssignValue   The assigned value to Set/Enter and Verify
* @param strAssignAttribute   The Attribute to update. NOTE: SET TO GBLNULL TO USE STANDARD SET/EDIT METHOD after clicking the object.
* @param boolVerifyValue  Verify the value in the EditBox after Set/Enter? true/false
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 05/16/2022 JTR-13967 Task and JTR-13524 Module
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyEditBoxValueAttribute (strLocation, strEditFullPath, strEditName,
		strValueEditFullPath, strAssignValue, strAssignAttribute, boolVerifyValue) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//let gblHTMLLF = GVars.GblHTMLLineFeed('Value'); // Declared but not used
	//Create the output values
	// Declare the variables
	let strTestObjType = 'EditBox';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weEditBox;	 // Type inferred
	//let strTestObjText; // Declared but unused
	//Return the element
	weEditBox = CWCore.returnWebElement(strEditFullPath);
	//Process the element
	if (weEditBox != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weEditBox, 'Edit', strEditName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsEditState = {}; // Mimic Groovy Map
		mapResultsEditState = CWCore.objVerifyState(weEditBox, strEditName, true, true);
		//Output results
		let boolEditStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsEditState.boolPassed);
		let strVerEditStateResults = mapResultsEditState.strMethodDetails;
		//SetVerify the values
		if (boolEditStatePassed == true) {
			//Click on the parent edit to activate the value edit field
			weEditBox.click();
			//Return the value edit field
			let weValueEditBox = CWCore.returnWebElement(strValueEditFullPath);
			if (weValueEditBox != null) {
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weValueEditBox, 'Edit', strEditName);
				}
				//Check the element state (Enabled, Visible)
				let mapResultsValueEditState = {}; // Mimic Groovy Map
				mapResultsValueEditState = CWCore.objVerifyState(weValueEditBox, strEditName, true, true);
				//Output results
				let boolValEditStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsValueEditState.boolPassed);
				let strVerValEditStateResults = mapResultsValueEditState.strMethodDetails;
				//Check the value edit field state
				if (boolValEditStatePassed == true) {
					let mapSetEditValue = {}; // Mimic Groovy Map
					//Set the value try using standard edit set first if not then attribute
					if (strAssignAttribute == gblNull) {
						mapSetEditValue = CWCore.objSetEditBoxValue(weValueEditBox, strEditName, strAssignValue, false, boolVerifyValue, false);
					}
					else {
						mapSetEditValue = CWCore.objSetEditBoxAttibuteValue(weValueEditBox, strEditName, strAssignValue, strAssignAttribute, boolVerifyValue);
					}
					let boolValEditSetVerPassed = StringsAndNumbers.JComm_StringToBoolean (mapSetEditValue.boolPassed);
					let strValEditSetVerResults = mapSetEditValue.strMethodDetails;
					//Update step details
					if (boolValEditSetVerPassed == true) {
						strMethodDetails = "The location name: " + strLocation + " and element name '" +
								strEditName + "' is set to the specified value of: '" + strAssignValue + "'.";
					}
					else
					{
						boolPassed = false;
						strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
								strEditName + "' is NOT SET TO THE ASSIGNED VALUE!!! see details below:" + gblLineFeed + strValEditSetVerResults;
					}
				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element Xpath '" +
							strValueEditFullPath + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strVerValEditStateResults;
				}
			}
			else
			{
				boolPassed = false;
				strMethodDetails = "The location '" + strLocation + "' '" + strValueEditFullPath + "' DID NOT RETURN an ELEMENT.";
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strEditName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strVerEditStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strEditName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  VerifyEditBoxValue  -----------------------------------
* Verify the EditBox contains/displays the assigned value
* @param strLocation		The web page location value
* @param strElemFullPath	The EditBox full XPath
* @param strElemName		The meaningful name of the EditBox
* @param strAssignValue   The assigned value to Set/Enter and Verify
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 04/20/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyEditBoxValue (strLocation, strElemFullPath, strElemName,
		strAssignValue) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'EditBox';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weEditBox;	 // Type inferred
	//let strTestObjText; // Declared but unused

	//Return the element
	weEditBox = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weEditBox != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weEditBox, strTestObjType, strElemName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsEditState = {}; // Mimic Groovy Map
		mapResultsEditState = CWCore.objVerifyState(weEditBox, strElemName, true, true);
		//Output results
		let boolEditStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsEditState.boolPassed);
		let strVerEditStateResults = mapResultsEditState.strMethodDetails;
		//SetVerify the values
		if (boolEditStatePassed == true) {
			let mapResultsEditVer = {}; // Mimic Groovy Map
			mapResultsEditVer = CWCore.objVerifyEditBoxValue(weEditBox, strElemName, strAssignValue);
			let boolEditVerPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsEditVer.boolPassed);
			let strEditVerResults = mapResultsEditVer.strMethodDetails;
			//Update step details
			if (boolEditVerPassed == true) {
				strMethodDetails = "The location name: " + strLocation + " and element name '" +
						strElemName + "' is displaying the specified value of: '" + strAssignValue + "'.";
			}
			else
			{
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strElemName + "' is NOT DISPLAYING THE ASSIGNED VALUE!!! see details below:" + gblLineFeed + strEditVerResults;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strVerEditStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  VerifyElementValue  -----------------------------------
* Verify the element values based on object property
* @param strLocation		The web page location value
* @param strElemFullPath	The EditBox full XPath
* @param strElemName		The meaningful name of the EditBox
* @param strObjectType	  The type of object i.e. button, link, etc.
* @param strAssignValue   The assigned value to Set/Enter and Verify
* @param strElementProperty   The element property that contains the value to verify
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 08/03/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyElementValue (strLocation, strElemFullPath, strElemName,
		strObjectType, strAssignValue, strElementProperty) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = strObjectType;
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let we;		 // Type inferred
	//let strTestObjText; // Declared but unused
	//Return the element
	we = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (we != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(we, strTestObjType, strElemName);
		}
		//Verify the element value
		let strTempValue; // Type inferred
		if (strElementProperty == 'Text') {
			strTempValue = we.getText();
		}
		else {
			let boolAttIsPresent; // Type inferred
			boolAttIsPresent = CWCore.isAttribtuePresent(we, strElementProperty);
			if (boolAttIsPresent == true) {
				//Return the value
				strTempValue = we.getAttribute(strElementProperty);
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location '" + strLocation + "' '" + strElemName +
				"' DOES NOT HAVE THE ATTRIBUTE OF: " + strElementProperty + ".";
			}
		}
		if (strTempValue == strAssignValue) {
			strMethodDetails = "The " + strElemName + " value displayed matches the assigned value of: " + strAssignValue + ".";
		}
		else {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The " + strElemName + " value displayed of: " + strTempValue +
			" DOES NOT MATCH the ASSIGNED: " + strAssignValue + "!!!";
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "FAILED!!! The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
* -------------------------------------  VerifyElementMultiLineTextValue  -----------------------------------
* Verify the values displayed in an element as Multiline Text such as labels.
*
* @param mapInputValues	   The map which contains the assigned values.
* @param strLocation		  The web page location value
* @param strElemFullXPath	 The EditBox full XPath
* @param strElemName		  The meaningful name of the EditBox
* @parma strElemType		  The type of object i.e. label, editbox, textbox
* @param strElemLineXPath	 The XPath for the line of text if in a child element. Set to gblNull if not present
* @param strAssignValue   The assigned value verify separated by the gblDelimited value for each line
* @param strLineBreakChar The Line Break Character if present, for example <br>. Set to gblNull if not present
*
* @return mapResults		  The results showing Passed and method details.
*
* Developed for JTR-22008
*
* @author pkanaris
* @author Created: 09/15/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyElementMultiLineTextValue (mapInputValues) { // mapInputValues, return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	//Create the output values
	// Declare the variables
	let strLocation = mapInputValues.strLocation;
	let strElemFullXPath = mapInputValues.strElemFullXPath;
	let strElemName = mapInputValues.strElemName;
	let strElemType = mapInputValues.strElemType;
	let strElemLineXPath = mapInputValues.strElemLineXPath;
	let strAssignValue = mapInputValues.strAssignValue;
	let strLineBreakChar = mapInputValues.strLineBreakChar;
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let we;		 // Type inferred
	let strTempValue = gblNull; // Type inferred
	//Return the element
	we = CWCore.returnWebElement(strElemFullXPath);
	//Process the element
	if (we != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(we, strElemType, strElemName);
		}
		//Determine if we are looking for child elements that contain the value
		if (strElemLineXPath == gblNull && strLineBreakChar != gblNull) {
			//Process the element string using the LineBreakChar
			let strTempText = StringsAndNumbers.JComm_HandleNoData(we.getText());
			//Split the string into an array (Note: Original Groovy had empty split here)
			// Assuming there is a JComm_StringToArray function CommonWeb_that handles line breaks,
			// or a direct split using strLineBreakChar
			// Example: let arryLines = strTempText.split(strLineBreakChar);
			// This part of Groovy code is incomplete, "split the string into an array"
			// If it actually meant to re-evaluate equality with strAssignValue,
			// strTempValue would need to be re-assembled or strTempText directly compared.
			// For now, mirroring only what's explicitly there.
			strTempValue = strTempText; // For direct comparison, or it would be built up.

		}
		else if(strElemLineXPath != gblNull) {
			//Return the children elements and process them
			let lstChildTextLines = we.findElements(By.xpath(strElemLineXPath)); // findElements returns a List
			let intCntTextLines = lstChildTextLines.length; // Use .length for JS Array
			if (intCntTextLines > 0) {
				strTempValue = StringsAndNumbers.JComm_HandleNoData(lstChildTextLines[0].getText()); //Assign value zero since this is the first value
				//Process the list to create the tempValue
				for (let loopLines = 1; loopLines < intCntTextLines; loopLines++) {
					strTempValue = strTempValue + gblDelimiter + StringsAndNumbers.JComm_HandleNoData(lstChildTextLines[loopLines].getText());
				}
			}
		}
		//Compare the strTempValue to the expected.
		if (strAssignValue == strTempValue) {
			//Passed
			strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' correctly displayed the value: '" + strAssignValue + "'.";
		}
		else {
			//Fails
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location '" + strLocation + "' '" + strElemName + "' displayed value: '" + strTempValue +
			"' DOES NOT MATCH THE EXPECTED VALUE: '" + strAssignValue +"'!!!";
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "FAILED!!! The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/* MASKEDIT */
/**
* -------------------------------------  SetVerifyMaskEditBoxValue  -----------------------------------
* Set and verify the assigned value for the MaskedEditBox. Used for sensitive data such as passwords or account information.
* @param strLocation		The web page location value
* @param strElemFullPath	The MaskedEditBox full XPath
* @param strElemName		The meaningful name of the MaskedEditBox
* @param strAssignValue   The assigned value to Set/Enter and Verify
* @param boolValueIsEncrypted   The value Assigned is encrypted? true/false
* @param boolClearValue   Clear the value in the MaskedEditBox before Set/Enter? true/false
* @param boolVerifyValue  Verify the value in the MaskedEditBox after Set/Enter? true/false
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 04/20/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyMaskEditBoxValue (/**string*/strActionName,/**string*/strLocation, /**string*/strLocXPath, /**string*/strElemName,
		/**string*/strAssignValue,  /**boolean*/boolValueIsEncrypted,  /**boolean*/boolClearValue,  /**boolean*/boolVerifyValue) { // Return type any
	if (strAssignValue !== GblSkip || strAssignValue !== GblNotApplicable){ //Process the step only if either is not true
		//GlobalVars
		let gblNull = GblNull;
		let gblUndefined = GblUndefined;
		let gblLineFeed = GblLineFeed;
		//Create the output values
		// Declare the variables
		let strTestObjType = 'MaskedEditBox';
		let strStepAction =  strActionName;
		let strStepDescription = strActionName + ": *****Masked Value*****";
		let strStepExpectedResult = "The assigned value is entered and displayed as an encrypted value.";
		let strMethodDetails;
		let boolPassed = true;
		var weMaskEditBox; // No type is to be inferred
		//Return the Element Xpath and concantenate to the strLocXPath
		let strElemXpath = getORXPath(strElemName);
		if (boolDoDebug === true) {
			Tester.Message('The SetVerifyMaskEditBoxValue.strLocXPath: ' + strLocXPath)
		}
		if (boolDoDebug === true) {
			Tester.Message('The SetVerifyMaskEditBoxValue.Element ' + strElemName + ' XPath is: ' + strElemXpath)
		}
		let strElemFullPath = strLocXPath + strElemXpath
		if (boolDoDebug === true) {
			Tester.Message('The SetVerifyMaskEditBoxValue.FullXPath is: ' + strElemFullPath)
		}
		//Return the element
		weMaskEditBox = CommonWebCore.returnWebElement(strElemFullPath);	
		//Process the element
		if (weMaskEditBox != null) {
			//Highlight
			if (boolDoHighlight == true) {
				CommonWeb.HighlightElement(weMaskEditBox);
			}
			//Check the element state (Enabled, Visible)
			let mapResultsMaskEditState = {}; // Mimic Groovy Map
			mapResultsMaskEditState = CommonWebCore.objVerifyState(weMaskEditBox, strElemName, true, true);
			//Output results
			let boolMaskEditStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsMaskEditState.boolPassed);
			let strVerMaskEditStateResults = mapResultsMaskEditState.strMethodDetails;
			//SetVerify the values
			if (boolMaskEditStatePassed == true) {
				let mapResultsEditSet = {}; // Mimic Groovy Map
				// boolClearValue and boolVerifyValue from input are disregarded because CWCore.objSetMaskEditBoxValue
				// fixes them to true and true respectively.
				mapResultsEditSet = CommonWebCore.objSetMaskEditBoxValue(weMaskEditBox, strElemName, strAssignValue, true, boolValueIsEncrypted, true);
				boolPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsEditSet.boolPassed);
				strMethodDetails = mapResultsEditSet.strMethodDetails;
			}
			else
			{
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strVerMaskEditStateResults;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
		}
		//Update the TestStepResults
		TestStepResults.StepAction = strStepAction;
		TestStepResults.StepDesc = strStepDescription;
		TestStepResults.StepExpected = strStepExpectedResult;
		TestStepResults.StepActual = strMethodDetails;
		TestStepResults.StepData = strElemName + ": " + strAssignValue; 
		TestStepResults.StepPassed = boolPassed;
		TestExecReporting.ReportStepResults();
	}
}
/* Element */
/**
* -------------------------------------  ClickElement  -----------------------------------
* Click the specified element when visible and enabled.
* NOTE: Use this method when a click on elements not ordinarily clicked is required!!!
* @param strLocation		The web page location value
* @param strElemFullPath	The element full XPath
* @param strElemName		The meaningful name of the element
* @param strElemType		The element type (EditBox, Button, etc.)
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 02/07/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_ClickElement (/**string*/strLocation, /**string*/strElemFullPath, /**string*/strElemName, /**string*/strElemType) { 
	//GlobalVars
	let gblNull = GblNull;
	let gblUndefined = GblUndefined;
	let gblLineFeed = GblLineFeed;
	//Create the output values
	// Declare the variables
	let strTestObjType = strElemType;
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weElement;	 // Type inferred
	//Return the element
	weElement = CommonWebCore.returnWebElement(strElemFullPath)
	//Process the element
	if (weElement != null) {
		//Highlight
		if (boolDoHighlight == true) {
			CommonWeb.HighlightElement(weElement);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsElementState = {}; // Mimic Groovy Map
		mapResultsElementState = CommonWebCore.objVerifyState(weElement, strElemName, true, true);
		//Output results
		let boolElementStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsElementState.boolPassed);
		let strVerElementStateResults = mapResultsElementState.strMethodDetails;
		//SetVerify the values
		if (boolElementStatePassed == true) {
			Tester.SuppressReport(true)
			weElement.DoClick();
			Tester.SuppressReport(false)
			strMethodDetails = "The location name: " + strLocation + " and element name '" +
					strElemName + "' " + strElemType + " was clicked.";
			if (boolDoDebug == true) {
				Tester.Message(strMethodDetails);
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strVerElementStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "'" + strElemType + " DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  ClickWebElement  -----------------------------------
* Click the specified web element when visible and enabled.
* NOTE: Use this method when an element is found in a parent object instead of using full XPath!!!
* @param weElement		  The web element to click
* @param strElemName		The meaningful name of the element
* @param strElemType		The element type (EditBox, Button, etc.)
* @param boolCheckElemState		Check the element state? true/false
* @param boolDoStepReporting		Execute the Step Reporting in this function? true/false
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 07/14/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_ClickWebElement (weElement, strElemName, strElemType, /**boolean*/boolCheckElemState, /**boolean*/boolDoStepReporting) { 
	//GlobalVars
	if (boolDoStepReporting == null){
		boolDoStepReporting = true;
	}
	let gblNull = GblNull;
	let gblUndefined = GblUndefined;
	let gblLineFeed = GblLineFeed;
	//Create the output values
	// Declare the variables
	let strTestObjType = strElemType;
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let boolElementStatePassed
	//Process the element
	if (weElement != null) {
		//Highlight
		if (boolDoHighlight == true) {
			CommonWeb.HighlightElement(weElement);
		}
		if (boolCheckElemState == true){
			//Check the element state (Enabled, Visible)
			let mapResultsElementState = {}; // Mimic Groovy Map
			mapResultsElementState = CommonWebCore.objVerifyState(weElement, strElemName, true, true);
			//Output results
			boolElementStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsElementState.boolPassed);
			let strVerElementStateResults = mapResultsElementState.strMethodDetails;
		}else{
			boolElementStatePassed = true //Always true if we do not check the state.
		}			
		//SetVerify the values
		if (boolElementStatePassed == true) {
			Tester.SuppressReport(true)
			weElement.DoClick();
			Tester.SuppressReport(false)
			strMethodDetails = "The web element assigned '" + strElemName + "' " + strElemType + " was clicked.";
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The web element assigned '" +
					strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strVerElementStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The web element assigned '" + strElemName + "'" + strElemType + " WAS NULL!!!";
	}if (boolDoStepReporting == true){
		//Update the TestStepResults
		TestStepResults.StepActual = strMethodDetails;
		TestStepResults.StepData = GblNotApplicable; 
		TestStepResults.StepPassed = boolPassed;
		TestExecReporting.ReportStepResults();
	}else{
		//Update the map
		mapResults.boolPassed = boolPassed.toString();
		mapResults.strMethodDetails = strMethodDetails;
		return mapResults;
	}
}
/* BUTTON */
/**
* -------------------------------------  ClickButton  -----------------------------------
* Click the specified button when visible and enabled.
* @param strActionName		The name of the step action
* @param strElemFullPath	The web page location value
* @param strLocXPath		The location/page Xpath
* @param strElemName		The meaningful name of the element as it is in the object repository script
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 04/20/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_ClickButton (/**string*/strActionName,/**string*/strLocation, /**string*/strLocXPath, /**string*/strElemName, /**boolean*/ boolClickBtn){ // Return type any
	if (boolClickBtn == true){
		//GlobalVars
		let gblNull = GblNull;
		let gblUndefined = GblUndefined;
		let gblLineFeed = GblLineFeed;
		//Create the output values
		// Declare the variables
		let strTestObjType = 'Button';
		let strStepAction =  strActionName;
		let strStepDescription = strActionName + ": " + boolClickBtn;
		let strStepExpectedResult = strElemName + " is selected when true.";
		let strMethodDetails; // Type inferred
		let boolPassed = true;
		let weButton;
		//Return the Element Xpath and concantenate to the strLocXPath
		let strElemXpath = getORXPath(strElemName);
		if (boolDoDebug === true) {
			Tester.Message('The CommonWeb_ClickButton.strLocXPath: ' + strLocXPath)
		}
		if (boolDoDebug === true) {
			Tester.Message('The CommonWeb_ClickButton.Element ' + strElemName + ' XPath is: ' + strElemXpath)
		}
		let strElemFullPath = strLocXPath + strElemXpath
		if (boolDoDebug === true) {
			Tester.Message('The CommonWeb_ClickButton.FullXPath is: ' + strElemFullPath)
		}
		//Return the element
		weButton = CommonWebCore.returnWebElement(strElemFullPath);
		//Process the element
		if (weButton != null) {
			//Move to the element if needed
			let mapMoveToElem = CommonWeb.MoveToWebElement(strLocation, weButton, strElemName); // Call to global func
			//Highlight
			if (boolDoHighlight == true) {
				CommonWeb.HighlightElement(weButton);
			}
			//Check the element state (Enabled, Visible)
			let mapResultsBtnState = {}; // Mimic Groovy Map
			mapResultsBtnState = CommonWebCore.objVerifyState(weButton, strElemName, true, true);
			//Output results
			let boolBtnStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsBtnState.boolPassed);
			let strVerBtnStateResults = mapResultsBtnState.strMethodDetails;
			//Attempt to click the button
			if (boolBtnStatePassed == true) {
				Tester.SuppressReport(true)
				weButton.DoClick();
				Tester.SuppressReport(false)
				strMethodDetails = "The location name: " + strLocation + " and element name '" +
						strElemName + "' was clicked.";
			}
			else
			{
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strVerBtnStateResults;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
		}
		//Update the TestStepResults
		TestStepResults.StepAction = strStepAction;
		TestStepResults.StepDesc = strStepDescription;
		TestStepResults.StepExpected = strStepExpectedResult;
		TestStepResults.StepActual = strMethodDetails;
		TestStepResults.StepData = "boolClickBtn: " + boolClickBtn; 
		TestStepResults.StepPassed = boolPassed;
		TestExecReporting.ReportStepResults();
	}
}

/* LINK */
/**
* -------------------------------------  ClickLINK  -----------------------------------
* Click the specified link when visible and enabled.
* @param strLocation		The web page location value
* @param strElemFullPath	The Button full XPath
* @param strElemName		The meaningful name of the Button
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 12/17/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_ClickLink (strLocation, strElemFullPath, strElemName) { // Return type any
	//GlobalVars
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'Link';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weLink;		// Type inferred
	//Return the element
	weLink = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weLink != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weLink, strTestObjType, strElemName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsLinkState = {}; // Mimic Groovy Map
		mapResultsLinkState = CWCore.objVerifyState(weLink, strElemName, true, true);
		//Output results
		let boolLinkStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsLinkState.boolPassed);
		let strVerLinkStateResults = mapResultsLinkState.strMethodDetails;
		//SetVerify the values
		if (boolLinkStatePassed == true) {
			//Catch any exception for the click
			//Add try catch
			try {
				weLink.click();
				strMethodDetails = "The location name: " + strLocation + " and element name '" +
						strElemName + "' was clicked.";
			}
			catch (eLnkClick) {
				Tester.Message(ExceptionUtils.getStackTrace(eLnkClick));
				strMethodDetails = 'Exception occurred!!! SEE ERROR STACK TRACE: ' + ExceptionUtils.getStackTrace(eLnkClick) +
						" Will try to capture using object XPath";
				boolPassed = false;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strVerLinkStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/* MESSAGE BOX */
/**
* -------------------------------------  VerifyMsgBoxText  -----------------------------------
* Verify the text displayed in the message box.
* @param strActionName		The name of the step action
* @param strLocation		The web page location value
* @param strLocXPath	The location XPath
* @param strElemName		The meaningful name of the Message Box
* @param strMsgText   The assigned text displayed
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 05/03/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyMsgBoxText (strActionName,strLocation, strLocXPath, strElemName, strMsgText) { // Return type any
	if (strMsgText !== GblSkip || strMsgText !== GblNotApplicable){ //Process the step only if either is not true
		//GlobalVars
		let gblNull = GblNull;
		let gblUndefined = GblUndefined;
		let gblLineFeed = GblLineFeed;
		//Create the output values
		// Declare the variables
		let strTestObjType = 'MsgBox';
		let strStepAction =  strActionName;
		let strStepDescription = "Verify the system is displaying the process message specified.   i.e. 'Invalid login, please try again.' '\*N/A*' if no message is to be displayed.";
		let strStepExpectedResult = "The system displays the appropriate process message of: " + strMsgText;
		let strMethodDetails;
		let boolPassed = true;
		let weMsgBox;	  // Type inferred
		//Return the Element Xpath and concantenate to the strLocXPath
		let strElemXpath = getORXPath(strElemName);
		if (boolDoDebug === true) {
			Tester.Message('The SetVerifyEditBoxValue.strLocXPath: ' + strLocXPath)
		}
		if (boolDoDebug === true) {
			Tester.Message('The SetVerifyEditBoxValue.Element ' + strElemName + ' XPath is: ' + strElemXpath)
		}
		let strElemFullPath = strLocXPath + strElemXpath
		if (boolDoDebug === true) {
			Tester.Message('The SetVerifyEditBoxValue.FullXPath is: ' + strElemFullPath)
		}
		//Return the element
		weMsgBox = CommonWebCore.returnWebElement(strElemFullPath);
		//Process the element
		if (weMsgBox != null) {
			//Highlight
			if (boolDoHighlight == true) {
				CommonWeb.HighlightElement(weMsgBox);
				//CommonWeb.HighlightElement(weEditBox);
			}
			//Check the element state (Enabled, Visible)
			let mapResultsMsgBoxState = {}; // Mimic Groovy Map
			mapResultsMsgBoxState = CommonWebCore.objVerifyState(weMsgBox, strElemName, true, true);
			//Output results
			let boolMsgBoxStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsMsgBoxState.boolPassed);
			let strVerMsgBoxStateResults = mapResultsMsgBoxState.strMethodDetails;
			//Verify the value
			if (boolMsgBoxStatePassed == true) {
				let mapResultsVerMsgBox = {}; // Mimic Groovy Map
				mapResultsVerMsgBox = CommonWebCore.objVerifyMsgBoxValue(weMsgBox, strElemName, strMsgText);
				boolPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsVerMsgBox.boolPassed);
				strMethodDetails = mapResultsVerMsgBox.strMethodDetails;
	
			}
			else
			{
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strVerMsgBoxStateResults;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
		}
		//Update the TestStepResults
		TestStepResults.StepAction = strStepAction;
		TestStepResults.StepDesc = strStepDescription;
		TestStepResults.StepExpected = strStepExpectedResult;
		TestStepResults.StepActual = strMethodDetails;
		TestStepResults.StepData = "strMsgText: " + strMsgText; 
		TestStepResults.StepPassed = boolPassed;
		TestExecReporting.ReportStepResults();
	}
}
/* CHECKBOX */
/**
* -------------------------------------  SetVerifyCheckBox  -----------------------------------
* Set and verify the CheckBox to the correct state
* @param strLocation		The web page location value
* @param strElemFullPath	The CheckBox full XPath
* @param strElemName		The meaningful name of the EditBox
* @param boolAssignedChecked   The checkbox is to be checked? true/false
* @param boolVerifyChecked	 Verify the checked state? true/false
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 12/31/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyCheckBox (strLocation, strElemFullPath, strElemName,
		boolAssignedChecked, boolVerifyChecked) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'CheckBox';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weCheckBox;	// Type inferred
	//Return the element
	weCheckBox = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weCheckBox != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weCheckBox, 'CheckBox', strElemName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsWEState = {}; // Mimic Groovy Map
		mapResultsWEState = CWCore.objVerifyState(weCheckBox, strElemName, true, true);
		//Output results
		let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
		let strStateResults = mapResultsWEState.strMethodDetails;
		//SetVerify the values
		if (boolStatePassed == true) {
			let mapResultsCkBoxSet = {}; // Mimic Groovy Map
			mapResultsCkBoxSet = CWCore.objSetCheckBox(weCheckBox, strElemName, boolAssignedChecked);
			let strSetChkBoxResults = mapResultsCkBoxSet.strMethodDetails;
			let weNewCkBox = CWCore.returnWebElement(strElemFullPath); //Return the checkbox after checking to reduce risk of stale objects.
			//Update step details
			if (weNewCkBox == null) {
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strElemName + "' was not returned from the setcheckbox method.";
				boolPassed = false;
			}
			else
			{
				strMethodDetails = strSetChkBoxResults;
				if (boolVerifyChecked == true){
					let boolVerChkPassed = CWCore.objVerifyCheckBoxChecked(weNewCkBox, boolAssignedChecked);
					if (boolVerChkPassed == true) {
						strMethodDetails = "The checkbox '" + strElemName + "' was correctly set to the checked state of: " + boolAssignedChecked + ".";
					}
					else {
						strMethodDetails = "The checkbox '" + strElemName + "' was NOT SET to the checked state of: " + boolAssignedChecked + "!!!";
						boolPassed = false;
					}
				}
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  SetVerifyCheckBoxWebElement  -----------------------------------
* Set and verify the CheckBox to the correct state
* @param strLocation			   The web page location value
* @param WebElement weCheckBox   The CheckBox webelement
* @param strElemName			   The meaningful name of the EditBox
* @param boolAssignedChecked	   The checkbox is to be checked? true/false
* @param boolVerifyChecked		 Verify the checked state? true/false
* @return mapResults			   The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 12/31/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyCheckBoxWebElement (strLocation, weCheckBox, strElemName,
		boolAssignedChecked, boolVerifyChecked) { // weCheckBox is `any`, return `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'CheckBox';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	// Process the element
	if (weCheckBox != null) { // Note: strTestObjText is declared but unused in Groovy
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weCheckBox, 'CheckBox', strElemName); // Check strTestObjType usage
		}
		//Check the element state (Enabled, Visible)
		let mapResultsWEState = {}; // Mimic Groovy Map
		mapResultsWEState = CWCore.objVerifyState(weCheckBox, strElemName, true, true);
		//Output results
		let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
		let strStateResults = mapResultsWEState.strMethodDetails;
		//SetVerify the values
		if (boolStatePassed == true) {
			let mapResultsCkBoxSet = {}; // Mimic Groovy Map
			mapResultsCkBoxSet = CWCore.objSetCheckBox(weCheckBox, strElemName, boolAssignedChecked);
			let strSetChkBoxResults = mapResultsCkBoxSet.strMethodDetails;
			let weNewCkBox = mapResultsCkBoxSet.webElement; //Return the checkbox after checking to reduce risk of stale objects.
			//Update step details
			if (weNewCkBox == null) {
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strElemName + "' was not returned from the setcheckbox method.";
				boolPassed = false;
			}
			else
			{
				strMethodDetails = strSetChkBoxResults;
				if (boolVerifyChecked == true){
					let boolVerChkPassed = CWCore.objVerifyCheckBoxChecked(weNewCkBox, boolAssignedChecked);
					if (boolVerChkPassed == true) {
						strMethodDetails = "The checkbox '" + strElemName + "' was correctly set to the checked state of: " + boolAssignedChecked + ".";
					}
					else {
						strMethodDetails = "The checkbox '" + strElemName + "' was NOT SET to the checked state of: " + boolAssignedChecked + "!!!";
						boolPassed = false;
					}
				}
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  SetVerifyCheckBoxGroupItems  -----------------------------------
* Set and verify the CheckBox(es) in the checkbox group to the assigned state
* @param mapChkBoxGroupData		The map containing all of the values below
* @param strPageLoc 				The web page location value
* @param strPageXpath				The web page XPath
* @param strChkBoxGrpName 			The meaningful name of the Checkbox group
* @param strChkboxGrpXPath			The checkbox group Xpath
* @param strChkboxItemXPath		The generic checkbox XPath that will find all checkboxes in the group
* @param strChkboxItemLblXPath		The generic checkbox label XPath for all labels in the checkbox group
* @param strChkBoxGroupData		The data including the label,isChecked pair for each item in a delimited string
* @param boolChkBoxGroupMatchAll	Match the order when true or match the item data as long as input count is less than or equal to group checkbox count
* @param boolChkboxVerifySetValue 	Verify the checked state at selection? true/false
* @return mapResults 			The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 12/31/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyCheckBoxGroupItems (mapChkBoxGroupData) { // mapChkBoxGroupData is `any`, return `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	// Declare the variables
	let strTestObjType = 'ChkboxGroup';
	let strChkboxGrpFullXpath;
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	//Return the values from the map
	let strPageLoc = mapChkBoxGroupData.strPageLoc;
	let strPageXPath = mapChkBoxGroupData.strLocXPath;
	let strChkBoxGrpName = mapChkBoxGroupData.strCkkboxGrpName;
	let strChkboxGrpXPath = mapChkBoxGroupData.strChkboxGroupXPath;
	let strChkboxSelItemXPath = mapChkBoxGroupData.strChkboxSelItemXPath; // Added in this section
	let strChkboxItemXPath = mapChkBoxGroupData.strChkboxItemXPath;
	let strChkboxItemLblXPath = mapChkBoxGroupData.strChkBoxItemLblXPath;
	let strChkBoxGroupData = mapChkBoxGroupData.strChkboxGroupInputData;
	let boolChkBoxGroupMatchAll = mapChkBoxGroupData.boolChkBoxGroupMatchAll;
	let boolChkboxVerifySetValue = mapChkBoxGroupData.boolChkBoxVerifySetValue;
	//Return the element
	let weChkboxGroup;
	if (strPageXPath == gblNull || strPageXPath == null) {
		strChkboxGrpFullXpath = strChkboxGrpXPath;
	}
	else {
		strChkboxGrpFullXpath = strPageXPath + strChkboxGrpXPath;
	}
	weChkboxGroup = CWCore.returnWebElement(strChkboxGrpFullXpath);
	//Process the element
	if (weChkboxGroup != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weChkboxGroup, strTestObjType, strChkBoxGrpName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsWEState = {}; // Mimic Groovy Map
		mapResultsWEState = CWCore.objVerifyState(weChkboxGroup, strChkBoxGrpName, true, true);
		//Output results
		let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
		let strStateResults = mapResultsWEState.strMethodDetails;
		//SetVerify the values
		if (boolStatePassed == true) {
			let intCntChkboxSelectors;
			let intCntChkboxes;
			let intCntChkboxLabels;
			/**
			List<WebElement> lstChkboxSelectors = weChkboxGroup.findElements(By.xpath(strChkboxSelItemXPath))
			intCntChkboxSelectors = lstChkboxSelectors.size()
			List<WebElement> lstChkboxes = weChkboxGroup.findElements(By.xpath(strChkboxItemXPath))
			intCntChkboxes = lstChkboxSelectors.size()
			List<WebElement> lstChkboxLabels = weChkboxGroup.findElements(By.xpath(strChkboxItemLblXPath))
			intCntChkboxLabels = lstChkboxSelectors.size()
			*/
			//Issues with the Checkbox Group not returning only the child objects required using the driver and full XPath OGK 03/14/2022
			let lstChkboxSelectors = TCObj.tcDriver.findElements(By.xpath(strChkboxGrpFullXpath + strChkboxSelItemXPath)); // Using tcDriver
			intCntChkboxSelectors = lstChkboxSelectors.length; // Use .length for JS Array
			let lstChkboxes = TCObj.tcDriver.findElements(By.xpath(strChkboxGrpFullXpath + strChkboxItemXPath)); // Using tcDriver
			intCntChkboxes = lstChkboxes.length; // Use .length for JS Array
			let lstChkboxLabels = TCObj.tcDriver.findElements(By.xpath(strChkboxGrpFullXpath + strChkboxItemLblXPath)); // Using tcDriver
			intCntChkboxLabels = lstChkboxLabels.length; // Use .length for JS Array
			if (intCntChkboxSelectors == intCntChkboxes && intCntChkboxSelectors == intCntChkboxLabels) {
				//Split the input value to begin the processing
				let mapCheckboxItemsValues = StringsAndNumbers.JComm_StringToArray(strChkBoxGroupData, gblDelimiter);
				let arryCheckboxItemsValues = mapCheckboxItemsValues.ArryOfValues;
				let intCntCheckboxItemsValues = arryCheckboxItemsValues.length; // Use .length
				if (intCntChkboxSelectors == intCntCheckboxItemsValues && boolChkBoxGroupMatchAll == true) {
					strMethodDetails = "The user has provided input values with " + intCntCheckboxItemsValues + " label,checked pairs that matches the count of Checkbox items and must match order.";
				}
				else if (intCntChkboxSelectors == intCntCheckboxItemsValues && boolChkBoxGroupMatchAll == false) {
					strMethodDetails = "The user has provided input values with " + intCntCheckboxItemsValues + " label,checked pairs that matches the count of Checkbox items and may be in a different order.";
				}
				else if (intCntChkboxSelectors != intCntCheckboxItemsValues && boolChkBoxGroupMatchAll == false) {
					strMethodDetails = "The user has provided input values with " + intCntCheckboxItemsValues + " label,checked pairs that does not match the count of " +
					intCntChkboxSelectors + " Checkbox items. Match values is false so we will set checked state on item where the label matches.";
				}
				else {
					//Fail
					boolPassed = false;
					strMethodDetails = "FAILED!!! The system is displaying " + intCntChkboxSelectors + " checkboxes and the user provided " + // Corrected count
					intCntCheckboxItemsValues + " label,checked pairs WHICH DOES NOT MATCH and the USER SPECIFIED MATCH VALUES AS TRUE!!!";
				}
				if (boolPassed == true) {
					//Process the input values
					let boolLabelMatch;
					let boolIsChecked;
					let strTempValPair;
					let strAssgLabel;
					let strTempLabel;
					let weTempLabel;
					strMethodDetails = "Processing the Checkbox Group '" + strChkBoxGrpName + "' results.";
					for (let loopInput = 0; loopInput < intCntCheckboxItemsValues; loopInput++) {
						strTempValPair = arryCheckboxItemsValues[loopInput]; // Use array indexing
						boolLabelMatch = false;
						//return the label text and checked state
						//needed to have a way for commas to be present in the value use "\,"
						if (strTempValPair.includes("/,") == true) { // Use .includes
							let tempValue = strTempValPair.replace("/,", "/-/");
							strAssgLabel = StringsAndNumbers.JComm_GetLeftTextInString(tempValue, ",");
							boolIsChecked = StringsAndNumbers.JComm_StringToBoolean(StringsAndNumbers.JComm_GetRightTextInString(tempValue, ","));
							strAssgLabel = strAssgLabel.replace("/-/", ",");
						}
						else {
							strAssgLabel = StringsAndNumbers.JComm_GetLeftTextInString(strTempValPair, ",");
							boolIsChecked = StringsAndNumbers.JComm_StringToBoolean(StringsAndNumbers.JComm_GetRightTextInString(strTempValPair, ","));
						}
						if (boolChkBoxGroupMatchAll == true) {
							//return the label for the checkbox that matches the loopInput and check if match
							weTempLabel = lstChkboxLabels[loopInput]; // Use array indexing
							//Highlight
							if (TCExecParams.getBoolDoHighlight() == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weTempLabel, 'Label', strChkBoxGrpName);
							}
							strTempLabel = weTempLabel.getText();
							if (strTempLabel == strAssgLabel) {
								//Attempt to set the item
								//Put the values into a map
								let mapChkBoxInput = {}; // Mimic Groovy Map
								mapChkBoxInput.strLocation = strPageLoc;
								mapChkBoxInput.weCheckSelector = lstChkboxSelectors[loopInput]; // Use array indexing
								mapChkBoxInput.weCheckbox = lstChkboxes[loopInput]; // Use array indexing
								mapChkBoxInput.strAssocText = strAssgLabel;
								mapChkBoxInput.boolAssignedChecked = boolIsChecked;
								mapChkBoxInput.boolVerifyChecked = boolChkboxVerifySetValue;
								let mapSetVerifyChkbox = {}; // Mimic Groovy Map
								mapSetVerifyChkbox = SetVerifyGraphicCheckBox(mapChkBoxInput); // Call to global func
								//Check the results and mark pass or fail
								boolPassed = StringsAndNumbers.JComm_StringToBoolean (mapSetVerifyChkbox.boolPassed);
								//Update the results
								strMethodDetails = strMethodDetails + gblLineFeed + mapSetVerifyChkbox.strMethodDetails;
								// break; // Original had break here, but it should probably continue processing input loop
							}
							else {
								boolPassed = false;
								strMethodDetails = strMethodDetails + "FAILED!!! The Label value for item: " + loopInput +
								" of: '" + strTempLabel + "' DOES NOT MATCH THE ASSIGNED VALUE OF: '" + strAssgLabel + "'!!!";
								break; // Break input for loop
							}
						}
						else {
							//Loop through the checkbox item labels to find a match
							for (let loopLabels = 0; loopLabels < intCntChkboxLabels; loopLabels++) {
								weTempLabel = lstChkboxLabels[loopLabels]; // Use array indexing
								//Highlight
								if (TCExecParams.getBoolDoHighlight() == true) {
									let mapHighlight = {}; // Mimic Groovy Map
									mapHighlight = CWCore.objHighlightElementJS(weTempLabel, 'Label', strChkBoxGrpName);
								}
								strTempLabel = weTempLabel.getText();
								if (strTempLabel == strAssgLabel) {
									boolLabelMatch = true;
									//Attempt to set the item
									//Put the values into a map
									let mapChkBoxInput = {}; // Mimic Groovy Map
									mapChkBoxInput.strLocation = strPageLoc;
									mapChkBoxInput.weCheckSelector = lstChkboxSelectors[loopLabels]; // Use array indexing
									mapChkBoxInput.weCheckbox = lstChkboxes[loopLabels]; // Use array indexing
									mapChkBoxInput.strAssocText = strAssgLabel;
									mapChkBoxInput.boolAssignedChecked = boolIsChecked;
									mapChkBoxInput.boolVerifyChecked = boolChkboxVerifySetValue;
									let mapSetVerifyChkbox = {}; // Mimic Groovy Map
									mapSetVerifyChkbox = SetVerifyGraphicCheckBox(mapChkBoxInput); // Call to global func
									//Check the results and mark pass or fail
									boolPassed = StringsAndNumbers.JComm_StringToBoolean (mapSetVerifyChkbox.boolPassed);
									//Update the results
									strMethodDetails = strMethodDetails + gblLineFeed + mapSetVerifyChkbox.strMethodDetails;
									break; // Break from loopLabels
								}
							}
							if (boolPassed == false) {
								break; //Exit the input loop
							}
							else if (boolLabelMatch == false) {
								boolPassed = false;
								strMethodDetails = strMethodDetails + "FAILED!!! The Label value assigned of: " + strAssgLabel +
								" DID NOT MATCH ANY OF THE DISPLAYED CHECKBOXES";
								break; // Break input loop
							}
						}
						if (boolPassed == false) {
							break; //Exit the input loop
						}
					}
				}
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location name: " + strPageLoc + " and element name '" +
				strChkBoxGrpName + " COUNT of child elements DID NOT MATCH!!!" + gblLineFeed +
				"intCntChkboxSelectors: " + intCntChkboxSelectors + gblLineFeed +
				"intCntChkboxes: " + intCntChkboxes + gblLineFeed +
				"intCntChkboxLabels: " + intCntChkboxLabels;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strPageLoc + " and element name '" +
					strChkBoxGrpName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strPageLoc + "' '" + strChkBoxGrpName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------   SetVerifyCheckBoxJavaScript   -----------------------------------
* Set and verify the CheckBox to the correct state using JavaScript.
* NOTE: Use this method when others have failed
* @param strLocation 			The web page location value
* @param strElemFullPath 		The CheckBox full XPath
* @param strElemName 			The meaningful name of the EditBox
* @param boolAssignedChecked	The checkbox is to be checked? true/false
* @param boolVerifyChecked		Verify the checked state? true/false
* @return mapResults 			The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 12/31/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyCheckBoxJavaScript (strLocation, strElemFullPath, strElemName,
		boolAssignedChecked, boolVerifyChecked){
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'CheckBox';
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails;
	let boolPassed = true;
	let weCheckBox;
	// let strTestObjText; // Declared but unused in original Groovy

	//Return the element
	weCheckBox = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weCheckBox != null) {
		let mapResultsCkBoxSet = {}; // Mimic Groovy Map
		mapResultsCkBoxSet = CWCore.objSetCheckBoxJavaScript(weCheckBox, strElemName, boolAssignedChecked);
		let strSetChkBoxResults = mapResultsCkBoxSet.strMethodDetails;
		let weNewCkBox = mapResultsCkBoxSet.webElement; //Return the checkbox after checking to reduce risk of stale objects.
		//Update step details
		if (weNewCkBox == null) {
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemName + "' was not returned from the setcheckbox method.";
			boolPassed = false;
		}
		else
		{
			strMethodDetails = strSetChkBoxResults;
			if (boolVerifyChecked == true){
				let boolVerChkPassed = CWCore.objVerifyCheckBoxChecked(weNewCkBox, boolAssignedChecked);
				if (boolVerChkPassed == true) {
					strMethodDetails = "The checkbox '" + strElemName + "' was correctly set to the checked state of: " + boolAssignedChecked  + ".";
				}
				else {
					strMethodDetails = "The checkbox '" + strElemName + "' was NOT SET to the checked state of: " + boolAssignedChecked  + "!!!";
					boolPassed = false;
				}
			}
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString(); // Using direct assignment over `put`
	mapResults.strMethodDetails = strMethodDetails; // Ensure this is set for all paths
	return mapResults;
}
/**
* -------------------------------------   SetVerifyGraphicCheckBox   -----------------------------------
* Set and verify the Graphic CheckBox to the correct state. Unchecked visual is a checkbox checked is a check mark image
* The input type checkbox state will provide the checked and unchecked state.
* @param mapGraphicCheckbox 	Contains the object information for the form check and the checkbox and checked state
* @param strLocation 			The web page location value
* @param weCheckSelector 		The webelement that must be clicked to change the state
* @param weCheckbox			The checkbox element that will reflect the checked state
* @param strAssocText 			The text associated if any to the checkbox
* @param boolAssignedChecked	The checkbox is to be checked? true/false
* @param boolVerifyChecked		Verify the checked state? true/false
* @return mapResults 			The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 12/31/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyGraphicCheckBox (mapGraphicCheckbox){
	//GlobalVars
	let gblNull = GVars.gblNull('Value');
	let gblUndefined = GVars.gblUndefined('Value');
	let gblLineFeed = GVars.gblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails;
	let boolPassed = true;
	//Assign the map values to variables
	let strLocation = mapGraphicCheckbox.strLocation;
	let weCheckSelector = mapGraphicCheckbox.weCheckSelector;
	let weCheckbox = mapGraphicCheckbox.weCheckbox;
	let strAssocText = mapGraphicCheckbox.strAssocText;
	let boolAssignedChecked = mapGraphicCheckbox.boolAssignedChecked;
	let boolVerifyChecked = mapGraphicCheckbox.boolVerifyChecked;
	// let strTestObjText; // Declared but unused in original Groovy

	//Create actions to allow moving to an object in case it is not displayed.
	let actions = new Actions(TCObj.getTcDriver());
	//Process the element
	if (weCheckSelector != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weCheckSelector, 'CheckBoxSelector', strAssocText);
		}
		actions.moveToElement(weCheckSelector); //Move to the element so we have them on screen
		//Check the element state (Enabled, Visible)
		let mapResultsWEState = {}; // Mimic Groovy Map
		mapResultsWEState = CWCore.objVerifyState(weCheckSelector, strAssocText, true, true);
		//Output results
		let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
		let strStateResults = mapResultsWEState.strMethodDetails;
		//SetVerify the values
		if (boolStatePassed == true) {
			let mapResultsCkBoxSet = {}; // Mimic Groovy Map
			mapResultsCkBoxSet = CWCore.objSetGraphicCheckBox(weCheckSelector, weCheckbox, strAssocText, boolAssignedChecked);
			let strSetChkBoxResults = mapResultsCkBoxSet.strMethodDetails;
			let weNewCkBox = mapResultsCkBoxSet.webElement; //Return the checkbox after checking to reduce risk of stale objects.
			//Update step details
			if (weNewCkBox == null) {
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strAssocText + "' was not returned from the setcheckbox method.";
				boolPassed = false;
			}
			else
			{
				strMethodDetails = strSetChkBoxResults;
				if (boolVerifyChecked == true){
					let boolVerChkPassed = CWCore.objVerifyCheckBoxChecked(weCheckbox, boolAssignedChecked);
					if (boolVerChkPassed == true) {
						strMethodDetails = "The checkbox '" + strAssocText + "' was correctly set to the checked state of: " + boolAssignedChecked  + ".";
					}
					else {
						strMethodDetails = "The checkbox '" + strAssocText + "' was NOT SET to the checked state of: " + boolAssignedChecked  + "!!!";
						boolPassed = false;
					}
				}
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strAssocText + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strAssocText + "' WAS NOT DEFINED IN THE CALLING TEST!!!";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString(); // Using direct assignment over `put`
	mapResults.strMethodDetails = strMethodDetails; // Ensure this is set for all paths
	return mapResults;
}
/* Radio Group */
/**
* -------------------------------------  SetVerifyRadioGroupOption  -----------------------------------
* Set and verify the Radio Group option to the specified state of selected
* @param strLocation				   The web page location value
* @param strElemRadioGroupFullPath	 The Radio Group full XPath
* @param strElemRadioGroupName		 The meaningful name of the Radio Group
* @param strElemRadioBtnXPath			The Xpath for the radio button/option
* @param strOptionToSelect			The Radio Button/Option to select
* @param boolVerifySeleted			Verify the selected state? true/false
* @return mapResults				   The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 02/12/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyRadioGroupOption (strLocation, strElemRadioGroupFullPath, strElemRadioGroupName,
		strElemRadioOptXPath, strElemRadioOptBtnXPath, strOptionToSelect, boolVerifySelected) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'RadioGroup';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weRadioGroup;  // Type inferred
	//let strTempOptionText; // Declared but unused

	//Return the element
	weRadioGroup = CWCore.returnWebElement(strElemRadioGroupFullPath);
	//Process the element
	if (weRadioGroup != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weRadioGroup, strTestObjType, strElemRadioGroupName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsWEState = {}; // Mimic Groovy Map
		mapResultsWEState = CWCore.objVerifyState(weRadioGroup, strElemRadioGroupName, true, true);
		//Output results
		let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
		let strStateResults = mapResultsWEState.strMethodDetails;
		//SetVerify the values
		if (boolStatePassed == true) {
			//Set the Radio Group option
			let mapSetRadioGrpOption = {}; // Mimic Groovy Map
			mapSetRadioGrpOption = CWCore.objSetRadioGroupOption(weRadioGroup, strElemRadioGroupName, strElemRadioOptXPath, strElemRadioGroupFullPath, strElemRadioOptBtnXPath, strOptionToSelect, boolVerifySelected);
			boolPassed = StringsAndNumbers.JComm_StringToBoolean (mapSetRadioGrpOption.boolMethodPassed);
			strMethodDetails = mapSetRadioGrpOption.strMethodDetails;
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemRadioGroupName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemRadioGroupName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  SetVerifyCustomRadioGroupOption  -----------------------------------
* Set and verify the Radio Group option to the specified state of selected using JavaGroup for Selection as Radio button have ::before and ::after
* Base Example is JTR-12580 Keyword created against JTR-14690
* @param mapMethodInfo			The map which contains the following values:
* @param strStartLoc			 The web page location value
* @param strRadioGrpParentName			The parent object name that contains the option
* @param strRadioGrpParentXpath		   The Xpath for the parent object
* @param strOptionName			The name of the option to select
* @param strElemRadioOptXpath			 The Radio option Xpath
* @param strElemRadioOptValueXpath	 The element xpath under the OptXpath that contain the option value.
* @param strElemRadioBtnXpath			The Xpath for the radio button/option
* @param strBoolUseJavaScript			Use Java Script to select the option based on psuedo code for the label as in ::before and ::after
* @param strElemRadioOptSelXpath		The Xpath for the element that reflects the selected state
* @param strElemValues			Contains a delimited set of values for the OptionParentProperty|OptionParentPropertyValue|LabelOptionText
* @param boolVerifySelected			Verify the selected state? true/false
* @return mapResults				   The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 06/12/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyCustomRadioGroupOption (mapMethodInfo) { // mapMethodInfo is `any`, return `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'RadioGroupParent';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weRDOParent;   // Type inferred
	//let strTempOptionText; // Declared but unused
	//Return the values from the map
	let strLocation = mapMethodInfo.strStartLoc;
	let strRadioGrpParentName = mapMethodInfo.strRadioGrpParentName;
	let strRadioGrpParentXpath = mapMethodInfo.strRadioGrpParentXpath;
	let strElemRadioOptXpath = mapMethodInfo.strElemRadioOptXpath;
	let strElemRadioOptValueXpath = mapMethodInfo.strElemRadioOptValueXpath;
	let strElemRadioOptBtnXpath = mapMethodInfo.strElemRadioOptBtnXpath;
	let boolUseJavaScript = mapMethodInfo.boolUseJavaScript;
	let strElemRadioOptSelXpath = mapMethodInfo.strElemRadioOptSelXpath;
	let strElemValues = mapMethodInfo.strElemValues;
	let boolVerifySelected = mapMethodInfo.boolVerifySelected;
	//Split strElemValues
	let mapSplitValues = {}; // Mimic Groovy Map
	mapSplitValues = StringsAndNumbers.JComm_StringToArray(strElemValues, GVars.GblDelimiter('Value'));
	if (StringsAndNumbers.JComm_StringToInteger(mapSplitValues.intItemCount) == 3){
		let strOptParentProperty; // Type inferred
		let strOptParentPropertyValue; // Type inferred
		let strOptionItemText; // Type inferred
		let arryElemValues = mapSplitValues.ArryOfValues; // Inferred to array
		strOptParentProperty = StringsAndNumbers.JComm_HandleNoData(arryElemValues[0]); // Access array by index
		strOptParentPropertyValue = StringsAndNumbers.JComm_HandleNoData(arryElemValues[1]);
		strOptionItemText = StringsAndNumbers.JComm_HandleNoData(arryElemValues[2]);
		//Return the element
		weRDOParent = CWCore.returnWebElement(strRadioGrpParentXpath);
		//Process the element
		if (weRDOParent != null) {
			//Highlight
			if (TCExecParams.getBoolDoHighlight() == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weRDOParent, strTestObjType, strRadioGrpParentName);
			}
			//Check the element state (Enabled, Visible)
			let mapResultsWEState = {}; // Mimic Groovy Map
			mapResultsWEState = CWCore.objVerifyState(weRDOParent, strRadioGrpParentName, true, true);
			//Output results
			let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
			let strStateResults = mapResultsWEState.strMethodDetails;
			//SetVerify the values
			if (boolStatePassed == true) {
				//Return the options for the RDOParent
				let lstChildRDOObjects = weRDOParent.findElements(By.xpath(strElemRadioOptXpath)); // findElements returns a List
				let intCntRDOChildObj = lstChildRDOObjects.length; // Use .length for JS Array
				if (intCntRDOChildObj > 0) {
					let weRDOChild;	 // Type inferred
					let strRDOChildPropValue; // Type inferred
					let boolValMatch = false; // Type inferred
					//Find the radio option by the strElemRadioOptXPath, property name and value
					for (let intRDOChildItem = 0; intRDOChildItem < intCntRDOChildObj; intRDOChildItem ++) {
						weRDOChild = lstChildRDOObjects[intRDOChildItem]; // Access array by index
						//Highlight
						if (TCExecParams.getBoolDoHighlight() == true) {
							let mapHighlight = {}; // Mimic Groovy Map
							mapHighlight = CWCore.objHighlightElementJS(weRDOChild, 'Radio Option', strRadioGrpParentName);
						}
						// Check if the attribute is present *before* getting it.
						if (strOptParentProperty !== 'text' && CWCore.isAttribtuePresent(weRDOChild, strOptParentProperty) == true) { // Added 'text' check and isAttributePresent
							strRDOChildPropValue = StringsAndNumbers.JComm_HandleNoData(weRDOChild.getAttribute(strOptParentProperty));
							//Check if the value matches
							if (strRDOChildPropValue == strOptParentPropertyValue) {
								boolValMatch = true;
								break;
							}
						}
						else if(strOptParentProperty == gblNA || strOptParentProperty === 'text') { // If property is not applicable or it's text
							//Return the option text
							let lstChildOptionValueElem = weRDOChild.findElements(By.xpath(strElemRadioOptValueXpath));
							let intCntRDOValueChildObj = lstChildOptionValueElem.length; // Use .length for JS Array
							if (intCntRDOValueChildObj == 1) {
								let strTempOptText = lstChildOptionValueElem[0].getText(); // Access array by index
								//Check if the value matches
								if (strOptionItemText == strTempOptText) {
									boolValMatch = true;
									break;
								}
							} else { // Handle case where value element is not found or not unique
								boolPassed = false;
								strMethodDetails = `FAILED!!! Radio Option Item contained ${intCntRDOValueChildObj} value elements, expected 1.`;
								break;
							}
						}
						else {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The Radio Option Item specified DOES NOT HAVE THE ATTRIBUTE '" + strOptParentProperty + "'!!!";
							break;
						}
					}
					if (boolValMatch == true && boolPassed == true) { // Added boolPassed check
						//Find the child item that contain the text and verify value
						let lstChildOptionValueElem = weRDOChild.findElements(By.xpath(strElemRadioOptValueXpath));
						let intCntRDOValueChildObj = lstChildOptionValueElem.length; // Use .length for JS Array
						if (intCntRDOValueChildObj == 1) {
							let weChildOptText = lstChildOptionValueElem[0]; // Access array by index
							//Highlight
							if (TCExecParams.getBoolDoHighlight() == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weChildOptText, 'Radio Option Text', strRadioGrpParentName);
							}
							//Check the value for a match
							let strOptionText = StringsAndNumbers.JComm_HandleNoData(weChildOptText.getText());
							if (strOptionText == strOptionItemText) {
								//Select/Click the button
								//Return the child selected item and return the value
								let lstChildOptionSelElem = weRDOChild.findElements(By.xpath(strElemRadioOptSelXpath));
								let intCntRDOSelectChildObj = lstChildOptionSelElem.length; // Use .length for JS Array
								if (intCntRDOSelectChildObj == 1) {
									let weChildObjOptionSel = lstChildOptionSelElem[0]; // Access array by index
									//Highlight
									if (TCExecParams.getBoolDoHighlight() == true) {
										let mapHighlight = {}; // Mimic Groovy Map
										mapHighlight = CWCore.objHighlightElementJS(weChildObjOptionSel, 'Radio Option Selected', strRadioGrpParentName);
									}
									//Return the selected state
									// `getAttribute("checked")` usually returns "true" or null when not a boolean `checked` prop is used.
									// Or, for `input[type='radio']`, it's a boolean property.
									let boolSelected = weChildObjOptionSel.getAttribute("checked"); // Will be "true" or null if not checked, need to coerce.
									if (typeof boolSelected === 'string' && boolSelected.toLowerCase() === 'true') {
										boolSelected = true;
									} else {
										boolSelected = false; // Convert to boolean
									}

									if (boolSelected == false) {
										//If not selected, return the button element
										let lstChildOptionBtnElem = weRDOChild.findElements(By.xpath(strElemRadioOptBtnXpath));
										let intCntRDOBtnChildObj = lstChildOptionBtnElem.length; // Use .length for JS Array
										if (intCntRDOBtnChildObj == 1) {
											//Click the button element or use Javascript to click the element
											let weOptBtn = lstChildOptionBtnElem[0]; // Access array by index
											//Highlight
											if (TCExecParams.getBoolDoHighlight() == true) {
												let mapHighlight = {}; // Mimic Groovy Map
												mapHighlight = CWCore.objHighlightElementJS(weOptBtn, 'Radio Option Button', strRadioGrpParentName);
											}
											if (boolUseJavaScript == true) {
												// Casts to `any` to allow direct execution.
												TCObj.tcDriver.executeScript("arguments[0].click();", weOptBtn);
											}
											else {
												weOptBtn.click();
											}
											if (boolVerifySelected == true) {
												// re-check selected state after click
												let boolIsChecked = weChildObjOptionSel.getAttribute("checked");
												if (typeof boolIsChecked === 'string' && boolIsChecked.toLowerCase() === 'true') {
													boolIsChecked = true;
												} else {
													boolIsChecked = false;
												}

												if (boolIsChecked == true) {
													strMethodDetails = "The '" + strRadioGrpParentName + "' Option '" + strOptionItemText +"' was clicked and is in the selected state.";
												}
												else {
													boolPassed = false;
													strMethodDetails = "FAILED!!! The '" + strRadioGrpParentName + "' Option '" + strOptionItemText +"' IS NOT IN THE SELECTED STATE!!!";
												}
											}
										}
										else {
											boolPassed = false;
											strMethodDetails = "FAILED!!! The Radio Option Item specified contained '" + intCntRDOBtnChildObj + "' button element(s), THAT DOES NOT MATCH THE '1' count REQUIRED!!!";
										}
									}
									else {
										strMethodDetails = "The option '" + strOptionItemText + "' is already selected.";
									}
								}
								else {
									boolPassed = false;
									strMethodDetails = "FAILED!!! The Radio Option Item specified contained '" + intCntRDOSelectChildObj + "' select elements, THAT DOES NOT MATCH THE '1' count REQUIRED!!!";
								}
							}
							else {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The Radio Option Text value of: '" + strOptionText + "' DOES NOT MATCH THE EXPECTED VALUE: '" + strOptionItemText + "'!!!";
							}
						}
						else {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The Radio Option Item specified contained '" + intCntRDOValueChildObj + "' value elements, THAT DOES NOT MATCH THE '1' count REQUIRED!!!";
						}
					}
				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The Radio Options Parent: " + strRadioGrpParentName + " DID NOT RETURN any OPTION ELEMENTS!!!";
				}
			}
			else
			{
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strRadioGrpParentName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location '" + strLocation + "' '" + strRadioGrpParentName + "' DID NOT RETURN an ELEMENT.";
		}
	}
	else {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The element values '" + strElemValues + "' DOES NOT CONTAIN '3' VALUES as REQUIRED!!!";
	}

	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  VerifyCustomRadioGroupOptionSelected  -----------------------------------
* Verify the Radio Group option assigned is in the correct state. The control is custom is that we need to find the option using a parent object to the selected element.
* Base Example is JTR-12580 Keyword created against JTR-14690
* @param mapMethodInfo			The map which contains the following values:
* @param strStartLoc			 The web page location value
* @param strRadioGrpParentName			The parent object name that contains the options. Should contain all possible options.
* @param strRadioGrpParentXpath		   The Xpath for the parent object
* @param strOptionName			The name of the option to select
* @param strElemRadioOptXpath			 The Radio option Xpath
* @param strElemRadioOptValueXpath	 The element xpath under the OptXpath that contain the option value.
* @param strElemRadioOptSelXpath		The Xpath for the element that reflects the selected state
* @param strElemValues			Contains a delimited set of values for the OptionParentProperty|OptionParentPropertyValue|LabelOptionText
*
* @return mapResults				   The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 07/07/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyCustomRadioGroupOptionSelected (mapMethodInfo) { // mapMethodInfo is `any`, return `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'RadioGroupParent';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weRDOParent;   // Type inferred
	//let strTempOptionText; // Declared but unused
	//Return the values from the map
	let strLocation = mapMethodInfo.strStartLoc;
	let strRadioGrpParentName = mapMethodInfo.strRadioGrpParentName;
	let strRadioGrpParentXpath = mapMethodInfo.strRadioGrpParentXpath;
	let strOptionName = mapMethodInfo.strOptionName;
	let strElemRadioOptXpath = mapMethodInfo.strElemRadioOptXpath;
	let strElemRadioOptValueXpath = mapMethodInfo.strElemRadioOptValueXpath;
	let strElemRadioOptSelXpath = mapMethodInfo.strElemRadioOptSelXpath;
	let strElemValues = mapMethodInfo.strElemValues;
	//Split strElemValues
	let mapSplitValues = {}; // Mimic Groovy Map
	mapSplitValues = StringsAndNumbers.JComm_StringToArray(strElemValues, GVars.GblDelimiter('Value'));
	if (StringsAndNumbers.JComm_StringToInteger(mapSplitValues.intItemCount) == 3){
		let strOptParentProperty; // Type inferred
		let strOptParentPropertyValue; // Type inferred
		let strOptionItemText; // Type inferred
		let arryElemValues = mapSplitValues.ArryOfValues; // Inferred to array
		strOptParentProperty = StringsAndNumbers.JComm_HandleNoData(arryElemValues[0]); // Access array by index
		strOptParentPropertyValue = StringsAndNumbers.JComm_HandleNoData(arryElemValues[1]);
		strOptionItemText = StringsAndNumbers.JComm_HandleNoData(arryElemValues[2]);
		//Return the element
		weRDOParent = CWCore.returnWebElement(strRadioGrpParentXpath);
		//Process the element
		if (weRDOParent != null) {
			//Highlight
			if (TCExecParams.getBoolDoHighlight() == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weRDOParent, strTestObjType, strRadioGrpParentName);
			}
			//Check the element state (Enabled, Visible)
			let mapResultsWEState = {}; // Mimic Groovy Map
			mapResultsWEState = CWCore.objVerifyState(weRDOParent, strRadioGrpParentName, true, true);
			//Output results
			let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
			let strStateResults = mapResultsWEState.strMethodDetails;
			//SetVerify the values
			if (boolStatePassed == true) {
				//Return the options for the RDOParent
				let lstChildRDOObjects = weRDOParent.findElements(By.xpath(strElemRadioOptXpath)); // findElements returns a List
				let intCntRDOChildObj = lstChildRDOObjects.length; // Use .length for JS Array
				if (intCntRDOChildObj > 0) {
					let weRDOChild;	 // Type inferred
					let strRDOChildPropValue; // Type inferred
					let boolValMatch = false; // Type inferred
					//Find the radio option by the strElemRadioOptXPath, property name and value
					for (let intRDOChildItem = 0; intRDOChildItem < intCntRDOChildObj; intRDOChildItem ++) {
						weRDOChild = lstChildRDOObjects[intRDOChildItem]; // Access array by index
						//Highlight
						if (TCExecParams.getBoolDoHighlight() == true) {
							let mapHighlight = {}; // Mimic Groovy Map
							mapHighlight = CWCore.objHighlightElementJS(weRDOChild, 'Radio Option', strRadioGrpParentName);
						}
						// Check if attribute exists before getting it.
						if (strOptParentProperty !== 'text' && CWCore.isAttribtuePresent(weRDOChild, strOptParentProperty) == true) {
							strRDOChildPropValue = StringsAndNumbers.JComm_HandleNoData(weRDOChild.getAttribute(strOptParentProperty));
							//Check if the value matches
							if (strRDOChildPropValue == strOptParentPropertyValue) {
								boolValMatch = true;
								break;
							}
						}
						else if (strOptParentProperty === 'text') { // If property to check is text itself
							let lstChildOptionValueElem = weRDOChild.findElements(By.xpath(strElemRadioOptValueXpath));
							if (lstChildOptionValueElem.length === 1) {
								let strTempOptText = lstChildOptionValueElem[0].getText();
								if (strOptionItemText === strTempOptText) {
									boolValMatch = true;
									break;
								}
							}
						}
						else {
							// If boolValMatch is false here, it means no match was found for the property or text within this child,
							// or the attribute wasn't present.
						}
					}
					if (boolValMatch == true) {
						//Find the child item that contain the text and verify value
						let lstChildOptionValueElem = weRDOChild.findElements(By.xpath(strElemRadioOptValueXpath));
						let intCntRDOValueChildObj = lstChildOptionValueElem.length; // Use .length for JS Array
						if (intCntRDOValueChildObj == 1) {
							let weChildOptText = lstChildOptionValueElem[0]; // Access array by index
							//Highlight
							if (TCExecParams.getBoolDoHighlight() == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weChildOptText, 'Radio Option Text', strRadioGrpParentName);
							}
							//Check the value for a match
							let strOptionText = StringsAndNumbers.JComm_HandleNoData(weChildOptText.getText());
							if (strOptionText == strOptionItemText) {
								//Select/Click the button (This is a Verify function, no click)
								//Return the child selected item and return the value
								let lstChildOptionSelElem = weRDOChild.findElements(By.xpath(strElemRadioOptSelXpath));
								let intCntRDOSelectChildObj = lstChildOptionSelElem.length; // Use .length for JS Array
								if (intCntRDOSelectChildObj == 1) {
									let weChildObjOptionSel = lstChildOptionSelElem[0]; // Access array by index
									//Highlight
									if (TCExecParams.getBoolDoHighlight() == true) {
										let mapHighlight = {}; // Mimic Groovy Map
										mapHighlight = CWCore.objHighlightElementJS(weChildObjOptionSel, 'Radio Option Selected', strRadioGrpParentName);
									}
									//Return the selected state
									let boolIsChecked = weChildObjOptionSel.getAttribute("checked");
									if (typeof boolIsChecked === 'string' && boolIsChecked.toLowerCase() === 'true') {
										boolIsChecked = true;
									} else {
										boolIsChecked = false;
									}
									if (boolIsChecked == true) {
										strMethodDetails = "The '" + strRadioGrpParentName + "' Option '" + strOptionName +"' is in the selected state.";
									}
									else {
										boolPassed = false;
										strMethodDetails = "FAILED!!! The '" + strRadioGrpParentName + "' Option '" + strOptionName +"' IS NOT IN THE SELECTED STATE!!!";
									}
								}
								else {
									boolPassed = false;
									strMethodDetails = "FAILED!!! The Radio Option Item specified contained '" + intCntRDOSelectChildObj + "' select elements, THAT DOES NOT MATCH THE '1' count REQUIRED!!!";
								}
							}
							else {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The Radio Option Text value of: '" + strOptionText + "' DOES NOT MATCH THE EXPECTED VALUE: '" + strOptionItemText + "'!!!";
							}
						}
						else {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The Radio Option Item specified contained '" + intCntRDOValueChildObj + "' value elements, THAT DOES NOT MATCH THE '1' count REQUIRED!!!";
						}
					}
				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The Radio Options Parent: " + strRadioGrpParentName + " DID NOT RETURN any OPTION ELEMENTS!!!";
				}
			}
			else
			{
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strRadioGrpParentName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location '" + strLocation + "' '" + strRadioGrpParentName + "' DID NOT RETURN an ELEMENT.";
		}
	}
	else {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The element values '" + strElemValues + "' DOES NOT CONTAIN '3' VALUES as REQUIRED!!!";
	}

	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/* Standard Single Select Combo/Dropdown Box */
/**
* -------------------------------------  objSelectVerifyStandardSingleValueList  -----------------------------------
* Set and verify the Single Value List/Dropdown to the correct value
* @param strLocation			   The web page location value
* @param strElemFullPath		   The EditBox full XPath
* @param strElemName			   The meaningful name of the EditBox
* @param strValueToSelect		  The value to be selected
* @param boolItemTextMustMatch	 The item text must match? true/false
* @param boolSelectItemByIndex	 Select the value by index? true/false
* @param boolVerifySelectedItem	Verify the value selected? true/false
* @return mapResults			   The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 01/12/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_objSelectVerifyStandardSingleValueList (strLocation, strElemFullPath, strElemName,
		strValueToSelect, boolItemTextMustMatch, boolSelectItemByIndex, boolVerifySelectedItem) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'DropDown';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weDropDown;	// Type inferred
	//Return the element
	weDropDown = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weDropDown != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weDropDown, strTestObjType, strElemName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsWEState = {}; // Mimic Groovy Map
		mapResultsWEState = CWCore.objVerifyState(weDropDown, strElemName, true, true);
		//Output results
		let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
		let strStateResults = mapResultsWEState.strMethodDetails;
		//SetVerify the values
		if (boolStatePassed == true) {
			//Attempt to select the value
			let mapSelectItem = {}; // Mimic Groovy Map
			mapSelectItem = CWCore.objSelectPickListItemByOption(weDropDown, strElemFullPath, strElemName, strValueToSelect, boolItemTextMustMatch, boolSelectItemByIndex);
			//Output results
			let boolSelectPassed = StringsAndNumbers.JComm_StringToBoolean (mapSelectItem.boolPassed);
			let strSelectResults = mapSelectItem.strMethodDetails;
			let weNewDropDown = mapSelectItem.objNewDropBox;
			if (boolSelectPassed == true) {
				weDropDown = CWCore.returnWebElement(strElemFullPath); // Re-get element after action
				if (weNewDropDown != null) {
					if (boolVerifySelectedItem == true) {
						//Verify the value selected
						let mapVerifySelectedValue = {}; // Mimic Groovy Map
						mapVerifySelectedValue = CWCore.objSelectVerifyItemSelectedByOption(weDropDown, strElemName, strValueToSelect, false);
						//Output results
						let boolSelVerPassed = StringsAndNumbers.JComm_StringToBoolean (mapVerifySelectedValue.boolPassed); // Fixed from mapSelectItem.boolPassed
						let strSelVerResults = mapVerifySelectedValue.strMethodDetails; // Fixed from mapSelectItem.strMethodDetails
						if (boolSelVerPassed == true) {
							strMethodDetails = "The value: " + strValueToSelect + " was selected from '" + strElemName + "'.";
						}
						else {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The value: " + strValueToSelect + " was selected from '" + strElemName +
									"', HOWEVER, the DISPLAYED VALUE IS INCORRECT!!!. See Details: " + GVars.GblLineFeed("Value") + strSelVerResults;
						}
					}
					else {
						//Report the results
						strMethodDetails = strSelectResults;
					}
				}
				else if (weNewDropDown == null) {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The '" + strElemName + "' WAS NOT RETURNED AFTER SELECTION!!!";
				}
			}
			else {
				boolPassed = false;
				strMethodDetails = strSelectResults;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure final detail string
	return mapResults;
}
/* Checkbox Group */
//TODO update the method comments
/**
* -------------------------------------  SetVerifyGraphicChkBoxGroupItems  -----------------------------------
* Set and verify the Graphic CheckBox to the correct state. Unchecked visual is a checkbox checked is a check mark image
* The input type checkbox state will provide the checked and unchecked state.
* @param mapGraphicCheckbox 	Contains the object information for the form check and the checkbox and checked state
* @param strLocation 			The web page location value
* @param weCheckSelector 		The webelement that must be clicked to change the state
* @param weCheckbox			The checkbox element that will reflect the checked state
* @param arrayOfValues			The arrayOfValues consist of a value pair separated by the gblDelimiter.
* 								The value pair is the value to match. i.e "MyVAlue,true" pairing. If your value contains a "," replace in your value with "/," i.e. "Last,Name,true" is "Last/,Name,true"
* @param boolVerifyChecked		Verify the checked state? true/false
* @return mapResults 			The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 12/31/2021
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyGraphicChkBoxGroupItems (mapSetVerChkboxGrpData) { // mapSetVerChkboxGrpData is `any`, return `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	// Declare the variables
	let strTestObjType = 'ChkboxGroup';
	let strChkboxGrpFullXpath;
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	//Return the values from the map
	let strPageLoc = mapSetVerChkboxGrpData.strPageLoc;
	let strPageXPath = mapSetVerChkboxGrpData.strLocXPath;
	let strChkBoxGrpName = mapSetVerChkboxGrpData.strCkkboxGrpName;
	let strChkboxGrpXPath = mapSetVerChkboxGrpData.strChkboxGroupXPath;
	let strChkboxSelItemXPath = mapSetVerChkboxGrpData.strChkboxSelItemXPath;
	let strChkboxItemXPath = mapSetVerChkboxGrpData.strChkboxItemXPath;
	let strChkboxItemLblXPath = mapSetVerChkboxGrpData.strChkBoxItemLblXPath;
	let strChkBoxGroupData = mapSetVerChkboxGrpData.strChkboxGroupInputData;
	let boolChkBoxGroupMatchAll = mapSetVerChkboxGrpData.boolChkBoxGroupMatchAll;
	let boolChkboxVerifySetValue = mapSetVerChkboxGrpData.boolChkBoxVerifySetValue;
	//Return the element
	let weChkboxGroup;
	if (strPageXPath == gblNull || strPageXPath == null) {
		strChkboxGrpFullXpath = strChkboxGrpXPath;
	}
	else {
		strChkboxGrpFullXpath = strPageXPath + strChkboxGrpXPath;
	}
	weChkboxGroup = CWCore.returnWebElement(strChkboxGrpFullXpath);
	//Process the element
	if (weChkboxGroup != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weChkboxGroup, strTestObjType, strChkBoxGrpName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsWEState = {}; // Mimic Groovy Map
		mapResultsWEState = CWCore.objVerifyState(weChkboxGroup, strChkBoxGrpName, true, true);
		//Output results
		let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
		let strStateResults = mapResultsWEState.strMethodDetails;
		//SetVerify the values
		if (boolStatePassed == true) {
			let intCntChkboxSelectors;
			let intCntChkboxes;
			let intCntChkboxLabels;
			/**
			List<WebElement> lstChkboxSelectors = weChkboxGroup.findElements(By.xpath(strChkboxSelItemXPath))
			intCntChkboxSelectors = lstChkboxSelectors.size()
			List<WebElement> lstChkboxes = weChkboxGroup.findElements(By.xpath(strChkboxItemXPath))
			intCntChkboxes = lstChkboxSelectors.size()
			List<WebElement> lstChkboxLabels = weChkboxGroup.findElements(By.xpath(strChkboxItemLblXPath))
			intCntChkboxLabels = lstChkboxSelectors.size()
			*/
			//Issues with the Checkbox Group not returning only the child objects required using the driver and full XPath OGK 03/14/2022
			let lstChkboxSelectors = TCObj.tcDriver.findElements(By.xpath(strChkboxGrpFullXpath + strChkboxSelItemXPath)); // Using tcDriver
			intCntChkboxSelectors = lstChkboxSelectors.length; // Use .length for JS Array
			let lstChkboxes = TCObj.tcDriver.findElements(By.xpath(strChkboxGrpFullXpath + strChkboxItemXPath)); // Using tcDriver
			intCntChkboxes = lstChkboxes.length; // Use .length for JS Array
			let lstChkboxLabels = TCObj.tcDriver.findElements(By.xpath(strChkboxGrpFullXpath + strChkboxItemLblXPath)); // Using tcDriver
			intCntChkboxLabels = lstChkboxLabels.length; // Use .length for JS Array
			if (intCntChkboxSelectors == intCntChkboxes && intCntChkboxSelectors == intCntChkboxLabels) {
				//Split the input value to begin the processing
				let mapCheckboxItemsValues = StringsAndNumbers.JComm_StringToArray(strChkBoxGroupData, gblDelimiter);
				let arryCheckboxItemsValues = mapCheckboxItemsValues.ArryOfValues;
				let intCntCheckboxItemsValues = arryCheckboxItemsValues.length; // Use .length
				if (intCntChkboxSelectors == intCntCheckboxItemsValues && boolChkBoxGroupMatchAll == true) {
					strMethodDetails = "The user has provided input values with " + intCntCheckboxItemsValues + " label,checked pairs that matches the count of Checkbox items and must match order.";
				}
				else if (intCntChkboxSelectors == intCntCheckboxItemsValues && boolChkBoxGroupMatchAll == false) {
					strMethodDetails = "The user has provided input values with " + intCntCheckboxItemsValues + " label,checked pairs that matches the count of Checkbox items and may be in a different order.";
				}
				else if (intCntChkboxSelectors != intCntCheckboxItemsValues && boolChkBoxGroupMatchAll == false) {
					strMethodDetails = "The user has provided input values with " + intCntCheckboxItemsValues + " label,checked pairs that does not match the count of " +
					intCntChkboxSelectors + " Checkbox items. Match values is false so we will set checked state on item where the label matches.";
				}
				else {
					//Fail
					boolPassed = false;
					strMethodDetails = "FAILED!!! The system is displaying " + intCntChkboxSelectors + " checkboxes and the user provided " + // Corrected count
					intCntCheckboxItemsValues + " label,checked pairs WHICH DOES NOT MATCH and the USER SPECIFIED MATCH VALUES AS TRUE!!!";
				}
				if (boolPassed == true) {
					//Process the input values
					let boolLabelMatch;
					let boolIsChecked;
					let strTempValPair;
					let strAssgLabel;
					let strTempLabel;
					let weTempLabel;
					strMethodDetails = "Processing the Checkbox Group '" + strChkBoxGrpName + "' results.";
					for (let loopInput = 0; loopInput < intCntCheckboxItemsValues; loopInput++) {
						strTempValPair = arryCheckboxItemsValues[loopInput]; // Use array indexing
						boolLabelMatch = false;
						//return the label text and checked state
						//needed to have a way for commas to be present in the value use "\,"
						if (strTempValPair.includes("/,")) { // Use .includes
							let tempValue = strTempValPair.replace("/,", "/-/");
							strAssgLabel = StringsAndNumbers.JComm_GetLeftTextInString(tempValue, ",");
							boolIsChecked = StringsAndNumbers.JComm_StringToBoolean(StringsAndNumbers.JComm_GetRightTextInString(tempValue, ","));
							strAssgLabel = strAssgLabel.replace("/-/", ",");
						}
						else {
							strAssgLabel = StringsAndNumbers.JComm_GetLeftTextInString(strTempValPair, ",");
							boolIsChecked = StringsAndNumbers.JComm_StringToBoolean(StringsAndNumbers.JComm_GetRightTextInString(strTempValPair, ","));
						}
						if (boolChkBoxGroupMatchAll == true) {
							//return the label for the checkbox that matches the loopInput and check if match
							weTempLabel = lstChkboxLabels[loopInput]; // Use array indexing
							//Highlight
							if (TCExecParams.getBoolDoHighlight() == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weTempLabel, 'Label', strChkBoxGrpName);
							}
							strTempLabel = weTempLabel.getText();
							if (strTempLabel == strAssgLabel) {
								//Attempt to set the item
								//Put the values into a map
								let mapChkBoxInput = {}; // Mimic Groovy Map
								mapChkBoxInput.strLocation = strPageLoc;
								mapChkBoxInput.weCheckSelector = lstChkboxSelectors[loopInput]; // Use array indexing
								mapChkBoxInput.weCheckbox = lstChkboxes[loopInput]; // Use array indexing
								mapChkBoxInput.strAssocText = strAssgLabel;
								mapChkBoxInput.boolAssignedChecked = boolIsChecked;
								mapChkBoxInput.boolVerifyChecked = boolChkboxVerifySetValue;
								let mapSetVerifyChkbox = {}; // Mimic Groovy Map
								mapSetVerifyChkbox = SetVerifyGraphicCheckBox(mapChkBoxInput); // Call to global func
								//Check the results and mark pass or fail
								boolPassed = StringsAndNumbers.JComm_StringToBoolean (mapSetVerifyChkbox.boolPassed);
								//Update the results
								strMethodDetails = strMethodDetails + gblLineFeed + mapSetVerifyChkbox.strMethodDetails;
								// break; // Original had break here, but it should probably continue processing input loop
							}
							else {
								boolPassed = false;
								strMethodDetails = strMethodDetails + "FAILED!!! The Label value for item: " + loopInput +
								" of: '" + strTempLabel + "' DOES NOT MATCH THE ASSIGNED VALUE OF: '" + strAssgLabel + "'!!!";
								break; // Break input for loop
							}
						}
						else {
							//Loop through the checkbox item labels to find a match
							for (let loopLabels = 0; loopLabels < intCntChkboxLabels; loopLabels++) {
								weTempLabel = lstChkboxLabels[loopLabels]; // Use array indexing
								//Highlight
								if (TCExecParams.getBoolDoHighlight() == true) {
									let mapHighlight = {}; // Mimic Groovy Map
									mapHighlight = CWCore.objHighlightElementJS(weTempLabel, 'Label', strChkBoxGrpName);
								}
								strTempLabel = weTempLabel.getText();
								if (strTempLabel == strAssgLabel) {
									boolLabelMatch = true;
									//Attempt to set the item
									//Put the values into a map
									let mapChkBoxInput = {}; // Mimic Groovy Map
									mapChkBoxInput.strLocation = strPageLoc;
									mapChkBoxInput.weCheckSelector = lstChkboxSelectors[loopLabels]; // Use array indexing
									mapChkBoxInput.weCheckbox = lstChkboxes[loopLabels]; // Use array indexing
									mapChkBoxInput.strAssocText = strAssgLabel;
									mapChkBoxInput.boolAssignedChecked = boolIsChecked;
									mapChkBoxInput.boolVerifyChecked = boolChkboxVerifySetValue;
									let mapSetVerifyChkbox = {}; // Mimic Groovy Map
									mapSetVerifyChkbox = SetVerifyGraphicCheckBox(mapChkBoxInput); // Call to global func
									//Check the results and mark pass or fail
									boolPassed = StringsAndNumbers.JComm_StringToBoolean (mapSetVerifyChkbox.boolPassed);
									//Update the results
									strMethodDetails = strMethodDetails + gblLineFeed + mapSetVerifyChkbox.strMethodDetails;
									break; // Break from loopLabels
								}
							}
							if (boolPassed == false) {
								break; //Exit the input loop
							}
							else if (boolLabelMatch == false) {
								boolPassed = false;
								strMethodDetails = strMethodDetails + "FAILED!!! The Label value assigned of: " + strAssgLabel +
								" DID NOT MATCH ANY OF THE DISPLAYED CHECKBOXES";
								break; // Break input loop
							}
						}
						if (boolPassed == false) {
							break; //Exit the input loop
						}
					}
				}
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location name: " + strPageLoc + " and element name '" +
				strChkBoxGrpName + " COUNT of child elements DID NOT MATCH!!!" + gblLineFeed +
				"intCntChkboxSelectors: " + intCntChkboxSelectors + gblLineFeed +
				"intCntChkboxes: " + intCntChkboxes + gblLineFeed +
				"intCntChkboxLabels: " + intCntChkboxLabels;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strPageLoc + " and element name '" +
					strChkBoxGrpName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strPageLoc + "' '" + strChkBoxGrpName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure this is set for all paths
	return mapResults;
}
/* Radio Group */
/**
* -------------------------------------  SetVerifyRadioGroupOption  -----------------------------------
* Set and verify the Radio Group option to the specified state of selected
* @param strLocation				   The web page location value
* @param strElemRadioGroupFullPath	 The Radio Group full XPath
* @param strElemRadioGroupName		 The meaningful name of the Radio Group
* @param strElemRadioBtnXPath			The Xpath for the radio button/option
* @param strOptionToSelect			The Radio Button/Option to select
* @param boolVerifySeleted			Verify the selected state? true/false
* @return mapResults				   The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 02/12/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyRadioGroupOption (strLocation, strElemRadioGroupFullPath, strElemRadioGroupName,
		strElemRadioOptXPath, strElemRadioOptBtnXPath, strOptionToSelect, boolVerifySelected) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'RadioGroup';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weRadioGroup;  // Type inferred
	//let strTempOptionText; // Declared but unused

	//Return the element
	weRadioGroup = CWCore.returnWebElement(strElemRadioGroupFullPath);
	//Process the element
	if (weRadioGroup != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weRadioGroup, strTestObjType, strElemRadioGroupName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsWEState = {}; // Mimic Groovy Map
		mapResultsWEState = CWCore.objVerifyState(weRadioGroup, strElemRadioGroupName, true, true);
		//Output results
		let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
		let strStateResults = mapResultsWEState.strMethodDetails;
		//SetVerify the values
		if (boolStatePassed == true) {
			//Set the Radio Group option
			let mapSetRadioGrpOption = {}; // Mimic Groovy Map
			mapSetRadioGrpOption = CWCore.objSetRadioGroupOption(weRadioGroup, strElemRadioGroupName, strElemRadioOptXPath, strElemRadioGroupFullPath, strElemRadioOptBtnXPath, strOptionToSelect, boolVerifySelected);
			boolPassed = StringsAndNumbers.JComm_StringToBoolean (mapSetRadioGrpOption.boolMethodPassed);
			strMethodDetails = mapSetRadioGrpOption.strMethodDetails;
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemRadioGroupName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemRadioGroupName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  SetVerifyCustomRadioGroupOption  -----------------------------------
* Set and verify the Radio Group option to the specified state of selected using JavaGroup for Selection as Radio button have ::before and ::after
* Base Example is JTR-12580 Keyword created against JTR-14690
* @param mapMethodInfo			The map which contains the following values:
* @param strStartLoc			 The web page location value
* @param strRadioGrpParentName			The parent object name that contains the option
* @param strRadioGrpParentXpath		   The Xpath for the parent object
* @param strOptionName			The name of the option to select
* @param strElemRadioOptXpath			 The Radio option Xpath
* @param strElemRadioOptValueXpath	 The element xpath under the OptXpath that contain the option value.
* @param strElemRadioBtnXpath			The Xpath for the radio button/option
* @param strBoolUseJavaScript			Use Java Script to select the option based on psuedo code for the label as in ::before and ::after
* @param strElemRadioOptSelXpath		The Xpath for the element that reflects the selected state
* @param strElemValues			Contains a delimited set of values for the OptionParentProperty|OptionParentPropertyValue|LabelOptionText
* @param boolVerifySelected			Verify the selected state? true/false
* @return mapResults				   The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 06/12/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyCustomRadioGroupOption (mapMethodInfo) { // mapMethodInfo is `any`, return `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'RadioGroupParent';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weRDOParent;   // Type inferred
	//let strTempOptionText; // Declared but unused
	//Return the values from the map
	let strLocation = mapMethodInfo.strStartLoc;
	let strRadioGrpParentName = mapMethodInfo.strRadioGrpParentName;
	let strRadioGrpParentXpath = mapMethodInfo.strRadioGrpParentXpath;
	let strElemRadioOptXpath = mapMethodInfo.strElemRadioOptXpath;
	let strElemRadioOptValueXpath = mapMethodInfo.strElemRadioOptValueXpath;
	let strElemRadioOptBtnXpath = mapMethodInfo.strElemRadioOptBtnXpath;
	let boolUseJavaScript = mapMethodInfo.boolUseJavaScript;
	let strElemRadioOptSelXpath = mapMethodInfo.strElemRadioOptSelXpath;
	let strElemValues = mapMethodInfo.strElemValues;
	let boolVerifySelected = mapMethodInfo.boolVerifySelected;
	//Split strElemValues
	let mapSplitValues = {}; // Mimic Groovy Map
	mapSplitValues = StringsAndNumbers.JComm_StringToArray(strElemValues, GVars.GblDelimiter('Value'));
	if (StringsAndNumbers.JComm_StringToInteger(mapSplitValues.intItemCount) == 3){
		let strOptParentProperty; // Type inferred
		let strOptParentPropertyValue; // Type inferred
		let strOptionItemText; // Type inferred
		let arryElemValues = mapSplitValues.ArryOfValues; // Inferred to array
		strOptParentProperty = StringsAndNumbers.JComm_HandleNoData(arryElemValues[0]); // Access array by index
		strOptParentPropertyValue = StringsAndNumbers.JComm_HandleNoData(arryElemValues[1]);
		strOptionItemText = StringsAndNumbers.JComm_HandleNoData(arryElemValues[2]);
		//Return the element
		weRDOParent = CWCore.returnWebElement(strRadioGrpParentXpath);
		//Process the element
		if (weRDOParent != null) {
			//Highlight
			if (TCExecParams.getBoolDoHighlight() == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weRDOParent, strTestObjType, strRadioGrpParentName);
			}
			//Check the element state (Enabled, Visible)
			let mapResultsWEState = {}; // Mimic Groovy Map
			mapResultsWEState = CWCore.objVerifyState(weRDOParent, strRadioGrpParentName, true, true);
			//Output results
			let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
			let strStateResults = mapResultsWEState.strMethodDetails;
			//SetVerify the values
			if (boolStatePassed == true) {
				//Return the options for the RDOParent
				let lstChildRDOObjects = weRDOParent.findElements(By.xpath(strElemRadioOptXpath)); // findElements returns a List
				let intCntRDOChildObj = lstChildRDOObjects.length; // Use .length for JS Array
				if (intCntRDOChildObj > 0) {
					let weRDOChild;	 // Type inferred
					let strRDOChildPropValue; // Type inferred
					let boolValMatch = false; // Type inferred
					//Find the radio option by the strElemRadioOptXPath, property name and value
					for (let intRDOChildItem = 0; intRDOChildItem < intCntRDOChildObj; intRDOChildItem ++) {
						weRDOChild = lstChildRDOObjects[intRDOChildItem]; // Access array by index
						//Highlight
						if (TCExecParams.getBoolDoHighlight() == true) {
							let mapHighlight = {}; // Mimic Groovy Map
							mapHighlight = CWCore.objHighlightElementJS(weRDOChild, 'Radio Option', strRadioGrpParentName);
						}
						// Check if the attribute is present *before* getting it, and handle 'text' as a property
						if (strOptParentProperty === 'text') {
							let lstChildOptionValueElem = weRDOChild.findElements(By.xpath(strElemRadioOptValueXpath));
							if (lstChildOptionValueElem.length === 1) {
								let strTempOptText = lstChildOptionValueElem[0].getText();
								if (strOptionItemText === strTempOptText) {
									boolValMatch = true;
									break;
								}
							} else {
								boolPassed = false;
								strMethodDetails = `FAILED!!! Radio Option Item contained ${lstChildOptionValueElem.length} value elements, expected 1.`;
								break;
							}
						} else if(CWCore.isAttribtuePresent(weRDOChild, strOptParentProperty) == true) { // General attribute
							strRDOChildPropValue = StringsAndNumbers.JComm_HandleNoData(weRDOChild.getAttribute(strOptParentProperty));
							//Check if the value matches
							if (strRDOChildPropValue == strOptParentPropertyValue) {
								boolValMatch = true;
								break;
							}
						}
						else { // Attribute not found or not handled explicitly
							boolPassed = false;
							strMethodDetails = "FAILED!!! The Radio Option Item specified DOES NOT HAVE THE ATTRIBUTE '" + strOptParentProperty + "'!!!";
							break;
						}
					}
					if (boolValMatch == true && boolPassed == true) { // Added boolPassed check
						//Find the child item that contain the text and verify value
						let lstChildOptionValueElem = weRDOChild.findElements(By.xpath(strElemRadioOptValueXpath));
						let intCntRDOValueChildObj = lstChildOptionValueElem.length; // Use .length for JS Array
						if (intCntRDOValueChildObj == 1) {
							let weChildOptText = lstChildOptionValueElem[0]; // Access array by index
							//Highlight
							if (TCExecParams.getBoolDoHighlight() == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weChildOptText, 'Radio Option Text', strRadioGrpParentName);
							}
							//Check the value for a match
							let strOptionText = StringsAndNumbers.JComm_HandleNoData(weChildOptText.getText());
							if (strOptionText == strOptionItemText) {
								//Select/Click the button
								//Return the child selected item and return the value
								let lstChildOptionSelElem = weRDOChild.findElements(By.xpath(strElemRadioOptSelXpath));
								let intCntRDOSelectChildObj = lstChildOptionSelElem.length; // Use .length for JS Array
								if (intCntRDOSelectChildObj == 1) {
									let weChildObjOptionSel = lstChildOptionSelElem[0]; // Access array by index
									//Highlight
									if (TCExecParams.getBoolDoHighlight() == true) {
										let mapHighlight = {}; // Mimic Groovy Map
										mapHighlight = CWCore.objHighlightElementJS(weChildObjOptionSel, 'Radio Option Selected', strRadioGrpParentName);
									}
									//Return the selected state
									let boolSelected = weChildObjOptionSel.getAttribute("checked"); // Get return of getAttribute
									boolSelected = (typeof boolSelected === 'string' && boolSelected.toLowerCase() === 'true'); // Coerce to boolean.

									if (boolSelected == false) {
										//If not selected, return the button element
										let lstChildOptionBtnElem = weRDOChild.findElements(By.xpath(strElemRadioOptBtnXpath));
										let intCntRDOBtnChildObj = lstChildOptionBtnElem.length; // Use .length for JS Array
										if (intCntRDOBtnChildObj == 1) {
											//Click the button element or use Javascript to click the element
											let weOptBtn = lstChildOptionBtnElem[0]; // Access array by index
											//Highlight
											if (TCExecParams.getBoolDoHighlight() == true) {
												let mapHighlight = {}; // Mimic Groovy Map
												mapHighlight = CWCore.objHighlightElementJS(weOptBtn, 'Radio Option Button', strRadioGrpParentName);
											}
											if (boolUseJavaScript == true) {
												// Casting to `any` to allow direct execution.
												TCObj.tcDriver.executeScript("arguments[0].click();", weOptBtn);
											}
											else {
												weOptBtn.click();
											}
											if (boolVerifySelected == true) {
												// re-check selected state after click
												let boolIsChecked = weChildObjOptionSel.getAttribute("checked");
												boolIsChecked = (typeof boolIsChecked === 'string' && boolIsChecked.toLowerCase() === 'true'); // Coerce to boolean.
												if (boolIsChecked == true) {
													strMethodDetails = "The '" + strRadioGrpParentName + "' Option '" + strOptionItemText +"' was clicked and is in the selected state.";
												}
												else {
													boolPassed = false;
													strMethodDetails = "FAILED!!! The '" + strRadioGrpParentName + "' Option '" + strOptionItemText +"' IS NOT IN THE SELECTED STATE!!!";
												}
											}
										}
										else {
											boolPassed = false;
											strMethodDetails = "FAILED!!! The Radio Option Item specified contained '" + intCntRDOBtnChildObj + "' button element(s), THAT DOES NOT MATCH THE '1' count REQUIRED!!!";
										}
									}
									else {
										strMethodDetails = "The option '" + strOptionItemText + "' is already selected.";
									}
								}
								else {
									boolPassed = false;
									strMethodDetails = "FAILED!!! The Radio Option Item specified contained '" + intCntRDOSelectChildObj + "' select elements, THAT DOES NOT MATCH THE '1' count REQUIRED!!!";
								}
							}
							else {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The Radio Option Text value of: '" + strOptionText + "' DOES NOT MATCH THE EXPECTED VALUE: '" + strOptionItemText + "'!!!";
							}
						}
						else { // If value element not found or not unique
							boolPassed = false;
							strMethodDetails = "FAILED!!! The Radio Option Item specified contained '" + intCntRDOValueChildObj + "' value elements, THAT DOES NOT MATCH THE '1' count REQUIRED!!!";
						}
					}
				}
				else { // No RDO option elements found matching the XPath
					boolPassed = false;
					strMethodDetails = "FAILED!!! The Radio Options Parent: " + strRadioGrpParentName + " DID NOT RETURN any OPTION ELEMENTS!!!";
				}
			}
			else
			{
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strRadioGrpParentName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location '" + strLocation + "' '" + strRadioGrpParentName + "' DID NOT RETURN an ELEMENT.";
		}
	}
	else { // strElemValues not 3 items
		boolPassed = false;
		strMethodDetails = "FAILED!!! The element values '" + strElemValues + "' DOES NOT CONTAIN '3' VALUES as REQUIRED!!!";
	}

	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  VerifyCustomRadioGroupOptionSelected  -----------------------------------
* Verify the Radio Group option assigned is in the correct state. The control is custom is that we need to find the option using a parent object to the selected element.
* Base Example is JTR-12580 Keyword created against JTR-14690
* @param mapMethodInfo			The map which contains the following values:
* @param strStartLoc			 The web page location value
* @param strRadioGrpParentName			The parent object name that contains the options. Should contain all possible options.
* @param strRadioGrpParentXpath		   The Xpath for the parent object
* @param strOptionName			The name of the option to select
* @param strElemRadioOptXpath			 The Radio option Xpath
* @param strElemRadioOptValueXpath	 The element xpath under the OptXpath that contain the option value.
* @param strElemRadioOptSelXpath		The Xpath for the element that reflects the selected state
* @param strElemValues			Contains a delimited set of values for the OptionParentProperty|OptionParentPropertyValue|LabelOptionText
*
* @return mapResults				   The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 07/07/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyCustomRadioGroupOptionSelected (mapMethodInfo) { // mapMethodInfo is `any`, return `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'RadioGroupParent';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weRDOParent;   // Type inferred
	//let strTempOptionText; // Declared but unused
	//Return the values from the map
	let strLocation = mapMethodInfo.strStartLoc;
	let strRadioGrpParentName = mapMethodInfo.strRadioGrpParentName;
	let strRadioGrpParentXpath = mapMethodInfo.strRadioGrpParentXpath;
	let strOptionName = mapMethodInfo.strOptionName;
	let strElemRadioOptXpath = mapMethodInfo.strElemRadioOptXpath;
	let strElemRadioOptValueXpath = mapMethodInfo.strElemRadioOptValueXpath;
	let strElemRadioOptSelXpath = mapMethodInfo.strElemRadioOptSelXpath;
	let strElemValues = mapMethodInfo.strElemValues;
	//Split strElemValues
	let mapSplitValues = {}; // Mimic Groovy Map
	mapSplitValues = StringsAndNumbers.JComm_StringToArray(strElemValues, GVars.GblDelimiter('Value'));
	if (StringsAndNumbers.JComm_StringToInteger(mapSplitValues.intItemCount) == 3){
		let strOptParentProperty; // Type inferred
		let strOptParentPropertyValue; // Type inferred
		let strOptionItemText; // Type inferred
		let arryElemValues = mapSplitValues.ArryOfValues; // Inferred to array
		strOptParentProperty = StringsAndNumbers.JComm_HandleNoData(arryElemValues[0]); // Access array by index
		strOptParentPropertyValue = StringsAndNumbers.JComm_HandleNoData(arryElemValues[1]);
		strOptionItemText = StringsAndNumbers.JComm_HandleNoData(arryElemValues[2]);
		//Return the element
		weRDOParent = CWCore.returnWebElement(strRadioGrpParentXpath);
		//Process the element
		if (weRDOParent != null) {
			//Highlight
			if (TCExecParams.getBoolDoHighlight() == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weRDOParent, strTestObjType, strRadioGrpParentName);
			}
			//Check the element state (Enabled, Visible)
			let mapResultsWEState = {}; // Mimic Groovy Map
			mapResultsWEState = CWCore.objVerifyState(weRDOParent, strRadioGrpParentName, true, true);
			//Output results
			let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
			let strStateResults = mapResultsWEState.strMethodDetails;
			//SetVerify the values
			if (boolStatePassed == true) {
				//Return the options for the RDOParent
				let lstChildRDOObjects = weRDOParent.findElements(By.xpath(strElemRadioOptXpath)); // findElements returns a List
				let intCntRDOChildObj = lstChildRDOObjects.length; // Use .length for JS Array
				if (intCntRDOChildObj > 0) {
					let weRDOChild;	 // Type inferred
					let strRDOChildPropValue; // Type inferred
					let boolValMatch = false; // Type inferred
					//Find the radio option by the strElemRadioOptXPath, property name and value
					for (let intRDOChildItem = 0; intRDOChildItem < intCntRDOChildObj; intRDOChildItem ++) {
						weRDOChild = lstChildRDOObjects[intRDOChildItem]; // Access array by index
						//Highlight
						if (TCExecParams.getBoolDoHighlight() == true) {
							let mapHighlight = {}; // Mimic Groovy Map
							mapHighlight = CWCore.objHighlightElementJS(weRDOChild, 'Radio Option', strRadioGrpParentName);
						}
						// Check if attribute exists before getting it.
						if (strOptParentProperty === 'text') { // Custom check for text property
							let lstChildOptionValueElem = weRDOChild.findElements(By.xpath(strElemRadioOptValueXpath));
							if (lstChildOptionValueElem.length === 1) {
								let strTempOptText = lstChildOptionValueElem[0].getText();
								if (strOptionItemText === strTempOptText) {
									boolValMatch = true;
									break; // Found matching option
								}
							}
						} else if(CWCore.isAttribtuePresent(weRDOChild, strOptParentProperty) == true) { // General attribute
							strRDOChildPropValue = StringsAndNumbers.JComm_HandleNoData(weRDOChild.getAttribute(strOptParentProperty));
							//Check if the value matches
							if (strRDOChildPropValue == strOptParentPropertyValue) {
								boolValMatch = true;
								break;
							}
						}
						// If boolValMatch is false here, it means no match was found for the property or text within this child,
						// or the attribute wasn't present.
					}
					if (boolValMatch == true) {
						//Find the child item that contain the text and verify value
						let lstChildOptionValueElem = weRDOChild.findElements(By.xpath(strElemRadioOptValueXpath));
						let intCntRDOValueChildObj = lstChildOptionValueElem.length; // Use .length for JS Array
						if (intCntRDOValueChildObj == 1) {
							let weChildOptText = lstChildOptionValueElem[0]; // Access array by index
							//Highlight
							if (TCExecParams.getBoolDoHighlight() == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weChildOptText, 'Radio Option Text', strRadioGrpParentName);
							}
							//Check the value for a match
							let strOptionText = StringsAndNumbers.JComm_HandleNoData(weChildOptText.getText());
							if (strOptionText == strOptionItemText) {
								//Select/Click the button
								//Return the child selected item and return the value
								let lstChildOptionSelElem = weRDOChild.findElements(By.xpath(strElemRadioOptSelXpath));
								let intCntRDOSelectChildObj = lstChildOptionSelElem.length; // Use .length for JS Array
								if (intCntRDOSelectChildObj == 1) {
									let weChildObjOptionSel = lstChildOptionSelElem[0]; // Access array by index
									//Highlight
									if (TCExecParams.getBoolDoHighlight() == true) {
										let mapHighlight = {}; // Mimic Groovy Map
										mapHighlight = CWCore.objHighlightElementJS(weChildObjOptionSel, 'Radio Option Selected', strRadioGrpParentName);
									}
									//Return the selected state
									let boolIsChecked = weChildObjOptionSel.getAttribute("checked");
									boolIsChecked = (typeof boolIsChecked === 'string' && boolIsChecked.toLowerCase() === 'true'); // Coerce to boolean.
									if (boolIsChecked == true) {
										strMethodDetails = "The '" + strRadioGrpParentName + "' Option '" + strOptionName +"' is in the selected state.";
									}
									else {
										boolPassed = false;
										strMethodDetails = "FAILED!!! The '" + strRadioGrpParentName + "' Option '" + strOptionName +"' IS NOT IN THE SELECTED STATE!!!";
									}
								}
								else {
									boolPassed = false;
									strMethodDetails = "FAILED!!! The Radio Option Item specified contained '" + intCntRDOSelectChildObj + "' select elements, THAT DOES NOT MATCH THE '1' count REQUIRED!!!";
								}
							}
							else {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The Radio Option Text value of: '" + strOptionText + "' DOES NOT MATCH THE EXPECTED VALUE: '" + strOptionItemText + "'!!!";
							}
						}
						else {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The Radio Option Item specified contained '" + intCntRDOValueChildObj + "' value elements, THAT DOES NOT MATCH THE '1' count REQUIRED!!!";
						}
					}
				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The Radio Options Parent: " + strRadioGrpParentName + " DID NOT RETURN any OPTION ELEMENTS!!!";
				}
			}
			else
			{
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strRadioGrpParentName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location '" + strLocation + "' '" + strRadioGrpParentName + "' DID NOT RETURN an ELEMENT.";
		}
	}
	else { // strElemValues not 3 items
		boolPassed = false;
		strMethodDetails = "FAILED!!! The element values '" + strElemValues + "' DOES NOT CONTAIN '3' VALUES as REQUIRED!!!";
	}

	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure final detail string
	return mapResults;
}
/* Standard Single Select Combo/Dropdown Box */
/**
* -------------------------------------  objSelectVerifyStandardSingleValueList  -----------------------------------
* Set and verify the Single Value List/Dropdown to the correct value
* @param strLocation			   The web page location value
* @param strElemFullPath		   The EditBox full XPath
* @param strElemName			   The meaningful name of the EditBox
* @param strValueToSelect		  The value to be selected
* @param boolItemTextMustMatch	 The item text must match? true/false
* @param boolSelectItemByIndex	 Select the value by index? true/false
* @param boolVerifySelectedItem	Verify the value selected? true/false
* @return mapResults			   The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 01/12/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_objSelectVerifyStandardSingleValueList (strLocation, strElemFullPath, strElemName,
		strValueToSelect, boolItemTextMustMatch, boolSelectItemByIndex, boolVerifySelectedItem) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'DropDown';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weDropDown;	// Type inferred
	//Return the element
	weDropDown = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weDropDown != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weDropDown, strTestObjType, strElemName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsWEState = {}; // Mimic Groovy Map
		mapResultsWEState = CWCore.objVerifyState(weDropDown, strElemName, true, true);
		//Output results
		let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
		let strStateResults = mapResultsWEState.strMethodDetails;
		//SetVerify the values
		if (boolStatePassed == true) {
			//Attempt to select the value
			let mapSelectItem = {}; // Mimic Groovy Map
			mapSelectItem = CWCore.objSelectPickListItemByOption(weDropDown, strElemFullPath, strElemName, strValueToSelect, boolItemTextMustMatch, boolSelectItemByIndex);
			//Output results
			let boolSelectPassed = StringsAndNumbers.JComm_StringToBoolean (mapSelectItem.boolPassed);
			let strSelectResults = mapSelectItem.strMethodDetails;
			let weNewDropDown = mapSelectItem.objNewDropBox;
			if (boolSelectPassed == true) {
				weDropDown = CWCore.returnWebElement(strElemFullPath); // Re-get element after action
				if (weNewDropDown != null) {
					if (boolVerifySelectedItem == true) {
						//Verify the value selected
						let mapVerifySelectedValue = {}; // Mimic Groovy Map
						mapVerifySelectedValue = CWCore.objSelectVerifyItemSelectedByOption(weDropDown, strElemName, strValueToSelect, false);
						//Output results
						let boolSelVerPassed = StringsAndNumbers.JComm_StringToBoolean (mapVerifySelectedValue.boolPassed); // Fixed from mapSelectItem.boolPassed
						let strSelVerResults = mapVerifySelectedValue.strMethodDetails; // Fixed from mapSelectItem.strMethodDetails
						if (boolSelVerPassed == true) {
							strMethodDetails = "The value: " + strValueToSelect + " was selected from '" + strElemName + "'.";
						}
						else {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The value: " + strValueToSelect + " was selected from '" + strElemName +
									"', HOWEVER, the DISPLAYED VALUE IS INCORRECT!!!. See Details: " + gblLineFeed + strSelVerResults; // Changed GVars.GblHTMLLineFeed to gblLineFeed
						}
					}
					else {
						//Report the results
						strMethodDetails = strSelectResults;
					}
				}
				else if (weNewDropDown == null) {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The '" + strElemName + "' WAS NOT RETURNED AFTER SELECTION!!!";
				}
			}
			else {
				boolPassed = false;
				strMethodDetails = strSelectResults;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure final detail string
	return mapResults;
}
//Edit Multi Value Filter Listbox see Supplier Advanced Search Enter the 'Country of Origin' filter and select item. for example
/**
* -------------------------------------  objSetVerifyeditMultiValueFilterListBox  -----------------------------------
* Select an item from a ListBox multi-selectable)
* @param mapInputValues			 The map containing all of the input values for the method
* @param strStartLoc				The web page location value
* @param strValueFilter			 The filter value(s) to set/verify in filter edit
* @param strValueItems			 The value(s) to select and add to the filter multi-select edit.
* @param boolSelctItemExactMatch	The item text must match? true/false
* @param strXPath					 The XPaths for the multiple objects required for the method see below:

* @param strElemEditBoxName		   The edit box name where the filter value is entered
* @param strElemEditBoxXPath		  The edit box XPath where the filter value is entered
* @param strElemEditListBoxName	   The Edit List Name of items added to the edit field once values are selected
* @param strElemEditListBoxXPath	  The Edit List Xpath of items added to the edit field once values are selected
* @param strElemEditListBoxItemBtnName	   The Edit List Button Item Name, is the item element including the text and delete that is a button in the edit box list
* @param strElemEditListBoxItemBtnXPath	  The Edit List Item Xpath, is the item element including the text and delete that is a button in the edit box list
* @param strElemEditListBoxItemTextName	  The Edit List Item Text Name is the item element in the edit list
* @param strElemEditListBoxItemTextXPath	 The Edit List Item Text Xpath is the item element in the edit list
* @param strElemEditListBoxItemBtnDeleteName The Edit List Item Delete Name is the item element delete button in the edit list
* @param strElemEditListBoxItemBtnDeleteXPath The Edit List Item Delete Xpath is the item element delete button in the edit list
* @param strElemListBoxName		   The filter list box name
* @param strElemListBoxXPath		  The filter list box Xpath
* @param strElemListBoxItemName	   The filter list box item name
* @param strElemListBoxItemXPath	  The filter list box item Xpath
* @param strBoolClearValue			Alternate boolean used to clear the filter edit before entering value.By Default == false
* @param strBoolVerifyValue		  Alternate boolean used to determine if the value entered in the filter should be verified.By Default == false
* @param strBoolTabOut				Alternate boolean used to determine if tab is selected after value is entered.By Default == false
*
* @return mapResults				The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 02/03/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_objSetVerifyeditMultiValueFilterListBox (mapInputValues) { // mapInputValues is `any`, return type `any`
	//Set the global variables
	let gblNull = GVars.GblNull("Value");
	let gblUndefined = GVars.GblUndefined("Value");
	let gblLineFeed = GVars.GblLineFeed("Value");
	let gblSkipStep = GVars.GblSkip('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	//Create the method variables
	let boolMethodPassed = true;
	let strMethodDetails = gblNull;
	let mapResults = {}; // Mimic Groovy Map
	//Return the map variables
	//Set the name and Xpaths for each of the objects
	let strStartLoc = mapInputValues.strStartLoc;
	let strValueFilter = mapInputValues.strValueFilter;
	let strValueItems = mapInputValues.strValueItems;
	let strElemEditBoxName = mapInputValues.strElemEditBoxName;
	let strElemEditBoxXPath = mapInputValues.strElemEditBoxXPath;
	let strElemEditListBoxName = mapInputValues.strElemEditListBoxName;
	let strElemEditListBoxXPath = mapInputValues.strElemEditListBoxXPath;
	let strElemEditListBoxItemBtnName = mapInputValues.strElemEditListBoxItemBtnName;
	let strElemEditListBoxItemBtnXPath = mapInputValues.strElemEditListBoxItemBtnXPath;
	let strElemEditListBoxItemTextName = mapInputValues.strElemEditListBoxItemTextName;
	let strElemEditListBoxItemTextXPath = mapInputValues.strElemEditListBoxItemTextXPath;
	let strElemEditListBoxItemBtnDeleteName = mapInputValues.strElemEditListBoxItemBtnDeleteName;
	let strElemEditListBoxItemBtnDeleteXPath = mapInputValues.strElemEditListBoxItemBtnDeleteXPath;
	let strElemListBoxName = mapInputValues.strElemListBoxName;
	let strElemListBoxXPath = mapInputValues.strElemListBoxXPath;
	let strElemListBoxItemName = mapInputValues.strElemListBoxItemName;
	let strElemListBoxItemXPath = mapInputValues.strElemListBoxItemXPath; // Fixed from original
	//Check for keys to set the clear value before entry, verify value and tabout
	let boolClearValue = false;
	let boolVerifyValue = false;
	let boolTabOut = false;
	//Update the alternate values if assigned in the map
	//ClearValue
	if (mapInputValues.hasOwnProperty('strBoolClearValue')) {
		boolClearValue = StringsAndNumbers.JComm_StringToBoolean(mapInputValues.strBoolClearValue);
	}
	//VerifyValue
	if (mapInputValues.hasOwnProperty('strBoolVerifyValue')) {
		boolVerifyValue = StringsAndNumbers.JComm_StringToBoolean(mapInputValues.strBoolVerifyValue);
	}
	//Tabout
	if (mapInputValues.hasOwnProperty('strBoolTabOut')) {
		boolTabOut = StringsAndNumbers.JComm_StringToBoolean(mapInputValues.strBoolTabOut);
	}
	//Check the filter values by creating new arrays
	let arryFilterItems;
	let arrySelectItems;
	let mapFilterItems = StringsAndNumbers.JComm_StringToArray(strValueFilter, gblDelimiter);
	arryFilterItems = mapFilterItems.ArryOfValues;
	let cntFilterItems = arryFilterItems.length; // Use .length
	let mapSelectItems = StringsAndNumbers.JComm_StringToArray(strValueItems, gblDelimiter);
	arrySelectItems = mapSelectItems.ArryOfValues;
	let cntSelectItems = arrySelectItems.length; // Use .length
	if (cntFilterItems == cntSelectItems && cntFilterItems > 0) {
		let strFilterValue;
		let strSelectValue;
		//let boolLoopPassed; // Declared but unused
		let boolEditPassed;
		let boolSelectPassed;
		//let strLoopDetails; // Declared but unused
		let strEditResults;
		let strSelectResults;
		let strSelectedItemValues; // Should be initialized before loop for correct concat
		let mapResultsSetEdit; // Type inferred
		let mapResultsSelectItem; // Type inferred
		//Create the loop for setting the filter and selecting items
		for (let loopArray = 0; loopArray < cntFilterItems; loopArray++) {
			//Return each of the values
			strFilterValue = StringsAndNumbers.JComm_HandleNoData(arryFilterItems[loopArray]).trim(); // Access by index
			strSelectValue = StringsAndNumbers.JComm_HandleNoData(arrySelectItems[loopArray]).trim(); // Access by index
			//Process the EditBox
			mapResultsSetEdit = {}; // Mimic Groovy Map
			mapResultsSetEdit = SetVerifyEditBoxValue(strStartLoc, strElemEditBoxXPath, strElemEditBoxName, strFilterValue, boolClearValue, boolVerifyValue, boolTabOut);
			//Check the results and mark pass or fail
			boolEditPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsSetEdit.boolPassed);
			strEditResults = mapResultsSetEdit.strMethodDetails;
			if (boolEditPassed == true) {
				//Process the selection
				mapResultsSelectItem = {}; // Mimic Groovy Map
				mapResultsSelectItem= objPickItemFromListBox(strStartLoc, strElemListBoxXPath, strElemListBoxName, strElemListBoxItemXPath, strSelectValue);
				//Check the results and mark pass or fail
				boolSelectPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsSelectItem.boolPassed);
				strSelectResults = mapResultsSelectItem.strMethodDetails;
				strMethodDetails = strMethodDetails + gblLineFeed + strEditResults + gblLineFeed + strSelectResults;
				if (boolSelectPassed == true) {
					if (loopArray == 0) {
						strSelectedItemValues = strSelectValue;
					}
					else {
						strSelectedItemValues = strSelectedItemValues + gblDelimiter + strSelectValue;
					}
				}

			}
			else {
				//Report edit failed
				if (strMethodDetails == null) {
					strMethodDetails = strEditResults; // Initialize if null
				}
				else {
					strMethodDetails = strMethodDetails + gblLineFeed + strEditResults;
				}
			}
			if(boolEditPassed == true && boolSelectPassed == true) {
				//TODO report all the values as output for single step
				//TODO adding to the strMethodDetails with dynamic steps when available.
				strMethodDetails = "Successfully set the 'Filter value(s) of: " + strValueFilter +
						" and selected the item(s): " + strValueItems;
			}
			else {
				//FAil the step and break out of the loop
				boolMethodPassed = false;
				strMethodDetails = 'FAILED!!! on setting the filter and selecting the value!!! See Details: ' +
						gblLineFeed + strMethodDetails;
				break;
			}

		}
		//Check the item(s) displayed as selected.
		if(boolMethodPassed == true) {
			//Return the field element
			let weField = CWCore.returnWebElement(strElemEditListBoxXPath);
			//Process the element
			if (weField != null) {
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weField, 'Edit Field', strElemEditListBoxName);
				}
				//Return the child elements
				let weListBox = CWCore.returnWebElement(strElemEditListBoxXPath + "//ul"); // This looks for a specific UL child
				//Check if null
				if (weListBox != null) {
					//Cannot be highlighted?
					if (TCExecParams.getBoolDoHighlight() == true) {
						let mapHighlight = {}; // Mimic Groovy Map
						mapHighlight = CWCore.objHighlightElementJS(weListBox, 'Edit Field Listbox', strElemEditListBoxName + ' ListBox');
					}
					//Create list of child button //li
					//Return the option list
					let lstChildObjects = weListBox.findElements(By.xpath("./child::*")); // findElements returns a List
					let strFieldValues; // Type inferred
					//Count the options
					let intCntButtons = lstChildObjects.length; // Use .length for JS Array
					if (intCntButtons > 0) {
						let weButton;
						// let weButtonDelete; // Declared but unused
						let strButtonText;
						let strButtonDelText;
						let strButtonItemValue;
						//Create the loop for setting the filter and selecting items
						for (let loopButtons = 0; loopButtons < intCntButtons; loopButtons++) {
							//Return each button an it's text
							weButton = lstChildObjects[loopButtons]; // Access by index
							strButtonText = weButton.getText();
							//Assume the button will have a delete button and only one button
							//Return the delete item text
							strButtonDelText = weButton.findElement(By.xpath(strElemEditListBoxItemBtnDeleteXPath)).getText();
							//Trim the delete text if present
							strButtonItemValue = StringsAndNumbers.JComm_GetLeftTextInString(strButtonText, strButtonDelText);
							//Add the values
							if (loopButtons == 0) {
								strFieldValues = strButtonItemValue;
							}
							else {
								strFieldValues = strFieldValues + gblDelimiter + strButtonItemValue;
							}
						}
						//Compare the values found with the assigned
						if (strFieldValues == strValueItems) { // Direct comparison to strValueItems, not strSelectedItemValues
							strMethodDetails = "Successfully set the field values to: " + strValueItems + '. Details below: ' +
									gblLineFeed + strMethodDetails;
						}
						else {
							boolMethodPassed = false;
							strMethodDetails = "FAILED!!! after setting the filter and selecting the value!!! THE FIELD value of: " + strFieldValues +
									"DOES NOT MATCH THE EXPECTED of: " + strValueItems + "!!! See Details: " + // Assuming strValueItems is the target
									gblLineFeed + strMethodDetails;

						}
					}
				}
				else { // weListBox was null
					boolMethodPassed = false;
					strMethodDetails = 'FAILED!!! after setting the filter and selecting the value!!! THE FIELD DOES NOT CONTAIN A LISTBOX!!! See Details: ' +
							gblLineFeed + strMethodDetails;
				}
			}
			else { // weField was null
				boolMethodPassed = false;
				strMethodDetails = 'FAILED!!! after setting the filter and selecting the value!!! THE FIELD CONTAINING VALUES IS NOT PRESENT!!! See Details: ' +
						gblLineFeed + strMethodDetails;
			}
		}
	}
	else { // cntFilterItems != cntSelectItems OR cntFilterItems == 0
		boolMethodPassed = false;
		strMethodDetails = "FAILED!!!! Unable to process setting filter and selecting results. " +
				"The 'Filter' count: " + cntFilterItems + " DOES NOT MATCH the 'Select' item count of: " + cntSelectItems +
				" OR THERE ARE NO ITEMS IN THE ARRAY!!!";
	}
	//Update the map
	mapResults.boolPassed = boolMethodPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  objSetVerifyEditMultiValueWithUniqueSelectValueFilterListBox  -----------------------------------
* Select an item from a ListBox multi-selectable, where select value is unique from selected value.
*
* --------------------------------------------------NOTICE-----------------------------------------------------------------------------------
* --------------------- strValueItems for the selected value if a linefeed is present must have the GVar.LineFeed value. i.e. Contract Type|False|Contract Type%nIn: Contract Type
* Designed using JTR-12728 as a base. Allows for adding a match value to the strValue assigned allowing the user to select where the value contains the value.
* @param mapInputValues			 The map containing all of the input values for the method
* @param strStartLoc				The web page location value
* @param strValueFilter			 The filter value(s) to set/verify in filter edit.
* @param strValueItems			 The value(s) to select and add to the filter multi-select edit. Consist of SelectValue|BoolMatchSelect|SelectedValue for each item separated by ';'
* @param boolSelctItemExactMatch	The item text must match? true/false
* @param strXPath					 The XPaths for the multiple objects required for the method see below:

* @param strElemEditBoxName		   The edit box name where the filter value is entered
* @param strElemEditBoxXPath		  The edit box XPath where the filter value is entered
* @param strElemEditListBoxName	   The Edit List Name of items added to the edit field once values are selected
* @param strElemEditListBoxXPath	  The Edit List Xpath of items added to the edit field once values are selected
* @param strElemEditListBoxItemBtnName	   The Edit List Button Item Name, is the item element including the text and delete that is a button in the edit box list
* @param strElemEditListBoxItemBtnXPath	  The Edit List Item Xpath, is the item element including the text and delete that is a button in the edit box list
* @param strElemEditListBoxItemTextName	  The Edit List Item Text Name is the item element in the edit list
* @param strElemEditListBoxItemTextXPath	 The Edit List Item Text Xpath is the item element in the edit list
* @param strElemEditListBoxItemBtnDeleteName The Edit List Item Delete Name is the item element delete button in the edit list
* @param strElemEditListBoxItemBtnDeleteXPath The Edit List Item Delete Xpath is the item element delete button in the edit list
* @param strElemListBoxName		   The filter list box name
* @param strElemListBoxXPath		  The filter list box Xpath
* @param strElemListBoxItemName	   The filter list box item name
* @param strElemListBoxItemXPath	  The filter list box item Xpath
* @param strBoolClearValue			Alternate boolean used to clear the filter edit before entering value.By Default == false
* @param strBoolVerifyValue		  Alternate boolean used to determine if the value entered in the filter should be verified.By Default == false
* @param strBoolTabOut				Alternate boolean used to determine if tab is selected after value is entered.By Default == false
* @param strboolDeletePrePopValue	Alternate boolean used to determine if the pre-populated value is first deleted if present.

* @return mapResults				The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 04/08/2022
* @author Last Edited: 11/12/2022
* @author Last Edited By: PGK
* @author Edit Comments: (Include email, date and details)
* @author Updated to meet the requirements for JTR-17320
*/
function CommonWeb_objSetVerifyEditMultiValueWithUniqueSelectValueFilterListBox (mapInputValues) { // mapInputValues is `any`, return type `any`
	//Set the global variables
	let gblNull = GVars.GblNull("Value");
	let gblUndefined = GVars.GblUndefined("Value");
	let gblLineFeed = GVars.GblLineFeed("Value");
	let gblSkipStep = GVars.GblSkip('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	//Create the method variables
	let boolMethodPassed = true;
	let strMethodDetails = gblNull;
	let mapResults = {}; // Mimic Groovy Map
	//Return the map variables
	//Set the name and Xpaths for each of the objects
	let strStartLoc = mapInputValues.strStartLoc;
	let strValueFilter = mapInputValues.strValueFilter;
	let strValueItems = mapInputValues.strValueItems; //Must be in the format of SelectValue|strBoolMatchValue|SelectedValue
	//let boolSelctItemExactMatch = mapInputValues.boolSelctItemExactMatch; // Not used
	let strElemEditBoxName = mapInputValues.strElemEditBoxName;
	let strElemEditBoxXPath = mapInputValues.strElemEditBoxXPath;
	let strElemEditListBoxName = mapInputValues.strElemEditListBoxName;
	let strElemEditListBoxXPath = mapInputValues.strElemEditListBoxXPath;
	let strElemEditListBoxItemBtnName = mapInputValues.strElemEditListBoxItemBtnName;
	let strElemEditListBoxItemBtnXPath = mapInputValues.strElemEditListBoxItemBtnXPath;
	let strElemEditListBoxItemTextName = mapInputValues.strElemEditListBoxItemTextName;
	let strElemEditListBoxItemTextXPath = mapInputValues.strElemEditListBoxItemTextXPath;
	let strElemEditListBoxItemBtnDeleteName = mapInputValues.strElemEditListBoxItemBtnDeleteName;
	let strElemEditListBoxItemBtnDeleteXPath = mapInputValues.strElemEditListBoxItemBtnDeleteXPath;
	let strElemListBoxName = mapInputValues.strElemListBoxName;
	let strElemListBoxXPath = mapInputValues.strElemListBoxXPath;
	let strElemListBoxItemName = mapInputValues.strElemListBoxItemName;
	let strElemListBoxItemXPath = mapInputValues.strElemListBoxItemXPath;

	//Check for keys to set the clear value before entry, verify value and tabout
	let boolClearValue = false;
	let boolVerifyValue = false;
	let boolTabOut = false;
	let boolDeletePrePopValue = false;
	//Update the alternate values if assigned in the map
	//ClearValue
	if (mapInputValues.hasOwnProperty('strBoolClearValue')) {
		boolClearValue = StringsAndNumbers.JComm_StringToBoolean(mapInputValues.strBoolClearValue);
	}
	//VerifyValue
	if (mapInputValues.hasOwnProperty('strBoolVerifyValue')) {
		boolVerifyValue = StringsAndNumbers.JComm_StringToBoolean(mapInputValues.strBoolVerifyValue);
	}
	//Tabout
	if (mapInputValues.hasOwnProperty('strBoolTabOut')) {
		boolTabOut = StringsAndNumbers.JComm_StringToBoolean(mapInputValues.strBoolTabOut);
	}
	//Delete Pre Populated Value
	if (mapInputValues.hasOwnProperty('strboolDeletePrePopValue')) {
		// Note: Original Groovy had boolDeletePrePopValue = StringsAndNumbers.JComm_StringToBoolean(mapInputValues.strBoolTabOut)
		// This looks like a copy-paste error, assuming it should be strboolDeletePrePopValue
		boolDeletePrePopValue = StringsAndNumbers.JComm_StringToBoolean(mapInputValues.strboolDeletePrePopValue);
	}


	//Check the filter values by creating new arrays
	let arryFilterItems;
	let arrySelectItems;
	let mapFilterItems = StringsAndNumbers.JComm_StringToArray(strValueFilter, gblDelimiter);
	arryFilterItems = mapFilterItems.ArryOfValues;
	let cntFilterItems = arryFilterItems.length; // Use .length
	let mapSelectItems = StringsAndNumbers.JComm_StringToArray(strValueItems, ';'); // Split by semicolon
	arrySelectItems = mapSelectItems.ArryOfValues;
	let cntSelectItems = arrySelectItems.length; // Use .length
	if (cntFilterItems == cntSelectItems && cntFilterItems > 0) {
		let strFilterValue;
		let strSelectValue;
		let strSelectedValue; // This value is returned from objPickItemFromListBoxUniqueSelectValue
		//let boolLoopPassed; // Declared but unused
		let boolEditPassed;
		let boolSelectPassed;
		//let strLoopDetails; // Declared but unused
		let strEditResults;
		let strSelectResults;
		let strSelectedItemValues; // Will be built throughout the loop.
		let mapResultsSetEdit; // Type inferred
		let mapResultsSelectItem; // Type inferred
		//Check if we must click the delete button to clear the value. We assume if the button is present that it must be clicked if clear is specified
		if (boolClearValue == true) {
			//Return the element
			let weDelete = CWCore.returnWebElement(strElemEditListBoxItemBtnDeleteXPath);
			if (weDelete != null) {
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weDelete, 'Button', 'edit delete');
				}
				weDelete.click();
				sleep(100); // Mimic Thread.sleep
			}
		}
		//Create the loop for setting the filter and selecting items
		for (let loopArray = 0; loopArray < cntFilterItems; loopArray++) {
			//Return each of the values
			strFilterValue = StringsAndNumbers.JComm_HandleNoData(arryFilterItems[loopArray]).trim(); // Access by index
			strSelectValue = StringsAndNumbers.JComm_HandleNoData(arrySelectItems[loopArray]).trim(); // Access by index
			//Process the EditBox
			mapResultsSetEdit = {}; // Mimic Groovy Map
			mapResultsSetEdit = SetVerifyEditBoxValue(strStartLoc, strElemEditBoxXPath, strElemEditBoxName, strFilterValue, boolClearValue, boolVerifyValue, boolTabOut);
			//Check the results and mark pass or fail
			boolEditPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsSetEdit.boolPassed);
			strEditResults = mapResultsSetEdit.strMethodDetails;
			if (boolEditPassed == true) {
				DateTime.WaitSecs(TCExecParams.getIntViewDelaySecs());
				//Process the selection
				mapResultsSelectItem = {}; // Mimic Groovy Map
				mapResultsSelectItem = objPickItemFromListBoxUniqueSelectValue(strStartLoc, strElemListBoxXPath, strElemListBoxName, strElemListBoxItemXPath, strSelectValue);
				//Check the results and mark pass or fail
				boolSelectPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsSelectItem.boolPassed);
				strSelectResults = mapResultsSelectItem.strMethodDetails;
				strSelectedValue = mapResultsSelectItem.strSelectedValue; // Value now returned from unique select
				strMethodDetails = strMethodDetails + gblLineFeed + strEditResults + gblLineFeed + strSelectResults;
				if (boolSelectPassed == true) {
					if (loopArray == 0) {
						strSelectedItemValues = strSelectedValue;
					}
					else {
						strSelectedItemValues = strSelectedItemValues + gblDelimiter + strSelectedValue;
					}
				}

			}
			else {
				//Report edit failed
				if (strMethodDetails == null) {
					strMethodDetails = strEditResults; // Initialize if null
				}
				else {
					strMethodDetails = strMethodDetails + gblLineFeed + strEditResults;
				}
			}
			if(boolEditPassed == true && boolSelectPassed == true) {
				//TODO report all the values as output for single step
				//TODO adding to the strMethodDetails with dynamic steps when available.
				strMethodDetails = "Successfully set the 'Filter value(s) of: " + strValueFilter +
						" and selected the item(s): " + strValueItems; // Uses strValueItems, not strSelectedItemValues
			}
			else {
				//FAil the step and break out of the loop
				boolMethodPassed = false;
				strMethodDetails = 'FAILED!!! on setting the filter and selecting the value!!! See Details: ' +
						gblLineFeed + strMethodDetails;
				break;
			}

		}
		//Check the item(s) displayed as selected.
		if(boolMethodPassed == true) {
			//Return the field element
			let weField = CWCore.returnWebElement(strElemEditListBoxXPath);
			//Process the element
			if (weField != null) {
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weField, 'Edit Field', strElemEditListBoxName);
				}
				//Return the child elements
				let weListBox = CWCore.returnWebElement(strElemEditListBoxXPath + "//ul"); // This looks for a specific UL child
				//Check if null
				if (weListBox != null) {
					//Cannot be highlighted?
					if (TCExecParams.getBoolDoHighlight() == true) {
						let mapHighlight = {}; // Mimic Groovy Map
						mapHighlight = CWCore.objHighlightElementJS(weListBox, 'Edit Field Listbox', strElemEditListBoxName + ' ListBox');
					}
					//Create list of child button //li
					//Return the option list
					let lstChildObjects = weListBox.findElements(By.xpath("./child::*")); // findElements returns a List
					let strFieldValues; // Type inferred
					//Count the options
					let intCntButtons = lstChildObjects.length; // Use .length for JS Array
					if (intCntButtons > 0) {
						let weButton;
						// let weButtonDelete; // Declared but unused
						let strButtonText;
						let strButtonDelText;
						let strButtonItemValue;
						//Create the loop for setting the filter and selecting items
						for (let loopButtons = 0; loopButtons < intCntButtons; loopButtons++) {
							//Return each button an it's text
							weButton = lstChildObjects[loopButtons]; // Access by index
							strButtonText = weButton.getText();
							//Assume the button will have a delete button and only one button
							//Return the delete item text
							strButtonDelText = weButton.findElement(By.xpath(strElemEditListBoxItemBtnDeleteXPath)).getText();
							//Trim the delete text if present
							strButtonItemValue = StringsAndNumbers.JComm_GetLeftTextInString(strButtonText, strButtonDelText);
							//Add the values
							if (loopButtons == 0) {
								strFieldValues = strButtonItemValue;
							}
							else {
								strFieldValues = strFieldValues + gblDelimiter + strButtonItemValue;
							}
						}
						//Compare the values found with the assigned
						let strTempValue = StringsAndNumbers.JComm_ReplaceLineFeedWithGblLineFeed(strFieldValues);
						if (strTempValue == strSelectedItemValues) { // Comparing built strTempValue with built strSelectedItemValues
							strMethodDetails = "Successfully set the field values to: " + strSelectedItemValues + '. Details below: ' +
									gblLineFeed + strMethodDetails;
						}
						else {
							boolMethodPassed = false;
							strMethodDetails = "FAILED!!! after setting the filter and selecting the value!!! THE FIELD value of: " + strFieldValues +
									"DOES NOT MATCH THE EXPECTED of: " + strSelectedItemValues + "!!! See Details: " +
									gblLineFeed + strMethodDetails;

						}
					}
				}
				else { // weListBox was null
					boolMethodPassed = false;
					strMethodDetails = 'FAILED!!! after setting the filter and selecting the value!!! THE FIELD DOES NOT CONTAIN A LISTBOX!!! See Details: ' +
							gblLineFeed + strMethodDetails;
				}
			}
			else { // weField was null
				boolMethodPassed = false;
				strMethodDetails = 'FAILED!!! after setting the filter and selecting the value!!! THE FIELD CONTAINING VALUES IS NOT PRESENT!!! See Details: ' +
						gblLineFeed + strMethodDetails;
			}
		}
	}
	else { // cntFilterItems != cntSelectItems OR cntFilterItems == 0
		boolMethodPassed = false;
		strMethodDetails = "FAILED!!!! Unable to process setting filter and selecting results. " +
				"The 'Filter' count: " + cntFilterItems + " DOES NOT MATCH the 'Select' item count of: " + cntSelectItems +
				" OR THERE ARE NO ITEMS IN THE ARRAY!!!";
	}
	//Update the map
	mapResults.boolPassed = boolMethodPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  objPickItemFromListBox  -----------------------------------
* Select an item from a ListBox (May be multi-selectable)
* @param strLocation			   The web page location value
* @param strElemFullPath		   The List box full XPath
* @param strElemName			   The meaningful name of the element
* @param strItemXpath			 The XPath for the items
* @param strValueToSelect		  The value to be selected
* @return mapResults			   The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 02/03/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_objPickItemFromListBox (strLocation, strListBoxFullXpath, strListBoxName,
		strItemXpath, strValueToSelect) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'Listbox';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weListBox;	 // Type inferred
	//let strTestObjText; // Declared but unused

	//Return the element
	weListBox = CWCore.returnWebElement(strListBoxFullXpath);
	//Process the element
	if (weListBox != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weListBox, 'ListBox', strListBoxName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsWEState = {}; // Mimic Groovy Map
		mapResultsWEState = CWCore.objVerifyState(weListBox, strListBoxName, true, true);
		//Output results
		let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
		let strStateResults = mapResultsWEState.strMethodDetails;
		//SetVerify the values
		if (boolStatePassed == true) {
			//Attempt to select the value
			let mapSelectListBoxItem = {}; // Mimic Groovy Map
			//TODO create new core for listbox item
			mapSelectListBoxItem = CWCore.objSelectListBoxItemByChild(weListBox, strListBoxName, strItemXpath, strValueToSelect);
			//Output results
			let boolSelectPassed = StringsAndNumbers.JComm_StringToBoolean (mapSelectListBoxItem.boolPassed);
			let strSelectResults = mapSelectListBoxItem.strMethodDetails;
			if (boolSelectPassed == false) {
				boolPassed = false;
			}
			strMethodDetails = strSelectResults;
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strListBoxName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strListBoxName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
* -------------------------------------  objVerfiyListBoxItem  -----------------------------------
* Verify list box item.
*
* NOTE!!!!! provide the zero based index to verify order or ENTER '-1' if ORDER does NOT MATTER.
*
* @param mapListBoxInput	  The map containing the following
* @param strLocation		  The web page location value
* @param strElemFullPath	  The Listbox full XPath
* @param strElemName		  The meaningful name of the element
* @param strItemXpath		 The XPath for the items
* @param strValueToVerify	 The value to be verified
* @param intItemIndex		 The item order index. SET to -1 if order does not matter
* @param intItemCnt		 The number of items expected. Set to -1 if it will be ignored.
*
* @return mapResults		  The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 11/03/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_objVerfiyListBoxItem (mapListBoxInput) { // mapListBoxInput is `any`, return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Retrieve the values from the map
	let strLocation = mapListBoxInput.strLocation;
	let strListBoxFullXpath = mapListBoxInput.strListBoxFullXpath;
	let strListBoxName = mapListBoxInput.strListBoxName;
	let strItemXpath = mapListBoxInput.strItemXpath;
	let strValueToVerify = mapListBoxInput.strValueToVerify;
	let intItemIndex = mapListBoxInput.intItemIndex;
	let intItemCnt = mapListBoxInput.intItemCnt;
	//Create the output values
	// Declare the variables
	let strTestObjType = 'Listbox';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weListBox;	 // Type inferred
	let intVerItemIndex; // Type inferred
	//Return the element
	weListBox = CWCore.returnWebElement(strListBoxFullXpath);
	//Process the element
	if (weListBox != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weListBox, 'ListBox', strListBoxName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsWEState = {}; // Mimic Groovy Map
		mapResultsWEState = CWCore.objVerifyState(weListBox, strListBoxName, true, true);
		//Output results
		let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
		let strStateResults = mapResultsWEState.strMethodDetails;
		//SetVerify the values
		if (boolStatePassed == true) {
			let mapLstItemData = {}; // Mimic Groovy Map
			mapLstItemData.weListBox = weListBox;
			mapLstItemData.strObjName = strListBoxName;
			mapLstItemData.strItemXPath = strItemXpath;
			mapLstItemData.strItemText = strValueToVerify;
			mapLstItemData.intItemIndex = intItemIndex;
			mapLstItemData.intItemCnt = intItemCnt;
			//Attempt to Verify the value(s)
			let mapVerifyListBoxItem = {}; // Mimic Groovy Map
			mapVerifyListBoxItem = CWCore.objVerifyListBoxItem(mapLstItemData);
			//Output results
			boolPassed = StringsAndNumbers.JComm_StringToBoolean (mapVerifyListBoxItem.boolPassed);
			strMethodDetails = mapVerifyListBoxItem.strMethodDetails;
			intVerItemIndex = mapVerifyListBoxItem.intItemIndex;
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strListBoxName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strListBoxName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.intVerItemIndex = intVerItemIndex;
	return mapResults;
}

/**
* -------------------------------------  objPickItemFromListBoxSeperateSelectItem  -----------------------------------
* Select an item from a ListBox using a separate select element
* @param strLocation			   The web page location value
* @param strElemFullPath		   The EditBox full XPath
* @param strElemName			   The meaningful name of the element
* @param strItemXpath			 The XPath for the items
* @param strItemChildSelectXpath   The XPath for the child item that will be used to select a value
* @param strValueToSelect		  The value to be selected
* @return mapResults			   The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 02/03/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_objPickItemFromListBoxSeperateSelectItem (strLocation, strListBoxFullXpath, strListBoxName,
		strItemXpath, strItemChildSelectXpath, strValueToSelect) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'Listbox';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weListBox;	 // Type inferred
	//let strTestObjText; // Declared but unused

	//Return the element
	weListBox = CWCore.returnWebElement(strListBoxFullXpath);
	//Process the element
	if (weListBox != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weListBox, 'ListBox', strListBoxName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsWEState = {}; // Mimic Groovy Map
		mapResultsWEState = CWCore.objVerifyState(weListBox, strListBoxName, true, true);
		//Output results
		let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
		let strStateResults = mapResultsWEState.strMethodDetails;
		//SetVerify the values
		if (boolStatePassed == true) {
			//Attempt to select the value
			let mapSelectListBoxItem = {}; // Mimic Groovy Map
			mapSelectListBoxItem = CWCore.objSelectListBoxItemSeparateSelectElemByChild(weListBox, strListBoxName, strItemXpath, strItemChildSelectXpath, strValueToSelect);
			//Output results
			let boolSelectPassed = StringsAndNumbers.JComm_StringToBoolean (mapSelectListBoxItem.boolPassed);
			let strSelectResults = mapSelectListBoxItem.strMethodDetails;
			if (boolSelectPassed == false) {
				boolPassed = false;
			}
			strMethodDetails = strSelectResults;
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strListBoxName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strListBoxName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  objPickItemFromListBoxUniqueSelectValue  -----------------------------------
* Select an item from a ListBox (Usually multi-selectable)
* @param strLocation			   The web page location value
* @param strElemFullPath		   The EditBox full XPath
* @param strElemName			   The meaningful name of the element
* @param strValueToSelect		  The value to be selected. Consist of SelectValue|BoolMatchSelect|SelectedValue for each item separated by ';'
* @param strItemXpath			 The XPath for the items
* @param boolItemTextMustMatch	 The item text must match? true/false
* @param boolSelectItemByIndex	 Select the value by index? true/false
* @return mapResults			   The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 04/11/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_objPickItemFromListBoxUniqueSelectValue (strLocation, strListBoxFullXpath, strListBoxName,
		strItemXpath, strValueToSelect) { // Return type any
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'Listbox';
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weListBox;	 // Type inferred
	//let strTestObjText; // Declared but unused
	let strSelectValue; // Type inferred
	let strSelectedValue = gblNull; // Initialized to null for now, set later.
	let boolMatchSelectValue; // Type inferred
	//Return the separate values for strValueToSelect
	let mapSelectValues = StringsAndNumbers.JComm_StringToArray(strValueToSelect, gblDelimiter);
	let arrySelectValues = mapSelectValues.ArryOfValues;
	let cntSelectValues = arrySelectValues.length; // Use .length
	if (cntSelectValues == 3) {
		//Assign the values
		strSelectValue = StringsAndNumbers.JComm_HandleNoData(arrySelectValues[0]).trim();
		let strBollMatchSelect = StringsAndNumbers.JComm_HandleNoData(arrySelectValues[1]).trim();
		if (StringsAndNumbers.JComm_StringIsBoolean(strBollMatchSelect) == true) {
			boolMatchSelectValue = StringsAndNumbers.JComm_StringToBoolean(strBollMatchSelect);
		}
		else {
			boolMatchSelectValue = false;
		}
		strSelectedValue = StringsAndNumbers.JComm_HandleNoData(arrySelectValues[2]).trim();
		//Return the element
		weListBox = CWCore.returnWebElement(strListBoxFullXpath);
		//Process the element
		if (weListBox != null) {
			//Highlight
			if (TCExecParams.getBoolDoHighlight() == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weListBox, 'ListBox', strListBoxName);
			}
			//Check the element state (Enabled, Visible)
			let mapResultsWEState = {}; // Mimic Groovy Map
			mapResultsWEState = CWCore.objVerifyState(weListBox, strListBoxName, true, true);
			//Output results
			let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
			let strStateResults = mapResultsWEState.strMethodDetails;
			//SetVerify the values
			if (boolStatePassed == true) {
				//Attempt to select the value
				let mapSelectListBoxItem = {}; // Mimic Groovy Map
				if (boolMatchSelectValue == true) {
					//Select the value based on matching text
					mapSelectListBoxItem = CWCore.objSelectListBoxItemByChild(weListBox, strListBoxName, strItemXpath, strSelectValue);
				}
				else {
					//Select the value based on item contains value
					mapSelectListBoxItem = CWCore.objSelectListBoxItemByChildContains(weListBox, strListBoxName, strItemXpath, strSelectValue);
				}
				//Output results
				let boolSelectPassed = StringsAndNumbers.JComm_StringToBoolean (mapSelectListBoxItem.boolPassed);
				let strSelectResults = mapSelectListBoxItem.strMethodDetails;
				if (boolSelectPassed == false) {
					boolPassed = false;
				}
				strMethodDetails = strSelectResults;
			}
			else
			{
				boolPassed = false;
				strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
						strListBoxName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "The location '" + strLocation + "' '" + strListBoxName + "' DID NOT RETURN an ELEMENT.";
		}
	}
	else { // cntSelectValues != 3 (the value for strValueToSelect)
		boolPassed = false;
		strMethodDetails = "FAILED!!! The assigned value of: '" + strValueToSelect + "' contained '" + cntSelectValues + "' WHICH DOES NOT MATCH THE REQUIRED '3' VALUES!!!";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.strSelectedValue = strSelectedValue; // This value is now correctly set
	return mapResults;
}
/**
* -------------------------------------  objSetVerifyEditSingleValueFilterListBox  -----------------------------------
* Enter a filter value and select one item from the list box
* @param mapInputValues			 The map containing all of the input values for the method as follows:
* @param strStartLoc				The web page location value
* @param strFilterValue			 The filter value(s) to set/verify in filter edit
* @param strSelectValue			 The value(s) to select from the list box
* @param strValstrElemEditBoxName   The name of the filter edit box
* @param strElemEditBoxXPath		The Xpath for the filter edit box
* @param strElemListBoxName		 The name of the list box
* @param strElemListBoxXPath		The list box Xpath
* @param strElemListBoxItemXPath	The XPath for the items
* @param boolItemTextMustMatch	  The item text must match? true/false
* @param boolSelectItemByIndex	  Select the value by index? true/false
* @return mapResults				   The results showing Passed and method details.
*
* @author kpluedde
* @author Created: 02/15/2022
* @author Last Edited: 05/31/2023
* @author Last Edited By:PKanaris
* @author Edit Comments: (Include email, date and details)
* @author Added new variable from map strSelectedValue for the edit box value where the select value does not match the final set value for the edit box.
*/
function CommonWeb_objSetVerifyEditSingleValueFilterListBox (mapInputValues) { // mapInputValues is `any`, return type `any`
	//Set the global variables
	let gblNull = GVars.GblNull("Value");
	let gblUndefined = GVars.GblUndefined("Value");
	let gblLineFeed = GVars.GblLineFeed("Value");
	let gblSkipStep = GVars.GblSkip('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	//Create the method variables
	let boolMethodPassed = true;
	let strMethodDetails = gblNull;
	let mapResults = {}; // Mimic Groovy Map
	//Return the map variables
	//Set the name and Xpaths for each of the objects
	let strStartLoc = mapInputValues.strStartLoc;
	let strFilterValue = mapInputValues.strFilterValue;
	let strSelectValue = mapInputValues.strSelectValue;
	let strElemEditBoxName = mapInputValues.strElemEditBoxName;
	let strElemEditBoxXPath = mapInputValues.strElemEditBoxXPath;
	let strElemListBoxName = mapInputValues.strElemListBoxName;
	let strElemListBoxXPath = mapInputValues.strElemListBoxXPath;
	let strElemListBoxItemXPath = mapInputValues.strElemListBoxItemXPath; // Fixed from original

	//Alternate value to hold the selected value when the select value from list will not match the selected outcome
	let strSelectedValue = gblNull;
	if (mapInputValues.hasOwnProperty('strSelectedValue')) {
		strSelectedValue = mapInputValues.strSelectedValue;
	}
	let boolEditPassed;
	let boolSelectPassed;
	let strEditResults;
	let strSelectResults;
	//let strSelectedItemValues; // Declared but never used

	let mapResultsSetEdit; // Type inferred
	let mapResultsSelectItem; // Type inferred
	let boolTabOut = false;
	let boolElemStale = false;
	if (strSelectValue == gblSkipStep) {
		boolTabOut = true; //Set so we tab out of the field to close the select list.
	}
	//Process the EditBox
	mapResultsSetEdit = {}; // Mimic Groovy Map
	mapResultsSetEdit = SetVerifyEditBoxValue(strStartLoc, strElemEditBoxXPath, strElemEditBoxName, strFilterValue, true, false, boolTabOut);
	//Return the webelement and check if stale
	let weTemp = CWCore.returnWebElementCustomTiming(strElemEditBoxXPath, 5, 100);
	if (weTemp != null) {
		try {
			weTemp.toString(); // arbitrary method call
		}
		catch (e) {
			boolElemStale = true;
		}
	}
	if (boolElemStale == true) {
		boolMethodPassed = false; // Fixed == false to = false
		strMethodDetails = "FAILED!!! The location name: " + strStartLoc + " and element name '" +
		strElemListBoxName + " WAS FOUND TO BE STALE!!!!";
	}
	else {
		//Check the results and mark pass or fail
		boolEditPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsSetEdit.boolPassed);
		strEditResults = mapResultsSetEdit.strMethodDetails;
		if (boolEditPassed == true) {
			//Check if we are selecting
			if (strSelectValue == gblSkipStep) {
				strMethodDetails = 'The user has specified to skip selecting items from the list';
			}
			else {
				//Process the selection
				mapResultsSelectItem = {}; // Mimic Groovy Map
				mapResultsSelectItem = objPickItemFromListBox(strStartLoc, strElemListBoxXPath, strElemListBoxName, strElemListBoxItemXPath, strSelectValue);
				//Check the results and mark pass or fail
				boolSelectPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsSelectItem.boolPassed);
				strSelectResults = mapResultsSelectItem.strMethodDetails;
				strMethodDetails = strMethodDetails + gblLineFeed + strEditResults + gblLineFeed + strSelectResults;
			}
		}
		else {
			//Report edit failed
			if (strMethodDetails == null) {
				strMethodDetails = strEditResults; // Initialize if null
			}
			else {
				strMethodDetails = strMethodDetails + gblLineFeed + strEditResults;
			}
		}
		if(boolEditPassed == true && boolSelectPassed == true) { // If both edit and select parts passed
			//TODO report all the values as output for single step
			//TODO adding to the strMethodDetails with dynamic steps when available.
			strMethodDetails = "Successfully set the 'Filter value(s) of: " + strFilterValue +
					" and selected the item(s): " + strSelectValue;
		}
		else if(boolEditPassed == true && strSelectValue == gblSkipStep){ // If edit passed, and select was skipped
			strMethodDetails = "Successfully set the 'Filter value(s) of: " + strFilterValue + " and 'Skipped' the selection as specified.";
		}
		else { // Otherwise, an explicit failure
			//Fail the step
			boolMethodPassed = false;
			strMethodDetails = 'FAILED!!! on setting the filter and selecting the value!!! See Details: ' +
					gblLineFeed + strMethodDetails;
		}

		//Check the edit field is displaying the assigned valued.
		if(boolMethodPassed == true) {
			let strValueAssigned; // Type inferred
			if (strSelectValue == gblSkipStep) {
				strValueAssigned = strFilterValue;
			}
			else if (strSelectedValue != gblNull) { // Use strSelectedValue if provided and not null
				strValueAssigned = strSelectedValue;
			}
			else {
				strValueAssigned = strSelectValue; // Default to strSelectValue if no specific selected value
			}
			//Verify the displayed values matches the assigned
			let mapVerifySetValue = {}; // Mimic Groovy Map
			mapVerifySetValue = VerifyEditBoxValue(strStartLoc, strElemEditBoxXPath, strElemEditBoxName, strValueAssigned);
			//Check if the verify value passed
			boolMethodPassed = StringsAndNumbers.JComm_StringToBoolean (mapVerifySetValue.boolPassed);
			let strVerifyResults = mapVerifySetValue.strMethodDetails;
			if (boolMethodPassed == true) {
				strMethodDetails = "The location name: " + strStartLoc + " and element name '" +
									strElemListBoxName + "' is set to the specified value of: '" + strSelectValue + "'.";
			}
			else {
				strMethodDetails = "FAILED!!! The location name: " + strStartLoc + " and element name '" +
				strElemListBoxName + "' is NOT SET TO THE ASSIGNED VALUE!!! see details below:" + gblLineFeed + strVerifyResults;
			}
		}
	}
	//Update the map
	mapResults.boolPassed = boolMethodPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  objSetVerifyEditSingleValueFilterListBoxSeperateSelectElement  -----------------------------------
* Enter a filter value and select one item from the list box which has a unique element containing the value and another element containing the selection element
* See JTR-14510 Module: AP Reports Cycle Time Invoice Create to Invoice Export and Invoice Submit to Invoice Export Filter Report
* Step Enter the 'Supplier' filter and select item.
* @param mapInputValues			 The map containing all of the input values for the method as follows:
* @param strStartLoc				The web page location value
* @param strFilterValue			 The filter value(s) to set/verify in filter edit
* @param strSelectValue			 The value(s) to select from the list box
* @param strValstrElemEditBoxName   The name of the filter edit box
* @param strElemEditBoxXPath		The Xpath for the filter edit box
* @param strElemListBoxName		 The name of the list box
* @param strElemListBoxXPath		The list box Xpath
* @param strElemListBoxItemXPath	The XPath for the items that contain text to match in order to determine select item.
* @param strElemListBoxSelectItemXpath   The XPath to be clicked in order to select the item
* @param boolItemTextMustMatch	  The item text must match? true/false
* @param boolSelectItemByIndex	  Select the value by index? true/false
* @return mapResults				   The results showing Passed and method details.
*
* @author kpluedde
* @author Created: 02/15/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_objSetVerifyEditSingleValueFilterListBoxSeperateSelectElement (mapInputValues) { // mapInputValues is `any`, return type `any`
	//Set the global variables
	let gblNull = GVars.GblNull("Value");
	let gblUndefined = GVars.GblUndefined("Value");
	let gblLineFeed = GVars.GblLineFeed("Value");
	let gblSkipStep = GVars.GblSkip('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	//Create the method variables
	let boolMethodPassed = true;
	let strMethodDetails = gblNull;
	let mapResults = {}; // Mimic Groovy Map
	//Return the map variables
	//Set the name and Xpaths for each of the objects
	let strStartLoc = mapInputValues.strStartLoc;
	let strFilterValue = mapInputValues.strFilterValue;
	let strSelectValue = mapInputValues.strSelectValue;
	//let strValstrElemEditBoxName = mapInputValues.strValstrElemEditBoxName; // Declared but unused
	let strElemEditBoxName = mapInputValues.strElemEditBoxName; // Assuming this from docs to be consistent
	let strElemEditBoxXPath = mapInputValues.strElemEditBoxXPath;
	let strElemListBoxName = mapInputValues.strElemListBoxName;
	let strElemListBoxXPath = mapInputValues.strElemListBoxXPath;
	let strElemListBoxItemXPath = mapInputValues.strElemListBoxItemXPath;
	let strElemListBoxSelectItemXpath = mapInputValues.strElemListBoxSelectItemXpath; // This seems to be set to strElemListBoxItemXPath again

	//let boolItemTextMustMatch = mapInputValues.boolItemTextMustMatch; // Not used
	//let boolSelectItemByIndex = mapInputValues.boolSelectItemByIndex; // Not used

	let boolEditPassed;
	let boolSelectPassed;
	let strEditResults;
	let strSelectResults;
	//let strSelectedItemValues; // Declared but unused in current function

	let mapResultsSetEdit; // Type inferred
	let mapResultsSelectItem; // Type inferred
	let boolTabOut = false;
	if (strSelectValue == gblSkipStep) {
		boolTabOut = true; //Set so we tab out of the field to close the select list.
	}
	//Process the EditBox
	mapResultsSetEdit = {}; // Mimic Groovy Map
	mapResultsSetEdit = SetVerifyEditBoxValue(strStartLoc, strElemEditBoxXPath, strElemEditBoxName, strFilterValue, true, false, boolTabOut);
	//Check the results and mark pass or fail
	boolEditPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsSetEdit.boolPassed);
	strEditResults = mapResultsSetEdit.strMethodDetails;
	if (boolEditPassed == true) {
		//Check if we are selecting
		if (strSelectValue == gblSkipStep) {
			strMethodDetails = 'The user has specified to skip selecting items from the list';
		}
		else {
			//Process the selection
			mapResultsSelectItem = {}; // Mimic Groovy Map
			mapResultsSelectItem = objPickItemFromListBoxSeperateSelectItem(strStartLoc, strElemListBoxXPath, strElemListBoxName, strElemListBoxItemXPath, strElemListBoxSelectItemXpath, strSelectValue); // Using this specific function
			//Check the results and mark pass or fail
			boolSelectPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsSelectItem.boolPassed);
			strSelectResults = mapResultsSelectItem.strMethodDetails;
			strMethodDetails = strMethodDetails + gblLineFeed + strEditResults + gblLineFeed + strSelectResults;
		}
	}
	else {
		//Report edit failed
		if (strMethodDetails == null) {
			strMethodDetails = strEditResults; // Initialize if null
		}
		else {
			strMethodDetails = strMethodDetails + gblLineFeed + strEditResults;
		}
	}
	if(boolEditPassed == true && boolSelectPassed == true) {
		//TODO report all the values as output for single step
		//TODO adding to the strMethodDetails with dynamic steps when available.
		strMethodDetails = "Successfully set the 'Filter value(s) of: " + strFilterValue +
				" and selected the item(s): " + strSelectValue;
	}
	else if(boolEditPassed == true && strSelectValue == gblSkipStep){
		strMethodDetails = "Successfully set the 'Filter value(s) of: " + strFilterValue + " and 'Skipped' the selection as specified.";
	}
	else {
		//Fail the step
		boolMethodPassed = false;
		strMethodDetails = 'FAILED!!! on setting the filter and selecting the value!!! See Details: ' +
				gblLineFeed + strMethodDetails;
	}

	//Check the edit field is displaying the assigned valued.
	if(boolMethodPassed == true) {
		let strValueAssigned; // Type inferred
		if (strSelectValue == gblSkipStep) {
			strValueAssigned = strFilterValue;
		}
		else {
			strValueAssigned = strSelectValue;
		}
		//Verify the displayed values matches the assigned
		let mapVerifySetValue = {}; // Mimic Groovy Map
		mapVerifySetValue = VerifyEditBoxValue(strStartLoc, strElemEditBoxXPath, strElemEditBoxName, strValueAssigned);
		//Check if the verify value passed
		boolMethodPassed = StringsAndNumbers.JComm_StringToBoolean (mapVerifySetValue.boolPassed);
		let strVerifyResults = mapVerifySetValue.strMethodDetails;
		if (boolMethodPassed == true) {
			strMethodDetails = "The location name: " + strStartLoc + " and element name '" +
								strElemListBoxName + "' is set to the specified value of: '" + strSelectValue + "'.";
		}
		else {
			strMethodDetails = "FAILED!!! The location name: " + strStartLoc + " and element name '" +
			strElemListBoxName + "' is NOT SET TO THE ASSIGNED VALUE!!! see details below:" + gblLineFeed + strVerifyResults;
		}
	}
	//Update the map
	mapResults.boolPassed = boolMethodPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  VerifyDropDownSelectedValue  -----------------------------------
* Verify the selected item in a drop down.
* @param strLocation		The web page location value
* @param strElemFullPath	The Message Box full XPath
* @param strElementName	 The meaningful name of the Message Box
* @param strValue	   The assigned text displayed
* @oaran boolMatchValue   Match the text exactly? true/false if false must contain the text.
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 02/25/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyDropDownSelectedValue (strLocation, strElemFullPath, strElementName, strValue,
										boolMatchValue, boolVisible, boolEnabled) { // Return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	//Create the output values
	// Declare the variables
	let strTestObjType = 'MsgBox'; // Looks like a copy-paste error from MsgBox
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weDropDown;	// Type inferred
	//Return the element
	weDropDown = CWCore.returnWebElement(strElemFullPath);
	//Process the element
	if (weDropDown != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weDropDown, strTestObjType, strElementName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsDropDownState = {}; // Mimic Groovy Map
		mapResultsDropDownState = CWCore.objVerifyState(weDropDown, strElementName, boolVisible, boolEnabled);
		//Output results
		let boolDropDownStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsDropDownState.boolPassed);
		let strVerDropDownStateResults = mapResultsDropDownState.strMethodDetails;
		//Verify the value
		if (boolDropDownStatePassed == true) {
			let mapVerifyValue = CWCore.objSelectVerifyItemSelectedByOption(weDropDown, strElementName, strValue, boolMatchValue);
			boolPassed = StringsAndNumbers.JComm_StringToBoolean (mapVerifyValue.boolPassed);
			strMethodDetails = mapVerifyValue.strMethodDetails;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strElementName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  SelectItemBasedOnObjectPropertyInParentObject  -----------------------------------
* Select/Click the assigned object based on the match for a object in the parent usually in row element with cells contain 2 or more objects
* The selection is dependent about the property value of Element 1 and then selection of the assigned item. The selection element can be the same as item 1 if needed.
* Developed against JTR-15094
* @param mapSelItemInfo	   The map containing the details required for selection.
* @param strParentName	   The name of the parent object
* @param strParentXPath	   The parent object XPath
* @param strChildRowXPath	   The XPath for the child/row that contains the elements
* @param strElem1Name		 The name of Element 1
* @param strElem1XPath		The XPath of Element 1
* @param strElem1Property	   The property of Element 1 to check
* @param strElem1PropValue	   The property value of Element 1
* @param strElem2Name		 The name of Element 2
* @param strElem2XPath		The XPath of Element 2
*
* @return mapResults		  The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 7/13/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SelectItemBasedOnObjectPropertyInParentObject(mapSelItemInfo) { // mapSelItemInfo is `any`, return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let boolDoDebug = TCExecParams.boolDoDebug;
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails; // Type inferred from assignment
	let boolPassed = true;
	let boolProcSelect = false; // Type inferred
	let intChildCount; // Type inferred
	//Return the values from the map
	let strParentName = mapSelItemInfo.strParentName;
	let strParentXPath = mapSelItemInfo.strParentXPath;
	let strChildRowXPath = mapSelItemInfo.strChildRowXPath;
	let strElem1Name = mapSelItemInfo.strElem1Name;
	let strElem1XPath = mapSelItemInfo.strElem1XPath;
	let strElem1Property = mapSelItemInfo.strElem1Property;
	let strElem1PropValue = mapSelItemInfo.strElem1PropValue;
	let strElem2Name = mapSelItemInfo.strElem2Name;
	let strElem2XPath = mapSelItemInfo.strElem2XPath;
	//Return the element
	let weParent; // Type inferred
	let weChild;  // Type inferred
	let weElem1;  // Type inferred
	let weElem2;  // Type inferred
	let strTempPropValue; // Type inferred
	weParent = CWCore.returnWebElement(strParentXPath);
	if (weParent != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weParent, 'Element', strParentName);
		}
		//Return the child rows and count them
		let lstChildObjs = weParent.findElements(By.xpath(strChildRowXPath)); // findElements returns a List
		intChildCount = lstChildObjs.length; // Use .length for JS Array
		if (intChildCount == 0) {
			//Error no children found
			boolPassed = false;
			strMethodDetails = "FAILED!!! The Parent '" + strParentName + "' RETURNED NO CHILD OBJECTS matching XPath: '" + strChildRowXPath + "'!!!";
		}
		else {
			for (let intChild = 0; intChild < intChildCount; intChild++) {
				weChild = lstChildObjs[intChild]; // Access array by index
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weChild, 'Child Element', strParentName + ' Child');
				}
				weElem1 = weChild.findElement(By.xpath(strElem1XPath)); // findElement returns a WebElement
				if (weElem1 != null) {
					//Highlight
					if (TCExecParams.getBoolDoHighlight() == true) {
						let mapHighlight = {}; // Mimic Groovy Map
						mapHighlight = CWCore.objHighlightElementJS(weElem1, 'Child Element1', strParentName + ' Child1');
					}
					//Return the property value and see if we match
					if (strElem1Property == "text") {
						strTempPropValue = StringsAndNumbers.JComm_HandleNoData(weElem1.getText());
						if (boolDoDebug === true) {
							Tester.Message("The value of the element is: " + strTempPropValue);
						}
					}
					else {
						//Check if the attribute is present
						let boolAttribPresent = CWCore.isAttribtuePresent(weElem1, strElem1Property);
						if (boolAttribPresent == true) {
							strTempPropValue = StringsAndNumbers.JComm_HandleNoData(weElem1.getAttribute(strElem1Property));
						}
						else {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The element '" + strParentName + "' Child object, DOES NOT HAVE THE PROPERTY: " + strElem1Property;
						}
					}
					if (boolPassed == true && strTempPropValue != null && strTempPropValue != gblNull) {
						//Does the value match the assigned value
						if (strTempPropValue == strElem1PropValue) {
							//Return the select element and attempt to click
							weElem2 = weChild.findElement(By.xpath(strElem2XPath)); // findElement returns a WebElement
							if (weElem2 != null) {
								let mapWEClick = {}; // Mimic Groovy Map
								mapWEClick = ClickWebElement(weElem2, strParentName + " child selection", "Element"); // Call to global func
								boolPassed = StringsAndNumbers.JComm_StringToBoolean (mapWEClick.boolPassed);
								strMethodDetails = mapWEClick.strMethodDetails;
								boolProcSelect = true;
							}
							else {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The element '" + strParentName + "' Child object for SELECTION was NOT FOUND BY XPath: " + strElem2XPath;
							}
							break; // Found and processed the matching row, break out of the loop
						}
					}
				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The element '" + strParentName + "' Child object, DID NOT RETURN AN ELEMENT USING XPath: " + strElem1XPath;
				}
			}

		}
	}
	else {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The element '" + strParentName + "' DID NOT RETURN AN ELEMENT USING XPath: " + strParentXPath;
	}
	if (boolProcSelect == false) { // If no element was selected/clicked in the loop
		boolPassed = false;
		strMethodDetails = "FAILED!!! The element '" + strParentName + "' DID NOT RETURN AN ELEMENT which MATCHED the value to select of: " +
		strElem1PropValue + " for the element property: " + strElem1Property + ".";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure this is set for all paths
	return mapResults;
}
/**
* -------------------------------------  VerifyMultiErrorsOnPage  -----------------------------------
* Verify the Error Messages
* @param mapErrorbox		  Contains the object information for the Error box
* @param strLocation		  The web page location value
* @param weErrorbox		 The error box element that will contain the error
* @param strAssocText		 The text associated if any to the error box
* @param strParentXpath	   The Parent XPath usually the page xpath
* @param strChildXPath		The Child XPath which is the generic error message path that will return all errors found on the page.
* @param intExpErrorCnt	   The Error Count expected
* @param strValue			 The error message values containing strFieldName|strElementNaem|strErrorMsgText with each error message seperated by a ';'
* @return mapResults		  The results showing Passed and method details.
*
* @author kplueddemann
* @author Created: 04/06/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyMultiErrorsOnPage (mapSetVerErrorGrpData) { // mapSetVerErrorGrpData is `any`, return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	// Declare the variables
	let strTestObjType = 'errorBox';
	//let strErrorboxGrpFullXpath; // Declared but unused directly
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let intChildCount; // Type inferred
	//Return the values from the map
	let strLocation = mapSetVerErrorGrpData.strLocation;
	let strParentXpath = mapSetVerErrorGrpData.strLocXPath;
	let strChildXPath = mapSetVerErrorGrpData.strErrorMsgXPath;
	let intExpErrorCnt = mapSetVerErrorGrpData.intErrorBoxCnt;
	let strValue = mapSetVerErrorGrpData.strValue;
	let boolErrBoxCntMatchErrorCnt = true; // Default value from Groovy
	if (mapSetVerErrorGrpData.hasOwnProperty('strBoolErrBoxCntMatchErrorCnt')) { // Use hasOwnProperty
		boolErrBoxCntMatchErrorCnt = StringsAndNumbers.JComm_StringToBoolean(mapSetVerErrorGrpData.strBoolErrBoxCntMatchErrorCnt);
	}
	let intVisibleErrCnt = 0; // Type inferred
	//Return the element
	let weParent; // Type inferred
	weParent = CWCore.returnWebElement(strParentXpath);
	if (weParent != null) {
		//Return the count of child elements
		let lstChildObjs = TCObj.tcDriver.findElements(By.xpath(strParentXpath + strChildXPath)); // findElements returns a List
		intChildCount = lstChildObjs.length; // Use .length for JS Array
		if (intChildCount == 0) {
			//Error no messages found
			boolPassed = false; // Fixed `==` to `=` assignment
			strMethodDetails = "FAILED!!! No Child objects found matching XPath: '" + strParentXpath + strChildXPath + "'!!!";
		}
		else if (intChildCount == intExpErrorCnt || boolErrBoxCntMatchErrorCnt == false) {
			if (StringsAndNumbers.JComm_HandleNoData(strValue) == gblNull) {
				//Fail no values
				boolPassed = false; // Fixed `==` to `=` assignment
				strMethodDetails = "FAILED!!! No items found in the assigned value: '" + StringsAndNumbers.JComm_HandleNoData(strValue) + "'!!!";
			}
			else {
				//Place the value into an array and make sure the count is correct
				let mapItemsValues = StringsAndNumbers.JComm_StringToArray(strValue, ';');
				let arryItemValues = mapItemsValues.ArryOfValues;
				let intCntItemValues = arryItemValues.length; // Use .length for JS Array
				if (intCntItemValues <= intChildCount){
					let weMsg; // Type inferred
					let strTempElemID; // Type inferred
					let strTempElemText; // Type inferred
					let strTempAryValue; // Type inferred
					let boolTempVisible; // Type inferred
					//Process the messages
					if (StringsAndNumbers.JComm_HandleNoData(strMethodDetails) == gblNull) {
						strMethodDetails = "Process page error messages: " + gblLineFeed;
					}
					for (let loopErrMsgs = 0; loopErrMsgs < intChildCount; loopErrMsgs ++) {
						//Return the split value into three items

						//Return the errorMsgLine object
						weMsg = lstChildObjs[loopErrMsgs]; // Access array by index
						if (weMsg == null) {
							//Fail since we did not return and object
							boolPassed = false;
							strMethodDetails = "FAILED!!! DID NOT RETURN a CHILD OBJECT at INDEX: " + loopErrMsgs + "!!!";
							break; // Break loop
						}
						else {
							//Return the visible state
							boolTempVisible = weMsg.isDisplayed();
							if (boolTempVisible == true) {
								intVisibleErrCnt ++;
								if (TCExecParams.getBoolDoHighlight() == true) {
									let mapHighlight = {}; // Mimic Groovy Map
									mapHighlight = CWCore.objHighlightElementJS(weMsg, 'Error Message', strLocation);
								}
								//Get the errorMsgLine ID to show which error message it is
								strTempElemID = StringsAndNumbers.JComm_HandleNoData(weMsg.getAttribute("ID"));
								//Get the errorMsgText
								strTempElemText = StringsAndNumbers.JComm_HandleNoData(weMsg.getText());
								//Return the value from the array and split it to the temp values
								strTempAryValue = StringsAndNumbers.JComm_HandleNoData(arryItemValues[intVisibleErrCnt-1]); // Access array by index
								if (strTempAryValue == gblNull) {
									//Fail since we did not return a value
									boolPassed = false;
									strMethodDetails = "FAILED!!! DID NOT RETURN a Value at INDEX: " + loopErrMsgs + "!!!";
									break; // Break loop
								}
								else {
									let mapTempArrValue = StringsAndNumbers.JComm_StringToArray(strTempAryValue, gblDelimiter);
									let arryTempValues = mapTempArrValue.ArryOfValues;
									let intCntTempItemValues = arryTempValues.length; // Use .length for JS Array
									let boolValuesMatch = true; // Assume true initially
									if (intCntTempItemValues == 3) {
										//Are the values correct
										//Message element ID
										if (arryTempValues[1] != strTempElemID) { // Access array by index
											boolValuesMatch = false;
										}
										//Message element text
										if (arryTempValues[2] != strTempElemText) { // Access array by index
											boolValuesMatch = false;
										}
										if (boolValuesMatch == true) {
											strMethodDetails = strMethodDetails + gblLineFeed + arryTempValues[0] + " Field Error Message '" + loopErrMsgs + "' matched the assigned values of: " + strTempAryValue;
										}
										else {
											boolPassed = false;
											strMethodDetails = strMethodDetails + gblLineFeed + arryTempValues[0] + " FAILED Error Message '" + loopErrMsgs + "' DID NOT match the assigned values of: " +
											strTempAryValue + " INSTEAD Element ID: " + strTempElemID + " and Element Text: " + strTempElemText + " are present for the element error message.";
										}
									}
									else { // If intCntTempItemValues is not 3 AND boolValuesMatch is true (from outside previous if)
										// This else block is a bit tricky in Groovy, potentially implying `boolValuesMatch` could be true here,
										// but it's only set to false inside the `if (intCntTempItemValues == 3)` block.
										// Assuming the Groovy `else if` was a typo `else` as boolValuesMatch is not used outside of the scope
										// before the `if (intCntTempItemValues == 3)` condition.
										// Here, it must be the case where intCntTempItemValues is not 3.
										boolPassed = false; // Fail because value did not contain 3 items
										strMethodDetails = "FAILED!!! The temp array value of: ' " + strTempAryValue + "' contained " + intCntTempItemValues +
										" WHICH DOES NOT MATCH THE REQUIRED '3' for temp item array index: " + loopErrMsgs + "!!!";
										break; // Break loop
									}
								}
							}
						}
					}
					//Recheck if the visible matched the expected counts
					if (intVisibleErrCnt != intCntItemValues) {
						boolPassed = false;
						strMethodDetails = "FAILED!!! The EXPECTED ERROR COUNT: " + intCntItemValues + " DOES NOT MATCH THE VISIBLE ERROR COUNT: " + intVisibleErrCnt + "!!!";
					}
				}
				else {
					//Error value item count does not match expected
					boolPassed = false;
					strMethodDetails = "FAILED!!! The The EXPECTED ERROR COUNT: " + intCntItemValues + " DOES NOT MATCH THE DISPLAYED ERROR COUNT: " + intChildCount + "!!!";
				}
			}
		}
		else {
			//Error child count does not match expected
			boolPassed = false;
			strMethodDetails = "FAILED!!! The DISPLAYED ERROR COUNT: " + intChildCount + " DOES NOT MATCH THE EXPECTED ERROR BOX COUNT: " + intExpErrorCnt + "!!!";
		}
	}
	else {
		boolPassed = false; // Fixed `==` to `=` assignment
		strMethodDetails = "FAILED!!! The location '" + strLocation + "' DID NOT RETURN AN ELMENT USING XPath: " + strParentXpath;
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString(); // Using TS object assignment over `put`
	mapResults.strMethodDetails = strMethodDetails; // Ensure this is set for all paths
	return mapResults;
}
//Table methods
//Get row count
/**
* -------------------------------------  tblGetRowCount  -----------------------------------
* Return the number of rows in the table
* @param strXPathForTable   The XPath for the table WebElement
* @param strXPathForRow	 The XPath for the row. Header Rows are usually //thead//tr table body is //tbody/tr
*
* @return intRowCnt		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 04/13/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_tblGetRowCount (strXPathForTable, strXPathForRow) { // Return type `number`
	let intRowCnt = -1;
	let lstTblRows = TCObj.tcDriver.findElements(By.xpath(strXPathForTable + strXPathForRow)); // findElements returns a List
	intRowCnt = lstTblRows.length; // Use .length for JS Array
	return intRowCnt;
}
//Get row column count (Each row may contain a different number of columns
//Return row
//Return cell

/*Treeview Methods
* Based on Workgroups JTR-13492
* Task JTR-14413
*/
/**
* -------------------------------------  ExpandCollapseTreeviewObject  -----------------------------------
*
* Select the link for the treeview item when specified or set the checkbox for the treeview item to the assigned state
* @param mapMethodData			The map containing all of the data values for the method call.
* @param strLocation			The location of the treeview
* @param strMethodData		 The string of objects and values as a pipe delimited value pair seperated by a semicolon.
* @param EXAMPLE data			parentitem1|*N/A*; parentItem2|Expanded
* @param strTreeViewFullPath	 The Treeview full XPath
* @param strTreeViewElemName	 The meaningful name of the TreeView
* @param strTreeViewParentXpath   The parent object XPath. i.e. ul
* @param strTreeViewBranchXpath   The object XPath for the branch i.e. li
* @param strTreeViewElemValue	The delimited value for the branch
* @param boolTreeViewElemIsParent The delimited values is a parent element? true/false
* @param strTreeViewItemLinkXPath   The item link Xpath

* @return mapResults		  The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 08/17/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
* @author NOTE:				   ENTER "*EXPANDALL*" to expand all branches
*								 ENTER "*COLLAPESALL*" to collapse all branches
*/
//TODO update the comments when completed
function CommonWeb_ExpandCollapseTreeviewObject (mapMethodData) { // mapMethodData is `any`, return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	//Create the output values
	// Declare the variables
	//let strTestObjType = 'CheckBox'; // not used
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weTV;		 // Type inferred
	//let strTestObjText; // Declared but unused

	//Return the map values
	let strLocation = mapMethodData.strLocation;
	let strTreeViewFullPath = mapMethodData.strTreeViewFullPath;
	let strTreeViewElemName = mapMethodData.strTreeViewElemName;
	let strValues = mapMethodData.strTreeViewExpCollValue;
	let strTreeItemXpath = mapMethodData.strTreeItemXpath;
	let strTreeItemExpCollElemProp = mapMethodData.strTreeItemExpCollElemProp;
	let strTreeItemTextXPath = mapMethodData.strTreeItemTextXPath;
	let strTreeItemExpXpath = mapMethodData.strTreeItemExpXpath;
	let strTreeItemCollapseXpath = mapMethodData.strTreeItemCollapseXpath;
	let strTreeItemGroupXpath = mapMethodData.strTreeItemGroupXpath;
	let strTreeItemLastProperty = mapMethodData.strTreeItemLastProperty;
	let strTreeItemLastPropValue = mapMethodData.strTreeItemLastPropValue;
	//Return the element
	weTV = CWCore.returnWebElement(strTreeViewFullPath);
	//Process the element
	if (weTV != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weTV, 'TreeView', strTreeViewElemName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsWEState = {}; // Mimic Groovy Map
		mapResultsWEState = CWCore.objVerifyState(weTV, strTreeViewElemName, true, true);
		//Output results
		let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
		let strStateResults = mapResultsWEState.strMethodDetails;
		//SetVerify the values
		if (boolStatePassed == true) {
			//Start Processing the treeview
			//Count the branches off of the root
			let lstChildBranches;
			let intCntRootBranches;
			//let strIFramePropValue; // Declared but unused
			let mapGetWEChildren = {}; // Mimic Groovy Map
			mapGetWEChildren = CWCore.returnChildElements(weTV, "./child::*");
			lstChildBranches = mapGetWEChildren.lstChildObjects;
			intCntRootBranches = mapGetWEChildren.cntChildObjs;
			if (intCntRootBranches > 0 ) {
				//Add code to Expand or Collapse all branches
				if (strValues == "*EXPANDALL*" || strValues == "*COLLAPESALL*") {
					//Process all the branches to either expand or collapse them
					let strTempExpColXpath;
					let strTempElemValue;
					let strTempBranchValue;
					let boolAssgExpanded;
					let boolIsExpanded;
					let boolPropExpPresent = true;
					let weTemp;	  // Type inferred
					let weItemText;  // Type inferred
					let weExpColl;   // Type inferred
					let weItemGrp;   // Type inferred
					let weItem;	  // Type inferred
					let lstChildElems; // Type inferred
					let listChildItemTextElems; // Type inferred
					let intChildElemCnt; // Type inferred
					let intChildItemTextElemCnt; // Type inferred
					let mapGetWEItemChildren = {}; // Mimic Groovy Map
					//Set the assigned expand or collapsed
					if (strValues == "*EXPANDALL*") {
						boolAssgExpanded = true;
					}
					else {
						boolAssgExpanded = false;
					}
					//Process the treeview
					weTemp = weTV; //Assign the base treeview to begin the process
					//let bool; // Declared but unused in Groovy
					//Return the child items from the current element
					mapGetWEChildren = CWCore.returnChildElements(weTemp, "./child::*");
					lstChildElems = mapGetWEChildren.lstChildObjects;
					intChildElemCnt = mapGetWEChildren.cntChildObjs;
					if (intChildElemCnt > 0) {
						//Loop through the current level of items
						for (let loopTVItem = 0; loopTVItem < intChildElemCnt; loopTVItem++) {
							strTempBranchValue = null;
							//Assign the lstChildItem to a temp WE
							weTemp = lstChildElems[loopTVItem]; // Access array by index
							//Highlight
							if (TCExecParams.getBoolDoHighlight() == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weTemp, 'TreeViewItem', strTreeViewElemName);
							}
							//Return the child items from the current element
							mapGetWEItemChildren = CWCore.returnChildElements(weTemp, strTreeItemTextXPath);
							listChildItemTextElems = mapGetWEItemChildren.lstChildObjects;
							intChildItemTextElemCnt = mapGetWEItemChildren.cntChildObjs;
							weItemText = listChildItemTextElems[0]; //We only want the first one since it has the current item text.
							//Highlight
							if (TCExecParams.getBoolDoHighlight() == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weItemText, 'TreeViewItemText', strTreeViewElemName);
							}
							//Check if the assigned property is present
							boolPropExpPresent = CWCore.isAttribtuePresent(weTemp, strTreeItemExpCollElemProp);
							while (boolPropExpPresent == true && boolPassed == true) { // Added boolPassed check
								strTempElemValue = weItemText.getText();
								if (strTempBranchValue == null) {
									strTempBranchValue = strTempElemValue;
								}
								else {
									strTempBranchValue = strTempBranchValue + gblDelimiter + strTempElemValue;
								}
								//Return the expanded and check if we need to change it.
								boolIsExpanded = StringsAndNumbers.JComm_StringToBoolean(weTemp.getAttribute(strTreeItemExpCollElemProp));
								if (boolAssgExpanded != boolIsExpanded) {
									//Return the Exp/Collapse WE and click it
									if (boolIsExpanded == true) {
										strTempExpColXpath = strTreeItemExpXpath;
										weExpColl = CWCore.returnChildElement(weTemp, strTreeItemExpXpath);
									}
									else {
										strTempExpColXpath = strTreeItemCollapseXpath;
										weExpColl = CWCore.returnChildElement(weTemp, strTreeItemCollapseXpath);
									}
									if (weExpColl != null) {
										//Highlight
										if (TCExecParams.getBoolDoHighlight() == true) {
											let mapHighlight = {}; // Mimic Groovy Map
											mapHighlight = CWCore.objHighlightElementJS(weExpColl, 'TreeViewItemExpColl', strTreeViewElemName);
										}
										//Click the element
										weExpColl.click();
										DateTime.WaitSecs(TCExecParams.getIntViewDelaySecs());
										//Check if the branch is in the correct exp/coll state
										boolIsExpanded = StringsAndNumbers.JComm_StringToBoolean(weTemp.getAttribute(strTreeItemExpCollElemProp));
										if (boolAssgExpanded != boolIsExpanded) {
											boolPassed = false;
											strMethodDetails = strMethodDetails + gblLineFeed + "FAILED!!! The '" + strTempBranchValue + "'item Expand/Collapse XPath of: '" + strTreeItemCollapseXpath + "' DID NOT RETURN A WEBELEMENT!!!";
											break; // break from while loop
										}
										else {
											//Check if there is a group
											weItemGrp = CWCore.returnChildElement(weTemp, strTreeItemGroupXpath);
											if (weItemGrp != null) {
												//Highlight
												if (TCExecParams.getBoolDoHighlight() == true) {
													let mapHighlight = {}; // Mimic Groovy Map
													mapHighlight = CWCore.objHighlightElementJS(weItemGrp, 'TreeViewItemGroup', 'Group " + ');
												}
												//Return the item
												weItem = CWCore.returnChildElement(weTemp, strTreeItemXpath);
												if (weItem != null) {
													//Highlight
													if (TCExecParams.getBoolDoHighlight() == true) {
														let mapHighlight = {}; // Mimic Groovy Map
														mapHighlight = CWCore.objHighlightElementJS(weItemGrp, 'TreeViewItemGroup', 'Group " + ');
													}
													weTemp = weItem; // Update weTemp to the child item
												}
											}
											boolPropExpPresent = CWCore.isAttribtuePresent(weTemp, strTreeItemExpCollElemProp);
											if (TCExecParams.getBoolDoDebug() == true) {
												Tester.Message("The Expand Collapse Property present is: " + boolPropExpPresent);
											}
										}
									}
									else {
										boolPassed = false;
										strMethodDetails = strMethodDetails + gblLineFeed + "FAILED!!! The Exp/Collapse Xpath did NOT RETURN and ELEMENT!!!";
										break; // break from while loop
									}
								} else { // boolAssgExpanded == boolIsExpanded
									// If already in desired state, set boolPropExpPresent to false to exit the while loop.
									boolPropExpPresent = false;
								}
							}
						}
					}
				}
				else { // strValues does not contain *EXPANDALL* or *COLLAPSEALL*
					//Check each branch to see if it can be expanded and based on the data expand or collapse the items.
					//Split the input value to begin the processing
					let mapGetItemValues = StringsAndNumbers.JComm_StringToArray(strValues, ';');
					let arryItemValues = mapGetItemValues.ArryOfValues;
					let intItemValueCnt = arryItemValues.length; // Use .length
					if (intItemValueCnt > 0) {
						let mapGetItemData = {}; // Mimic Groovy Map
						let intItemDataCnt;
						let intItemDataIndex;
						let strTempValue;
						let strTempItemValue;
						let strTempElemValue;
						let strTempLastValue;
						let strTempExpColXpath;
						let boolAssgExpanded;
						let boolIsExpanded;
						let arryItemData;
						let weTemp;	 // Type inferred
						let weItemText; // Type inferred
						let weExpColl;  // Type inferred
						let lstChildElems; // Type inferred
						let mapGetWEItemChildren; // Type inferred
						let listChildItemTextElems; // Type inferred
						let intChildElemCnt; // Type inferred
						let intChildItemTextElemCnt; // Type inferred
						//Build a loop for the values
						for (let loopItems = 0; loopItems < intItemValueCnt; loopItems++) {
							strTempValue = arryItemValues[loopItems]; // Access array by index
							//Split each value into the branches and leaves as well as checking the last value is boolean
							mapGetItemData = StringsAndNumbers.JComm_StringToArray(strTempValue, gblDelimiter);
							arryItemData = mapGetItemData.ArryOfValues;
							intItemDataCnt = arryItemData.length; // Use .length
							//Check the number of data items is at least 2
							if (intItemDataCnt > 1) {
								//Process the treeview
								weTemp = weTV; //Assign the base treeview to begin the process
								//Return the last value and verify it is boolean and convert to boolean value
								strTempLastValue = arryItemData[intItemDataCnt - 1]; // Access array by index
								if (StringsAndNumbers.JComm_StringIsBoolean(strTempLastValue) == true) {
									boolAssgExpanded = StringsAndNumbers.JComm_StringToBoolean(strTempLastValue);
									intItemDataIndex = 0;
									while (intItemDataIndex < intItemDataCnt - 1 && boolPassed == true) { // Added boolPassed check
										//Return the child items from the current element
										mapGetWEChildren = CWCore.returnChildElements(weTemp, "./child::*");
										lstChildElems = mapGetWEChildren.lstChildObjects;
										intChildElemCnt = mapGetWEChildren.cntChildObjs;
										if (intChildElemCnt > 0) {
											//Return the value to match
											strTempItemValue = arryItemData[intItemDataIndex]; // Access array by index
											//Loop through the current level of items to match the values
											let foundBranch = false; // Flag to indicate if matching branch is found at current level
											for (let loopTVItem = 0; loopTVItem < intChildElemCnt; loopTVItem++) {
												if (boolPassed == false) { // Exit early if already failed
													break;
												}
												//Assign the lstChildItem to a temp WE
												weTemp = lstChildElems[loopTVItem]; // Access array by index
												//Highlight
												if (TCExecParams.getBoolDoHighlight() == true) {
													let mapHighlight = {}; // Mimic Groovy Map
													mapHighlight = CWCore.objHighlightElementJS(weTemp, 'TreeViewItem', strTreeViewElemName);
												}
												//Return the child items from the current element
												mapGetWEItemChildren = CWCore.returnChildElements(weTemp, strTreeItemTextXPath);
												listChildItemTextElems = mapGetWEItemChildren.lstChildObjects;
												intChildItemTextElemCnt = mapGetWEItemChildren.cntChildObjs;
												weItemText = listChildItemTextElems[0]; //We only want the first one since it has the current item text.
												//Highlight
												if (TCExecParams.getBoolDoHighlight() == true) {
													let mapHighlight = {}; // Mimic Groovy Map
													mapHighlight = CWCore.objHighlightElementJS(weItemText, 'TreeViewItemText', strTreeViewElemName);
												}
												strTempElemValue = weItemText.getText();
												if (strTempItemValue == strTempElemValue) { // Found matching branch at current level
													foundBranch = true;
													//Are we on the same level as the data element
													if (intItemDataIndex == intItemDataCnt - 2) {//Subtracted two since index is zero based and boolean is last value
														//If no more items see if this item can be expanded and if so set the expanded or collapse correctly
														//Check if the assigned property is present
														if (CWCore.isAttribtuePresent(weTemp, strTreeItemExpCollElemProp) == true) {
															//Return the expanded and check if we need to change it.
															boolIsExpanded = StringsAndNumbers.JComm_StringToBoolean(weTemp.getAttribute(strTreeItemExpCollElemProp));
															if (boolAssgExpanded != boolIsExpanded) {
																//Return the Exp/Collapse WE and click it
																if (boolIsExpanded == true) {
																	strTempExpColXpath = strTreeItemExpXpath;
																	weExpColl = CWCore.returnChildElement(weTemp, strTreeItemExpXpath);
																}
																else {
																	strTempExpColXpath = strTreeItemCollapseXpath;
																	weExpColl = CWCore.returnChildElement(weTemp, strTreeItemCollapseXpath);
																}
																if (weExpColl != null) {
																	//Highlight
																	if (TCExecParams.getBoolDoHighlight() == true) {
																		let mapHighlight = {}; // Mimic Groovy Map
																		mapHighlight = CWCore.objHighlightElementJS(weExpColl, 'TreeViewItemExpColl', strTreeViewElemName);
																	}
																	//Click the element
																	weExpColl.click();
																	DateTime.WaitSecs(TCExecParams.getIntViewDelaySecs());
																	intItemDataIndex++; // Move to next level in input data
																	//Refresh the item element (Not strictly needed due to re-assignment in outer loop)
																	//weTemp = lstChildElems[loopTVItem];
																	//Highlight
																	if (TCExecParams.getBoolDoHighlight() == true) {
																		let mapHighlight = {}; // Mimic Groovy Map
																		mapHighlight = CWCore.objHighlightElementJS(weItemText, 'TreeViewItemText', strTreeViewElemName);
																	}
																	boolIsExpanded = StringsAndNumbers.JComm_StringToBoolean(weTemp.getAttribute(strTreeItemExpCollElemProp));
																	if (boolAssgExpanded == boolIsExpanded) {
																		strMethodDetails = strMethodDetails + gblLineFeed + "Clicked on the '" + strTempValue + " item Expand/Collapse element to set the expand/collapse state specified.";
																		//break; // Break from loopTVItem if successful, continue to next input item
																	}
																	else {
																		boolPassed = false;
																		strMethodDetails = strMethodDetails + gblLineFeed + "FAILED!!! The '" + strTempValue + "'item Expand/Collapse XPath of: '" + strTreeItemCollapseXpath + "' DID NOT RETURN A WEBELEMENT!!!";
																		break; // Break from loopTVItem
																	}
																}
																else {
																	boolPassed = false;
																	strMethodDetails = strMethodDetails + gblLineFeed + "FAILED!!! The Exp/Collapse Xpath did NOT RETURN and ELEMENT!!!"; // Changed from 'LAST ITEM VALUE IS NOT BOOLEAN'
																	break; // Break from loopTVItem
																}
															}
															else { // Already in assigned state
																strMethodDetails = strMethodDetails + gblLineFeed + "The item '" + strTempValue + " is in the assigned expand/collapse state.";
																//break; // Break from loopTVItem, continue to next input item
															}
														}
													} else { // Not the last item in the path yet.
														intItemDataIndex++; // Move to next level in input data
														// Fall-through to re-enter while loop with updated weTemp for next level.
													}
													break; // Break from loopTVItem (inner loop) once found matching branch
												}
											}
											if(!foundBranch) { // If after looping all child elements, no match found for strTempItemValue
												boolPassed = false;
												strMethodDetails = strMethodDetails + gblLineFeed + "FAILED!!! Item '" + strTempItemValue + "' not found at this level.";
												break; // Break from while loop.
											}
										}
									}
								}
								else {
									boolPassed = false;
									strMethodDetails = strMethodDetails + gblLineFeed + "FAILED!!! The '" + strTempValue + "' LAST ITEM VALUE IS NOT BOOLEAN!!!";
									break; // Break from loopItems
								}
							}
							else {
								boolPassed = false;
								strMethodDetails = strMethodDetails + gblLineFeed + "FAILED!!! The '" + strValues + "' Item Index: " + loopItems + " DOES NOT CONTAIN AT LEAST '2' ITEMS!!!" +
								"The item values is: " + strTempValue;
								break; // Break from loopItems
							}
						}
					}
					else {
						boolPassed = false;
						strMethodDetails = "FAILED!!! The '" + strValues + "' DOES NOT CONTAIN ANY ITEMS!!!";
					}
				}
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The '" + strTreeViewElemName + "' treeview does not appear to have any child objects using item XPath: " + strTreeItemXpath + "!!!";
			}
		}
		else
		{
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strTreeViewElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
		}
	}
	else
	{
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strTreeViewElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString(); // Using TS object assignment over `put`
	mapResults.strMethodDetails = strMethodDetails; // Ensure this is set for all paths
	return mapResults;
}
/**
* -------------------------------------  SetVerifyTreeViewCheckbox  -----------------------------------
* Set the checkbox for the treeview item to the assigned state. Automatically expand the item if collapsed.
* @param strLocation				The location of the treeview
* @param mapMethodData			The map of objects and values
* @param strTreeViewFullPath		The Treeview full XPath
* @param strTreeViewElemName		The meaningful name of the TreeView
* @param strTreeViewParentXpath   The parent object XPath. i.e. ul
* @param strTreeViewBranchXpath   The object XPath for the branch i.e. li
* @param strTreeViewElemValue	The delimited value for the branch
* @param boolTreeViewElemIsParent The delimited values is a parent element? true/false
* @param strTreeViewItemLinkXPath   The item link Xpath
* @param boolLinkedClicked		The link assigned is clicked? true/false
* @param strTreeViewCheckboxXPath   The checkbox path. Set to null if no checkbox is present
* @param boolVerifyChecked		Verify the checked state? true/false

* @return mapResults		  The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 08/17/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SetVerifyTreeViewCheckbox (mapMethodData) { // mapMethodData is `any`, return `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	//Create the output values
	// Declare the variables
	//let strTestObjType = 'CheckBox'; // not used
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weTV;		 // Type inferred
	//let strTestObjText; // Declared but unused

	//Return the map values
	let strLocation = mapMethodData.strLocation;
	let strTreeViewFullPath = mapMethodData.strTreeViewFullPath;
	let strTreeViewElemName = mapMethodData.strTreeViewElemName;
	let strValues = mapMethodData.strTreeViewCkBoxValues;
	let boolVerChecked = mapMethodData.boolVerChecked;
	let strTreeItemXpath = mapMethodData.strTreeItemXpath;
	let strTreeItemExpCollElemProp = mapMethodData.strTreeItemExpCollElemProp;
	let strTreeItemTextXPath = mapMethodData.strTreeItemTextXPath;
	let strTreeItemExpXpath = mapMethodData.strTreeItemExpXpath;
	let strTreeItemCollapseXpath = mapMethodData.strTreeItemCollapseXpath;
	let strTreeItemChkBoxXpath = mapMethodData.strTreeItemChkBoxXpath;
	let strTreeItemGroupXpath = mapMethodData.strTreeItemGroupXpath;
	let strTreeItemLastProperty = mapMethodData.strTreeItemLastProperty;
	let strTreeItemLastPropValue = mapMethodData.strTreeItemLastPropValue;
	//Return the element
	weTV = CWCore.returnWebElement(strTreeViewFullPath);
	//Process the element
	if (weTV != null) {
		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weTV, 'TreeView', strTreeViewElemName);
		}
		//Check the element state (Enabled, Visible)
		let mapResultsWEState = {}; // Mimic Groovy Map
		mapResultsWEState = CWCore.objVerifyState(weTV, strTreeViewElemName, true, true);
		//Output results
		let boolStatePassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsWEState.boolPassed);
		let strStateResults = mapResultsWEState.strMethodDetails;
		//SetVerify the values
		if (boolStatePassed == true) {
			//Start Processing the treeview
			//Count the branches off of the root
			let lstChildBranches;
			let intCntRootBranches;
			//let strIFramePropValue; // Declared but unused
			let mapGetWEChildren = {}; // Mimic Groovy Map
			mapGetWEChildren = CWCore.returnChildElements(weTV, "./child::*");
			lstChildBranches = mapGetWEChildren.lstChildObjects;
			intCntRootBranches = mapGetWEChildren.cntChildObjs;
			if (intCntRootBranches > 0 ) {
				//Check each branch to see if it can be expanded and based on the data expand or collapse the items.
				//Split the input value to begin the processing
				let mapGetItemValues = StringsAndNumbers.JComm_StringToArray(strValues, ';');
				let arryItemValues = mapGetItemValues.ArryOfValues;
				let intItemValueCnt = arryItemValues.length; // Use .length
				if (intItemValueCnt > 0) {
					let mapGetItemData = {}; // Mimic Groovy Map
					let intItemDataCnt;
					let intItemDataIndex;
					let strTempValue;
					let strTempItemValue;
					let strTempElemValue;
					let strTempLastValue;
					let strTempExpColXpath;
					let boolAssgExpanded;
					let boolIsExpanded;
					let boolAssgChecked;
					// let boolIsLastItemValue; // Declared but unused
					let arryItemData;
					let weTemp; // Type inferred
					let weItemText; // Type inferred
					let weExpColl; // Type inferred
					let weCkBox; // Type inferred
					let lstChildElems; // Type inferred
					let mapGetWEItemChildren; // Type inferred
					let listChildItemTextElems; // Type inferred
					let listChildItemCkBoxElems; // Type inferred
					let intChildElemCnt; // Type inferred
					let intChildItemTextElemCnt; // Type inferred
					let intChildItemCkBoxElemCnt; // Type inferred

					boolAssgExpanded = true; //We will always expand the element if not the last item
					//Build a loop for the values
					for (let loopItems = 0; loopItems < intItemValueCnt; loopItems++) {
						strTempValue = arryItemValues[loopItems]; // Access array by index
						//Split each value into the branches and leaves as well as checking the last value is boolean
						mapGetItemData = StringsAndNumbers.JComm_StringToArray(strTempValue, gblDelimiter);
						arryItemData = mapGetItemData.ArryOfValues;
						intItemDataCnt = arryItemData.length; // Use .length
						//Check the number of data items is at least 2
						if (intItemDataCnt > 1) {
							//Process the treeview
							weTemp = weTV; //Assign the base treeview to begin the process
							//Return the last value and verify it is boolean and convert to boolean value
							strTempLastValue = arryItemData[intItemDataCnt - 1]; // Access array by index
							if (StringsAndNumbers.JComm_StringIsBoolean(strTempLastValue) == true) {
								boolAssgChecked = StringsAndNumbers.JComm_StringToBoolean(strTempLastValue);
								intItemDataIndex = 0;
								while (intItemDataIndex < intItemDataCnt - 1 && boolPassed == true) { // Added boolPassed check
									//Return the child items from the current element
									mapGetWEChildren = CWCore.returnChildElements(weTemp, "./child::*");
									lstChildElems = mapGetWEChildren.lstChildObjects;
									intChildElemCnt = mapGetWEChildren.cntChildObjs;
									if (intChildElemCnt > 0) {
										//Return the value to match
										strTempItemValue = arryItemData[intItemDataIndex]; // Access array by index
										//Loop through the current level of items to match the values
										let foundItem = false;
										for (let loopTVItem = 0; loopTVItem < intChildElemCnt; loopTVItem++) {
											if (boolPassed == false) { // Exit early if already failed
												break;
											}
											//Assign the lstChildItem to a temp WE
											weTemp = lstChildElems[loopTVItem]; // Access array by index
											//Highlight
											if (TCExecParams.getBoolDoHighlight() == true) {
												let mapHighlight = {}; // Mimic Groovy Map
												mapHighlight = CWCore.objHighlightElementJS(weTemp, 'TreeViewItem', strTreeViewElemName);
											}
											//Return the child items from the current element
											mapGetWEItemChildren = CWCore.returnChildElements(weTemp, strTreeItemTextXPath);
											listChildItemTextElems = mapGetWEItemChildren.lstChildObjects;
											intChildItemTextElemCnt = mapGetWEItemChildren.cntChildObjs;
											weItemText = listChildItemTextElems[0]; //We only want the first one since it has the current item text.
											//TODO should we verify the exlement exist
											//Highlight
											if (TCExecParams.getBoolDoHighlight() == true) {
												let mapHighlight = {}; // Mimic Groovy Map
												mapHighlight = CWCore.objHighlightElementJS(weItemText, 'TreeViewItemText', strTreeViewElemName);
											}
											strTempElemValue = weItemText.getText();
											if (strTempItemValue == strTempElemValue) { // Found matching element for current level
												foundItem = true;
												//Are we on the same level as the data element
												if (intItemDataIndex == intItemDataCnt - 2) {//Subtracted two since index is zero based and boolean is last value
													//We are at the same level return the checkbox and set to the assigned state.
													intItemDataIndex++; // Move to next level in input data after processing checkbox
													//Return the child items from the current element
													mapGetWEItemChildren = CWCore.returnChildElements(weTemp, strTreeItemChkBoxXpath);
													listChildItemCkBoxElems = mapGetWEItemChildren.lstChildObjects;
													intChildItemCkBoxElemCnt = mapGetWEItemChildren.cntChildObjs;
													weCkBox = listChildItemCkBoxElems[0]; //We only want the first one since it has the current item checkbox.
													if (weCkBox != null) {
														//Highlight
														if (TCExecParams.getBoolDoHighlight() == true) {
															let mapHighlight = {}; // Mimic Groovy Map
															mapHighlight = CWCore.objHighlightElementJS(weCkBox, 'TreeViewItemCheckBox', strTreeViewElemName);
														}
														//SetVerfiy the checkbox
														let mapSetVerCkBox = {}; // Mimic Groovy Map
														mapSetVerCkBox = SetVerifyCheckBoxWebElement(strLocation, weCkBox, "TreeViewItem:" + arryItemValues[loopItems], boolAssgChecked, boolVerChecked); // Call to global func
														//Verify set/verify passed
														let boolSetVerPassed = StringsAndNumbers.JComm_StringToBoolean(mapSetVerCkBox.boolPassed);
														//Always add the output
														strMethodDetails = strMethodDetails + gblLineFeed + mapSetVerCkBox.strMethodDetails;
														if (boolSetVerPassed == false) {
															boolPassed = false;
															break; // break from loopTVItem
														}
													} else { // Checkbox not found
														boolPassed = false;
														strMethodDetails = strMethodDetails + gblLineFeed + "FAILED!!! Checkbox for item '" + strTempItemValue + "' not found.";
														break; // break from loopTVItem
													}
													break; // Move to the next value (break from loopTVItem)
												}
												else { // Not the last item in the path (need to expand to go deeper)
													intItemDataIndex++; // Move to next level in input data
													//Else we will need to auto expand to get to the next level
													//If no more items see if this item can be expanded and if so set the expanded or collapse correctly
													//Check if the assigned property is present
													if (CWCore.isAttribtuePresent(weTemp, strTreeItemExpCollElemProp) == true) {
														//Return the expanded and check if we need to change it.
														boolIsExpanded = StringsAndNumbers.JComm_StringToBoolean(weTemp.getAttribute(strTreeItemExpCollElemProp));
														if (boolAssgExpanded != boolIsExpanded) {
															//Return the Exp/Collapse WE and click it
															if (boolIsExpanded == true) {
																strTempExpColXpath = strTreeItemExpXpath;
																weExpColl = CWCore.returnChildElement(weTemp, strTreeItemExpXpath);
															}
															else {
																strTempExpColXpath = strTreeItemCollapseXpath;
																weExpColl = CWCore.returnChildElement(weTemp, strTreeItemCollapseXpath);
															}
															if (weExpColl != null) {
																//Highlight
																if (TCExecParams.getBoolDoHighlight() == true) {
																	let mapHighlight = {}; // Mimic Groovy Map
																	mapHighlight = CWCore.objHighlightElementJS(weExpColl, 'TreeViewItemExpColl', strTreeViewElemName);
																}
																//Click the element
																weExpColl.click();
																DateTime.WaitSecs(TCExecParams.getIntViewDelaySecs());
																//intItemDataIndex++; // This would double increment if placed here after success. It is already incremented once above.
																//Refresh the item element (re-get weTemp if it refers to context before click)
																//weTemp = lstChildElems[loopTVItem];
																//Highlight
																if (TCExecParams.getBoolDoHighlight() == true) {
																	let mapHighlight = {}; // Mimic Groovy Map
																	mapHighlight = CWCore.objHighlightElementJS(weItemText, 'TreeViewItemText', strTreeViewElemName);
																}
																boolIsExpanded = StringsAndNumbers.JComm_StringToBoolean(weTemp.getAttribute(strTreeItemExpCollElemProp));
																if (boolAssgExpanded == boolIsExpanded) {
																	strMethodDetails = strMethodDetails + gblLineFeed + "Clicked on the '" + strTempValue + " item Expand/Collapse element to set the expand/collapse state specified.";
																	break; // Break from loopTVItem (found and expanded), proceed to next level in while loop
																}
																else {
																	boolPassed = false;
																	strMethodDetails = strMethodDetails + gblLineFeed + "FAILED!!! The '" + strTempValue + "'item Expand/Collapse XPath of: '" + strTreeItemCollapseXpath + "' DID NOT RETURN A WEBELEMENT!!!"; // XPath might be wrong for expanded vs collapsed
																	break; // Break from loopTVItem
																}
															}
															else {
																boolPassed = false;
																strMethodDetails = strMethodDetails + gblLineFeed + "FAILED!!! The Exp/Collapse Element was NULL for '" + strTempItemValue + "'!!!"; // Changed error message
																break; // Break from loopTVItem
															}
														}
														else { // Already in assigned state
															strMethodDetails = strMethodDetails + gblLineFeed + "The item '" + strTempValue + " is in the assigned expand/collapse state.";
															break; // Break from loopTVItem (found and already expanded), proceed to next level in while loop
														}
													}
												}
											} // End of inner loop (loopTVItem)
											if (!foundItem) { // If no matching item was found for strTempItemValue
												boolPassed = false;
												strMethodDetails = strMethodDetails + gblLineFeed + "FAILED!!! Item '" + strTempItemValue + "' not found at this level.";
												break; // Break from while loop.
											}
										}
									}
									else { // No child elements for the current branch
										boolPassed = false;
										strMethodDetails = strMethodDetails + gblLineFeed + "FAILED!!! No child elements found for '" + strTempItemValue + "' at this level.";
										break; // Break from while loop.
									}
								}
							}
							else { // Last value (boolean) is not boolean
								boolPassed = false;
								strMethodDetails = strMethodDetails + gblLineFeed + "FAILED!!! The '" + strTempValue + "' LAST ITEM VALUE IS NOT BOOLEAN!!!";
								break; // Break from loopItems
							}
						} // End loopItems
						else { // intItemDataCnt <= 1 (not enough items in value string)
							boolPassed = false;
							strMethodDetails = strMethodDetails + gblLineFeed + "FAILED!!! The '" + strValues + "' Item Index: " + loopItems + " DOES NOT CONTAIN AT LEAST '2' ITEMS!!!" +
							"The item values is: " + strTempValue;
							break; // Break from loopItems
						}
					}
				}
				else { // arryItemValues is empty
					boolPassed = false;
					strMethodDetails = "FAILED!!! The '" + strValues + "' DOES NOT CONTAIN ANY ITEMS!!!";
				}
			} // End if (intCntRootBranches > 0) else
			else { // No root branches found
				boolPassed = false;
				strMethodDetails = "FAILED!!! The '" + strTreeViewElemName + "' treeview does not appear to have any child objects using item XPath: " + strTreeItemXpath + "!!!";
			}
		}
		else { // weTV state not good
			boolPassed = false;
			strMethodDetails = "FAILED!!! The location name: " + strLocation + " and element name '" +
					strTreeViewElemName + "' is NOT in the CORRECT STATE!!! see details below:" + gblLineFeed + strStateResults;
		}
	}
	else { // weTV is null
		boolPassed = false;
		strMethodDetails = "The location '" + strLocation + "' '" + strTreeViewElemName + "' DID NOT RETURN an ELEMENT.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure this is set for all paths
	return mapResults;
}

/* Tab Objects */
/**
* -------------------------------------  OutputTabData  -----------------------------------
* Output the tab data using the mapped values
* @param 		mapInputValues 			The input values used in the method
* @param 		strStartLocation 		The assigned feature start location
* @param 		strOutputSheetName 		The name of the output sheet
* @param 		strTabControlName 		The tab control element name
* @param 		strTabControlXpath 		The Xpath for the tab control
* @param 		strTabLevelXPath 		The Xpath defining a row level of tabs. Set to gblNA if not applicable
* @param 		strTabLevelProp 		The attribute name to return the tab level value. Set to gblNA if not applicable
* @param 		strTabLevelValue 		The tab level value used to determine the level number or id. Usually consist of a expression and value combination. For example: '%Right%|tabset'
* @param 		strTabLevelNoTabMsg 	Message text displayed if no tabs are present. Set to gblNA if not applicable
* @param 		strTabItemXpath 		The Xpath for the tab item. Should containing a leading '.' to find only sub children
* @param 		strTabSelectElemXPath 	The Xpath for the tab element which contains the Selected state. Set to gblNA if the tab element has the attribute and not another object.
* @param 		strTabSelectedProperty 	The attribute name to return the tab level value.
* @param 		strTabSelectedValue 	The tab selected value used to determine the selected state. Usually consist of a expression and value combination. For example: '%Cont%|selected'
* @param 		strTabImageXpath 		The Xpath for the image within the tab element. Set to gblNA if not applicable
* @param 		boolTabImageHasText 	Does the image contain text that we will need to return? True/False frequently must be parsed from the tab text to show on the tab text.
* @param 		strImageAttribute		The attribute name to return the value used to parse and determine the image that is present.
* @param 		strImageValue			The image value used to parse the Image specified attribute and determine the image. Usually consist of a expression and value combination. For example: '%Right%|*GBLSPACE*'
*
* @return mapResults 		The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 01/19/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_OutputTabData (mapInputValues) { // mapInputValues is `any`, return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value');
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	let boolDoHighlight = TCExecParams.getBoolDoHighlight();
	//let intViewDelay = TCExecParams.getIntViewDelaySecs(); // Declared but unused
	//Declare the variables
	let mapResults = {}; // Mimic Groovy Map
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	//Create variables and assign values from the mapInputValues
	let strStartLocation = mapInputValues.strStartLocation;
	let strOutputSheetName = mapInputValues.strOutputSheetName;
	let strTabControlName = mapInputValues.strTabControlName;
	let strTabControlXpath = mapInputValues.strTabControlXpath;
	let strTabLevelXPath = mapInputValues.strTabLevelXPath;
	let strTabLevelProp = mapInputValues.strTabLevelProp;
	let strTabLevelValue = mapInputValues.strTabLevelValue;
	let strTabLevelNoTabMsg = mapInputValues.strTabLevelNoTabMsg;
	let strTabItemXpath = mapInputValues.strTabItemXpath;
	let strTabSelectElemXPath = mapInputValues.strTabSelectElemXPath;
	let strTabSelectedProperty = mapInputValues.strTabSelectedProperty;
	let strTabSelectedValue = mapInputValues.strTabSelectedValue;
	let strTabImageXpath = mapInputValues.strTabImageXpath;
	let boolTabImageHasText = mapInputValues.boolTabImageHasText;
	let strImageAttribute = mapInputValues.strImageAttribute;
	let strImageValue = mapInputValues.strImageValue;
	//Return the tab control element
	let weTabCtrl = CWCore.returnWebElement(strTabControlXpath);
	if (weTabCtrl == null ) {
		boolPassed = false;
		strMethodDetails = "The location '" + strStartLocation + "' '" + strTabControlName + "' USING XPath: " + strTabControlXpath + " DID NOT RETURN an ELEMENT.";
	}
	else {
		//Process the tab control
		//Define the variables for the output
		let strTabValue;
		let strTabSelValue;
		let strTabImage;
		let strTabImageValue;
		let strTabRow;
		let strTabOrder;
		//let TabSelected; // Declared but unused

		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weTabCtrl, 'TabControl', strTabControlName);
		}
		let weTabRow;
		let intTabRowElemCnt;
		let lstTabRowElements;
		//Return the tab levels
		if (strTabLevelXPath == gblNA) {
			intTabRowElemCnt = 1;
		}
		else {
			let mapGetTabRowElements = {}; // Mimic Groovy Map
			mapGetTabRowElements = CWCore.returnChildElements(weTabCtrl, strTabLevelXPath);
			lstTabRowElements = mapGetTabRowElements.lstChildObjects;
			intTabRowElemCnt = mapGetTabRowElements.cntChildObjs;
			if (intTabRowElemCnt < 1 && strTabLevelXPath != gblNA) {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The '" + strTabControlName + "' DOES NOT contain any Tab Levels!!!";
			}
		}
		if (boolPassed == true) {
			//Declare the variables
			let intOutputRow = 0;
			let mapGetTabElements = {}; // Mimic Groovy Map
			let lstTabElements;
			let intTabElemCnt;
			let intImgColID;
			let strCellTheme = 'OutputDataCentered'; //Set to output always
			let mapSetCellValue;
			let strColor;
			let strBckgColor;
			let boolCustomTheme;
			//Process the tab levels
			for (let loopTabRows = 0; loopTabRows < intTabRowElemCnt; loopTabRows++) {
				strTabRow = gblNull;
				if (strTabLevelXPath == gblNA) {
					weTabRow = weTabCtrl;
				}
				else {
					weTabRow = lstTabRowElements[loopTabRows]; // Access array by index
				}
				//Highlight
				if (TCExecParams.getBoolDoHighlight() == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weTabRow, 'TabRow' + loopTabRows , strTabControlName);
				}
				//Return the count of tab elements
				mapGetTabElements = CWCore.returnChildElements(weTabRow, strTabItemXpath);
				lstTabElements = mapGetTabElements.lstChildObjects;
				intTabElemCnt = mapGetTabElements.cntChildObjs;
				if (intTabElemCnt < 1) {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The '" + strTabControlName + "' TabRow" + loopTabRows + " DOES NOT contain any TABS!!!";
				}
				else {
					//Check the tab level
					//Return the level if applicable
					if (strTabLevelProp == gblNull ||strTabLevelProp == gblNA) {
						strTabRow = gblNA;
					}
					else {
						let mapSplitElementString = {}; // Mimic Groovy Map
						mapSplitElementString = StringsAndNumbers.JComm_StringToArray(strTabLevelValue, gblDelimiter);
						let intValueItemCnt = StringsAndNumbers.JComm_StringToInteger(mapSplitElementString.intItemCount);
						if (intValueItemCnt == 2) {
							let arryTabLvlValues = mapSplitElementString.ArryOfValues;
							//Determine the row number by attribute value
							//Is the attribute present
							if (CWCore.isAttribtuePresent(weTabRow, strTabLevelProp) == true) {
								//Return the value
								let strTabLvlPropValue = StringsAndNumbers.JComm_HandleNoData(weTabRow.getAttribute(strTabLevelProp));
								//Return the level value
								let mapReturnTabLvl = {}; // Mimic Groovy Map
								mapReturnTabLvl = StringsAndNumbers.JComm_ReturnValueFromString(strTabLvlPropValue, arryTabLvlValues[1], arryTabLvlValues[0]);
								if (StringsAndNumbers.JComm_StringToBoolean (mapReturnTabLvl.boolTextIsPresent) == true) {
									strTabRow = (StringsAndNumbers.JComm_StringToInteger(mapReturnTabLvl.strValue) + 1).toString(); //Assumes the tabs are zero index.
								}
								else {
									boolPassed = false;
									strMethodDetails = mapReturnTabLvl.strMethodDetails;
								}
							}
							else {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The Tab Level ATTRIBUTE of: " + strTabLevelProp + " DOES NOT EXIST for the ELEMENT!!!";
							}
						}
						else {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The Tab Level Value of: " + strTabLevelValue + " DOES NOT CONTAIN TWO VALUES REQUIRED!!!";
						}
					}
					if (boolPassed == true) {
						//Process the current tab level and output the results for each tab.
						//Declare the variables
						let weTab;
						let weTabElem; // Declared but unused
						let strActiveTab = gblNull; // Declared but unused
						// let boolTabSelected; // Declared but unused
						// let boolHasImage; // Declared but unused

						for (let loopTabItems = 0; loopTabItems < intTabElemCnt; loopTabItems++) {
							//Set the output row
							intOutputRow++;
							strColor = gblNull;
							strBckgColor = gblNull;
							boolCustomTheme = false;
							intImgColID = -1;
							//Set the
							strTabValue = gblNull;
							strTabSelValue = gblNull;
							strTabImage = gblNull;
							strTabImageValue = gblNull;
							strTabOrder = (loopTabItems + 1).toString(); //Zero index converted to 1 based for users.
							strActiveTab = gblNull; // Also re-declared in loop
							weTab = lstTabElements[loopTabItems]; // Access array by index
							//Highlight
							if (TCExecParams.getBoolDoHighlight() == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weTab, 'TabRow' + loopTabRows + " Tab " + loopTabItems, strTabControlName);
							}
							//Return the tab value
							strTabValue = StringsAndNumbers.JComm_HandleNoData(weTab.getText()); //Return the text for the tab
							//Is the tab selected
							let mapSplitElementString = {}; // Mimic Groovy Map
							mapSplitElementString = StringsAndNumbers.JComm_StringToArray(strTabSelectedValue, gblDelimiter);
							let intValueItemCnt = StringsAndNumbers.JComm_StringToInteger(mapSplitElementString.intItemCount);
							if (intValueItemCnt == 2) {
								let arryTabSelValues = mapSplitElementString.ArryOfValues;
								//Check if the property is present
								if (CWCore.isAttribtuePresent(weTab, strTabSelectedProperty) == true) {
									//Return the value
									let strTabSelPropValue = StringsAndNumbers.JComm_HandleNoData(weTab.getAttribute(strTabSelectedProperty));
									//Return the level value
									let boolTextPresent = StringsAndNumbers.JComm_VerifyTextPresent(strTabSelPropValue, arryTabSelValues[1], arryTabSelValues[0]);
									strTabSelValue = boolTextPresent.toString();
								}
								else {
									boolPassed = false;
									strMethodDetails = "FAILED!!! The Tab Level ATTRIBUTE of: " + strTabSelectedProperty + " DOES NOT EXIST for the TAB ELEMENT!!!";
								}
							}
							else if(strTabSelectedProperty == 'activetab') {
								let strCurrTabID;
								//Custom code for the PhoenixNavLink tab control found in requisitions
								let weTemp = CWCore.returnWebElement(strTabSelectElemXPath);
								if (weTemp == null) {
									boolPassed = false;
									strMethodDetails = "FAILED!!! The Xpath for the element of: " + strTabSelectElemXPath + "DID NOT RETURN and ELEMENT for the ActiveTab validation!!!";
								}
								//Check if the property is present
								if (CWCore.isAttribtuePresent(weTemp, strTabSelectedProperty) == true) {
									//Return the value
									let strTabSelPropValue = StringsAndNumbers.JComm_HandleNoData(weTemp.getAttribute(strTabSelectedProperty)); //Will be the ID of the tab
									if (CWCore.isAttribtuePresent(weTab, 'id') == true) {
										strCurrTabID = StringsAndNumbers.JComm_HandleNoData(weTab.getAttribute('id'));
									}
									else {
										strCurrTabID = StringsAndNumbers.JComm_HandleNoData(weTab.findElement(By.xpath("./*")).getAttribute('id')); //Developed to fit the purchase order
									}
									if (strTabSelPropValue == strCurrTabID) {
										strTabSelValue = 'true';
									}
									else {
										strTabSelValue = 'false';
									}
								}
								else {
									boolPassed = false;
									strMethodDetails = "FAILED!!! The Tab ATTRIBUTE of: " + strTabSelectedProperty + " DOES NOT EXIST for the 'ACTIVETAB' for the TAB ELEMENT!!!";
								}
							}
							else {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The Tab Selected VALUE of: " + strTabSelectedValue + " DOES NOT CONTAIN TWO VALUES REQUIRED!!!";
							}
							//Process the image
							if (strTabImageXpath == gblNull || strTabImageXpath == gblNA) {
								//No image is specified
								strTabImage = gblNA;
								strTabImageValue = gblNA;
							}
							else {
								//
								//If image returned check for text
								let weTabImage;
								let mapGetTabImages = CWCore.returnChildElements(weTab, strTabImageXpath);
								//let lstTabImages = mapGetTabImages.lstChildObjects; // Declared but unused
								let intTabImgCnt = mapGetTabImages.cntChildObjs;
								if (intTabImgCnt == 1) {
									weTabImage = mapGetTabImages.lstChildObjects[0]; // Access array by index
								}
								if (weTabImage == null) {
									if (boolTabImageHasText == gblNA || boolTabImageHasText == gblNull) {
										strTabImageValue = gblNA;
									}
									else {
										strTabImageValue = gblNull; //Tab may have an image but this tab does not have an image
									}
								}
								else {
									//Highlight
									if (TCExecParams.getBoolDoHighlight() == true) {
										let mapHighlight = {}; // Mimic Groovy Map
										mapHighlight = CWCore.objHighlightElementJS(weTabImage, 'TabImage for Tab ' + loopTabItems, strTabControlName);
									}
									//Return the image text value
									strTabImageValue = StringsAndNumbers.JComm_HandleNoData(weTabImage.getText());
									//Remove the text from the tab value since it contains all the text
									if (strTabImageValue != gblNull) {
										strTabValue = StringsAndNumbers.JComm_GetLeftTextInString(strTabValue, strTabImageValue);
									}
									let strImageAtrValue = weTabImage.getAttribute(strImageAttribute);
									//Parse the image
									let mapSplitString = {}; // Mimic Groovy Map
									mapSplitString = StringsAndNumbers.JComm_StringToArray(strImageValue, gblDelimiter);
									let intSplitItemCnt = StringsAndNumbers.JComm_StringToInteger(mapSplitString.intItemCount);
									let mapGetImageInfo = {}; // Mimic Groovy Map
									if (intSplitItemCnt == 2) {
										let arryTabSelValues = mapSplitString.ArryOfValues;
										mapGetImageInfo = StringsAndNumbers.JComm_ReturnValueFromString(strImageAtrValue, arryTabSelValues[1], arryTabSelValues[0]);
									}
									else {
										mapGetImageInfo.boolTextIsPresent = 'false';
										mapGetImageInfo.strMethodDetails = "FAILED!!! strImageValue: " + strImageValue + " return " + intSplitItemCnt + " item(s) INSTEAD of 2";
									}
									if (mapGetImageInfo.boolTextIsPresent == 'true') {
										strTabImage = mapGetImageInfo.strValue;
									}
									else {
										strTabImage = mapGetImageInfo.strMethodDetails;
									}
									//TODO do we need the color
									strColor = StringsAndNumbers.JComm_HandleNoData(weTabImage.getCssValue("color"));
									strBckgColor = StringsAndNumbers.JComm_HandleNoData(weTabImage.getCssValue("background-color"));
									//Return the property and add to the strTempCellElements
									let strTempStatusColor = "Color:" + strColor + "Background Color:"+ strBckgColor;
									strTabImage = strTabImage + gblDelimiter + strTempStatusColor;
									boolCustomTheme = true;
									intImgColID = ExcelData.excelGetColIndexByColName(TCObj.getObjSheetTCActiveOutputSheet(), "TabImage").ColIndex;
								}
							}
							//Output the data to the spreadsheet. Columns: TabValue	TabSelected	TabImage	TabImageValue	TabRow	TabOrder
							//Output all of the values
							mapSetCellValue = ExcelData.excelSetCellValueByRowNumberColName(intOutputRow, "TabValue", strTabValue, strCellTheme);
							mapSetCellValue = ExcelData.excelSetCellValueByRowNumberColName(intOutputRow, "TabSelected", strTabSelValue, strCellTheme);
							if (boolCustomTheme == false) {
								mapSetCellValue = ExcelData.excelSetCellValueByRowNumberColName(intOutputRow, "TabImage", strTabImage, strCellTheme);
							}
							else {
								//Process the output with a custom them
								let mapCustomTheme = {}; // Mimic Groovy Map
								mapCustomTheme.ExcelRow = intOutputRow;
								mapCustomTheme.ExcelCol = intImgColID;
								mapCustomTheme.CellValue = strTabImage;
								mapCustomTheme.RGBAColor = strColor;
								mapCustomTheme.RGBABkColor = strBckgColor;
								mapCustomTheme.Alignment = 'Center'; // Assuming 'Center' is recognized
								//Call the set cell value with customtheme method
								mapSetCellValue = ExcelData.excelSetCellValueByRowColNumberCustTheme(mapCustomTheme);
							}
							mapSetCellValue = ExcelData.excelSetCellValueByRowNumberColName(intOutputRow, "TabImageValue", strTabImageValue, strCellTheme);
							mapSetCellValue = ExcelData.excelSetCellValueByRowNumberColName(intOutputRow, "TabRow", strTabRow, strCellTheme);
							mapSetCellValue = ExcelData.excelSetCellValueByRowNumberColName(intOutputRow, "TabOrder", strTabOrder, strCellTheme);
						}
					}
				}
			}
			//Update the column width to fit
			if (boolPassed == true) {
				//TODO create a new function CommonWeb_to autofit the column
				let mapAutoFit = {}; // Mimic Groovy Map
				mapAutoFit = ExcelData.excelAutofitCols();
				boolPassed = StringsAndNumbers.JComm_StringToBoolean(mapAutoFit.boolPassed);
				if (boolPassed == true) {
					//TODO do we need any statistics?
					strMethodDetails = "Successfully output the tab items.";
				}
				else {
					strMethodDetails = mapAutoFit.strMethodDetails;
				}
			}
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure this is set for all paths
	return mapResults;
}
/**
* -------------------------------------  SelectTab  -----------------------------------
* Select the assigned tab
* @param mapInputValues		The input values used in the method
* @param strStartLocation	  The assigned feature start location
* @param strTabControlName	 The tab control element name
* @param strTabControlXpath	The Xpath for the tab control
* @param strTabLevelXPath	  The Xpath defining a row level of tabs.
* @param strTabLevelProp	   The attribute name to return the tab level value. Set to gblNA if not applicable
* @param strTabLevelValue	  The tab level value used to determine the level number or id. Usually consist of a expression and value combination. For example: '%Right%|tabset'
* @param strTabLevelNoTabMsg   Message text displayed if no tabs are present. Set to gblNA if not applicable
* @param strTabItemXpath	   The Xpath for the tab item. Should containing a leading '.' to find only sub children
* @param strSelectTabValue	The value of the tab to select which consist of StrRowNum|StrTabValue. One based Row Number set to gblNA if no row id is present.
* @param strTabSelectElemXPath   The Xpath for the tab element which contains the Selected state. Set to gblNA if the tab element has the attribute and not another object.
* @param strTabSelectedProperty   The attribute name to return the tab level value.
* @param strTabSelectedValue   The tab selected value used to determine the selected state. Usually consist of a expression and value combination. For example: '%Cont%|selected'
* @param strTabImageXpath	   The Xpath for the image within the tab element. Set to gblNA if not applicable
* @param boolTabImageHasText   Does the image contain text that we will need to return? True/False frequently must be parsed from the tab text to show on the tab text.
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 01/19/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SelectTab (mapInputValues) { // mapInputValues and return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value'); // Not used in this function
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	let boolDoHighlight = TCExecParams.getBoolDoHighlight();
	let intViewDelay = TCExecParams.getIntViewDelaySecs();
	//Declare the variables
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails; // Type inferred from assignment
	let boolPassed = true;
	//Create variables and assign values from the mapInputValues
	let strStartLocation = mapInputValues.strStartLocation;
	let strTabControlName = mapInputValues.strTabControlName;
	let strTabControlXpath = mapInputValues.strTabControlXpath;
	let strTabLevelXPath = mapInputValues.strTabLevelXPath;
	let strTabLevelProp = mapInputValues.strTabLevelProp;
	let strTabLevelValue = mapInputValues.strTabLevelValue;
	// let strTabLevelNoTabMsg = mapInputValues.strTabLevelNoTabMsg; // Not used
	let strTabItemXpath = mapInputValues.strTabItemXpath;
	let strSelectTabValue = mapInputValues.strSelectTabValue;
	let strTabSelectElemXPath = mapInputValues.strTabSelectElemXPath;
	let strTabSelectedProperty = mapInputValues.strTabSelectedProperty;
	let strTabSelectedValue = mapInputValues.strTabSelectedValue;
	let strTabImageXpath = mapInputValues.strTabImageXpath;
	// let boolTabImageHasText = mapInputValues.boolTabImageHasText; // Not used
	let weTab; // Type inferred (WebElement)
	//Return the tab control element
	let weTabCtrl = CWCore.returnWebElement(strTabControlXpath);
	if (weTabCtrl == null ) {
		boolPassed = false;
		strMethodDetails = "The location '" + strStartLocation + "' '" + strTabControlName + "' USING XPath: " + strTabControlXpath + " DID NOT RETURN an ELEMENT.";
	}
	else {
		//Process the tab control
		//Define the variables for the output
		let strTabValue;
		let strTabSelValue;
		let strTabImage;
		let strTabImageValue;
		let strTabRow;
		let strTabOrder;
		// let TabSelected; // Declared bool TabSelected but not used in groovy

		//Highlight
		if (TCExecParams.getBoolDoHighlight() == true) {
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weTabCtrl, 'TabControl', strTabControlName);
		}
		//Return the tab levels
		let weTabRow; // Type inferred (WebElement)
		let intTabRowElemCnt; // Type inferred (number)
		let strTabValueToSelect; // Type inferred (string)
		let lstTabRowElements; // Type inferred (List<WebElement>)
		let intTabRow = -1; // Type inferred (number)

		if (strTabLevelXPath == gblNA) {
			intTabRowElemCnt = 1;
		}
		else {
			let mapGetTabRowElements = {}; // Mimic Groovy Map
			mapGetTabRowElements = CWCore.returnChildElements(weTabCtrl, strTabLevelXPath);
			lstTabRowElements = mapGetTabRowElements.lstChildObjects;
			intTabRowElemCnt = mapGetTabRowElements.cntChildObjs;
			if (intTabRowElemCnt < 1 && strTabLevelXPath != gblNA) {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The '" + strTabControlName + "' DOES NOT contain any Tab Levels!!!";
			}
		}
		if (boolPassed == true) { // If all initial checks passed
			let intAsgnTabRow; // Type inferred (number)
			//Split the strSelectTabValue to determine which row we should select
			let mapSplitSelTabValString = {}; // Mimic Groovy Map
			mapSplitSelTabValString = StringsAndNumbers.JComm_StringToArray(strSelectTabValue, gblDelimiter);
			let intSelTabValCnt = StringsAndNumbers.JComm_StringToInteger(mapSplitSelTabValString.intItemCount);
			if (intSelTabValCnt == 2) {
				let strSelTabRow = mapSplitSelTabValString.ArryOfValues[0]; // Access array by index
				if (strSelTabRow == gblNA) {
					intAsgnTabRow = 1;
				}
				else {
					intAsgnTabRow = StringsAndNumbers.JComm_StringToInteger(strSelTabRow);
				}
				strTabValueToSelect = mapSplitSelTabValString.ArryOfValues[1]; // Access array by index
				if (intTabRowElemCnt >= intAsgnTabRow) {
					//Return the tabrow
					if (strTabLevelXPath == gblNA) {
						weTabRow = weTabCtrl;
					}
					else {
						weTabRow = lstTabRowElements[intAsgnTabRow - 1]; //Zero based index.
						//Return the tab control level ID (nested logic adapted from OutputTabData)
						let mapSplitElementString = {}; // Mimic Groovy Map
						mapSplitElementString = StringsAndNumbers.JComm_StringToArray(strTabLevelValue, gblDelimiter);
						let intValueItemCnt = StringsAndNumbers.JComm_StringToInteger(mapSplitElementString.intItemCount);
						if (intValueItemCnt == 2) {
							let arryTabLvlValues = mapSplitElementString.ArryOfValues;
							//Determine the row number by attribute value
							if (CWCore.isAttribtuePresent(weTabRow, strTabLevelProp) == true) {
								//Return the value
								let strTabLvlPropValue = StringsAndNumbers.JComm_HandleNoData(weTabRow.getAttribute(strTabLevelProp));
								//Return the level value
								let mapReturnTabLvl = {}; // Mimic Groovy Map
								mapReturnTabLvl = StringsAndNumbers.JComm_ReturnValueFromString(strTabLvlPropValue, arryTabLvlValues[1], arryTabLvlValues[0]);
								if (StringsAndNumbers.JComm_StringToBoolean (mapReturnTabLvl.boolTextIsPresent) == true) {
									strTabRow = (StringsAndNumbers.JComm_StringToInteger(mapReturnTabLvl.strValue) + 1).toString(); //Assumes the tabs are zero index.
								}
								else {
									boolPassed = false;
									strMethodDetails = mapReturnTabLvl.strMethodDetails;
								}
							}
							else {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The Tab Level ATTRIBUTE of: " + strTabLevelProp + " DOES NOT EXIST for the ELEMENT!!!";
							}
						} else { // intValueItemCnt not 2
							boolPassed = false;
							strMethodDetails = "FAILED!!! The Tab Level Value of: " + strTabLevelValue + " DOES NOT CONTAIN TWO VALUES REQUIRED!!!";
						}
					}
				}
				else { // intTabRowElemCnt < intAsgnTabRow
					boolPassed = false;
					strMethodDetails = "FAILED!!! The tab control contains " + intTabRowElemCnt + " WHICH IS LESS THAN the Assigned row of: " + intAsgnTabRow.toString() + "!!!";
				}
			}
			else { // intSelTabValCnt not 2
				boolPassed = false;
				strMethodDetails = "FAILED!!! The strSelectTabValue of: " + strSelectTabValue + " DOES NOT CONTAIN TWO VALUES REQUIRED!!!";
			}
		}
		else if (intTabRowElemCnt == 1 && strTabLevelProp == gblNA) { // Condition from original Groovy for single tab row with no level prop
			//Assign the element to the weTabRow
			weTabRow = lstTabRowElements[0]; // Assuming lstTabRowElements exists and has at least one element.
			intTabRow = 1;
		}
		else if (intTabRowElemCnt > 1 && strTabLevelProp == gblNA) { // Condition from original Groovy for multiple tab rows with no level prop
			boolPassed = false;
			strMethodDetails = "FAILED!!! The '" + strTabControlName + "' count of: " + intTabRowElemCnt + " which IS GREATER THAN 1, AND THE ROW ATTRIBUTE is " + gblNA + "!!!";
		}
		else { // weTabCtrl was null or other initial conditions failed
			boolPassed = false;
			strMethodDetails = "FAILED!!! The '" + strTabControlName + "' DOES NOT contain any Tab Levels!!!";
		}
		if (boolPassed == true) {
			//Process the assigned row to find the specified tab
			let intTabFoundIndex = -1; // Type inferred (number)
			//Highlight
			if (TCExecParams.getBoolDoHighlight() == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weTabRow, 'TabRow' , strTabControlName);
			}
			//Return the count of tab elements
			let boolTabFound = false; // Type inferred (boolean)
			let mapGetTabElements = CWCore.returnChildElements(weTabRow, strTabItemXpath);
			let lstTabElements = mapGetTabElements.lstChildObjects;
			let intTabElemCnt = mapGetTabElements.cntChildObjs;
			if (intTabElemCnt < 1) {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The '" + strTabControlName + "' TabRow" + intTabRow + " DOES NOT contain any TABS!!!";
			}
			else {
				//Process the tabs to return the tab value. Do not include any images or image text in the value.
				for (let intLoopTabs = 0; intLoopTabs < intTabElemCnt; intLoopTabs++) {
					weTab = lstTabElements[intLoopTabs]; // Access array by index
					//Highlight
					if (TCExecParams.getBoolDoHighlight() == true) {
						let mapHighlight = {}; // Mimic Groovy Map
						mapHighlight = CWCore.objHighlightElementJS(weTab, 'TabRow' + intTabRow + " Tab " + intLoopTabs, strTabControlName);
					}
					//Return the tab value. Only remove from the tab value if needed.
					strTabValue = StringsAndNumbers.JComm_HandleNoData(weTab.getText()); //Return the text for the tab
					//Process the image (from OutputTabData)
					if (strTabImageXpath != gblNull && strTabImageXpath != gblNA) {
						let weTabImage;
						let mapGetTabImages = CWCore.returnChildElements(weTab, strTabImageXpath);
						let intTabImgCnt = mapGetTabImages.cntChildObjs;
						if (intTabImgCnt == 1) {
							weTabImage = mapGetTabImages.lstChildObjects[0];
						}
						if (weTabImage != null) { // if image found
							//Return the image text value
							let strTabImageValueTemp = StringsAndNumbers.JComm_HandleNoData(weTabImage.getText());
							//Remove the text from the tab value since it contains all the text
							if (strTabImageValueTemp != gblNull) {
								strTabValue = StringsAndNumbers.JComm_GetLeftTextInString(strTabValue, strTabImageValueTemp); // Update strTabValue
							}
						}
					}
					if (strTabValue == strTabValueToSelect) {
						boolTabFound = true;
						intTabFoundIndex = intLoopTabs; // Capture index of found tab
						break; // Found the tab, exit loop
					}
				}
				if (boolTabFound == true) {
					//Click the tab and check the selected
					// Need to re-assign weTab to the found element using intTabFoundIndex
					weTab = lstTabElements[intTabFoundIndex]; // Get the actual element that was found
					weTab.click();
					DateTime.WaitSecs(TCExecParams.getIntViewDelaySecs());
					strMethodDetails = "The '" + strTabControlName + "' TabRow" + intTabRow + " contained the tab with the value:" + strTabValueToSelect + " and was clicked.";
					//NOTE: DO NOT TRY TO CHECK IF THE TABS ARE SELECTED AS THE PAGE AND ALL OBJECTS HAVE CHANGED
				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The '" + strTabControlName + "' TabRow" + intTabRow + " DOES NOT CONTAIN the tab that MATCHES value:" + strTabValueToSelect + "!!!";
				}
			}
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure this is set for all paths
	return mapResults;
}
/**
* -------------------------------------  VerifySelectedTab  -----------------------------------
* Verify the tab selected by displayed text
* @param mapInputValues		The input values used in the method
* @param strStartLocation	  The assigned feature start location
* @param strTabControlName	 The tab control element name
* @param strTabItemXpath	   The Xpath for the tab item. Should containing a leading '.' to find only sub children
* @param strSelectTabValue	The value of the tab to select which consist of StrRowNum|StrTabValue. One based Row Number set to gblNA if no row id is present.
* @param strTabSelectElemXPath   The Xpath for the tab element which contains the Selected state. Set to gblNA if the tab element has the attribute and not another object.
* @param strTabSelectedProperty   The attribute name to return the tab level value.
* @param strTabSelectedValue   The tab selected value used to determine the selected state. Usually consist of a expression and value combination. For example: '%Cont%|selected'
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 05/09/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifySelectedTab (mapInputValues) { // mapInputValues and return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value'); // Not used
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	// let boolDoHighlight = TCExecParams.getBoolDoHighlight(); // Declared but unused
	// let intViewDelay = TCExecParams.getIntViewDelaySecs(); // Declared but unused
	//Declare the variables
	let mapResults = {}; // Mimic Groovy Map Initialization
	let strMethodDetails; // Type inferred from assignment
	let boolPassed = true;
	//Create variables and assign values from the mapInputValues
	let strStartLocation = mapInputValues.strStartLocation;
	let strTabControlName = mapInputValues.strTabControlName;
	let strTabControlXpath = mapInputValues.strTabControlXpath;
	let strTabItemXpath = mapInputValues.strTabItemXpath;
	let strSelectTabValue = mapInputValues.strSelectTabValue; // Added this from OutputTabData param list
	let strTabSelectElemXPath = mapInputValues.strTabSelectElemXPath; // Added from OutputTabData param list
	let strTabSelectedProperty = mapInputValues.strTabSelectedProperty; // Added from OutputTabData param list
	let strTabSelectedValue = mapInputValues.strTabSelectedValue; // Added from OutputTabData param list

	//Return the tab control element
	let weTabCtrl = CWCore.returnWebElement(strTabControlXpath);
	if (weTabCtrl == null ) {
		boolPassed = false;
		strMethodDetails = "The location '" + strStartLocation + "' '" + strTabControlName + "' USING XPath: " + strTabControlXpath + " DID NOT RETURN an ELEMENT.";
	}
	else {
		//Highlight (using boolDoHighlight from outside this function's local scope, assuming it's available)
		if (TCExecParams.getBoolDoHighlight() == true) { // Use the global TCExecParams
			let mapHighlight = {}; // Mimic Groovy Map
			mapHighlight = CWCore.objHighlightElementJS(weTabCtrl, 'TabControl', strTabControlName);
		}
		//Return the count of tab elements
		let boolTabFound = false; // Type inferred from assignment
		let mapGetTabElements = CWCore.returnChildElements(weTabCtrl, strTabItemXpath);
		let lstTabElements = mapGetTabElements.lstChildObjects;
		let intTabElemCnt = mapGetTabElements.cntChildObjs;
		if (intTabElemCnt < 1) {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The '" + strTabControlName + " DOES NOT contain any TABS!!!";
		}
		else {
			// This is where the actual verification logic should go, based on the javadoc.
			// The original Groovy function CommonWeb_body ends here with an empty else.
			// Assuming the intent is to verify the selected tab against `strSelectTabValue`.
			// This logic is adapted from `SelectTab` and `OutputTabData` as implied by parameters.

			let intAsgnTabRow = 1; // Default to 1st tab row if not specified, based on `SelectTab` logic.
			let strTabValueToVerify;
			// Parse strSelectTabValue
			let mapSplitSelTabValString = StringsAndNumbers.JComm_StringToArray(strSelectTabValue, gblDelimiter);
			if (StringsAndNumbers.JComm_ToInteger(mapSplitSelTabValString.intItemCount) === 2) {
				let strSelTabRow = mapSplitSelTabValString.ArryOfValues[0];
				if (strSelTabRow !== gblNA) {
					intAsgnTabRow = StringsAndNumbers.JComm_StringToInteger(strSelTabRow);
				}
				strTabValueToVerify = mapSplitSelTabValString.ArryOfValues[1];
			} else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! strSelectTabValue: '" + strSelectTabValue + "' does not contain the expected 'row|value' format.";
			}

			if (boolPassed === true) {
				let weTargetTabRow = weTabCtrl; // Default to control itself if no levels
				if (strTabLevelXPath !== gblNA) {
					let mapGetTabRowElements = CWCore.returnChildElements(weTabCtrl, strTabLevelXPath);
					let lstTabRowElements = mapGetTabRowElements.lstChildObjects;
					if (lstTabRowElements.length >= intAsgnTabRow) {
						weTargetTabRow = lstTabRowElements[intAsgnTabRow - 1]; // Get the specific tab row element
					} else {
						boolPassed = false;
						strMethodDetails = "FAILED!!! Specified tab row " + intAsgnTabRow + " not found in control.";
					}
				}
			}


			if (boolPassed === true) {
				let foundSelectedTab = false;
				// Loop through tabs in the target tab row
				for (let loopTabItems = 0; loopTabItems < intTabElemCnt; loopTabItems++) {
					let weTab = lstTabElements[loopTabItems];
					let strTabValue = StringsAndNumbers.JComm_HandleNoData(weTab.getText());

					// Check if this tab is the currently selected one
					let isTabSelected = false;
					let mapSplitSelectedValString = StringsAndNumbers.JComm_StringToArray(strTabSelectedValue, gblDelimiter);
					if (StringsAndNumbers.JComm_ToInteger(mapSplitSelectedValString.intItemCount) === 2) { // Expected format: 'type|value'
						let arryTabSelValues = mapSplitSelectedValString.ArryOfValues;
						// Use isAttribtuePresent for safety before getAttribute
						if (CWCore.isAttribtuePresent(weTab, strTabSelectedProperty) === true) {
							let strTabSelPropValue = StringsAndNumbers.JComm_HandleNoData(weTab.getAttribute(strTabSelectedProperty));
							isTabSelected = StringsAndNumbers.JComm_VerifyTextPresent(strTabSelPropValue, arryTabSelValues[1], arryTabSelValues[0]);
						}
					} else if (strTabSelectedProperty === 'activetab' && strTabSelectElemXPath !== gblNA) { // Special 'activetab' logic
						let weTemp = CWCore.returnWebElement(strTabSelectElemXPath);
						if (weTemp && CWCore.isAttribtuePresent(weTemp, strTabSelectedProperty) === true) {
							let strTabSelPropValue = StringsAndNumbers.JComm_HandleNoData(weTemp.getAttribute(strTabSelectedProperty));
							let strCurrTabID = gblNull;
							if (CWCore.isAttribtuePresent(weTab, 'id') === true) {
								strCurrTabID = StringsAndNumbers.JComm_HandleNoData(weTab.getAttribute('id'));
							} else {
								let childElem = weTab.findElement(By.xpath("./*"));
								if (childElem && CWCore.isAttribtuePresent(childElem, 'id')) {
									strCurrTabID = StringsAndNumbers.JComm_HandleNoData(childElem.getAttribute('id'));
								}
							}
							isTabSelected = (strTabSelPropValue === strCurrTabID);
						}
					}

					if (isTabSelected) {
						foundSelectedTab = true; // Found the currently selected tab
						// Check if this selected tab matches the expected value
						if (strTabValue === strTabValueToVerify) {
							strMethodDetails = "Successfully verified the selected tab: '" + strTabValueToVerify + "'.";
						} else {
							boolPassed = false;
							strMethodDetails = "FAILED!!! Expected selected tab was '" + strTabValueToVerify + "', but found '" + strTabValue + "'.";
						}
						break; // Found the selected tab, no need to continue looping
					}
				}

				if (!foundSelectedTab) {
					boolPassed = false;
					strMethodDetails = "FAILED!!! No tab found to be selected on the control: '" + strTabControlName + "'.";
				}
			}
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
/**
* -------------------------------------  VerifyTabData  -----------------------------------
* Output the header table results
* @param mapInputValues		The input values used in the method
*
* @param strStartLocation	  The assigned feature start location
* @param strOutputSheetName	The name of the output sheet
* @param strInputSheetName	 The name of the input sheet
* @param strTabControlName	 The tab control element name
* @param strTabControlXpath	The Xpath for the tab control
* @param strTabLevelXPath	  The Xpath defining a row level of tabs. Set to gblNA if not applicable
* @param strTabLevelProp	   The attribute name to return the tab level value. Set to gblNA if not applicable
* @param strTabLevelValue	  The tab level value used to determine the level number or id. Usually consist of a expression and value combination. For example: '%Right%|tabset'
* @param strTabLevelNoTabMsg   Message text displayed if no tabs are present. Set to gblNA if not applicable
* @param strTabItemXpath	   The Xpath for the tab item. Should containing a leading '.' to find only sub children
* @param strTabSelectElemXPath   The Xpath for the tab element which contains the Selected state. Set to gblNA if the tab element has the attribute and not another object.
* @param strTabSelectedProperty   The attribute name to return the tab level value.
* @param strTabSelectedValue   The tab selected value used to determine the selected state. Usually consist of a expression and value combination. For example: '%Cont%|selected'
* @param strTabImageXpath	   The Xpath for the image within the tab element. Set to gblNA if not applicable
* @param boolTabImageHasText   Does the image contain text that we will need to return? True/False frequently must be parsed from the tab text to show on the tab text.
* @param strImageAttribute	  The attribute name to return the value used to parse and determine the image that is present.
* @param strImageValue		 The image value used to parse the Image specified attribute and determine the image. Usually consist of a expression and value combination. For example: '%Right%|*GBLSPACE*'
*
* @return mapResults		 The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 01/16/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyTabData (mapInputValues) { // mapInputValues and return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value'); // Not used
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	let boolDoHighlight = TCExecParams.getBoolDoHighlight();
	// let intViewDelay = TCExecParams.getIntViewDelaySecs(); // Declared but unused
	//Declare the variables
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails; // Type inferred from assignment
	let boolPassed = true;
	//Create variables and assign values from the mapInputValues
	let strStartLocation = mapInputValues.strStartLocation;
	let strInputSheetName = mapInputValues.strInputSheetName;
	let arryColName = mapInputValues.ArryOfValues; // Assumed to be populated, as per OutputTabData structure.
	let intExpectedInputColCnt = mapInputValues.intExpectedInputColCnt;
	let strOutputSheetName = mapInputValues.strOutputSheetName;
	let strTabControlName = mapInputValues.strTabControlName;
	let strTabControlXpath = mapInputValues.strTabControlXpath;
	let strTabLevelXPath = mapInputValues.strTabLevelXPath;
	let strTabLevelProp = mapInputValues.strTabLevelProp;
	let strTabLevelValue = mapInputValues.strTabLevelValue;
	// let strTabLevelNoTabMsg = mapInputValues.strTabLevelNoTabMsg; // Not used
	let strTabItemXpath = mapInputValues.strTabItemXpath;
	let strTabSelectElemXPath = mapInputValues.strTabSelectElemXPath;
	let strTabSelectedProperty = mapInputValues.strTabSelectedProperty;
	let strTabSelectedValue = mapInputValues.strTabSelectedValue;
	let strTabImageXpath = mapInputValues.strTabImageXpath;
	let boolTabImageHasText = mapInputValues.boolTabImageHasText;
	let strImageAttribute = mapInputValues.strImageAttribute;
	let strImageValue = mapInputValues.strImageValue;
	//Open and set the Inputsheet as the active inputput data.
	// let strTempInputValue; // Declared but unused

	//Load the inputsheet
	let intInputRowCnt;
	let intInputColCnt;
	// let intHdrColCnt; // Declared but unused
	let intCellPassed = 0;
	let intCellFailed = 0;
	let shInput; // Mimic XSSFSheet
	let mapOpenInputSheet = {}; // Mimic Groovy Map
	mapOpenInputSheet = ExcelData.excelGetSheetByName(TCObj.getObjWorkbook(), strInputSheetName);
	if (StringsAndNumbers.JComm_StringToBoolean(mapOpenInputSheet.boolPassed) == true) {
		shInput = mapOpenInputSheet.objWbSheet;
		//Return the row and column Count
		let mapSheetRowColCnt = {}; // Mimic Groovy Map
		mapSheetRowColCnt = ExcelData.excelGetRowAndColCount(shInput);
		if (StringsAndNumbers.JComm_StringToBoolean(mapSheetRowColCnt.boolPassed) == true) {
			intInputRowCnt = mapSheetRowColCnt.RowCount;
			intInputColCnt = mapSheetRowColCnt.ColCount;
			//Verify the input column count is equal to the output
			if (intExpectedInputColCnt != intInputColCnt) {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The input column count of: " + intInputColCnt + " does NOT MATCH the EXPECTED: " + intExpectedInputColCnt + " column(s).";
			}
		}
		else {
			boolPassed = false;
			strMethodDetails = mapSheetRowColCnt.strMethodDetails;
		}
	}
	else {
		boolPassed = false;
		strMethodDetails = mapOpenInputSheet.strMethodDetails;
	}
	if (boolPassed == true) {
		//Return the tab control element
		let weTabCtrl = CWCore.returnWebElement(strTabControlXpath);
		if (weTabCtrl == null ) {
			boolPassed = false;
			strMethodDetails = "The location '" + strStartLocation + "' '" + strTabControlName + "' USING XPath: " + strTabControlXpath + " DID NOT RETURN an ELEMENT.";
		}
		else {
			//Process the tab control
			//Define the variables for the output
			let strTabValue;
			let strTabSelValue;
			let strTabImage;
			let strTabImageValue;
			let strTabRow;
			let strTabOrder;
			// let TabSelected; // Declared boolean but unused

			//Highlight
			if (TCExecParams.getBoolDoHighlight() == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weTabCtrl, 'TabControl', strTabControlName);
			}
			let weTabRow;
			let intTabRowElemCnt;
			let lstTabRowElements;
			//Return the tab levels
			if (strTabLevelXPath == gblNA) {
				intTabRowElemCnt = 1;
			}
			else {
				let mapGetTabRowElements = {}; // Mimic Groovy Map
				mapGetTabRowElements = CWCore.returnChildElements(weTabCtrl, strTabLevelXPath);
				lstTabRowElements = mapGetTabRowElements.lstChildObjects;
				intTabRowElemCnt = mapGetTabRowElements.cntChildObjs;
				if (intTabRowElemCnt < 1 && strTabLevelXPath != gblNA) {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The '" + strTabControlName + "' DOES NOT contain any Tab Levels!!!";
				}
			}
			if (boolPassed == true) {
				//Declare the variables
				let intOutputRow = 0;
				let mapGetTabElements = {}; // Mimic Groovy Map
				let lstTabElements;
				let intTabElemCnt;
				let strCellTheme;
				let mapSetCellValue;
				let strColor;
				let strBckgColor;
				//Process the tab levels
				for (let loopTabRows = 0; loopTabRows < intTabRowElemCnt && boolPassed; loopTabRows++) { // Added boolPassed check
					strTabRow = gblNull;
					if (strTabLevelXPath == gblNA) {
						weTabRow = weTabCtrl;
					}
					else {
						weTabRow = lstTabRowElements[loopTabRows]; // Access array by index
					}
					//Highlight
					if (TCExecParams.getBoolDoHighlight() == true) {
						let mapHighlight = {}; // Mimic Groovy Map
						mapHighlight = CWCore.objHighlightElementJS(weTabRow, 'TabRow' + loopTabRows , strTabControlName);
					}
					//Return the count of tab elements
					mapGetTabElements = CWCore.returnChildElements(weTabRow, strTabItemXpath);
					lstTabElements = mapGetTabElements.lstChildObjects;
					intTabElemCnt = mapGetTabElements.cntChildObjs;
					if (intTabElemCnt < 1) {
						boolPassed = false;
						strMethodDetails = "FAILED!!! The '" + strTabControlName + "' TabRow" + loopTabRows + " DOES NOT contain any TABS!!!";
					}
					else {
						//Check the tab level
						//Return the level if applicable
						if (strTabLevelProp == gblNull ||strTabLevelProp == gblNA) {
							strTabRow = gblNA;
						}
						else {
							let mapSplitElementString = {}; // Mimic Groovy Map
							mapSplitElementString = StringsAndNumbers.JComm_StringToArray(strTabLevelValue, gblDelimiter);
							let intValueItemCnt = StringsAndNumbers.JComm_StringToInteger(mapSplitElementString.intItemCount);
							if (intValueItemCnt == 2) {
								let arryTabLvlValues = mapSplitElementString.ArryOfValues;
								//Determine the row number by attribute value
								//Is the attribute present
								if (CWCore.isAttribtuePresent(weTabRow, strTabLevelProp) == true) {
									//Return the value
									let strTabLvlPropValue = StringsAndNumbers.JComm_HandleNoData(weTabRow.getAttribute(strTabLevelProp));
									//Return the level value
									let mapReturnTabLvl = {}; // Mimic Groovy Map
									mapReturnTabLvl = StringsAndNumbers.JComm_ReturnValueFromString(strTabLvlPropValue, arryTabLvlValues[1], arryTabLvlValues[0]);
									if (StringsAndNumbers.JComm_StringToBoolean (mapReturnTabLvl.boolTextIsPresent) == true) {
										strTabRow = (StringsAndNumbers.JComm_StringToInteger(mapReturnTabLvl.strValue) + 1).toString(); //Assumes the tabs are zero index.
									}
									else {
										boolPassed = false;
										strMethodDetails = mapReturnTabLvl.strMethodDetails;
									}
								}
								else {
									boolPassed = false;
									strMethodDetails = "FAILED!!! The Tab Level ATTRIBUTE of: " + strTabLevelProp + " DOES NOT EXIST for the ELEMENT!!!";
								}
							}
							else {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The Tab Level Value of: " + strTabLevelValue + " DOES NOT CONTAIN TWO VALUES REQUIRED!!!";
							}
						}
						if (boolPassed == true) { // All current checks passed
							//Process the current tab level and output the results for each tab.
							let weTab;
							// let weTabElem; // Declared but unused
							// let boolTabSelected; // Declared but unused
							// let boolHasImage; // Declared but unused

							for (let loopTabItems = 0; loopTabItems < intTabElemCnt && boolPassed; loopTabItems++) { // Added boolPassed check
								//Set the output row
								intOutputRow ++; // Increment output row for each tab item
								//Set the
								strTabValue = gblNull;
								strTabSelValue = gblNull;
								strTabImage = gblNull;
								strTabImageValue = gblNull;
								strTabOrder = (loopTabItems + 1).toString(); //Zero index converted to 1 based for users.
								let strActiveTab = gblNull; // Re-declared for current tab, assuming it's related to activetab property
								weTab = lstTabElements[loopTabItems]; // Access array by index
								//Highlight
								if (TCExecParams.getBoolDoHighlight() == true) {
									let mapHighlight = {}; // Mimic Groovy Map
									mapHighlight = CWCore.objHighlightElementJS(weTab, 'TabRow' + loopTabRows + " Tab " + loopTabItems, strTabControlName);
								}
								//Return the tab value
								strTabValue = StringsAndNumbers.JComm_HandleNoData(weTab.getText()); //Return the text for the tab
								//Is the tab selected
								let mapSplitElementString = {}; // Mimic Groovy Map
								mapSplitElementString = StringsAndNumbers.JComm_StringToArray(strTabSelectedValue, gblDelimiter);
								let intValueItemCnt = StringsAndNumbers.JComm_StringToInteger(mapSplitElementString.intItemCount);
								if (intValueItemCnt == 2) {
									let arryTabSelValues = mapSplitElementString.ArryOfValues;
									//Check if the property is present
									if (CWCore.isAttribtuePresent(weTab, strTabSelectedProperty) == true) {
										//Return the value
										let strTabSelPropValue = StringsAndNumbers.JComm_HandleNoData(weTab.getAttribute(strTabSelectedProperty));
										//Return the level value
										let boolTextPresent = StringsAndNumbers.JComm_VerifyTextPresent(strTabSelPropValue, arryTabSelValues[1], arryTabSelValues[0]);
										strTabSelValue = boolTextPresent.toString();
									}
									else {
										boolPassed = false;
										strMethodDetails = "FAILED!!! The Tab Level ATTRIBUTE of: " + strTabSelectedProperty + " DOES NOT EXIST for the TAB ELEMENT!!!";
									}
								}
								else if(strTabSelectedProperty == 'activetab') {
									let strCurrTabID;
									//Custom code for the PhoenixNavLink tab control found in requisitions
									let weTemp = CWCore.returnWebElement(strTabSelectElemXPath);
									if (weTemp == null) {
										boolPassed = false;
										strMethodDetails = "FAILED!!! The Xpath for the element of: " + strTabSelectElemXPath + "DID NOT RETURN and ELEMENT for the ActiveTab validation!!!";
									}
									//Check if the property is present
									if (CWCore.isAttribtuePresent(weTemp, strTabSelectedProperty) == true) {
										//Return the value
										let strTabSelPropValue = StringsAndNumbers.JComm_HandleNoData(weTemp.getAttribute(strTabSelectedProperty)); //Will be the ID of the tab
										if (CWCore.isAttribtuePresent(weTab, 'id') == true) {
											strCurrTabID = StringsAndNumbers.JComm_HandleNoData(weTab.getAttribute('id'));
										}
										else {
											let childElem = weTab.findElement(By.xpath("./*")); // Find first child as per Groovy
											if (childElem) { // Check if findElement returned an element
												strCurrTabID = StringsAndNumbers.JComm_HandleNoData(childElem.getAttribute('id')); //Developed to fit the purchase order
											} else {
												strCurrTabID = gblNull; // No child with id found
											}
										}
										if (strTabSelPropValue == strCurrTabID) {
											strTabSelValue = 'true';
										}
										else {
											strTabSelValue = 'false';
										}
									}
									else {
										boolPassed = false;
										strMethodDetails = "FAILED!!! The Tab ATTRIBUTE of: " + strTabSelectedProperty + " DOES NOT EXIST for the 'ACTIVETAB' for the TAB ELEMENT!!!";
									}
								}
								else {
									boolPassed = false;
									strMethodDetails = "FAILED!!! The Tab Selected VALUE of: " + strTabSelectedValue + " DOES NOT CONTAIN TWO VALUES REQUIRED!!!";
								}
								//Process the image
								if (strTabImageXpath == gblNull || strTabImageXpath == gblNA) {
									//No image is specified
									strTabImage = gblNA;
									strTabImageValue = gblNA;
								}
								else {
									//
									//If image returned check for text
									let weTabImage;
									let mapGetTabImages = CWCore.returnChildElements(weTab, strTabImageXpath);
									//List<WebElement> lstTabImages = mapGetTabImages.lstChildObjects; // Declared but unused
									let intTabImgCnt = mapGetTabImages.cntChildObjs;
									if (intTabImgCnt == 1) {
										weTabImage = mapGetTabImages.lstChildObjects[0];
									}
									if (weTabImage == null) {
										if (boolTabImageHasText == gblNA || boolTabImageHasText == gblNull) { // Note: boolTabImageHasText is String, not boolean
											strTabImageValue = gblNA;
										}
										else {
											strTabImageValue = gblNull; //Tab may have an image but this tab does not have an image
										}
									}
									else {
										//Highlight
										if (TCExecParams.getBoolDoHighlight() == true) {
											let mapHighlight = {}; // Mimic Groovy Map
											mapHighlight = CWCore.objHighlightElementJS(weTabImage, 'TabImage for Tab ' + loopTabItems, strTabControlName);
										}
										//Return the image text value
										strTabImageValue = StringsAndNumbers.JComm_HandleNoData(weTabImage.getText());
										//Remove the text from the tab value since it contains all the text
										if (strTabImageValue != gblNull) {
											strTabValue = StringsAndNumbers.JComm_GetLeftTextInString(strTabValue, strTabImageValue);
										}
										let strImageAtrValue = weTabImage.getAttribute(strImageAttribute);
										//Parse the image
										let mapSplitString = {}; // Mimic Groovy Map
										mapSplitString = StringsAndNumbers.JComm_StringToArray(strImageValue, gblDelimiter);
										let intSplitItemCnt = StringsAndNumbers.JComm_StringToInteger(mapSplitString.intItemCount);
										let mapGetImageInfo = {}; // Mimic Groovy Map
										if (intSplitItemCnt == 2) {
											let arryTabSelValues = mapSplitString.ArryOfValues;
											mapGetImageInfo = StringsAndNumbers.JComm_ReturnValueFromString(strImageAtrValue, arryTabSelValues[1], arryTabSelValues[0]);
										}
										else {
											mapGetImageInfo.boolTextIsPresent = 'false';
											mapGetImageInfo.strMethodDetails = "FAILED!!! strImageValue: " + strImageValue + " return " + intSplitItemCnt + " item(s) INSTEAD of 2";
										}
										if (mapGetImageInfo.boolTextIsPresent == 'true') {
											strTabImage = mapGetImageInfo.strValue;
										}
										else {
											strTabImage = mapGetImageInfo.strMethodDetails;
										}

										//TODO do we need the color
										// For the template, these are not directly used in the output Excel structure here
										strColor = StringsAndNumbers.JComm_HandleNoData(weTabImage.getCssValue("color"));
										strBckgColor = StringsAndNumbers.JComm_HandleNoData(weTabImage.getCssValue("background-color"));
										//Return the property and add to the strTempCellElements
										let strTempStatusColor = "Color:" + strColor + "Background Color:"+ strBckgColor;
										strTabImage = strTabImage + gblDelimiter + strTempStatusColor;
										// boolCustomTheme is not used here (only for excel output function CommonWeb_call), for now.

										// intImgColID = ExcelData.excelGetColIndexByColName(TCObj.getObjSheetTCActiveOutputSheet(), "TabImage").ColIndex; // Index not used here.
									}
								}
								//Check the results against the input and update as needed
								let strTempColName;
								let strTempValue; // Renamed to avoid shadowing
								let strTempCompareResults;
								// Re-declaring intTempInputValue and ExcelData.excelGetCellValueByRowNumColName inside this loop for each comparison.
								let strTempInputValue; // Type inferred
								for (let intLoopOutput = 0; intLoopOutput < intExpectedInputColCnt; intLoopOutput++) {
									strTempColName = arryColName[intLoopOutput]; // Access array by index
									strTempInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, intOutputRow, strTempColName).CellValue;
									switch (intLoopOutput) {
										case 0:
											strTempValue = strTabValue;
											break;
										case 1:
											strTempValue = strTabSelValue;
											break;
										case 2:
											strTempValue = strTabImage;
											break;
										case 3:
											strTempValue = strTabImageValue;
											break;
										case 4:
											strTempValue = strTabRow;
											break;
										case 5:
											strTempValue = strTabOrder;
											break;
										default:
											strTempValue = gblNull; // Handle unexpected column index
											break;
									}
									if (strTempInputValue == strTempValue) {
										strCellTheme = 'TestRunPassStd';
										strTempCompareResults = strTempValue;
										intCellPassed++;
									}
									else {
										strCellTheme = 'TestRunFailStd';
										strTempCompareResults = `Expected: ${strTempInputValue} Displayed: ${strTempValue}`;
										boolPassed = false;
										intCellFailed++;
									}
									//Output the cell value
									mapSetCellValue = ExcelData.excelSetCellValueByRowNumberColName(intOutputRow, strTempColName, strTempCompareResults, strCellTheme);
								}
							}
						}
					}
				}
				//Autofit the cells
				let mapAutoFit = {}; // Mimic Groovy Map
				mapAutoFit = ExcelData.excelAutofitCols();
				if (StringsAndNumbers.JComm_StringToBoolean(mapAutoFit.boolPassed) == false) {
					strMethodDetails = strMethodDetails + mapAutoFit.strMethodDetails;
				}
				if (boolPassed == true) {
					strMethodDetails = "Successfully verified all tab data consisting of " + intCellPassed + " cell(s) of input data. See output sheet: " + strInputSheetName;
				}
				else {
					strMethodDetails = "FAILED!!! Process all tab data consisting of " + intCellPassed + " cell(s) passed and " + intCellFailed + " INCORRECT CELL(s) of input data. See output sheet: " + strInputSheetName;
				}
			}
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
//TODO Select and verify treeview PGK 09/11/2022 will be completed when JTR test cases require them.
/**
* -------------------------------------  SelectTreeviewObject  -----------------------------------
* Select the assigned treeview object. Note: usually this will navigate to another page or location.
* @param strLocation			   The location of the treeview
* @param mapMethodData			The map of objects and values
* @param strTreeViewFullPath		The Treeview full XPath
* @param strTreeViewElemName		The meaningful name of the TreeView
* @param strTreeViewParentXpath   The parent object XPath. i.e. ul
* @param strTreeViewBranchXpath   The object XPath for the branch i.e. li
* @param strTreeViewElemValue	The delimited value for the branch
* @param boolTreeViewElemIsParent The delimited values is a parent element? true/false
* @param strTreeViewItemLinkXPath   The item link Xpath
* @param boolLinkedClicked		The link assigned is clicked? true/false
* @param strTreeViewCheckboxXPath   The checkbox path. Set to null if no checkbox is present
* @param boolAssignedChecked	   The checkbox is to be checked? true/false
* @param boolVerifyChecked		Verify the checked state? true/false
* @return mapResults		  The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 08/17/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SelectTreeviewObject (mapMethodData) {
	// function CommonWeb_body is empty in the original Groovy, so it remains empty in the template.
	let mapResults = {};
	mapResults.boolPassed = 'false';
	mapResults.strMethodDetails = 'function CommonWeb_body is empty in original Groovy code.';
	return mapResults;
}
/**
* -------------------------------------  VerifyTreeviewObject  -----------------------------------
* Verify the treeview object
* @param strLocation				The location of the treeview
* @param mapMethodData			The map of objects and values
* @param strTreeViewFullPath		The Treeview full XPath
* @param strTreeViewElemName		The meaningful name of the TreeView
* @param strTreeViewParentXpath   The parent object XPath. i.e. ul
* @param strTreeViewBranchXpath   The object XPath for the branch i.e. li
* @param strTreeViewElemValue	The delimited value for the branch
* @param boolTreeViewElemIsParent The delimited values is a parent element? true/false
* @param strTreeViewItemLinkXPath   The item link Xpath
* @param boolLinkedClicked		The link assigned is clicked? true/false
* @param strTreeViewCheckboxXPath   The checkbox path. Set to null if no checkbox is present
* @param boolAssignedChecked	   The checkbox is to be checked? true/false
* @param boolVerifyChecked		Verify the checked state? true/false
* @return mapResults		  The results showing Passed and method details.
*
* @author pkanaris
* @author Created: 08/17/2022
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyTreeviewObject (mapMethodData) {
	// function CommonWeb_body is empty in the original Groovy, so it remains empty in the template.
	let mapResults = {};
	mapResults.boolPassed = 'false';
	mapResults.strMethodDetails = 'function CommonWeb_body is empty in original Groovy code.';
	return mapResults;
}
/**
*
* TreeViewVerifyItem
* TreeViewVerifyMultipleItems
* TreeViewSetVerifyCheckbox
* TreeViewSetVerifyMultipleCheckboxes
* TreeViewClickItem
* TreeViewExpandCollapseItem
*
* All of these methods will receive the treeview as an element within the map
*
*/

/* Side Bar Menu Items using UL
* Supports a title and single child menu items
* Based on the JI User Profile and Supplier Profile
*/
/** ------------------ ExpandCollapseSidebarMenuTitle -----------------------
* Expand or collapse the specified Title element
* @param mapInputValues The values to include title, expand/collapse, objects, and object XPaths
* @param strTitle The title value to expand or collapse
* @param boolIsExpanded Is the title expanded? true/false
*
* @param	strStartLocation									The Start Location for the side-bar menu
* @param	weSidebar											THe web element of the side-bar menu
* @param	strSideBarFullXPath									The XPath for the side-bar menu within the results
* @param	strSideBarName										The name of the side-bar
* @param	strSideBarMenuTitleName								The side-bar title name
* @param	strSideBarMenuTitleXpath							The side-bar title XPath
* @param	strSideBarMenuTitleTextXPath						The side-bar title text XPath
* @param	strSideBarMenuTitleExpandAttr						The Attribute for side-bar title. To see if the title is expanded
* @param 	boolAutoExpTitle									The Auto Expanded title is true/false
* @param	strSideBarMenuChildPanelXPath						The side-bar menu Child Panel XPath
* @param	strSideBarMenuChildPanelValueIsCollapsed			The child Panel is expandable true/false
* @param	strSideBarMenuChildItemXPath						The child Item XPath
* @param	strSideBarMenuChildItemSelectedProperty				The child Item Selected Property
* @param	strSideBarMenuChildItemSelectedPropertyValue		The selected Property Value of the Child Item
* @param	strSideBarMenuChildItemChildPropertyValue			The Property Value of the Child Item
* @param	strSideBarMenuChildItemElementXPath					The XPath of the Child Item element
* @param   strElementStatusProperty							The element Status Property
* @return mapResults Contains Passed/Failed and method details
*
* @author Created 01/09/2023
* @author pkanaris
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_ExpandCollapseSidebarMenuTitle (mapInputValues, strTitle, boolIsExpanded) { // mapInputValues is `any`, return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value'); // Not used
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	// let gblDelimiter = GVars.GblDelimiter('Value'); // Not used
	let gblTrueFalse = GVars.GBLTrueFalse('Value');
	let boolDoHighlight = TCExecParams.getBoolDoHighlight();
	let intViewDelay = TCExecParams.getIntViewDelaySecs();
	//Declare the variables
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails; // Type inferred from assignment
	let boolPassed = true;
	let weTitle; // Type inferred
	let weTitleChildItem; // Type inferred
	let intTitleElemCnt; // Type inferred
	let strTempTitleText; // Type inferred
	// let strTempTitleClass; // Not used
	// let strTempPanelClass; // Not used
	let strTempItemText; // Not used
	// let strTempObjText; // Not used
	let strTempItemPropertyValue; // Type inferred
	// let strTempElemType; // Not used
	// let strTempElemStatus; // Not used
	// let strTempElemText; // Not used
	let strTempChildItemText; // Type inferred
	let boolTempTitleExpanded; // Type inferred
	let boolItemSelected; // Type inferred

	//Return the values from the Map
	let strStartLocation = mapInputValues.strStartLocation;
	let weSidebar = mapInputValues.weSidebar;
	let strSideBarFullXPath = mapInputValues.strSideBarFullXPath; // Not used in this function, despite being in Javadoc
	let strSideBarName = mapInputValues.strSideBarName;
	let strSideBarMenuTitleName = mapInputValues.strSideBarMenuTitleName; // Not used
	let strSideBarMenuTitleXpath = mapInputValues.strSideBarMenuTitleXpath;
	let strSideBarMenuTitleTextXPath = mapInputValues.strSideBarMenuTitleTextXPath;
	let strSideBarMenuTitleExpandAttr = mapInputValues.strSideBarMenuTitleExpandAttr;
	let boolAutoExpTitle = mapInputValues.boolAutoExpTitle;
	let strSideBarMenuChildPanelXPath = mapInputValues.strSideBarMenuChildPanelXPath;
	// let strSideBarMenuChildPanelValueIsCollapsed = mapInputValues.strSideBarMenuChildPanelValueIsCollapsed; // Not used
	let strSideBarMenuChildItemXPath = mapInputValues.strSideBarMenuChildItemXPath;
	let strSideBarMenuChildItemSelectedProperty = mapInputValues.strSideBarMenuChildItemSelectedProperty;
	let strSideBarMenuChildItemSelectedPropertyValue = mapInputValues.strSideBarMenuChildItemSelectedPropertyValue;
	// let strSideBarMenuChildItemChildPropertyValue = mapInputValues.strSideBarMenuChildItemChildPropertyValue; // Not used
	// let strSideBarMenuChildItemElementXPath = mapInputValues.strSideBarMenuChildItemElementXPath; // Not used
	// let strElementStatusProperty = mapInputValues.strElementStatusProperty; // Not used
	let weTitleLink; // Type inferred

	//Return the count of title elements first.
	let mapGetTitleElements = {}; // Mimic Groovy Map
	let lstTitleElements; // Type inferred
	mapGetTitleElements = CWCore.returnChildElements(weSidebar, strSideBarMenuTitleXpath);
	lstTitleElements = mapGetTitleElements.lstChildObjects;
	intTitleElemCnt = mapGetTitleElements.cntChildObjs;
	if (intTitleElemCnt < 1) {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The '" + strSideBarName + "' DOES NOT contain any TITLES!!!";
	}
	else {
		//loop through the title items
		for (let loopTitleElems = 0; loopTitleElems < intTitleElemCnt && boolPassed; loopTitleElems++) { // Added boolPassed check
			//Return each child panel and expand it
			weTitle = lstTitleElements[loopTitleElems]; // Access array by index
			if (boolDoHighlight == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weTitle, 'Element', "Side Bar Title");
			}
			//Return the Expanded/Collapsed state
			if (strSideBarMenuTitleTextXPath != gblNA) {
					//Return the count of tab elements
					let mapGetLinkElements = CWCore.returnChildElements(weTitle, strSideBarMenuTitleTextXPath);
					let lstLinkElements = mapGetLinkElements.lstChildObjects;
					let intLinkElemCnt = mapGetLinkElements.cntChildObjs;
					if (intLinkElemCnt == 1) {
						weTitleLink = lstLinkElements[0]; // Access array by index
					}
					else {
						boolPassed = false;
						strMethodDetails = "FAILED!!! The title '" + strTempTitleText + "' DOES NOT APPEAR to contain a LINK ELEMENT!!!";
						break; // Break from loopTitleElems
					}
			}
			else { // If strSideBarMenuTitleTextXPath == gblNA, then weTitle itself is the link/expand element
				weTitleLink = weTitle;
			}
			if (weTitleLink != null) {
				//Return the title element value
				strTempTitleText = StringsAndNumbers.JComm_HandleNoData(weTitleLink.getText());
				if (boolDoHighlight == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weTitleLink, 'Element', "Title Link");
				}
				if (CWCore.isAttribtuePresent(weTitleLink, strSideBarMenuTitleExpandAttr)) {
					//Return value
					boolTempTitleExpanded = StringsAndNumbers.JComm_StringToBoolean(weTitleLink.getAttribute(strSideBarMenuTitleExpandAttr));
					Tester.Message("The expanded value is: " + boolTempTitleExpanded);
				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The ATTRIBUTE '" + strSideBarMenuTitleExpandAttr + "' WAS NOT FOUND in the title element!!!";
					break; // Break from loopTitleElems
				}
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The TITLE LINK was NOT RETURNED!!!";
				break; // Break from loopTitleElems
			}
			//Expand the panel if not already expanded
			if (boolAutoExpTitle == true && boolTempTitleExpanded == false && boolPassed == true) {
				// let mapMoveToElem = {}; // Mimic Groovy Map (commented out from original)
				// Use the existing global MoveToWebElement, assuming it takes webelement
				MoveToWebElement(strStartLocation, weTitleLink, strSideBarName); // Call to global func
				if (boolDoHighlight == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weTitleLink, 'Element', "Side Bar Title Link");
				}
				weTitle.click(); // Click the weTitle (main title element)
				DateTime.WaitSecs(intViewDelay);
				//Confirm in the correct state
				boolTempTitleExpanded = StringsAndNumbers.JComm_StringToBoolean(weTitleLink.getAttribute(strSideBarMenuTitleExpandAttr));
				if (boolTempTitleExpanded == false) { // If still not expanded
					boolPassed = false;
					strMethodDetails = "FAILED!!! Clicked the title '" + strTempTitleText + "' HOWEVER, it DOES NOT appear to be EXPANDED!!!";
					break; // Break from loopTitleElems
				}
			}
			//Return the panel
			let weTitlePanel; // Type inferred
			if (boolTempTitleExpanded == true && boolPassed == true) {
				//Return the count of title elements first.
				let mapGetPanelElements = {}; // Mimic Groovy Map
				let lstPanelElements; // Type inferred
				mapGetPanelElements = CWCore.returnChildElements(weTitle, strSideBarMenuChildPanelXPath);
				lstPanelElements = mapGetPanelElements.lstChildObjects;
				let intPanelElemCnt = mapGetPanelElements.cntChildObjs;
				if (intPanelElemCnt == 1) {
					weTitlePanel = lstPanelElements[0]; // Access array by index
					if (weTitlePanel == null) {
						boolPassed = false;
						strMethodDetails = "FAILED!!! The title '" + strTempTitleText + "' is expanded but does NOT APPEAR TO BE DISPLAYING A PANEL!!!";
						break; // Break from loopTitleElems
					}
					else {
						if (boolDoHighlight == true) {
							let mapHighlight = {}; // Mimic Groovy Map
							mapHighlight = CWCore.objHighlightElementJS(weTitlePanel, 'Element', "Side Bar Title Panel");
						}
					}
				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The title '" + strTempTitleText + "' panel count of: " + intPanelElemCnt + " DOES NOT MATCH the EXPECTED 1!!!";
					break; // Break from loopTitleElems
				}

			}
			//Return the child item count
			if (boolPassed == true) {
				let mapGetPanelItemElements = {}; // Mimic Groovy Map
				let lstPanelItemElements; // Type inferred
				mapGetPanelItemElements = CWCore.returnChildElements(weTitlePanel, strSideBarMenuChildItemXPath);
				lstPanelItemElements = mapGetPanelItemElements.lstChildObjects;
				let intPanelItemElemCnt = mapGetPanelItemElements.cntChildObjs;
				if (intPanelItemElemCnt >= 1) {
					//Loop through the child items
					if (intPanelItemElemCnt >= 1 && boolPassed == true) {
						for (let loopItems = 0; loopItems < intPanelItemElemCnt && boolPassed; loopItems++) { // Added boolPassed check
							strTempChildItemText = gblNA; // Used later in OutputSidebarMenu, here it's not set
							weTitleChildItem = lstPanelItemElements[loopItems]; // Access array by index
							// let mapMoveToItem = {}; // Declared but not used when ScrollToWebElement is internal
							// Call global func
							// The original Groovy code: mapMoveToItem = CommonWeb.ScrollToWebElement...
							// Here, calling directly using global func.
							ScrollToWebElement(strStartLocation, weTitleChildItem, strSideBarName); // Assuming strSideBarName is appropriate for element name
							if (boolDoHighlight == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weTitleChildItem, 'Element', "Side Bar Item");
							}
							//Return the text for the item
							strTempItemText = StringsAndNumbers.JComm_HandleNoData(weTitleChildItem.getText());
							//Check if the item is selected
							if (CWCore.isAttribtuePresent(weTitleChildItem, strSideBarMenuChildItemSelectedProperty)) {
								strTempItemPropertyValue = weTitleChildItem.getAttribute(strSideBarMenuChildItemSelectedProperty);
								if (strSideBarMenuChildItemSelectedPropertyValue == gblTrueFalse) { // Check for boolean string value
									//Check if boolean
									if (StringsAndNumbers.JComm_StringIsBoolean(strTempItemPropertyValue)) {
										boolItemSelected = StringsAndNumbers.JComm_StringToBoolean(strTempItemPropertyValue);
									}
									else {
										boolPassed = false;
										strMethodDetails = "FAILED!!! The value returned for the item selected of: " + strTempItemPropertyValue + " IS NOT A BOOLEAN value!!!";
										break; // Break from loopItems
									}
								}
								else { // Check if the value contains the selected text
									if ( StringsAndNumbers.JComm_VerifyTextPresent(strTempItemPropertyValue, strSideBarMenuChildItemSelectedPropertyValue, "%Con%") == true) {
										boolItemSelected = true;
									}
									else {
										boolItemSelected = false;
									}
								}
							}
							else {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The ATTRIBUTE '" + strSideBarMenuChildItemSelectedProperty + "' WAS NOT FOUND in the item element!!!";
								break; // Break from loopItems
							}
						}
					}
					// If the loop finished without boolPassed being false, then all checks passed here.
					// No specific boolPassed check needed for this block, as outer continue if (boolPassed) will catch it.
				}
			}
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure this is set for all paths
	return mapResults;
}

/** ------------------ OutputSidebarMenu -----------------------
* Output the Side bar Menu items and values
* @param mapInputValues The values to include input sheet, objects, and object xpaths
* @param NOTE: Each title will be expanded to identify each child panel list in the strResults
*
* @param	strStartLocation									The Start Location for the side-bar menu
* @param	weSidebar											THe web element of the side-bar menu
* @param	strSideBarFullXPath									The XPath for the side-bar menu within the results
* @param	strSideBarName										The name of the side-bar
* @param	strSideBarMenuTitleName								The side-bar title name
* @param	strSideBarMenuTitleXpath							The side-bar title XPath
* @param	strSideBarMenuTitleTextXPath						The side-bar title text XPath
* @param	strSideBarMenuTitleExpandAttr						The Attribute for side-bar title. To see if the title is expanded
* @param 	boolAutoExpTitle									The Auto Expanded title is true/false
* @param	strSideBarMenuChildPanelXPath						The side-bar menu Child Panel XPath
* @param	strSideBarMenuChildPanelValueIsCollapsed			The child Panel is expandable true/false
* @param	strSideBarMenuChildItemXPath						The child Item XPath
* @param	strSideBarMenuChildItemSelectedProperty				The child Item Selected Property
* @param	strSideBarMenuChildItemSelectedPropertyValue		The selected Property Value of the Child Item
* @param	strSideBarMenuChildItemChildPropertyValue			The Property Value of the Child Item
* @param	strSideBarMenuChildItemElementXPath					The XPath of the Child Item element
* @param	strSideBarMenuChildItemElementSvgXPath				The XPath for the Child Item
* @param	strSideBarMenuChildItemElementImageXPath			The Image XPath for the Child Item
* @param	strSideBarMenuChildItemElementBadgeXPath			The Badge XPath for the Child Item
* @param   strElementStatusProperty							The element Status Property
* @return mapResults Contains Passed/Failed and method details
*
* @author Created 01/03/2023
* @author PKanaris
* @author Last Edited: 01/25/2023
* @author Last Edited By: rbobodzh
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_OutputSidebarMenu (mapInputValues) { // mapInputValues is `any`, return `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value'); // Not used
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	let gblTrueFalse = GVars.GBLTrueFalse('Value');
	let boolDoHighlight = TCExecParams.getBoolDoHighlight();
	let intViewDelay = TCExecParams.getIntViewDelaySecs();
	//Declare the variables
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails; // Type inferred
	let boolPassed = true;
	let weTitle; // Type inferred
	let weTitleChildItemPanel; // Type inferred
	let weTitleChildItem; // Type inferred
	let weTitleChildItemText; // Type inferred
	let weTitleChildItemElem; // Type inferred
	let weTitleChildItemElemItem; // Type inferred
	let intTitleElemCnt; // Type inferred
	let intOutputRow; // Type inferred
	let strTempTitleText; // Type inferred
	let strTempTitleClass; // Type inferred (not defined in Groovy code provided)
	let strTempPanelClass; // Type inferred (not defined in Groovy code provided)
	let strTempItemText; // Type inferred
	// let strTempObjText; // Declared but unused
	let strTempItemPropertyValue; // Type inferred
	let strTempElemType; // Type inferred (not used to infer, but used in a fixed string directly in cwcore call)
	let strTempElemStatus; // Type inferred (not defined in Groovy code provided)
	let strTempElemText; // Type inferred (not defined in Groovy code provided)
	let strTempChildItemText; // Type inferred
	let strCellTheme; // Type inferred (Set below)
	//let mapSetCellValue = {}; // Mimic Groovy Map (Initialized below)
	let boolTempTitleExpanded; // Type inferred
	let boolIsChild; // Type inferred (not defined in Groovy code provided)
	let boolItemSelected; // Type inferred

	//Item Status object variables
	let strColor; // Type inferred
	let strBckgColor; // Type inferred
	let strTempStatusColor; // Type inferred (not defined in Groovy code provided)
	let boolCustomTheme; // Type inferred (not defined in Groovy code provided)

	//Return the values from the Map
	let strStartLocation = mapInputValues.strStartLocation;
	let weSidebar = mapInputValues.weSidebar;
	let strSideBarFullXPath = mapInputValues.strSideBarFullXPath; // Not used
	let strSideBarName = mapInputValues.strSideBarName;
	let strSideBarMenuTitleName = mapInputValues.strSideBarMenuTitleName;
	let strSideBarMenuTitleXpath = mapInputValues.strSideBarMenuTitleXpath;
	let strSideBarMenuTitleTextXPath = mapInputValues.strSideBarMenuTitleTextXPath;
	let strSideBarMenuTitleExpandAttr = mapInputValues.strSideBarMenuTitleExpandAttr;
	let boolAutoExpTitle = mapInputValues.boolAutoExpTitle;
	let strSideBarMenuChildPanelXPath = mapInputValues.strSideBarMenuChildPanelXPath;
	let strSideBarMenuChildPanelValueIsCollapsed = mapInputValues.strSideBarMenuChildPanelValueIsCollapsed; // Not used
	let strSideBarMenuChildItemXPath = mapInputValues.strSideBarMenuChildItemXPath;
	let strSideBarMenuChildItemSelectedProperty = mapInputValues.strSideBarMenuChildItemSelectedProperty;
	let strSideBarMenuChildItemSelectedPropertyValue = mapInputValues.strSideBarMenuChildItemSelectedPropertyValue;
	let strSideBarMenuChildItemChildPropertyValue = mapInputValues.strSideBarMenuChildItemChildPropertyValue; // Not used
	let strSideBarMenuChildItemElementXPath = mapInputValues.strSideBarMenuChildItemElementXPath;
	let strSideBarMenuChildItemElementSvgXPath = mapInputValues.strSideBarMenuChildItemElementSvgXPath;
	let strSideBarMenuChildItemElementImageXPath = mapInputValues.strSideBarMenuChildItemElementImageXPath;
	let strSideBarMenuChildItemElementBadgeXPath = mapInputValues.strSideBarMenuChildItemElementBadgeXPath;
	let strElementStatusProperty = mapInputValues.strElementStatusProperty;

	strCellTheme = 'OutputDataCentered'; //Set to output always
	let weTitleLink; // Type inferred

	//Return the count of title elements first.
	let mapGetTitleElements = {}; // Mimic Groovy Map
	let lstTitleElements; // Type inferred
	mapGetTitleElements = CWCore.returnChildElements(weSidebar, strSideBarMenuTitleXpath);
	lstTitleElements = mapGetTitleElements.lstChildObjects;
	intTitleElemCnt = mapGetTitleElements.cntChildObjs;
	if (intTitleElemCnt < 1) {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The '" + strSideBarName + "' DOES NOT contain any TITLES!!!";
	}
	else {
		intOutputRow = 0; //Used for the rows of output. Each title = 1 if no elements are present else for each item add 1
		//loop through the title items
		for (let loopTitleElems = 0; loopTitleElems < intTitleElemCnt && boolPassed; loopTitleElems++) { // Added boolPassed check
			//Return each child panel and expand it
			weTitle = lstTitleElements[loopTitleElems]; // Access array by index
			if (boolDoHighlight == true) {
				let mapHighlight = {}; // Mimic Groovy Map
				mapHighlight = CWCore.objHighlightElementJS(weTitle, 'Element', "Side Bar Title");
			}
			//Return the Expanded/Collapsed state
			if (strSideBarMenuTitleTextXPath != gblNA) {
					//Return the count of tab elements
					let mapGetLinkElements = CWCore.returnChildElements(weTitle, strSideBarMenuTitleTextXPath);
					let lstLinkElements = mapGetLinkElements.lstChildObjects;
					let intLinkElemCnt = mapGetLinkElements.cntChildObjs;
					if (intLinkElemCnt == 1) {
						weTitleLink = lstLinkElements[0]; // Access array by index
					}
					else {
						boolPassed = false;
						strMethodDetails = "FAILED!!! The title '" + strTempTitleText + "' DOES NOT APPEAR to contain a LINK ELEMENT!!!";
						break; // Break loopTitleElems
					}
			}
			else { // strSideBarMenuTitleTextXPath is gblNA
				weTitleLink = weTitle; //weTitle itself is the link/expand element
			}
			if (weTitleLink != null) {
				//Return the title element value
				strTempTitleText = StringsAndNumbers.JComm_HandleNoData(weTitleLink.getText());
				if (boolDoHighlight == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weTitleLink, 'Element', "Title Link");
				}
				if (CWCore.isAttribtuePresent(weTitleLink, strSideBarMenuTitleExpandAttr)) {
					//Return value
					boolTempTitleExpanded = StringsAndNumbers.JComm_StringToBoolean(weTitleLink.getAttribute(strSideBarMenuTitleExpandAttr));
					Tester.Message("The expanded value is: " + boolTempTitleExpanded);
				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The ATTRIBUTE '" + strSideBarMenuTitleExpandAttr + "' WAS NOT FOUND in the title element!!!";
					break; // Break loopTitleElems
				}
			}
			else {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The TITLE LINK was NOT RETURNED!!!";
				break; // Break loopTitleElems
			}
			//Expand the panel if not already expanded
			if (boolAutoExpTitle == true && boolTempTitleExpanded == false && boolPassed == true) {
				// mapMoveToElem commented out in original
				// CommonWeb.ScrollToWebElement uses WebElement arg, MoveToWebElement overload uses WebElement arg.
				// Assuming MoveToWebElement with WebElement is needed, as it is the only one without a fullxpath arg
				MoveToWebElement(strStartLocation, weTitleLink, strSideBarName); // Call to global func
				if (boolDoHighlight == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weTitleLink, 'Element', "Side Bar Title Link");
				}
				weTitle.click(); // Click the main title element
				DateTime.WaitSecs(intViewDelay);
				//Confirm in the correct state
				boolTempTitleExpanded = StringsAndNumbers.JComm_StringToBoolean(weTitleLink.getAttribute(strSideBarMenuTitleExpandAttr));
				if (boolTempTitleExpanded == false) { // If still not expanded after clicking
					boolPassed = false;
					strMethodDetails = "FAILED!!! Clicked the title '" + strTempTitleText + "' HOWEVER, it DOES NOT appear to be EXPANDED!!!";
					break; // Break loopTitleElems
				}
			}
			//Return the panel
			let weTitlePanel; // Type inferred
			if (boolTempTitleExpanded == true && boolPassed == true) {
				//Return the count of title elements first.
				let mapGetPanelElements = {}; // Mimic Groovy Map
				let lstPanelElements; // Type inferred
				mapGetPanelElements = CWCore.returnChildElements(weTitle, strSideBarMenuChildPanelXPath);
				lstPanelElements = mapGetPanelElements.lstChildObjects;
				let intPanelElemCnt = mapGetPanelElements.cntChildObjs;
				if (intPanelElemCnt == 1) {
					weTitlePanel = lstPanelElements[0]; // Access array by index
					if (weTitlePanel == null) {
						boolPassed = false;
						strMethodDetails = "FAILED!!! The title '" + strTempTitleText + "' is expanded but does NOT APPEAR TO BE DISPLAYING A PANEL!!!";
						break; // Break loopTitleElems
					}
					else {
						if (boolDoHighlight == true) {
							let mapHighlight = {}; // Mimic Groovy Map
							mapHighlight = CWCore.objHighlightElementJS(weTitlePanel, 'Element', "Side Bar Title Panel");
						}
					}
				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The title '" + strTempTitleText + "' panel count of: " + intPanelElemCnt + " DOES NOT MATCH the EXPECTED 1!!!";
					break; // Break loopTitleElems
				}

			}
			//Return the child item count
			if (boolPassed == true && weTitlePanel) { // Only proceed if weTitlePanel is defined
				let mapGetPanelItemElements = {}; // Mimic Groovy Map
				let lstPanelItemElements; // Type inferred
				mapGetPanelItemElements = CWCore.returnChildElements(weTitlePanel, strSideBarMenuChildItemXPath);
				lstPanelItemElements = mapGetPanelItemElements.lstChildObjects;
				let intPanelItemElemCnt = mapGetPanelItemElements.cntChildObjs;
				if (intPanelItemElemCnt >= 1) {
					//Loop through the child items
					if (intPanelItemElemCnt >= 1 && boolPassed == true) { // Added boolPassed check
						for (let loopItems = 0; loopItems < intPanelItemElemCnt && boolPassed; loopItems++) { // Added boolPassed check
							strTempChildItemText = gblNA; // Initialize for each item
							weTitleChildItem = lstPanelItemElements[loopItems]; // Access array by index
							// let mapMoveToItem = {}; // Used just here in original
							// Call global func. The original Groovy code: mapMoveToItem = CommonWeb.ScrollToWebElement...
							ScrollToWebElement(strStartLocation, weTitleChildItem, strSideBarName); // Assuming strSideBarName as ElemName

							if (boolDoHighlight == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weTitleChildItem, 'Element', "Side Bar Item");
							}
							//Return the text for the item
							strTempItemText = StringsAndNumbers.JComm_HandleNoData(weTitleChildItem.getText());
							//Check if the item is selected
							if (CWCore.isAttribtuePresent(weTitleChildItem, strSideBarMenuChildItemSelectedProperty)) {
								strTempItemPropertyValue = weTitleChildItem.getAttribute(strSideBarMenuChildItemSelectedProperty);
								if (strSideBarMenuChildItemSelectedPropertyValue == gblTrueFalse) { // Check for boolean string value
									//Check if boolean
									if (StringsAndNumbers.JComm_StringIsBoolean(strTempItemPropertyValue)) {
										boolItemSelected = StringsAndNumbers.JComm_StringToBoolean(strTempItemPropertyValue);
									}
									else {
										boolPassed = false;
										strMethodDetails = "FAILED!!! The value returned for the item selected of: " + strTempItemPropertyValue + " IS NOT A BOOLEAN value!!!";
										break; // Break from loopItems
									}
								}
								else { // Check if the value contains the selected text (e.g. '%Con%|selected')
									if ( StringsAndNumbers.JComm_VerifyTextPresent(strTempItemPropertyValue, strSideBarMenuChildItemSelectedPropertyValue, "%Con%") == true) {
										boolItemSelected = true;
									}
									else {
										boolItemSelected = false;
									}
								}

							}
							else {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The ATTRIBUTE '" + strSideBarMenuChildItemSelectedProperty + "' WAS NOT FOUND in the item element!!!";
								break; // Break from loopItems
							}
							intOutputRow++; // Increment output row here for each item
							//Output all of the values
							let mapSetCellValue = {}; // Mimic Groovy Map
							mapSetCellValue = ExcelData.excelSetCellValueByRowNumberColName(intOutputRow, "SideMenuTitle", strTempTitleText, strCellTheme);
							mapSetCellValue = ExcelData.excelSetCellValueByRowNumberColName(intOutputRow, "ExpandCollapse", boolTempTitleExpanded.toString(), strCellTheme);
							mapSetCellValue = ExcelData.excelSetCellValueByRowNumberColName(intOutputRow, "SideMenuItem", strTempItemText, strCellTheme);
							mapSetCellValue = ExcelData.excelSetCellValueByRowNumberColName(intOutputRow, "SideMenuItemSelected", boolItemSelected.toString(), strCellTheme);
						}
					}
				}
			} // End if (boolPassed == true && weTitlePanel) -- handled by outer loop boolPassed check

			//TODO do we need to do anything with images? (Original comment)
		} // End for loopTitleElems
		//Update the column width to fit
		if (boolPassed == true) {
			//TODO create a new function CommonWeb_to autofit the column
			let mapAutoFit = {}; // Mimic Groovy Map
			mapAutoFit = ExcelData.excelAutofitCols();
			boolPassed = StringsAndNumbers.JComm_StringToBoolean(mapAutoFit.boolPassed);
			if (boolPassed == true) {
				//TODO do we need any statistics?
				strMethodDetails = "Successfully output the menu items.";
			}
			else {
				strMethodDetails = mapAutoFit.strMethodDetails;
			}
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure this is set for all paths
	return mapResults;

}
/** ------------------ SelectSidebarMenuTitleItem -----------------------
* Select the Sidebar Menu Title Item based on Title title|menuitem|childItem
* @param mapInputValues The values to include title, item and childitem, objects, and object xpaths
* @param strValue The values of the title|item|childItem to select
*
* @param	strStartLocation									The Start Location for the side-bar menu
* @param	weSidebar											THe web element of the side-bar menu
* @param	strSideBarFullXPath									The XPath for the side-bar menu within the results
* @param	strSideBarName										The name of the side-bar
* @param	strSideBarMenuTitleName								The side-bar title name
* @param	strSideBarMenuTitleXpath							The side-bar title XPath
* @param	strSideBarMenuTitleTextXPath						The side-bar title text XPath
* @param	strSideBarMenuTitleExpandAttr						The Attribute for side-bar title. To see if the title is expanded
* @param 	boolAutoExpTitle									The Auto Expanded title is true/false
* @param	strSideBarMenuChildPanelXPath						The side-bar menu Child Panel XPath
* @param 	strSideBarMenuChildPanelValueIsCollapsed			The child Panel is expandable true/false
* @param	strSideBarMenuChildItemXPath						The child Item XPath
* @param	strSideBarMenuChildItemSelectedProperty				The child Item Selected Property
* @param	strSideBarMenuChildItemSelectedPropertyValue		The selected Property Value of the Child Item
* @param	strSideBarMenuChildItemChildPropertyValue			The Property Value of the Child Item
* @param	strSideBarMenuChildItemElementXPath					The XPath of the Child Item element
* @param  strElementStatusProperty							The element Status Property
*
* @return mapResults Contains Passed/Failed and method details
*
* @author Created 01/09/2023
* @author pkanaris
* @author Last Edited: 01/25/2023
* @author Last Edited By: rbobodzh
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_SelectSidebarMenuTitleItem (mapInputValues, strValue) { // mapInputValues is `any`, return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value'); // Not used in this function
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	// let gblTrueFalse = GVars.GBLTrueFalse('Value'); // Not used in this function
	let boolDoHighlight = TCExecParams.getBoolDoHighlight();
	let intViewDelay = TCExecParams.getIntViewDelaySecs();
	//Declare the variables
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails; // Type inferred from assignment
	let boolPassed = true;
	let weTitle; // Type inferred
	// let weTitleChildItemPanel; // Not used from source
	// let weTitleChildItem; // Not used from source
	// let weTitleChildItemText; // Not used from source
	// let weTitleChildItemElem; // Not used from source
	// let weTitleChildItemElemItem; // Not used from source
	let intTitleElemCnt; // Type inferred
	// let intOutputRow; // Not used
	let strTempTitleText; // Type inferred
	// let strTempTitleClass; // Not used
	// let strTempPanelClass; // Not used
	// let strTempItemText; // Not used
	// let strTempObjText; // Not used
	// let strTempItemPropertyValue; // Not used
	// let strTempElemType; // Not used
	// let strTempElemStatus; // Not used
	// let strTempElemText; // Not used
	// let strTempChildItemText; // Not used
	// let strCellTheme; // Not used
	// let mapSetCellValue = {}; // Not used
	let boolTempTitleExpanded; // Type inferred
	// let boolIsChild; // Not used
	// let boolItemSelected; // Not used
	// let boolCustomTheme; // Not used

	//Return the values from the Map
	let strStartLocation = mapInputValues.strStartLocation;
	let weSidebar = mapInputValues.weSidebar;
	// let strSideBarFullXPath = mapInputValues.strSideBarFullXPath; // Not used in this function
	let strSideBarName = mapInputValues.strSideBarName;
	// let strSideBarMenuTitleName = mapInputValues.strSideBarMenuTitleName; // Not used
	let strSideBarMenuTitleXpath = mapInputValues.strSideBarMenuTitleXpath;
	let strSideBarMenuTitleTextXPath = mapInputValues.strSideBarMenuTitleTextXPath;
	let strSideBarMenuTitleExpandAttr = mapInputValues.strSideBarMenuTitleExpandAttr;
	let boolAutoExpTitle = mapInputValues.boolAutoExpTitle;
	let strSideBarMenuChildPanelXPath = mapInputValues.strSideBarMenuChildPanelXPath;
	// let strSideBarMenuChildPanelValueIsCollapsed = mapInputValues.strSideBarMenuChildPanelValueIsCollapsed; // Not used
	let strSideBarMenuChildItemXPath = mapInputValues.strSideBarMenuChildItemXPath;
	// let strSideBarMenuChildItemSelectedProperty = mapInputValues.strSideBarMenuChildItemSelectedProperty; // Not used
	// let strSideBarMenuChildItemSelectedPropertyValue = mapInputValues.strSideBarMenuChildItemSelectedPropertyValue; // Not used
	// let strSideBarMenuChildItemChildPropertyValue = mapInputValues.strSideBarMenuChildItemChildPropertyValue; // Not used
	// let strSideBarMenuChildItemElementXPath = mapInputValues.strSideBarMenuChildItemElementXPath; // Not used
	// let strElementStatusProperty = mapInputValues.strElementStatusProperty; // Not used
	let weTitleLink; // Type inferred

	//Create Array from input values
	let arryElementNames = []; // Type inferred (Array)
	let intCntElemNames; // Type inferred (number)
	let mapSplitElementString = {}; // Mimic Groovy Map
	mapSplitElementString = StringsAndNumbers.JComm_StringToArray(strValue, gblDelimiter);
	intCntElemNames = StringsAndNumbers.JComm_StringToInteger(mapSplitElementString.intItemCount);
	if (intCntElemNames == 2) {
		arryElementNames = mapSplitElementString.ArryOfValues;
		//Starting with title, find the specified title and expand it.
		let strAssgTitle = StringsAndNumbers.JComm_HandleNoData(arryElementNames[0]);
		let strAssgChildItem = StringsAndNumbers.JComm_HandleNoData(arryElementNames[1]);
		//Return the title elements and attempt to expand the title assigned.
		let boolTitleFound = false; // Type inferred (boolean)
		let boolItemClicked = false; // Type inferred (boolean)
		//Capture the values checked for titles and child items for error messages
		let strTitles; // Type inferred (string)
		let strChildItems; // Type inferred (string)
		// let weTitleItem; // Declared but unused

		let mapGetTitleElements = {}; // Mimic Groovy Map
		let lstTitleElements; // Type inferred (List<WebElement>)
		let lstPanelItemElements; // Type inferred (List<WebElement>) // Used in later scope

		mapGetTitleElements = CWCore.returnChildElements(weSidebar, strSideBarMenuTitleXpath);
		lstTitleElements = mapGetTitleElements.lstChildObjects;
		intTitleElemCnt = mapGetTitleElements.cntChildObjs;
		if (intTitleElemCnt < 1) {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The '" + strSideBarName + "' DOES NOT contain any TITLES!!!";
		}
		else {
			// let intOutputRow = 0; // Not used in this particular function CommonWeb_// Used as comment in original
			//loop through the title items
			for (let loopTitleElems = 0; loopTitleElems < intTitleElemCnt && boolPassed; loopTitleElems++) { // Added boolPassed check
				//Return each child panel and expand it
				weTitle = lstTitleElements[loopTitleElems]; // Access array by index
				//This highlights the title and panel thus the text will be all the values PGK 03/31/2023
				if (boolDoHighlight == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weTitle, 'Element', "Side Bar Title");
				}
				//Must get the titlelink object so we can get the text
				// strTempTitleText assigned earlier, but will be re-assigned from weTitleLink.getText()
				strTempTitleText = StringsAndNumbers.JComm_HandleNoData(weTitle.getText());
				if (strSideBarMenuTitleTextXPath != gblNA) {
					//Return the count of elements
					let mapGetLinkElements = CWCore.returnChildElements(weTitle, strSideBarMenuTitleTextXPath);
					let lstLinkElements = mapGetLinkElements.lstChildObjects;
					let intLinkElemCnt = mapGetLinkElements.cntChildObjs;
					if (intLinkElemCnt == 1) {
						weTitleLink = lstLinkElements[0]; // Access array by index
					}
					else { // If link element not found or not unique
						boolPassed = false;
						strMethodDetails = "FAILED!!! The title '" + strTempTitleText + "' DOES NOT APPEAR to contain a LINK ELEMENT!!!";
						break; // Break from loopTitleElems
					}
				}
				else { // strSideBarMenuTitleTextXPath is gblNA
					weTitleLink = weTitle; //weTitle itself is the link/expand element
				}
				if (weTitleLink == null) {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The TITLE LINK was NOT RETURNED!!!";
					break; // Break from loopTitleElems
				}
				else {
					//Return the title element value
					strTempTitleText = StringsAndNumbers.JComm_HandleNoData(weTitleLink.getText());
					if (loopTitleElems == 0) {
						strTitles = strTempTitleText;
					}
					else {
						strTitles = strTitles + gblDelimiter + strTempTitleText;
					}
					//Does the temp title text match the assigned title
					if (strTempTitleText == strAssgTitle) {
						//Expand the item if it is not expanded
						boolTitleFound = true;
						// let mapMoveToElem = {}; // Original declared, but unused (ScrollToWebElement internal)
						//Move to the title which includes the link and panel
						MoveToWebElement(strStartLocation, weTitleLink, strSideBarName); // Call global func
						if (boolDoHighlight == true) {
							let mapHighlight = {}; // Mimic Groovy Map
							mapHighlight = CWCore.objHighlightElementJS(weTitleLink, 'Element', "Title Link");
						}
						//Return the Expanded/Collapsed state
						if (CWCore.isAttribtuePresent(weTitleLink, strSideBarMenuTitleExpandAttr)) {
							//Return value
							boolTempTitleExpanded = StringsAndNumbers.JComm_StringToBoolean(weTitleLink.getAttribute(strSideBarMenuTitleExpandAttr));
							Tester.Message("The expanded value is: " + boolTempTitleExpanded);
						}
						else {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The ATTRIBUTE '" + strSideBarMenuTitleExpandAttr + "' WAS NOT FOUND in the title element!!!";
							break; // Break from loopTitleElems
						}
						//Expand the panel if not already expanded
						if (boolAutoExpTitle == true && boolTempTitleExpanded == false && boolPassed == true) {
							if (boolDoHighlight == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weTitleLink, 'Element', "Side Bar Title Link");
							}
							weTitle.click(); // Click the weTitle (main title element)
							DateTime.WaitSecs(intViewDelay);
							//Confirm in the correct state
							boolTempTitleExpanded = StringsAndNumbers.JComm_StringToBoolean(weTitleLink.getAttribute(strSideBarMenuTitleExpandAttr));
							if (boolTempTitleExpanded == false) { // If still not expanded
								boolPassed = false;
								strMethodDetails = "FAILED!!! Clicked the title '" + strTempTitleText + "' HOWEVER, it DOES NOT appear to be EXPANDED!!!";
								break; // Break from loopTitleElems
							}
						}
						//Return the panel
						let weTitlePanel; // Type inferred
						if (boolTempTitleExpanded == true && boolPassed == true) { // Only if title is expanded and no previous errors
							//Return the count of title elements first.
							let mapGetPanelElements = {}; // Mimic Groovy Map
							let lstPanelElements; // Type inferred
							mapGetPanelElements = CWCore.returnChildElements(weTitle, strSideBarMenuChildPanelXPath);
							lstPanelElements = mapGetPanelElements.lstChildObjects;
							let intPanelElemCnt = mapGetPanelElements.cntChildObjs;
							if (intPanelElemCnt == 1) {
								weTitlePanel = lstPanelElements[0]; // Access array by index
								if (weTitlePanel == null) {
									boolPassed = false;
									strMethodDetails = "FAILED!!! The title '" + strTempTitleText + "' is expanded but does NOT APPEAR TO BE DISPLAYING A PANEL!!!";
									break; // Break from loopTitleElems
								}
								else {
									if (boolDoHighlight == true) {
										let mapHighlight = {}; // Mimic Groovy Map
										mapHighlight = CWCore.objHighlightElementJS(weTitlePanel, 'Element', "Side Bar Title Panel");
									}
								}
							}
							else {
								boolPassed = false;
								strMethodDetails = "FAILED!!! The title '" + strTempTitleText + "' panel count of: " + intPanelElemCnt + " DOES NOT MATCH the EXPECTED 1!!!";
								break; // Break from loopTitleElems
							}
						}
						//Return the child item count
						if (boolPassed == true && weTitlePanel) { // Only proceed if title is expanded and panel found
							let mapGetPanelItemElements = {}; // Mimic Groovy Map
							// let lstPanelItemElements; // Declared earlier.
							mapGetPanelItemElements = CWCore.returnChildElements(weTitlePanel, strSideBarMenuChildItemXPath);
							lstPanelItemElements = mapGetPanelItemElements.lstChildObjects; // Already defined above, but reassigned.
							let intPanelItemElemCnt = mapGetPanelItemElements.cntChildObjs;
							if (intPanelItemElemCnt >= 1) {
								//Loop through the child items
								if (intPanelItemElemCnt >= 1 && boolPassed == true) { // Added boolPassed check
									for (let loopItems = 0; loopItems < intPanelItemElemCnt && boolPassed; loopItems++) { // Added boolPassed check
										strTempChildItemText = gblNull; // Initialize for each item
										weTitleChildItem = lstPanelItemElements[loopItems]; // Access array by index
										// let mapMoveToItem = {}; // Declared but not used when ScrollToWebElement is internal
										// Call global func. The original Groovy code: mapMoveToItem = CommonWeb.MoveToWebElement...
										MoveToWebElement(strStartLocation, weTitleChildItem, strSideBarName); // Assuming strSideBarName as ElemName
										// Note: ScrollToWebElement not used in this path based on Groovy
										if (boolDoHighlight == true) {
											let mapHighlight = {}; // Mimic Groovy Map
											mapHighlight = CWCore.objHighlightElementJS(weTitleChildItem, 'Element', "Side Bar Item");
										}
										//Return the text for the item
										strTempChildItemText = StringsAndNumbers.JComm_HandleNoData(weTitleChildItem.getText());
										if (loopItems == 0) {
											strChildItems = strTempChildItemText;
										}
										else {
											strChildItems = strChildItems + gblDelimiter + strTempChildItemText;
										}
										//Does the child item text match the assigned item value
										if (strTempChildItemText == strAssgChildItem) {
											//Only move to the item if we will click it.
											// mapMoveToItem = CommonWeb.MoveToWebElement(strStartLocation, weTitleChildItem, strSideBarName); // Redundant if already moved
											//Click the item
											weTitleChildItem.click();
											//Show element was clicked
											boolItemClicked = true;
											DateTime.WaitSecs(intViewDelay);
											strMethodDetails = "Selected the '" + strValue + "'.";
											break; // Found and clicked, exit inner loop (loopItems)
										}
									}
									if (boolItemClicked == false) { // If loop over all items and didn't click
										boolPassed = false;
										strMethodDetails = "FAILED!!! Found the '"+ strAssgTitle + "' title, checked all '" + intPanelItemElemCnt + "' panel item(s)" +
										" but, DID NOT FIND THE CHILDITEM OF: " + strAssgChildItem +"!!!" + "Checked child items: " + strChildItems;
									}
								}
							}
						}
						if (boolTitleFound == true) { // Found title and processed item so break outer loop (loopTitleElems)
							break;
						}
					}
				}
				//Check if we found the title and report failure if not found
				if (boolTitleFound == false) {
					boolPassed = false;
					strMethodDetails = "FAILED!!! Checked the '"+ intTitleElemCnt + "' title(s) BUT DID NOT FIND THE ASSIGNED VALUE OF:" + strAssgTitle + "!!!" +
					"Check displayed titles of: " + strTitles;
				}
			}
		}
	}
	else { // intCntElemNames != 2
		boolPassed = false;
		strMethodDetails = "FAILED!!! THE COUNT OF ITEMS in the value '" + strValue + "' IS NOT 2!!!";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure this is set for all paths
	return mapResults;
}

/** ------------------ VerifySidebarMenu -----------------------
* Verify the Common Profile Sidebar menu items and values
* @param mapInputValues The values to include input sheet, objects, and object XPaths
* @param NOTE: Each cell will be checked for the elements listed in the strResultsObjNames
*
* @param 	INPUT SHEET Variables
* @param 	strInputShtData										The sheet name assigned for the input
*
* @param 	SIDEBAR Variables
* @param	strStartLocation									The Start Location for the side-bar menu
* @param	arryColName											The array containing the list of element names that can be in the results cells.
* @param	intExpectedInputColCnt								The expected count of elements where a cell contains more than one element. Currently only Title and Child item is available
* @param	strStartLocation									The Start Location for the side-bar menu
* @param	weSidebar											THe web element of the side-bar menu
* @param	strSideBarFullXPath									The XPath for the side-bar menu within the results
* @param	strSideBarName										The name of the side-bar
* @param	strSideBarMenuTitleName								The side-bar title name
* @param	strSideBarMenuTitleXpath							The side-bar title XPath
* @param	strSideBarMenuTitleTextXPath						The side-bar title text XPath
* @param	strSideBarMenuTitleExpandAttr						The Attribute for side-bar title. To see if the title is expanded
* @param 	boolAutoExpTitle									The Auto Expanded title is true/false
* @param	strSideBarMenuChildPanelXPath						The side-bar menu Child Panel XPath
* @param 	strSideBarMenuChildPanelValueIsCollapsed			The child Panel is expandable true/false
* @param	strSideBarMenuChildItemXPath						The child Item XPath
* @param	strSideBarMenuChildItemSelectedProperty				The child Item Selected Property
* @param	strSideBarMenuChildItemSelectedPropertyValue		The selected Property Value of the Child Item
* @param	strSideBarMenuChildItemElementXPath					The XPath of the Child Item element
* @return mapResults Contains Passed/Failed and method details
*
* @deprecated Use the newer version.
* @author Created 01/06/2023
* @author pkanaris
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifySidebarMenu (mapInputValues) { // mapInputValues and return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value'); // Not used
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblNA = GVars.GblNotApplicable('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	// let gblTrueFalse = GVars.GBLTrueFalse('Value'); // Not used
	let boolDoHighlight = TCExecParams.getBoolDoHighlight();
	let intViewDelay = TCExecParams.getIntViewDelaySecs();
	//Declare the variables
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails; // Type inferred from assignment
	let boolPassed = true;
	let weTitle; // Type inferred
	// let weTitleChildItemPanel; // Not used
	// let weTitleChildItem; // Not used
	// let weTitleChildItemText; // Not used
	// let weTitleChildItemElem; // Not used
	// let weTitleChildItemElemItem; // Not used
	let intTitleElemCnt; // Type inferred
	let intOutputRow; // Type inferred
	let strTempTitleText; // Type inferred
	// let strTempTitleClass; // Not used
	// let strTempPanelClass; // Not used
	let strTempItemText; // Type inferred
	// let strTempObjText; // Not used
	let strTempItemPropertyValue; // Type inferred
	// let strTempElemType; // Not used
	// let strTempElemStatus; // Not used
	// let strTempElemText; // Not used
	// let strTempChildItemText; // Not used
	let strCellTheme; // Type inferred
	// let mapSetCellValue = {}; // Mimic Groovy Map Initialization (initialized below)
	let boolTempTitleExpanded; // Type inferred
	// let boolIsChild; // Not used
	let boolItemSelected; // Type inferred

	//Item Status object variables
	// let strColor; // Not used
	// let strBckgColor; // Not used
	// let strTempStatusColor; // Not used
	// let strTempCellElements; // Not used
	// let boolCustomTheme; // Not used

	//Return the values from the Map
	let strStartLocation = mapInputValues.strStartLocation;
	let strInputShtData = mapInputValues.inputshtData;
	let arryColName = mapInputValues.ArryOfValues; // Assuming this means actual column names
	let intExpectedInputColCnt = mapInputValues.intExpectedInputColCnt;
	let weSidebar = mapInputValues.weSidebar;
	// let strSideBarFullXPath = mapInputValues.strSideBarFullXPath; // Not used
	let strSideBarName = mapInputValues.strSideBarName;
	// let strSideBarMenuTitleName = mapInputValues.strSideBarMenuTitleName; // Not used
	let strSideBarMenuTitleXpath = mapInputValues.strSideBarMenuTitleXpath;
	let strSideBarMenuTitleTextXPath = mapInputValues.strSideBarMenuTitleTextXPath;
	let strSideBarMenuTitleExpandAttr = mapInputValues.strSideBarMenuTitleExpandAttr; // Note: Groovy uses expandAttr for title but boolean `boolAutoExpTitle`
	let boolAutoExpTitle = mapInputValues.boolAutoExpTitle;
	let strSideBarMenuChildPanelXPath = mapInputValues.strSideBarMenuChildPanelXPath;
	// let strSideBarMenuChildPanelValueIsCollapsed = mapInputValues.strSideBarMenuChildPanelValueIsCollapsed; // Not used
	let strSideBarMenuChildItemXPath = mapInputValues.strSideBarMenuChildItemXPath;
	// let strSideBarMenuChildItemProperty = mapInputValues.strSideBarMenuChildItemProperty; // Not used
	let strSideBarMenuChildItemSelectedProperty = mapInputValues.strSideBarMenuChildItemSelectedProperty;
	let strSideBarMenuChildItemSelectedPropertyValue = mapInputValues.strSideBarMenuChildItemSelectedPropertyValue;
	let strSideBarMenuChildItemElementXPath = mapInputValues.strSideBarMenuChildItemElementXPath; // Not used
	// let strSideBarMenuChildItemElementSvgXPath = mapInputValues.strSideBarMenuChildItemElementSvgXPath; // Not used
	// let strSideBarMenuChildItemElementImageXPath = mapInputValues.strSideBarMenuChildItemElementImageXPath; // Not used
	// let strSideBarMenuChildItemElementBadgeXPath = mapInputValues.strSideBarMenuChildItemElementBadgeXPath; // Not used
	// let strElementStatusProperty = mapInputValues.strElementStatusProperty; // Not used
	let weTitleLink; // Type inferred

	//Open and set the Input sheet as the active input data.
	let strTempInputValue; // Type inferred

	//Load the input sheet
	let intInputRowCnt;
	let intInputColCnt;
	let intHdrColCnt; // Declared but unused

	let intCellPassed = 0;
	let intCellFailed = 0;
	let shInput; // Mimic XSSFSheet
	let mapOpenInputSheet = {}; // Mimic Groovy Map
	mapOpenInputSheet = ExcelData.excelGetSheetByName(TCObj.getObjWorkbook(), strInputShtData);
	if (StringsAndNumbers.JComm_StringToBoolean(mapOpenInputSheet.boolPassed) == true) {
		shInput = mapOpenInputSheet.objWbSheet;
		//Return the row and column Count
		let mapSheetRowColCnt = {}; // Mimic Groovy Map
		mapSheetRowColCnt = ExcelData.excelGetRowAndColCount(shInput);
		if (StringsAndNumbers.JComm_StringToBoolean(mapSheetRowColCnt.boolPassed) == true) {
			intInputRowCnt = mapSheetRowColCnt.RowCount;
			intInputColCnt = mapSheetRowColCnt.ColCount;
			//Verify the input column count is equal to the output
			if (intExpectedInputColCnt != intInputColCnt) {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The input column count of: " + intInputColCnt + " does NOT MATCH the EXPECTED: " + intExpectedInputColCnt + " column(s).";
			}
		}
		else {
			boolPassed = false;
			strMethodDetails = mapSheetRowColCnt.strMethodDetails;
		}
	}
	else {
		boolPassed = false;
		strMethodDetails = mapOpenInputSheet.strMethodDetails;
	}
	if (boolPassed == true) {
		//Return the count of title elements first.
		let mapGetTitleElements = {}; // Mimic Groovy Map
		let lstTitleElements; // Type inferred
		mapGetTitleElements = CWCore.returnChildElements(weSidebar, strSideBarMenuTitleXpath);
		lstTitleElements = mapGetTitleElements.lstChildObjects;
		intTitleElemCnt = mapGetTitleElements.cntChildObjs;
		if (intTitleElemCnt < 1) {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The '" + strSideBarName + "' DOES NOT contain any TITLES!!!";
		}
		else {
			intOutputRow = 0; //Used for the rows of output. Each title = 1 if no elements are present else for each item add 1
			//loop through the title items
			for (let loopTitleElems = 0; loopTitleElems < intTitleElemCnt && boolPassed; loopTitleElems++) { // Added boolPassed check
				//Return each child panel and expand it
				weTitle = lstTitleElements[loopTitleElems]; // Access array by index
				if (boolDoHighlight == true) {
					let mapHighlight = {}; // Mimic Groovy Map
					mapHighlight = CWCore.objHighlightElementJS(weTitle, 'Element', "Side Bar Title");
				}
				//Return the Expanded/Collapsed state
				if (strSideBarMenuTitleTextXPath != gblNA) {
					//Return the count of tab elements
					let mapGetLinkElements = CWCore.returnChildElements(weTitle, strSideBarMenuTitleTextXPath);
					let lstLinkElements = mapGetLinkElements.lstChildObjects;
					let intLinkElemCnt = mapGetLinkElements.cntChildObjs;
					if (intLinkElemCnt == 1) {
						weTitleLink = lstLinkElements[0]; // Access array by index
					}
					else {
						boolPassed = false;
						strMethodDetails = "FAILED!!! The title '" + strTempTitleText + "' DOES NOT APPEAR to contain a LINK ELEMENT!!!";
						break; // Break loopTitleElems
					}
				}
				else { // strSideBarMenuTitleTextXPath is gblNA
					weTitleLink = weTitle; //weTitle itself is the link/expand element
				}
				if (weTitleLink != null) {
					//Return the title element value
					strTempTitleText = StringsAndNumbers.JComm_HandleNoData(weTitleLink.getText()); // strTempTitleText might not be defined if strSideBarMenuTitleTextXPath was not gblNA
					if (boolDoHighlight == true) {
						let mapHighlight = {}; // Mimic Groovy Map
						mapHighlight = CWCore.objHighlightElementJS(weTitleLink, 'Element', "Title Link");
					}
					if (CWCore.isAttribtuePresent(weTitleLink, strSideBarMenuTitleExpandAttr)) {
						//Return value
						boolTempTitleExpanded = StringsAndNumbers.JComm_StringToBoolean(weTitleLink.getAttribute(strSideBarMenuTitleExpandAttr));
						Tester.Message("The expanded value is: " + boolTempTitleExpanded);
					}
					else {
						boolPassed = false;
						strMethodDetails = "FAILED!!! The ATTRIBUTE '" + strSideBarMenuTitleExpandAttr + "' WAS NOT FOUND in the title element!!!";
						break; // Break loopTitleElems
					}
				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! The TITLE LINK was NOT RETURNED!!!";
					break; // Break loopTitleElems
				}
				//Expand the panel if not already expanded
				if (boolAutoExpTitle == true && boolTempTitleExpanded == false && boolPassed == true) {
					// mapMoveToElem commented out // in original
					MoveToWebElement(strStartLocation, weTitleLink, strSideBarName); // Call global func
					if (boolDoHighlight == true) {
						let mapHighlight = {}; // Mimic Groovy Map
						mapHighlight = CWCore.objHighlightElementJS(weTitleLink, 'Element', "Side Bar Title Link");
					}
					weTitle.click(); // Click the weTitle (main title element)
					DateTime.WaitSecs(intViewDelay);
					//Confirm in the correct state
					boolTempTitleExpanded = StringsAndNumbers.JComm_StringToBoolean(weTitleLink.getAttribute(strSideBarMenuTitleExpandAttr));
					if (boolTempTitleExpanded == false) { // If still not expanded
						boolPassed = false;
						strMethodDetails = "FAILED!!! Clicked the title '" + strTempTitleText + "' HOWEVER, it DOES NOT appear to be EXPANDED!!!";
						break; // Break loopTitleElems
					}
				}
				//Return the panel
				let weTitlePanel; // Type inferred
				if (boolTempTitleExpanded == true && boolPassed == true) { // Only if title is expanded and no previous errors
					//Return the count of title elements first.
					let mapGetPanelElements = {}; // Mimic Groovy Map
					let lstPanelElements; // Type inferred
					mapGetPanelElements = CWCore.returnChildElements(weTitle, strSideBarMenuChildPanelXPath);
					lstPanelElements = mapGetPanelElements.lstChildObjects;
					let intPanelElemCnt = mapGetPanelElements.cntChildObjs;
					if (intPanelElemCnt == 1) {
						weTitlePanel = lstPanelElements[0]; // Access array by index
						if (weTitlePanel == null) {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The title '" + strTempTitleText + "' is expanded but does NOT APPEAR TO BE DISPLAYING A PANEL!!!";
							break; // Break loopTitleElems
						}
						else {
							if (boolDoHighlight == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weTitlePanel, 'Element', "Side Bar Title Panel");
							}
						}
					}
					else {
						boolPassed = false;
						strMethodDetails = "FAILED!!! The title '" + strTempTitleText + "' panel count of: " + intPanelElemCnt + " DOES NOT MATCH the EXPECTED 1!!!";
						break; // Break loopTitleElems
					}
				}
				//Return the child item count
				if (boolPassed == true && weTitlePanel) { // Only proceed if title is expanded and panel found
					let mapGetPanelItemElements = {}; // Mimic Groovy Map
					let lstPanelItemElements; // Type inferred
					mapGetPanelItemElements = CWCore.returnChildElements(weTitlePanel, strSideBarMenuChildItemXPath);
					lstPanelItemElements = mapGetPanelItemElements.lstChildObjects;
					let intPanelItemElemCnt = mapGetPanelItemElements.cntChildObjs;
					if (intPanelItemElemCnt >= 1) {
						//Loop through the child items
						if (intPanelItemElemCnt >= 1 && boolPassed == true) { // Added boolPassed check
							for (let loopItems = 0; loopItems < intPanelItemElemCnt && boolPassed; loopItems++) { // Added boolPassed check
								// strTempChildItemText, strTempItemText, strTempItemPropertyValue, boolItemSelected
								// Declared in outer scope.
								weTitleChildItem = lstPanelItemElements[loopItems]; // Access array by index
								// Call global func
								ScrollToWebElement(strStartLocation, weTitleChildItem, strSideBarName);
								if (boolDoHighlight == true) {
									let mapHighlight = {}; // Mimic Groovy Map
									mapHighlight = CWCore.objHighlightElementJS(weTitleChildItem, 'Element', "Side Bar Item");
								}
								//Return the text for the item
								strTempItemText = StringsAndNumbers.JComm_HandleNoData(weTitleChildItem.getText());
								//Check if the item is selected
								if (CWCore.isAttribtuePresent(weTitleChildItem, strSideBarMenuChildItemSelectedProperty)) {
									strTempItemPropertyValue = weTitleChildItem.getAttribute(strSideBarMenuChildItemSelectedProperty);
									if (strSideBarMenuChildItemSelectedPropertyValue == GVars.GBLTrueFalse('Value')) { // Correctly uses the globally defined gblTrueFalse
										//Check if boolean
										if (StringsAndNumbers.JComm_StringIsBoolean(strTempItemPropertyValue)) {
											boolItemSelected = StringsAndNumbers.JComm_StringToBoolean(strTempItemPropertyValue);
										}
										else {
											boolPassed = false;
											strMethodDetails = "FAILED!!! The value returned for the item selected of: " + strTempItemPropertyValue + " IS NOT A BOOLEAN value!!!";
											break; // Break from loopItems
										}
									}
									else { // Check if the value contains the selected text (e.g. '%Con%|selected')
										if ( StringsAndNumbers.JComm_VerifyTextPresent(strTempItemPropertyValue, strSideBarMenuChildItemSelectedPropertyValue, "%Con%") == true) {
											boolItemSelected = true;
										}
										else {
											boolItemSelected = false;
										}
									}
								}
								else {
									boolPassed = false;
									strMethodDetails = "FAILED!!! The ATTRIBUTE '" + strSideBarMenuChildItemSelectedProperty + "' WAS NOT FOUND in the item element!!!";
									break; // Break from loopItems
								}
								//Increment the excel row
								intOutputRow++;
								//Check the results against the input and update as needed
								let strTempColName;
								let strTempValue; // Renamed to avoid shadowing
								let strTempCompareResults;
								// Here, shInput and arryColName are from the outer function CommonWeb_scope. intOutputRow will be used.
								// strTempInputValue needs to be read.
								let strTempInputValue;

								let mapSetCellValue = {}; // Mimic Groovy Map initialization
								for (let intLoopOutput = 0; intLoopOutput < intExpectedInputColCnt && boolPassed; intLoopOutput++) { // Added boolPassed check
									strTempColName = arryColName[intLoopOutput]; // Access array by index
									if (intOutputRow <= intInputRowCnt) {
										strTempInputValue = ExcelData.excelGetCellValueByRowNumColName(shInput, intOutputRow, strTempColName).CellValue;
									}
									else { // Output row exceeds input row count
										strTempInputValue = "OUTPUTROW>INPUTROWCNT";
									}
									switch (intLoopOutput) {
										case 0: //ExpandCollapse
											strTempValue = boolTempTitleExpanded.toString();
											break;
										case 1: //SideMenuTitle
											strTempValue = strTempTitleText;
											break;
										case 2: //SideMenuItem
											strTempValue = strTempItemText;
											break;
										case 3: //SideMenuItemSelected
											strTempValue = boolItemSelected.toString();
											break;
										default: // For additional columns beyond these known 4
											strTempValue = gblNull; // Default or handle appropriately
											break;
									}
									if (strTempInputValue == strTempValue) {
										strCellTheme = 'TestRunPassStd';
										strTempCompareResults = strTempValue;
										intCellPassed++;
									}
									else {
										strCellTheme = 'TestRunFailStd';
										strTempCompareResults = `Expected: ${strTempInputValue} Displayed: ${strTempValue}`;
										intCellFailed++;
									}
									//Output the cell value
									mapSetCellValue = ExcelData.excelSetCellValueByRowNumberColName(intOutputRow, strTempColName, strTempCompareResults, strCellTheme);
								}
							}
						}
					}
				}
			}
			//Update the column width to fit
			let mapAutoFit = {}; // Mimic Groovy Map
			mapAutoFit = ExcelData.excelAutofitCols();
			boolPassed = StringsAndNumbers.JComm_StringToBoolean(mapAutoFit.boolPassed);
			if (intCellFailed == 0 && intCellPassed > 0 && boolPassed == true) {
				strMethodDetails = "Successfully process " + intOutputRow + " items totaling " + intCellPassed + " values.";
			}
			else if (intCellFailed > 0) {
				boolPassed = false;
				strMethodDetails = "FAILED DURING THE PROCESSING of: " + intOutputRow + " items totaling " + intCellFailed + " FAILED, and " + intCellPassed + " passed values.";
			}
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure this is set for all paths
	return mapResults;
}

/* Side Bar Menu Items using UL as Treeview
* Must support unlimited branches and leafs
*/
/**
* -------------------------------------  VerifyPermissionsLeftNavMenuItems  -----------------------------------
* Verify the Menu Items
* NOTE: Support the left navigation as identified in JTR-12848
* @param mapInputVariables The input variables that identify the header to include xpaths and properties
* @param Note: when strProperty* == gblNull we do not process the property.
* @param strInputShtName			The sheet name assigned for the input
* @param intInputDataStartRow		The input data row number to start using for verification
* @param intInputDataEndRow			The input data row number to end verification. Enter 999 for all rows
* @param boolMatchInputRowNumber	Must the order of displayed rows equal the input row order? true/false
* @param strColumnsExpected			The names of the columns that must be present in the input sheet
* @param strResultsOutShtName		The sheet name assigned for the output
*
* @return mapResults		 The results showing Passed and method details.
*
* @author kpluedde
* @author Created: 10/27/2023
* @author Last Edited:
* @author Last Edited By:
* @author Edit Comments: (Include email, date and details)
*/
function CommonWeb_VerifyPermissionsLeftNavMenuItems (mapInputVariables) { // mapInputVariables is `any`, return type `any`
	//GlobalVars
	let gblNull = GVars.GblNull('Value');
	let gblUndefined = GVars.GblUndefined('Value'); // Not used
	let gblLineFeed = GVars.GblLineFeed('Value');
	let gblDelimiter = GVars.GblDelimiter('Value');
	let gblSkip = GVars.GblSkip('Value'); // Not used
	let gblIgnoreData = GVars.GblIgnoreData('Value'); // Not used
	let gblNA = GVars.GblNotApplicable('Value');
	// let gblUseInputColName = GVars.GBLUseInputColName('Value'); // Not used
	let boolDoHighlight = TCExecParams.getBoolDoHighlight();
	let boolDoDebug = TCExecParams.getBoolDoDebug();
	// Declare the variables
	let mapResults = {}; // Mimic Groovy Map initialization
	let strMethodDetails; // Type inferred from assignment
	let boolPassed = true;
	//Return the method variables
	//Input Sheet variables
	let strStartLoc = mapInputVariables.strStartLoc; // Not used as weMenu is hardcoded
	let strInputSheetName = mapInputVariables.InputDataSheetName;
	let intInputDataStartRow = mapInputVariables.InputDataRowStart; // Not used
	let intInputDataEndRow = mapInputVariables.InputDataRowEnd; // Not used
	let boolMatchInputRowNumber = mapInputVariables.boolMatchInputRowNumber; // Not used
	let strColumnsExpected = mapInputVariables.strColumnsExpected;
	let strResultsOutShtName = mapInputVariables.strResultsOutShtName;

	// let intResultsRowCnt; // Not used
	// let boolProcCellElements; // Not used
	// let arryOutputColNames; // Not used

	//Check if the outputsheet name is valid
	let intLenOptResShName = strResultsOutShtName.length;
	if (intLenOptResShName < 32) {
		//Note, we open the file and do the work for each method and save/close the excel instance so save happens as we go.
		//Add the sheet to the input file
		let mapAddSheet = {}; // Mimic Groovy Map
		let strTCInputFilePath = TCObj.strTCInputFilePath;
		//TODO attempt to add the input stream to the test objects and use the input object to add the sheet in a new TObj WB and Sh
		mapAddSheet = ExcelData.excelAddSheetToWorkBook(strResultsOutShtName);
		let boolSheetAddedPassed = StringsAndNumbers.JComm_StringToBoolean (mapAddSheet.boolPassed);
		if (boolSheetAddedPassed == true) {
			//Add the column names
			// let strOutputColNames; // Not used
			// let strHdrColNames; // Not used
			let arryColName; // Type inferred
			//Add column names to the array
			let mapStringToArry = {}; // Mimic Groovy Map
			mapStringToArry = StringsAndNumbers.JComm_StringToArray(strColumnsExpected, gblDelimiter);
			arryColName = mapStringToArry.ArryOfValues;
			let mapAddColNames = {}; // Mimic Groovy Map
			mapAddColNames = ExcelData.excelCreateHeaderCols(strColumnsExpected);
			let boolAddColNames = StringsAndNumbers.JComm_StringToBoolean (mapAddColNames.boolPassed);
			if (boolAddColNames == true) {
				//Return the header elements
				//let lstColNames = mapAddColNames.lstColNames; // Declared but not used
				//Update the header column(s) to the assigned theme
				let strAssgTheme = 'OutputDataHdrStd';
				let mapUpdateHdrTheme = {}; // Mimic Groovy Map
				mapUpdateHdrTheme = ExcelData.excelSetRowCellFormatByTheme(0, strAssgTheme);
				let boolUPdHdrTheme = StringsAndNumbers.JComm_StringToBoolean (mapUpdateHdrTheme.boolPassed);
				if (boolUPdHdrTheme == false) {
					boolPassed = false;
					strMethodDetails = mapUpdateHdrTheme.strMethodDetails;
				}
				//Update the column width to fit
				let mapAutoFit = {}; // Mimic Groovy Map
				mapAutoFit = ExcelData.excelAutofitCols();
				if (StringsAndNumbers.JComm_StringToBoolean(mapAutoFit.boolPassed) == false) {
					strMethodDetails = strMethodDetails + mapAutoFit.strMethodDetails;
				}
			}
			else {
				boolPassed = false;
				strMethodDetails = mapAddColNames.strMethodDetails;
			}

			//Split to an array and get the count of input columns
			// let arryInputColNames = []; // Declared but used differently
			let mapSplitElementString = {}; // Mimic Groovy Map
			let intInputColExpCnt; // Type inferred
			let arryExpInputColNames; // Type inferred (string array)
			mapSplitElementString = StringsAndNumbers.JComm_StringToArray(strColumnsExpected, gblDelimiter);
			intInputColExpCnt = StringsAndNumbers.JComm_StringToInteger(mapSplitElementString.intItemCount);
			if (intInputColExpCnt == 0) {
				boolPassed = false;
				strMethodDetails = "FAILED!!! The INPUT COLUMNS EXPECTED OF: '" + strColumnsExpected + " contains ZERO COLUMN NAMES!!!";
			}
			else {
				//Assign the column names to an array
				arryExpInputColNames = mapSplitElementString.ArryOfValues;

				//Load the inputsheet
				let intInputRowCnt;
				let intInputColCnt;
				let shInput; // Mimic XSSFSheet
				let mapOpenInputSheet = {}; // Mimic Groovy Map
				mapOpenInputSheet = ExcelData.excelGetSheetByName(TCObj.getObjWorkbook(), strInputSheetName);
				if (StringsAndNumbers.JComm_StringToBoolean(mapOpenInputSheet.boolPassed) == true) {
					shInput = mapOpenInputSheet.objWbSheet;
					//Return the row and column Count
					let mapSheetRowColCnt = {}; // Mimic Groovy Map
					mapSheetRowColCnt = ExcelData.excelGetRowAndColCount(shInput);
					if (StringsAndNumbers.JComm_StringToBoolean(mapSheetRowColCnt.boolPassed) == true) {
						intInputRowCnt = mapSheetRowColCnt.RowCount;
						intInputColCnt = mapSheetRowColCnt.ColCount;
					}
					else {
						boolPassed = false;
						strMethodDetails = mapSheetRowColCnt.strMethodDetails;
					}
					//Return the input sheet column names and check if the match the output column names
					let strInputColValues;
					let mapGetInputColNames = {}; // Mimic Groovy Map
					mapGetInputColNames = ExcelData.excelGetHdrColNames(shInput);
					let boolInputColNames = StringsAndNumbers.JComm_StringToBoolean(mapGetInputColNames.boolPassed);
					if (boolInputColNames == true) {
						strInputColValues = mapGetInputColNames.ColValues;
						//Check if the match the output column names
						if (strInputColValues != strColumnsExpected) {
							boolPassed = false;
							strMethodDetails = "FAILED!!! The input column names '" + strInputColValues + "' DOES NOT MATCH the EXPECTED INPUT Columns '" + strColumnsExpected +"'!!!";
						}
						else {
							strMethodDetails = "The expected column names matched and the input file contains: '" + strInputColValues + "' column name(s).";
						}
					}
					else {
						boolPassed = false;
						strMethodDetails = mapGetInputColNames.strMethodDetails;
					}
				}
				else {
					boolPassed = false;
					strMethodDetails = mapOpenInputSheet.strMethodDetails;
				}

				//Return the menu object
				let strMenuFullXpath = "//div[@class='Menu' and @role = 'navigation' and @id = 'Menu']//ul[@class='MenuGroups' and @role='list']";
				// let strCategoryMenuButton = "//a[@role = 'button' and @aria-expanded='false']"; // Not used but in groovy as comment
				let strMenuCategoryItemXpath = "./li//a[contains(@class,'MenuItem')]";
				let strMenuCategoryMenu = "//div[contains(@class,'NavMenuContainer slideFromLeft') and contains (@style,'display') and contains(@style,'block')"; // Incomplete XPath
				let strMenuCategoryMenuSubCatItemXpath = "./div[contains(@class,'MenuSection Paged withHeader')]";
				let strMenuCatMenuSubCatItemTextXpath = "./a[@role='menuitem']//h2";
				let strMenuCatMenuSubCatFeatureMenuXpath = "//ul[@id="; // Incomplete XPath
				let strMenuCatMenuSubCatFeatureMenuItemXpath = "./descendant::a//span";
				let intInputRow = 0; // Initialize
				let weMenu; // Type inferred (WebElement)
				weMenu = CWCore.returnWebElement(strMenuFullXpath);
				if (weMenu == null) {
					//FAILS
					boolPassed = false;
					strMethodDetails = "FAILED!!! The Menu '" + strMenuFullXpath + "' DID NOT RETURN an ELEMENT!!!"; // Fixed string concatenation
				}
				else {
					//Do highlight
					if (boolDoHighlight == true) {
						let mapHighlight = {}; // Mimic Groovy Map
						mapHighlight = CWCore.objHighlightElementJS(weMenu, 'Element', 'Left Nav Menu');
					}
					//Category validate
					//Get the category items into an array
					let intCntCategoryItems; // Type inferred
					let intCntSubCatItems; // Type inferred
					// let intCurrentSubCatItem; // Not used
					let intCntFeatureItems; // Type inferred
					let intColIndex; // Type inferred (initialized inside loop)
					let weCatButton; // Type inferred
					let weCatMenu; // Type inferred
					let weSubCatMenu; // Type inferred
					let weSubCatButton; // Type inferred
					let weSubCatButtonText; // Type inferred
					let weFeatureButton; // Type inferred

					let strCategoryMenuXpath; // Type inferred
					let strCategoryItemXpath; // Not used
					let strSubCatFullXpath; // Not used
					let strCategoryItemSubMenuXpath; // Not used
					let strElemName; // Not used
					let strTempCatValue; // Type inferred
					let strTempInputCat; // Type inferred
					let strTempNextCatValue; // Type inferred
					let strTempSubCatValue; // Type inferred
					let strTempNextSubCatValue; // Type inferred
					let strTempInputSubCat; // Type inferred
					let strTempFeatureValue; // Type inferred
					let strTempNextFeatureValue; // Type inferred
					let strFeatMenuXPath; // Type inferred
					let strTempOutputCatValue; // Type inferred
					let strTempOutputSubCatValue; // Type inferred
					// Type inferred from OutputSidebarMenu, not used here
					// let strTempOutpuFeatureValue;
					let strTempOutputCatDisplayed; // Type inferred
					let strTempOutputSubCatDisplayed; // Type inferred
					let strTempOutputTheme; // Type inferred

					let boolTempInputCatDisplayed; // Type inferred
					let boolCatButtonPassed; // Type inferred
					let boolCatDoOutput; // Type inferred
					let boolSubCatDoOutput; // Type inferred
					let boolFeatureDoOutput; // Type inferred

					let boolTempInputSubCatDisplayed; // Type inferred
					let boolTempInputFeatureDisplayed; // Type inferred
					let boolSubCatButtonPassed; // Type inferred

					let lstMenuCategoryButtons; // Type inferred
					let lstMenuSubCatButtons; // Type inferred
					let lstMenuFeatureButtons; // Type inferred

					lstMenuCategoryButtons = weMenu.findElements(By.xpath(strMenuCategoryItemXpath)); // findElements returns a List
					intCntCategoryItems = lstMenuCategoryButtons.length; // Use .length for JS Array

					boolCatDoOutput = true;
					let boolBreak = false; // Type inferred
					let mapSetCellValue = {}; // Mimic Groovy Map initialization
					let boolSwitchCategory = false;
					boolCatDoOutput = true; // Redundant, already true

					//Loop through the category buttons and return the text
					for (let intLoopCat = 0; intLoopCat < intCntCategoryItems && boolPassed && !boolBreak; intLoopCat++) { // Added boolPassed & boolBreak checks
						//Refresh the menu items as it may be out of sync
						lstMenuCategoryButtons = weMenu.findElements(By.xpath(strMenuCategoryItemXpath));
						intCntCategoryItems = lstMenuCategoryButtons.length; // Use .length
						weCatButton = lstMenuCategoryButtons[intLoopCat]; // Access array by index
						if (weCatButton == null) {
							//FAIL
							boolPassed = false;
							strMethodDetails = "FAILED!!! The Catalog Button '" + lstMenuCategoryButtons[intLoopCat] + "' DID NOT RETURN an ELEMENT!!!"; // Fixed string concat
							break; // Break loop
						}
						else {
							//Do Highlight
							if (boolDoHighlight == true) {
								let mapHighlight = {}; // Mimic Groovy Map
								mapHighlight = CWCore.objHighlightElementJS(weCatButton, 'Element', "Catagory Instance: " + intLoopCat);
							}
							//Return the value
							strTempCatValue = StringsAndNumbers.JComm_HandleNoData(weCatButton.getAttribute("aria-label"));//Hard coded since this should not change
							//Loop through the input rows while the value matches the strTempCatValue
							let boolTempCatMatch = true;
							let boolCatBtnClicked = false;
							while (boolTempCatMatch == true && boolPassed == true && !boolBreak) { // Added boolPassed & boolBreak checks
								if (boolCatDoOutput == false) {
									boolCatDoOutput = true; // This branch is effectively skipped by boolSubCatDoOutput/boolFeatureDoOutput later
								}
								else {
									intInputRow = intInputRow + 1; // Increment for Excel row.
								}
								if (intInputRow > intInputRowCnt) { // Check before accessing Excel
									boolBreak = true; // Break from enclosing loops if no more input rows.
									break;
								}

								// Handle Excel read only once if successful
								let excelReadResultCat = ExcelData.excelGetCellValueByRowNumColName(shInput, intInputRow, arryExpInputColNames[0]);
								let excelReadResultCatDisplay = ExcelData.excelGetCellValueByRowNumColName(shInput, intInputRow, arryExpInputColNames[1]);

								strTempOutputCatDisplayed = gblNull; // Initialize
								strTempNextCatValue = gblNull;
								//Return the input category value
								strTempInputCat = StringsAndNumbers.JComm_HandleNoData(excelReadResultCat.CellValue);
								boolTempInputCatDisplayed = StringsAndNumbers.JComm_StringToBoolean(excelReadResultCatCatDisplay.CellValue);
								//Check if match and if they should match
								let mapResultsClickCategoryMenu = {}; // Mimic Groovy Map
								if (strTempInputCat == strTempCatValue && boolTempInputCatDisplayed == true) {
									//passes
									strTempOutputCatValue = strTempInputCat;
									strTempOutputCatDisplayed = "true";
									//Set the theme
									strTempOutputTheme = 'TestRunPassStd';
									Tester.Message("The value of the category of: " + strTempCatValue + " Passed");
									if (strTempCatValue != "Home") {
										//Attempt to click the category item
										boolCatBtnClicked = true;
										//Return the id and then get the name of the item
										let strTempCatBtnID = weCatButton.getAttribute("id");
										let strGetLeft = StringsAndNumbers.JComm_GetLeftTextInString(strTempCatBtnID, "_Invoker");
										//Create the path to the category element and sub-category menu
										strCategoryMenuXpath = strMenuFullXpath + strMenuCategoryMenu + " and contains(@id, '" + strGetLeft + "')]"; // Fixed XPath
										mapResultsClickCategoryMenu = ClickWebElement(weCatButton, strTempCatValue, "Category Button"); // Call global func
										DateTime.WaitSecs(TCExecParams.getIntViewDelaySecs());
										//Check the results and mark pass or fail
										boolCatButtonPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsClickCategoryMenu.boolPassed);
										if (boolCatButtonPassed == true) {
											Tester.Message("The Menu Button of the category of: " + strTempCatValue + " was Clicked");
										}
										else {
											strMethodDetails = "FAILED!!! The Menu Button of '" + strTempCatValue + " was not Clicked!!!"; // Fixed string concat
										}
									}
								}
								else if (strTempInputCat != strTempCatValue) {
									boolSwitchCategory = true;
									//passes
									//Check if the next button is a match
									//Are there more button left?
									let intCatNextButton = intLoopCat + 1;
									if (intCatNextButton < intCntCategoryItems) {
										strTempNextCatValue = StringsAndNumbers.JComm_HandleNoData(lstMenuCategoryButtons[intLoopCat + 1].getAttribute("aria-label")); // Access array by index
									}
									//Check if the next button is what we will look for
									if (strTempInputCat == strTempNextCatValue && boolTempInputCatDisplayed == true) {
										//passes
										boolCatDoOutput = false;
										boolTempCatMatch = false;
									}
									else if (strTempInputCat != strTempNextCatValue && strTempNextCatValue != "Out of Buttons" && boolTempInputCatDisplayed == true) { // Fixed "intCntSubCatItems" => "Out of Buttons"
										//FAILS output the data and break
										strTempOutputCatValue = "Expected: " + strTempInputCat + " ACTUAL: " + strTempCatValue;
										strTempOutputCatDisplayed = "Expected: true ACTUAL: false";
										//Set the theme
										strTempOutputTheme = 'TestRunFailStd';
										boolBreak = true;
										boolTempCatMatch = false;
										boolPassed = false;
										strMethodDetails = "FAILED!!! The InputCategory of '" + strTempInputCat +
										"' is the next category DISPLAYED but NEXT CATEGORY is '" + strTempNextCatValue + "'!!!";
									}
									else if (strTempInputCat != strTempNextCatValue && boolTempInputCatDisplayed == false) {
										//Pass category should not displayed
										strTempOutputCatValue = strTempInputCat;
										strTempOutputCatDisplayed = "false";
										//Set the theme
										strTempOutputTheme = 'TestRunPassStd';
									} else { // No specific matching or failing condition found
										// Covers cases where strTempInputCat != strTempNextCatValue but the next is "Out of Buttons" or boolTempInputCatDisplayed is false
										boolCatDoOutput = false;
										boolTempCatMatch = false;
									}
								}
								//Sub category
								if (strTempOutputCatDisplayed == "true") {
									// Handle Excel read only once if successful
									let excelReadResultSubCat = ExcelData.excelGetCellValueByRowNumColName(shInput, intInputRow, arryExpInputColNames[2]);
									let excelReadResultSubCatDisplay = ExcelData.excelGetCellValueByRowNumColName(shInput, intInputRow, arryExpInputColNames[3]);

									//Return the input sub category value
									strTempInputSubCat = StringsAndNumbers.JComm_HandleNoData(excelReadResultSubCat.CellValue);
									//Return the input sub category is displayed
									boolTempInputSubCatDisplayed = StringsAndNumbers.JComm_StringToBoolean(excelReadResultSubCatDisplay.CellValue);
									//Do the sub category by returning the object will need a counter
									//Is the menu present
									weCatMenu = null;
									weCatMenu = CWCore.returnWebElement(strCategoryMenuXpath); //Get the Menu
									if (weCatMenu == null && strTempInputSubCat == gblNA && boolTempInputSubCatDisplayed == false) {
										//Passes since the Category menu is not present
										//TODO add the output for showing this is correct
										//Pass category should not be displayed
										strTempOutputSubCatValue = strTempInputSubCat;
										strTempOutputSubCatDisplayed = "false";
										//Set the theme
										strTempOutputTheme = 'TestRunPassStd';
										boolSubCatDoOutput = true;
									}
									else if (weCatMenu == null && boolCatBtnClicked == true) {
										//FAIL
										boolPassed = false;
										strMethodDetails = "FAILED!!! The Catalog Menu '" + strCategoryMenuXpath + "' DID NOT RETURN an ELEMENT!!!"; // Fixed string concat
									}
									else if (weCatMenu != null) {
										strCategoryMenuXpath = null; // Clear to prevent re-use
										// let strMenuID = weCatMenu.getAttribute("id"); // Declared but not used when used here.

										boolSubCatDoOutput = false; //Set to false so we do not advance the input row. Will change to true in while.
										//Return the menu subcats
										lstMenuSubCatButtons = weCatMenu.findElements(By.xpath(strMenuCategoryMenuSubCatItemXpath));
										intCntSubCatItems = lstMenuSubCatButtons.length; // Use .length
										boolSubCatDoOutput = false; //Set to false so we do not advance the input row. Will change to true in while.
										//If we have SubCategory items loop through them
										let boolSwitchSubCat = false;
										for (let intLoopSubCat = 0; intLoopSubCat < intCntSubCatItems && boolPassed && !boolBreak; intLoopSubCat++) { // Added boolPassed & boolBreak checks
											// boolBreak = false; // Reset boolBreak for this inner loop (not used in groovy for this)
											//Refresh the list since we will be clicking the sub catagories
											lstMenuSubCatButtons = weCatMenu.findElements(By.xpath(strMenuCategoryMenuSubCatItemXpath)); // Re-get for fresh list
											intCntSubCatItems = lstMenuSubCatButtons.length; // Use .length
											weSubCatButton = lstMenuSubCatButtons[intLoopSubCat]; // Access array by index
											if (weSubCatButton == null) {
												//FAIL
												boolPassed = false;
												strMethodDetails = "FAILED!!! The Sub Catagory Button '" + lstMenuSubCatButtons[intLoopSubCat] + "' DID NOT RETURN an ELEMENT!!!"; // Fixed string concat
												break; // Break loop
											}
											else {
												//Do Highlight
												if (boolDoHighlight == true) {
													let mapHighlight = {}; // Mimic Groovy Map
													mapHighlight = CWCore.objHighlightElementJS(weSubCatButton, 'Element', "Sub Catagory Instance: " + intLoopSubCat);
												}
												//Return the element that holds the text
												weSubCatButtonText = CWCore.returnChildElement(weSubCatButton, strMenuCatMenuSubCatItemTextXpath);
												//Do Highlight
												if (boolDoHighlight == true) {
													let mapHighlight = {}; // Mimic Groovy Map
													mapHighlight = CWCore.objHighlightElementJS(weSubCatButtonText, 'Element', "Sub Catagory Instance Text: " + intLoopSubCat);
												}
												//Return the value
												strTempSubCatValue = StringsAndNumbers.JComm_HandleNoData(weSubCatButtonText.getText());
												let boolTempSubCatMatch = true;
												let boolSubCatBtnClicked = false;
												//Loop through the input rows while the value matches the strTempSubCatValue
												while (boolTempSubCatMatch == true && boolPassed == true && !boolBreak) { // Added boolPassed & boolBreak checks
													if (boolSubCatDoOutput == false) {
														boolSubCatDoOutput = true;
													}
													else {
														intInputRow = intInputRow + 1; // Increment for Excel row.
													}
													// Check for out of bounds
													if (intInputRow > intInputRowCnt) {
														boolBreak = true;
														break;
													}

													// Handle Excel read only once if successful
													let excelReadResultSubCatInner = ExcelData.excelGetCellValueByRowNumColName(shInput, intInputRow, arryExpInputColNames[2]);
													let excelReadResultSubCatDisplayInner = ExcelData.excelGetCellValueByRowNumColName(shInput, intInputRow, arryExpInputColNames[3]);

													strTempNextSubCatValue = gblNull;
													// if ( strTempSubCatValue=="Site Appearance and Behavior") {
													// 	Tester.Message("One more Sub-Cat should remain")
													// }
													// if (intInputRow == intInputRowCnt -1) {
													// 	Tester.Message("Process the row of input number:" + intInputRow)
													// }
													//Return the input sub category value
													strTempInputSubCat = StringsAndNumbers.JComm_HandleNoData(excelReadResultSubCatInner.CellValue);
													//Return the input category is displayed
													boolTempInputSubCatDisplayed = StringsAndNumbers.JComm_StringToBoolean(excelReadResultSubCatDisplayInner.CellValue);
													//Check if match and if they should match
													boolSubCatBtnClicked = false;
													let mapResultsClickSubCatMenu = {}; // Mimic Groovy Map
													if (strTempInputSubCat == strTempSubCatValue && boolTempInputSubCatDisplayed == true) {
														//passes
														boolSubCatBtnClicked = true;
														strTempOutputSubCatValue = strTempInputSubCat;
														strTempOutputSubCatDisplayed = "true";
														//Set the theme
														strTempOutputTheme = 'TestRunPassStd';
														Tester.Message("The value of the subcategory is: " + strTempSubCatValue + " Passed");
														//Return the button id
														let strSubCatBtn = weSubCatButton.getAttribute("id");
														//String strFeatMenuXPath = strCategoryMenuXpath + strMenuCatMenuSubCatFeatureMenuXpath + "'" + strSubCatBtn + "_groupMenu']"
														strFeatMenuXPath = `${strMenuCatMenuSubCatFeatureMenuXpath}'${strSubCatBtn}_groupMenu']`; // Using template literal
														//Click Sub Category Button
														mapResultsClickSubCatMenu = ClickWebElement(weSubCatButton, strTempSubCatValue, "SubCategory Button"); // Call global func
														//Check the results and mark pass or fail
														boolSubCatButtonPassed = StringsAndNumbers.JComm_StringToBoolean (mapResultsClickSubCatMenu.boolPassed);
														if (boolSubCatButtonPassed == true) {
															Tester.Message("The Menu Button of the subcategory of: " + strTempSubCatValue + " was Clicked");
														}
														else {
														strMethodDetails = "FAILED!!! The Menu Button of '" + strTempSubCatValue + " was not Clicked!!!"; // Fixed string concat
														}
													}
													else if (strTempInputSubCat != strTempSubCatValue) {
														//passes
														boolSwitchSubCat = true;
														boolTempSubCatMatch = false; //Set so we go to the next button by exiting the while loop
														boolSubCatBtnClicked = false;
														//Check if the next button is a match
														//Are there more button left?
														let intNextButton = intLoopSubCat + 1;
														if (intNextButton < intCntSubCatItems) {
															let weTempSubCatButton = lstMenuSubCatButtons[intNextButton]; // Access array by index
															strTempNextSubCatValue = StringsAndNumbers.JComm_HandleNoData(CWCore.returnChildElement(weTempSubCatButton,strMenuCatMenuSubCatItemTextXpath).getText());
														}
														else {
															strTempNextSubCatValue = "Out of Buttons";
														}
														//Check if the next button is what we will look for
														if (strTempInputSubCat == strTempNextSubCatValue && boolTempInputSubCatDisplayed == true) {
															//passes
															boolSubCatDoOutput = false;
															boolTempSubCatMatch = false;
															boolSwitchSubCat = false;
														}
														else if (strTempInputSubCat != strTempNextSubCatValue && strTempNextSubCatValue != "Out of Buttons" && boolTempInputSubCatDisplayed == true) {
															//FAILS output the data and break
															strTempOutputSubCatValue = `Expected: ${strTempInputSubCat} ACTUAL: ${strTempSubCatValue}`; // Using template literal
															strTempOutputSubCatDisplayed = "Expected: true ACTUAL: false";
															//Set the theme
															strTempOutputTheme = 'TestRunFailStd';
															boolBreak = true; // Set outer loop break
															boolTempSubCatMatch = false;
															boolPassed = false;
															strMethodDetails = `FAILED!!! The Input Sub Category of '${strTempInputSubCat}' is the next sub category DISPLAYED but NEXT SUB CATEGORY is '${strTempNextSubCatValue}'!!!`; // Using template literal
															//intCntSubCatItems = intCntSubCatItems - 1 //subtract 1 so we stay on the same button for the next loop // Original comment
															//intInputRow = intInputRow + 1 // Original comment
															//boolDoOutput = true // Original comment
														}
														else if (strTempInputSubCat != strTempNextSubCatValue && boolTempInputSubCatDisplayed == false) {
															//Pass category should not displayed
															strTempOutputSubCatValue = strTempInputSubCat;
															strTempOutputSubCatDisplayed = "false";
															//Set the theme
															strTempOutputTheme = 'TestRunPassStd';
															boolTempSubCatMatch = false;
															//intInputRow = intInputRow + 1 // Original comment
															//boolDoOutput = true // Original comment
														}
														else {
															boolSubCatDoOutput = false;
															boolTempSubCatMatch = false;
														}
													}
													//Add the feature Here
													if (strTempOutputSubCatDisplayed == "true" && boolSubCatBtnClicked == true && boolPassed == true && !boolBreak) { // Added boolPassed & boolBreak checks
														// Handle Excel read only once if successful
														let excelReadResultFeatureInner = ExcelData.excelGetCellValueByRowNumColName(shInput, intInputRow, arryExpInputColNames[4]);
														let excelReadResultFeatureDisplayInner = ExcelData.excelGetCellValueByRowNumColName(shInput, intInputRow, arryExpInputColNames[5]);

														//Return the input Feature value
														strTempInputFeature = StringsAndNumbers.JComm_HandleNoData(excelReadResultFeatureInner.CellValue);
														//Return the input feature is displayed
														boolTempInputFeatureDisplayed = StringsAndNumbers.JComm_StringToBoolean(excelReadResultFeatureDisplayInner.CellValue);
														if (weSubCatButton == null && strTempInputFeature == gblNA && boolTempInputFeatureDisplayed == false) {
															//Passes since the Category menu is not present
															//TODO add the output for showing this is correct
															//Pass sub category should not be displayed
															strTempOutputFeatureValue = strTempInputFeature;
															strTempOutputFeatureDisplayed = "false";
															//Set the theme
															strTempOutputTheme = 'TestRunPassStd';
															//boolBreak = true // Original comment
															boolFeatureDoOutput = true;
															//boolTempSubCatMatch = false // Original comment
														}
														else if (weSubCatButton == null && boolSubCatBtnClicked == true) { // If weSubCatButton is null but it was supposed to be clicked
															//FAIL
															boolPassed = false;
															strMethodDetails = "FAILED!!! The Sub Catagory Button '" + lstMenuSubCatButtons[intLoopSubCat] + "' DID NOT RETURN an ELEMENT!!!"; // Fixed string concat
														}
														else if (weSubCatButton != null) {
															boolSubCatDoOutput = false; // Original
															//Return the Feature menu
															//WebElement weFeatureMenu = CWCore.returnWebElement("//ul[contains(@style,'z-index: 5;')]") // Original commented out
															let weFeatureMenu = CWCore.returnWebElement(strFeatMenuXPath); // Using strFeatMenuXPath
															if (weFeatureMenu != null) {
																strFeatMenuXPath = null; // Clear XPath reference
																//Do not highlight as it will change the menu properties and not allow you to find the children
																lstMenuFeatureButtons = weFeatureMenu.findElements(By.xpath(strMenuCatMenuSubCatFeatureMenuItemXpath));
																intCntFeatureItems = lstMenuFeatureButtons.length; // Use .length
																boolFeatureDoOutput = false; //Set to false so we do not advance the input row. Will change to true in while.
															}
															else { // Feature menu not found
																boolPassed = false;
																strMethodDetails = "FAILED!!! THE FEATURE MENU '" + strSubCatBtn + "' XPath DID NOT RETURN AND ELEMENT"; // Fixed string concat
															}
															//If we have Features items loop through them
															for (let intLoopFeature = 0; intLoopFeature < intCntFeatureItems && boolPassed && !boolBreak; intLoopFeature++) { // Added boolPassed & boolBreak checks
																//Refresh the list
																lstMenuFeatureButtons = weFeatureMenu.findElements(By.xpath(strMenuCatMenuSubCatFeatureMenuItemXpath));
																intCntFeatureItems = lstMenuFeatureButtons.length; // Use .length
																weFeatureButton = lstMenuFeatureButtons[intLoopFeature]; // Access array by index
																if (weFeatureButton == null) {
																	//FAIL
																	boolPassed = false;
																	strMethodDetails = "FAILED!!! The Feature Button '" + lstMenuFeatureButtons[intLoopFeature] + "' DID NOT RETURN an ELEMENT!!!"; // Fixed string concat
																	break; // Break loop
																}
																else {
																	//Do Highlight
																	if (boolDoHighlight == true) {
																		let mapHighlight = {}; // Mimic Groovy Map
																		mapHighlight = CWCore.objHighlightElementJS(weFeatureButton, 'Element', "Feature Instance: " + intLoopFeature);
																	}
																	//Return the value
																	strTempFeatureValue = StringsAndNumbers.JComm_HandleNoData(weFeatureButton.getText());
																	let boolTempFeatureMatch = true;
																	//Loop through the input rows while the value matches the strTempFeatureValue
																	while (boolTempFeatureMatch == true && boolPassed == true && !boolBreak) { // Added boolPassed & boolBreak check
																		if (boolFeatureDoOutput == false) {
																			boolFeatureDoOutput = true;
																		}
																		else {
																			intInputRow = intInputRow + 1; // Increment for Excel row.
																		}
																		strTempNextFeatureValue = gblNull;
																		if (intInputRow > intInputRowCnt) {
																			Tester.Message("The intInputRow is greater then the number of rows in the inputsheet.");
																			boolBreak = true; // Break from enclosing loops
																			break;
																		}
																		//Return the input Feature value
																		// Handle Excel read only once if successful
																		let excelReadResultFeatureFinal = ExcelData.excelGetCellValueByRowNumColName(shInput, intInputRow, arryExpInputColNames[4]);
																		let excelReadResultFeatureDisplayFinal = ExcelData.excelGetCellValueByRowNumColName(shInput, intInputRow, arryExpInputColNames[5]);


																		strTempInputFeature = StringsAndNumbers.JComm_HandleNoData(excelReadResultFeatureFinal.CellValue);
																		//Return the input category is displayed
																		boolTempInputFeatureDisplayed = StringsAndNumbers.JComm_StringToBoolean(excelReadResultFeatureDisplayFinal.CellValue);
																		//Check if match and if they should match
																		if (strTempInputFeature == strTempFeatureValue && boolTempInputFeatureDisplayed == true) {
																			//passes
																			strTempOutputFeatureValue = strTempInputFeature;
																			strTempOutputFeatureDisplayed = "true";
																			//Set the theme
																			strTempOutputTheme = 'TestRunPassStd';
																			Tester.Message("The value of the Feature is: " + strTempInputFeature + " Passed");
																		}
																		else if (strTempInputFeature != strTempFeatureValue) {
																			//passes
																			boolTempFeatureMatch = false; //Set so we go to the next button by exiting the while loop
																			//Check if the next button is a match
																			//Are there more button left?
																			let intNextButton = intLoopFeature + 1;
																			if (intNextButton < intCntFeatureItems) {
																				let weTempFeatureButton = lstMenuFeatureButtons[intNextButton]; // Acccess array by index
																				strTempNextFeatureValue = StringsAndNumbers.JComm_HandleNoData(weTempFeatureButton.getText());
																			}
																			else {
																				strTempNextFeatureValue = "Out of Buttons";
																			}
																			//Check if the next button is what we will look for
																			if (strTempInputFeature == strTempNextFeatureValue && boolTempInputFeatureDisplayed == true) {
																				//passes
																				boolFeatureDoOutput = false;
																				boolTempFeatureMatch = false;
																			}
																			else if (strTempInputFeature != strTempNextFeatureValue && strTempNextFeatureValue != "Out of Buttons" && boolTempInputFeatureDisplayed == true) {
																				//FAILS output the data and break
																				strTempOutputFeatureValue = `Expected: ${strTempInputFeature} ACTUAL: ${strTempFeatureValue}`; // Using template literal
																				strTempOutputFeatureDisplayed = "Expected: true ACTUAL: false";
																				//Set the theme
																				strTempOutputTheme = 'TestRunFailStd';
																				// PGK 01302024 boolBreak = true; // Original comment
																				boolTempFeatureMatch = false;
																				boolPassed = false;
																				strMethodDetails = `FAILED!!! The Input Feature of '${strTempInputFeature}' is the next Feature DISPLAYED but NEXT FEATURE is '${strTempNextFeatureValue}'!!!`; // Using template literal
																				//intLoopFeature = intLoopFeature - 1 //subtract 1 so we stay on the same button for the next loop // Original comment
																			}
																			else if (strTempInputFeature != strTempNextFeatureValue && boolTempInputFeatureDisplayed == false) {
																				//Pass Feature should not be displayed
																				strTempOutputFeatureValue = strTempInputFeature;
																				strTempOutputFeatureDisplayed = "false";
																				//Set the theme
																				strTempOutputTheme = 'TestRunPassStd';
																				boolTempFeatureMatch = false;
																				//intLoopFeature = intLoopFeature - 1 // Original comment
																			}
																			else { // General case, not match, not pass, not fail
																				boolFeatureDoOutput = false;
																				boolTempFeatureMatch = false;
																			}
																		}
																		//Output the feature level values
																		if (intInputRow <= intInputRowCnt){
																			if (boolFeatureDoOutput == true) {
																				//Output to the excel file
																				intColIndex = 0;
																				mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, strTempOutputCatValue, strTempOutputTheme);
																				if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
																					boolPassed = false;
																					strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
																				}
																				intColIndex = 1;
																				mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, strTempOutputCatDisplayed, strTempOutputTheme);
																				if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
																					boolPassed = false;
																					strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
																				}
																				//Output to the excel file
																				intColIndex = 2;
																				mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, strTempOutputSubCatValue, strTempOutputTheme);
																				if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
																					boolPassed = false;
																					strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
																				}
																				intColIndex = 3;
																				mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, strTempOutputSubCatDisplayed, strTempOutputTheme);
																				if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
																					boolPassed = false;
																					strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
																				}
																				intColIndex = 4;
																				mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, strTempOutputFeatureValue, strTempOutputTheme);
																				if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
																					boolPassed = false;
																					strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
																				}
																				intColIndex = 5;
																				mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, 'false', strTempOutputTheme);
																				if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) { // This is for `false`
																					boolPassed = false;
																					strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
																				}
																			}
																		}
																		else {
																			boolBreak = true;
																		}
																		if (boolBreak == true) {
																			break;
																		}
																	}
																}
																//Temp Code (original comment)
																// if (intLoopFeature == intCntFeatureItems -1) {
																// 	Tester.Message("The loop should exit")
																// }
															}
														}
													}
													// Output Sub Categories
													if (intInputRow <= intInputRowCnt){
														if (boolSubCatDoOutput == true ) {
															//Output to the excel file
															intColIndex = 0;
															mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, strTempOutputCatValue, strTempOutputTheme);
															if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
																boolPassed = false;
																strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
															}
															intColIndex = 1;
															mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, strTempOutputCatDisplayed, strTempOutputTheme);
															if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
																boolPassed = false;
																strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
															}
															//Output to the excel file
															intColIndex = 2;
															mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, strTempOutputSubCatValue, strTempOutputTheme);
															if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
																boolPassed = false;
																strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
															}
															intColIndex = 3;
															mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, strTempOutputSubCatDisplayed, strTempOutputTheme);
															if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
																boolPassed = false;
																strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
															}
															intColIndex = 4;
															mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, gblNA, strTempOutputTheme);
															if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
																boolPassed = false;
																strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
															}
															intColIndex = 5;
															mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, 'false', strTempOutputTheme);
															if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
																boolPassed = false;
																strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
															}
														}
													}
													else {
														boolBreak = true;
													}
													if (boolBreak == true) {
														break;
													}
												}
											}
										}
										boolTempCatMatch = false; // Original Groovy sets this here
									}
								}
								//PGK 02082024 for debug purposes (original comment)
								// if (intInputRow == intInputRowCnt) {
								// 	Tester.Message("The intInputRow equals the count of input rows.")
								// }
								// Output Categories
								if (intInputRow <= intInputRowCnt){
									if (boolCatDoOutput == true ){
										//Output to the excel file
										intColIndex = 0;
										mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, strTempOutputCatValue, strTempOutputTheme);
										if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
											boolPassed = false;
											strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
										}
										intColIndex = 1;
										mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, strTempOutputCatDisplayed, strTempOutputTheme);
										if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
											boolPassed = false;
											strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
										}
										intColIndex = 2;
										mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, strTempOutputSubCatValue, strTempOutputTheme);
										if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
											boolPassed = false;
											strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
										}
										intColIndex = 3;
										mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, strTempOutputSubCatDisplayed, strTempOutputTheme);
										if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
											boolPassed = false;
											strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
										}
										intColIndex = 4;
										mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, gblNA, strTempOutputTheme);
										if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
											boolPassed = false;
											strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
										}
										intColIndex = 5;
										mapSetCellValue = ExcelData.excelSetCellValueByRowColNumber(intInputRow, intColIndex, 'false', strTempOutputTheme);
										if (StringsAndNumbers.JComm_StringToBoolean(mapSetCellValue.boolPassed) == false) {
											boolPassed = false;
											strMethodDetails = strMethodDetails + mapSetCellValue.strMethodDetails;
										}
									}
								}
								else {
									boolBreak = true;
								}
								if (boolBreak == true) {
									break;
								}
							}
						}
					}
				}
			}
		}
		else {
			boolPassed = false;
			strMethodDetails = mapAddSheet.strMethodDetails; // mapAddSheet is from outer scope
		}
	}

	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails; // Ensure this is set for all paths
	return mapResults;
}