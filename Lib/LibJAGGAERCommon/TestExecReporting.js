// Describe the library purpose
/**
 * The library will process the reporting the to appropriate reporting methods below:
 * Extent Reports
 * XRay JSON and API
 * TestNG
 * HTML
 * XML
 * Excel
 */

SeSGlobalObject("TestExecReporting");

// Add Imports
// Note: org.apache.commons.lang3.exception.ExceptionUtils is not used in the direct logic of these methods.
// java.net.InetAddress is imported but not used in provided snippet.
// org.openqa.selenium.WebElement will need a Selenium/WebDriver client library in Node.js.

// --- Placeholder for ExtentReports equivalents ---
// There is no direct Node.js/TypeScript equivalent for ExtentReports.
// You would need to choose a new reporting library.
// For example, if using 'Allure':
// import { AllureRuntime, Status, Stage, StepInterface } from 'allure-js-commons';
// For the purpose of this template, we'll use generic placeholder types.
// You will replace these with actual types/interfaces from your chosen reporting library.
// type ExtendedReportType = any; // Represents ExtentReports
// type TestStepType = any;	 // Represents ExtentTest
// type StatusType = any;	   // Represents com.aventstack.extentreports.Status
// type ColorType = any;		// Represents ExtentColor
// type MarkupHelperType = any; // Represents MarkupHelper
// ---------------------------------------------------

// Add Shared Resources
//import * as CommonWeb from '../web/CommonWeb'; // Represents resources.web.CommonWeb

// Add Jaggaer Libs (assuming these will also be converted to TypeScript modules)
//import * as TCObj from './TestObjects'; // Represents resources.common.TestObjects
//import * as GVars from './GlobalVariables'; // Represents resources.common.GlobalVariables
//import * as StrNums from './StringsAndNumbers'; // Represents resources.common.StringsAndNumbers
//import * as TSExecParams from './TestCaseExecParams'; // Represents resources.common.TestCaseExecParams
// ExtentLogging will be a wrapper for your chosen reporting library
//import * as ExtLogging from './ExtentLogging'; // Represents resources.common.ExtentLogging
//import * as JiraXrayInt from './JiraXrayIntegration'; // Represents resources.common.JiraXrayIntegration

/**
 * Re-implementation of TestExecReporting class as a collection of functions.
 * Public static methods are converted to exported functions.
 */


/**
 * -------------------------------------  GetStepExecStatusAsString  -----------------------------------
 * Return the string values based on the status code ID:
 * @param intStatus
 * @return strStatus
 * Created: 04/21/2021
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function TestExecReporting_GetExecStatusAsString(intStatus) {
	let strStatus = '';
	switch (intStatus) {
		case 0:
			strStatus = 'PASS';
			break;
		case 1:
			strStatus = 'FAIL';
			break;
		default:
			strStatus = 'UNKNOWN STATUS CODE!!!!';
			break;
	}
	return strStatus;
}

/**
 * -------------------------------------  GetTestStepStatus  -----------------------------------
 * Return the string values based on the status code ID:
 * @param intStatus
 * @return strStatus
 * Created: MM/DD/YYYY
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function TestExecReporting_GetTestStepStatus(intStatus) {
	let strStatus = '';
	switch (intStatus) {
		case 0:
			strStatus = 'PASS';
			break;
		case 1:
			strStatus = 'FAIL';
			break;
		case 2:
			strStatus = 'TODO';
			break;
		case 3:
			strStatus = 'EXECUTING';
			break;
		case 4:
			strStatus = 'SKIPPED';
			break;
		case 5:
			strStatus = 'NOT-APPLICABLE';
			break;
		default:
			strStatus = 'UNKNOWN STATUS CODE!!!!';
			break;
	}
	return strStatus;
}

/**
 * -------------------------------------  ReportStepResults  -----------------------------------
 * Reports the test results using extent
 * @param mapStepInput (Placeholder for a Map/Object containing step details)
 * Created: 04/21/2021
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function TestExecReporting_ReportStepResults() {
//Remove the variable once we have the global variable working
	//Increment the step count ADD IN TEST OBJECTS SO WE CAN USE THEM
	//TCObj.setIntActionNumber(TCObj.getIntActionNumber() + 1);
	//Return the values from the map
	const strStepAction = TestStepResults.StepAction;
	const strStepDescription = TestStepResults.StepDesc;
	const strStepExpectedResult = TestStepResults.StepExpected;
	const strStepActualResults = TestStepResults.StepActual;
	const strSampleData = TestStepResults.StepData;
	const boolStepPassed = TestStepResults.StepPassed;
	//We need to add capture on error or test if the automatic is working.
	//Update the step counts
	if (boolStepPassed == true){
		TestCaseObject.intStepsPassed = TestCaseObject.intStepsPassed + 1; //The number of test case steps executed that have passed.
	}
	else{
		TestCaseObject.intStepsFailed = TestCaseObject.intStepsFailed + 1; //The number of test case steps that failed.
	}
	//Tester.Assert(strStepDescription, boolStepPassed, strSampleData,false);
	Tester.Assert(
		strStepDescription,
		boolStepPassed,
		strStepActualResults,
		 {"expectedResult":strStepExpectedResult, "sampleData":strSampleData, "action": strStepAction}
	);
}

function TestExecReporting_GenStepReportingMap (/*string*/strStepAction, /*string*/strStepDesc, /*string*/strStepExpect, /*string*/ strStepActual, /*string*/strStepData, /*boolean*/boolStepPassed){
	var TestStepResults = {
		StepAction : strStepAction,
		StepDesc : strStepDesc,
		StepExpected : strStepExpect,
		StepActual : strStepActual,
		StepData : strStepData,
		StepPassed : boolStepPassed,
	}
	return TestStepResults;
}

/**
 * ----------------------------  ReportDoSnapshot  ---------------------------------------
* Takes a snapshot adding to Extent report and the Jira Xray JSON output
* @param objExtRptMod (Placeholder for TestStepType, e.g., ExtentTest)
* @param mapStepInput (Placeholder for a Map/Object containing step details)
* @return mapReportingStepResults (Placeholder for a Map/Object containing results)
* Created: 04/21/2021
* Author: PGKanaris
* Last Edited:
* Last Edited By:
* Edit Comments: (Include email, date and details)
*/
function TestExecReporting_ReportDoSnapshot(objExtRptMod, mapStepInput) {
	const mapReportingStepResults = {}; //Store the return values for sending to the test step
	const mapExtReporting = ExtLogging.ExtRptDoSnapShot(objExtRptMod, mapStepInput); // Calls ExtLogging internal helper
	const boolExtRptStepPassed = StrNums.JComm_StringToBoolean(mapExtReporting.boolStepRptPassed);
	const strExtRptStepResults = mapExtReporting.strStepRptResults;
	//TODO remove once we create JSON where there will be used
	const strExtRptStepScrnCapturePath = mapExtReporting.strSnapshotFilePath;
	if (TSExecParams.getBoolDoDebug() === true) {
		console.log("The ReportDoSnapshot is showing a snapshotFilePath from Extent Screen Capture of: " + strExtRptStepScrnCapturePath);
	}
	const strExtRptStepScrnCaptureExt = mapExtReporting.strImageExt;
	if (boolExtRptStepPassed === true) {
		const mapJSONReporting = JiraXrayInt.XrayAttachSnapShot(mapStepInput, mapExtReporting); // Calls JiraXrayInt internal helper
		//New method that passes the Maps to ExtLogging.ExtRptTestStepResults(TCModOpenSCMWeb, mapStepDetails)
		//Next pass the mapStepDetails and the new map which contains the snapshot path if taken and the ext. to JSON reporting
		//JSON reporting for steps will convert the image to base 64 and attach if not null.
		// Original: mapReportingStepResults = mapExtReporting (shallow copy in Groovy)
		Object.assign(mapReportingStepResults, mapExtReporting); // Shallow copy
	}
	else {
		//Extent step reporting failed
		mapReportingStepResults.boolStepRptPassed = boolExtRptStepPassed;
		mapReportingStepResults.strStepRptResults = strExtRptStepResults;
	}
	return mapReportingStepResults;
}