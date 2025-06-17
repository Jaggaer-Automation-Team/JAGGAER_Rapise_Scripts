

var TestCaseObject = {
	tcDriver : null, //May not be required
	tcOriginalBrowserDriver : null, //We need to check if we need this
	strOriginalBrowserWindowsHandle : '',
	boolPopupBrowser : false,
	strLocName : '',
	boolSwitchedToFrame : false,
	strCommonFrame : "phx-overlay-iframe",
	weParentObject : null,//Assign the page object or section object as needed.
	strParentObjName : '',//Assign the page object or section object name as needed.
	strParentObjXPath : '', //Assign the object XPath for use in failure states.
	strSnapshotImages: '',
	strBUSnapShotElementName: '',//Assign a value from the step if strParentObjName is undefined
	strTOElementXpath : '',//The last element returned XPath used in the stale object check to return the element from CWCore.returnWebElement
	intFoundRetryWait : 0, //The amount of time to retry where the object is present in order to catch stale elements and resolve them
	intFoundRetryDelay : 0, //The number of milliseconds to wait between found retries
	intStaleObjMaxWait : 30, //Time is seconds for max wait time
	intStaleObjIntervalWait : 100, //Time in Milliseconds to wait between checks
	strLoadingObjXPath : "//div[@class= 'AjaxRegionLoadingSpinner']", //Update for any application during test start that does not use this indicator.
	intStepsPassed : 0, //The number of test case steps executed that have passed.
	intStepsFailed : 0, //The number of test case steps that failed.
	
	
//Excel objects
	objWorkbook : null,
	strDefaultDrvLtr : "C:",
	strDriveLetter : "C:",
	strTCInputFilePath : '',
	objSheetTCInputData : null,
	objSheetTCActiveOutputSheet : null,
	objSheetTCInputResultsHdr : null,
	objSheetTCOutputResultsHdr : null,
	objSheetTCInputResultsData : null,
	objSheetTCOutputResultsData : null,
	objExcelFileInputStream : null,	
}

var TestStepResults = {
	strStepAction : '',
	strStepDescription : '',
	strStepExpectedResult : '',
	strStepActualResults : '',
	strSampleData : '',
	boolStepPassed : true,
}

var mapLocValues = {
	valCategory : '',
	valSubCategory : '',
	valFeature : '',
	valSubFeature : '',
	valChildFeature : '',
	valGrandChildFeat : '',
	valSnapShotName : ''
}


/**
	updateExecutionsFolders : function (){

		let strTempFolder = (TestCaseExecParams.strTestExtentRptFolder || '').trim(),

		TestCaseExecParams.strTestExtentRptFolder = strCurDriveLtr + strTempFolder,

		strTempFolder = (TestCaseExecParams.strTestExtentRptFolder || '').trim(),
		console.log("Extent Report Dir Output: " + strTempFolder),

		//Check for output folder and create if needed
		if (!fs.existsSync(strTempFolder)) {
			fs.mkdirSync(strTempFolder, { recursive: true }),
		}

		//Update Images folder
		strTempFolder = (TestCaseExecParams.strTestScreenShotsFolder || '').trim(),
		TestCaseExecParams.strTestScreenShotsFolder = strCurDriveLtr + strTempFolder,

		strTempFolder = (TestCaseExecParams.strTestScreenShotsFolder || '').trim(),
		console.log("Images Dir Output: " + strTempFolder),

		//Check for output folder and create if needed
		if (!fs.existsSync(strTempFolder)) {
			fs.mkdirSync(strTempFolder, { recursive: true }),
		}
	}
}
// Test Cases
export class TestCaseVariables {
	tcVarCreateFilePath (strFileName, strTestRunFolder, strFileExtension, boolIsInput, boolTCDir) {
		return ""
	}
}

export class SpiraVariables{
	//Set variables
	strSpiraRoot = 'https://jaggaer.spiraservice.net/Services/v7_0/RestService.svc/'
	strSpiraAuthUserName = 'jtestauto'
	strSpiraAuthAPIKey = '{A2C0BB20-2A34-46D3-89A1-521A4D12D406}'
 }
*/