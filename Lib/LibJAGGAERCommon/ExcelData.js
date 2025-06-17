/**
 * @PageObject ExcelData description
 */
// Alexey: This is a global object. If you need to update metadata, press Shift+Ctrl+F5 in Rapise.
SeSGlobalObject("ExcelData");

/*
I. Analysis of Classes and Types
Given the large number of imports for Apache POI, the core functionality will be challenging to map directly to Node.js built-ins. You would likely need to find a robust TypeScript/Node.js library that provides similar Excel manipulation capabilities (e.g., exceljs, xlsx, node-xlsx). For this template, I will acknowledge that Workbook, Sheet, Row, Cell, CellStyle, Font, etc., are concepts that map to methods/objects within such a library, rather than direct 1:1 type conversions.

A. Classes/Types Defined in ExcelData.groovy (the provided source)
ExcelData: The main class itself.
B. Types Defined Externally but within the Project/Application (resources.common package)
These are types that would likely be converted to separate TypeScript modules within your new project structure.

TestObjects (aliased as TCObj)
GlobalVariables (aliased as GVars)
TestCaseExecParams (aliased as TSExecParams)
StringsAndNumbers (aliased as StrNums)
TestEnvironment (aliased as TSEnv)
DateTime
C. Types Defined Externally (Not found in the provided source or resources.common package)
These are standard Java/Groovy built-in types or classes from external libraries. In TypeScript, these will either have direct equivalents (string, number, boolean, Map to Object), or will need to be replaced with Node.js standard library functions (fs, path), or a third-party Excel manipulation library.

Java/Groovy Built-in Types:

String: (TypeScript equivalent: string)
Number: (TypeScript equivalent: number)
int: (TypeScript equivalent: number)
boolean: (TypeScript equivalent: boolean)
Map: (TypeScript equivalent: object or Map<string, any>). Given its usage with [:] and put, a plain JavaScript object is likely the closest.
List: (TypeScript equivalent: Array<any>).
ArrayList: (TypeScript equivalent: Array<any>).
File: From java.io.File. (TypeScript/Node.js equivalent functionality: primarily path module for path manipulation, fs module for file/directory operations.)
FileInputStream: From java.io.FileInputStream. (TypeScript/Node.js equivalent: fs.readFileSync or streaming with fs.createReadStream).
FileOutputStream: From java.io.FileOutputStream. (TypeScript/Node.js equivalent: fs.writeFileSync or streaming with fs.createWriteStream).
Exception: From java.lang.Exception. (TypeScript equivalent: Error).
FileNotFoundException: From java.io.FileNotFoundException. (TypeScript equivalent: NodeJS.ErrnoException with code: 'ENOENT').
IOException: From java.io.IOException. (TypeScript equivalent: Error or specific NodeJS.ErrnoException).
MessageFormat: From java.text.MessageFormat. (TypeScript equivalent: string templating, or a library like format-message).
Color: From java.awt.Color. (TypeScript equivalent: typically represented as hex strings, RGB objects, or numbers).
Arrays: From java.util.Arrays. Used for Arrays.asList. (TypeScript equivalent: array literals, split method on strings).
Thread.sleep: (TypeScript/Node.js equivalent: setTimeout or a custom sleep function if synchronous delay is needed, but typically avoided in Node.js for blocking the event loop).
Apache POI Library-Specific Types:
These are the most critical external types as they define the domain logic. A TypeScript re-implementation would require finding a Node.js library that provides similar Excel manipulation. I will use generic WorkbookType, SheetType, etc. in the template, assuming a chosen library would provide these interfaces/classes.

org.apache.commons.lang3.exception.ExceptionUtils: ExceptionUtils.getStackTrace. (TypeScript/Node.js: error.stack).
org.apache.poi.xssf.usermodel.XSSFColor
org.apache.poi.xssf.usermodel.XSSFFont
org.apache.poi.xssf.usermodel.XSSFWorkbook: (Main Workbook type for .xlsx)
org.apache.poi.xssf.usermodel.IndexedColorMap
org.apache.poi.xssf.usermodel.DefaultIndexedColorMap
org.apache.poi.ss.usermodel.WorkbookFactory
org.apache.poi.ss.usermodel.Workbook
org.apache.poi.ss.usermodel.Sheet
org.apache.poi.ss.usermodel.Row
org.apache.poi.ss.usermodel.Cell
org.apache.poi.ss.usermodel.Header
org.apache.poi.ss.usermodel.CellStyle
org.apache.poi.ss.usermodel.HorizontalAlignment
org.apache.poi.ss.usermodel.VerticalAlignment
org.apache.poi.ss.usermodel.Font
org.apache.poi.ss.usermodel.CreationHelper
org.apache.poi.ss.usermodel.DataFormatter
org.apache.poi.ss.usermodel.BuiltinFormats
org.apache.poi.ss.usermodel.IndexedColors
org.apache.poi.ss.usermodel.DataFormat
org.apache.poi.ss.usermodel.BorderStyle
org.apache.poi.ss.usermodel.FillPatternType
org.apache.poi.ss.usermodel.FormulaEvaluator
org.apache.poi.ss.usermodel.DateUtil
org.apache.poi.openxml4j.util.ZipSecureFile
II. TypeScript Template
This template uses placeholder types (WorkbookType, SheetType, etc.) for the Excel library components because the direct mapping to java.util.* and org.apache.poi.* types isn't feasible with Node.js built-ins. You would replace these with types from your chosen Node.js Excel library (exceljs, xlsx, etc.).

Important Considerations for Mapping POI to TypeScript:

File I/O: FileInputStream and FileOutputStream will map to fs.readFileSync, fs.writeFileSync, or fs.createReadStream/fs.createWriteStream.
Workbook/Sheet/Row/Cell Objects: These are the core abstraction. You'll need an external library. For example, exceljs offers Workbook, Worksheet, Row, Cell classes.
Cell Value Handling: getStringCellValue(), getNumericCellValue(), getBooleanCellValue(), FormulaEvaluator, DataFormatter, DateUtil all need mapping to the chosen library's cell value retrieval methods.
Styling: CellStyle, Font, Color, Alignment, BorderStyle, FillPatternType, IndexedColors, XSSFColor are all complex styling objects. This is where a robust Excel library is crucial. Direct mapping of java.awt.Color (RGB) will be new SomeExcelLibrary.Color(R, G, B).
ZipSecureFile.setMinInflateRatio(0): This is a security setting for POI. It's unlikely to have a direct equivalent in a Node.js library, as security for zip files is often handled differently or by default. It might not be needed.
Thread.sleep(100): In Node.js, Thread.sleep would block the event loop, which is generally bad practice. If a delay is needed, setTimeout or await new Promise(resolve => setTimeout(resolve, ms)) is used, but the latter involves async/await which is explicitly disallowed. Therefore, we'll note it as a blocking operation in the comments.

*/

// Describe the library purpose
/**
 * The library contains Excel operation methods
 */

// Add Imports
// External Node.js built-in modules for file system operations
//import * as fs from 'fs'; // For file system operations (e.g., fs.existsSync, fs.readFileSync, fs.writeFileSync)
//import * as path from 'path'; // For path manipulation (e.g., path.join, path.resolve)

// --- Placeholder for a chosen Excel library (e.g., exceljs, xlsx) ---
// For example, if using 'exceljs':
// import { Workbook, Worksheet, Row, Cell, Style, Font, Alignment, Border, Fill, Color } from 'exceljs';
// For the purpose of this template, we'll use generic placeholder types.
// You will replace these with actual types from your chosen Excel library.
// For example:
// let WorkbookType: any;
// let SheetType: any;
// let RowType: any;
// let CellType: any;
// let CellStyleType: any;
// let FontType: any;
// let ColorType: any;
// let IndexedColorsType: any;
// let BuiltinFormatsType: any;
// let BorderStyleType: any;
// let FillPatternType: any;
// let FormulaEvaluatorType: any;
// let DataFormatterType: any;
//
// And then define actual types like:
// type WorkbookType = typeof Workbook.prototype; // Or the actual class type
// type SheetType = typeof Worksheet.prototype;
// etc.
// -------------------------------------------------------------------

// Add Jaggaer Libs (assuming these will also be converted to TypeScript modules)
//import * as TCObj from './TestObjects'; // Represents resources.common.TestObjects
//import * as GVars from './GlobalVariables'; // Represents resources.common.GlobalVariables
//import * as TSExecParams from './TestCaseExecParams'; // Represents resources.common.TestCaseExecParams
//import * as StrNums from './StringsAndNumbers'; // Represents resources.common.StringsAndNumbers
//import * as TSEnv from './TestEnvironment'; // Represents resources.common.TestEnvironment
//import * as DateTime from './DateTime'; // Represents resources.common.DateTime

/**
 * Re-implementation of ExcelData class as a collection of functions.
 * Public static methods are converted to exported functions.
 */

// Helper to simulate Java's `Thread.sleep` (BLOCKING - generally avoid in Node.js)
function sleep(ms) {
	const start = Date.now();
	while (Date.now() < start + ms) {
		// Busy-wait
	}
}

// NOTE on POI MAPPING: All POI-specific classes (XSSFWorkbook, Sheet, Row, Cell, CellStyle, etc.)
// must be replaced with equivalent objects/methods from a Node.js Excel library (e.g., 'exceljs').
// The template will use comments to denote where these mappings are needed.

/**
 * -------------------------------------  excelOpenExcelWB  -----------------------------------
 * Opens the excel file and updates the testobject workbook object
 * @param strFilePath 		The full path for the file.
 *
 * @return mapResults 		The results showing Passed and method details.
 *
 * @author pkanaris
 * Created: 06/23/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelOpenExcelWB(strFilePath) {
	//GlobalVars
	const gblNull = GVars.GblNull('Value');
	const gblUndefined = GVars.GblUndefined('Value');
	const gblLineFeed = GVars.GblLineFeed('Value');
	// Declare the variables
	const mapResults = {};
	let strMethodDetails; // Will be assigned later
	let boolPassed = true;
	let wb; // Replace with your Excel library's Workbook type. e.g., let wb: Workbook;
	// Original: wb = new XSSFWorkbook()
	// You would initialize your workbook here, e.g., wb = new exceljs.Workbook();

	//Check if the file exist
	// Corresponding to new File(strFilePath) and tmpFile.exists()
	if (fs.existsSync(strFilePath)) {
		//Update the Inflate Ratio since these are only internal created files. Use caution
		// No direct equivalent for ZipSecureFile.setMinInflateRatio(0) in Node.js fs or common Excel libs.
		// It's a POI-specific security/performance setting. You might not need it.
		// https://stackoverflow.com/questions/33918382/zip-bomb-exception-while-writing-a-large-formatted-excel-xlsx
		try {
			// Original: FileInputStream fisTemp = ExcelData.excelOpenExcelFileInputStream(strFilePath)
			// Original: wb = WorkbookFactory.create(fisTemp)
			// In Node.js, you'd likely read the file buffer and load it:
			const fileBuffer = fs.readFileSync(strFilePath);
			// Example with 'exceljs':
			// wb = new exceljs.Workbook();
			// await wb.xlsx.load(fileBuffer); // 'async' is NOT allowed for this exercise.
			// You might need a synchronous loading method, or rethink the flow without Promises.
			// Some libraries might offer synchronous file loading.
			// For now, let's assume `excelOpenWBFromInputStream` handles this sync.
			// Placeholder:
			const fisTemp = excelOpenExcelFileInputStream(strFilePath); // Simulates FileInputStream
			wb = excelOpenWBFromInputStream(fisTemp); // Simulates WorkbookFactory.create

			strMethodDetails = 'Opened the input file: ' + strFilePath + '.';
			//Update the objWorkbook
			TCObj.objWorkbook = wb; // TCObj.objWorkbook should be the type of your Excel library's Workbook
		} catch (e) {
			boolPassed = false;
			// Original: ExceptionUtils.getStackTrace(e)
			strMethodDetails = "FAILED!!! UNABLED TO OPEN INPUT FILE!!! See exception: " + (e instanceof Error ? e.stack : String(e));
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!!The file: ' + strFilePath + ' DOES NOT EXIST';
	}

	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * -------------------------------------  excelOpenTCInputSheet  -----------------------------------
 * Opens the sheet within the TCInput File
 * @param strSheetName 		The sheetname
 *
 * @return mapResults 		The results showing Passed and method details.
 *
 * @author pkanaris
 * Created: 06/23/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelOpenTCInputSheet(strSheetName) {
	const mapResults = {};
	let strMethodDetails;
	let strFileSheetName;
	let boolPassed = true;
	let sht = null; // Placeholder for SheetType. e.g., let sht: Worksheet | null = null;
	const wb = TCObj.objWorkbook; // Assumes TCObj.objWorkbook holds the Workbook object

	// Return the count of sheets in the test case input file
	// Original: int intNumSheets = TCObj.objWorkbook.numberOfSheets
	let intNumSheets = 0;
	if (wb && typeof wb.numberOfSheets === 'number') { // Check if wb exists and has numberOfSheets
		intNumSheets = wb.numberOfSheets;
	} else if (wb && typeof wb.worksheets === 'object' && Array.isArray(wb.worksheets)) {
		// exceljs uses wb.worksheets for an array of worksheets
		intNumSheets = wb.worksheets.length;
	}


	if (intNumSheets > 0) {
		//Build loop
		for (let loopIteration = 0; loopIteration < intNumSheets; loopIteration++) {
			// Original: strFileSheetName = TCObj.objWorkbook.getSheetName(loopIteration)
			if (wb && typeof wb.getSheetName === 'function') { // POI-style getSheetName
				strFileSheetName = wb.getSheetName(loopIteration);
			} else if (wb && wb.worksheets && wb.worksheets[loopIteration]) { // exceljs-style
				strFileSheetName = wb.worksheets[loopIteration].name;
			}

			if (strFileSheetName === strSheetName) {
				// Original: sht = TCObj.objWorkbook.getSheetAt(loopIteration)
				if (wb && typeof wb.getSheetAt === 'function') { // POI-style getSheetAt
					sht = wb.getSheetAt(loopIteration);
				} else if (wb && wb.worksheets && wb.worksheets[loopIteration]) { // exceljs-style
					sht = wb.worksheets[loopIteration];
				}
				break;
			}
		}
	}
	//Check if we found the sheet
	if (sht === null) {
		boolPassed = false;
		strMethodDetails = "FAILED!!! Did NOT FIND SHEET: " + strSheetName + " in the TESTCASE INPUT file!!!";
	}
	else {
		strMethodDetails = "Found the sheet: " + strSheetName + " in the testcase input file and update TC TestCaseData sheet object.";
		TCObj.objSheetTCInputData = sht; // TCObj.objSheetTCInputData should be the type of your Excel library's Sheet
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * Returns the number of worksheets in the assigned workbook
 * @param wb The workbook object (placeholder for WorkbookType)
 * @param boolOutputDetails Output the method details? True/False
 * @param strReporterOutput The reporter output format (CSV, HTML, XHTML, DataSource)
 * @return intNumSheets The number of Excel sheets in the workbook
 * Created: MM/dd/YYYY
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetSheetCount(wb, boolOutputDetails, strReporterOutput) {
	// Original: int intNumSheets = wb.numberOfSheets
	if (wb && typeof wb.numberOfSheets === 'number') {
		return wb.numberOfSheets;
	} else if (wb && wb.worksheets && Array.isArray(wb.worksheets)) {
		// exceljs style
		return wb.worksheets.length;
	}
	return 0; // Default if not found or invalid
	//TODO create step to output the results to a reporter method
}

/**
 * Returns the sheet names from the assigned workbook
 * @param wb The workbook object (placeholder for WorkbookType)
 * @param boolOutputDetails Output the method details? True/False
 * @param strReporterOutput The reporter output format (CSV, HTML, XHTML, DataSource)
 * @return strSheetNames A delimited string with the names of the sheets
 * Created: 06/24/2021
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetAllSheetNames(wb, boolOutputDetails, strReporterOutput) {
	const mapResults = {};
	let strMethodDetails;
	let strSheetNames = '';
	let boolPassed = true;
	const gblLineFeed = GVars.GblLineFeed('Value');
	const gblDelimiter = GVars.GblDelimiter('Value');
	//Return the count of sheets
	const intNumSheets = excelGetSheetCount(wb, false, ''); // Reuse the excelGetSheetCount function
	if (intNumSheets > 0) {
		strMethodDetails = 'The workbook contained: ' + intNumSheets + ' sheet(s) with the following name(s):' + gblLineFeed;
		//Build loop
		for (let loopIteration = 0; loopIteration < intNumSheets; loopIteration++) {
			if (loopIteration > 0) {
				strSheetNames = strSheetNames + gblDelimiter;
				strMethodDetails = strMethodDetails + gblLineFeed;
			}
			//Return each sheet and add to strSheetNames
			let currentSheetName;
			if (wb && typeof wb.getSheetName === 'function') { // POI-style getSheetName
				currentSheetName = wb.getSheetName(loopIteration);
			} else if (wb && wb.worksheets && wb.worksheets[loopIteration]) { // exceljs-style
				currentSheetName = wb.worksheets[loopIteration].name;
			}
			strSheetNames = strSheetNames + currentSheetName;
			strMethodDetails = strMethodDetails + currentSheetName;
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! The workbook DOES NOT CONTAIN any SHEETS!!!';
	}

	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

function ExcelData_excelGetActiveWBSheetName(wb) {
	// This would depend heavily on the chosen Excel library.
	// POI Workbook.getActiveSheetIndex() or similar, then getSheetName(index).
	// In exceljs, active sheet is not explicitly exposed as a property, often it's managed externally
	// or you might loop through worksheets to find it if a specific property defines "active".
	return 'NOT IMPLEMENTED: Depends on Excel Library API for active sheet.';
}

function ExcelData_excelSetActiveWBSheetByName(wb, strSheetName) {
	// This would depend heavily on the chosen Excel library.
	// POI Workbook.setActiveSheet(int index) or similar.
	// In exceljs, you'd typically get the worksheet by name and then primarily interact with that sheet object.
	return 'NOT IMPLEMENTED: Depends on Excel Library API for setting active sheet.';
}

function ExcelData_excelGetSheetByID(wb, intSheetID) {
	// Original: wb.getSheetAt(intSheetID)
	if (wb && typeof wb.getSheetAt === 'function') {
		return wb.getSheetAt(intSheetID);
	} else if (wb && wb.worksheets && wb.worksheets[intSheetID]) {
		return wb.worksheets[intSheetID];
	}
	return null; // Return null if not found
}

function ExcelData_excelGetSheetIDByName(wb, strSheetName) {
	// Original: wb.getSheetIndex(strSheetName)
	// This calls `excelReturnSheetID` which does a loop, so we'll leave that in.
	return excelReturnSheetID(wb, strSheetName);
}

function ExcelData_excelReturnSheetID(wb, strSheetName) {
	let intSheetID = -1;
	//Return the count of sheets in the test case input file
	const intNumSheets = excelGetSheetCount(wb, false, '');
	if (intNumSheets > 0) {
		//Build loop
		for (let loopIteration = 0; loopIteration < intNumSheets; loopIteration++) {
			let strFileSheetName;
			if (wb && typeof wb.getSheetName === 'function') {
				strFileSheetName = wb.getSheetName(loopIteration);
			} else if (wb && wb.worksheets && wb.worksheets[loopIteration]) {
				strFileSheetName = wb.worksheets[loopIteration].name;
			}
			if (strFileSheetName === strSheetName) {
				intSheetID = loopIteration;
				break;
			}
		}
	}
	return intSheetID;
}

/**
 * -------------------------------------  excelGetSheetByName  -----------------------------------
 * Returns the sheet as part of the map from an excel workbook
 * @param strSheetName 		The sheetname
 *
 * @return mapResults 		The results showing Passed, method details, and the objSheet.
 *
 * @author pkanaris
 * Created: 06/23/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetSheetByName(wb, strSheetName) {
	const mapResults = {};
	let strMethodDetails; // Will be assigned later
	let strFileSheetName;
	let boolPassed = true;
	let sht = null; // Placeholder for SheetType
	let intSheetID;

	const intNumSheets = excelGetSheetCount(wb, false, '');
	if (intNumSheets > 0) {
		for (let loopIteration = 0; loopIteration < intNumSheets; loopIteration++) {
			if (wb && typeof wb.getSheetName === 'function') {
				strFileSheetName = wb.getSheetName(loopIteration);
			} else if (wb && wb.worksheets && wb.worksheets[loopIteration]) {
				strFileSheetName = wb.worksheets[loopIteration].name;
			}

			if (strFileSheetName === strSheetName) {
				if (wb && typeof wb.getSheetAt === 'function') {
					sht = wb.getSheetAt(loopIteration);
				} else if (wb && wb.worksheets && wb.worksheets[loopIteration]) {
					sht = wb.worksheets[loopIteration];
				}
				intSheetID = loopIteration;
				break;
			}
		}
	}
	//Check if we found the sheet
	if (sht === null) {
		boolPassed = false;
		strMethodDetails = "FAILED!!! Did NOT FIND SHEET: " + strSheetName + " in the assigned Workbook!!!";
	}
	else {
		strMethodDetails = "Found the sheet: " + strSheetName + " in the assigned Workbook.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.objWbSheet = sht; // Placeholder for SheetType
	mapResults.intSheetID = intSheetID;
	return mapResults;
}

/**
 * -------------------------------------  excelRenameSheet  -----------------------------------
 *
 * @param wb The workbook object (placeholder for WorkbookType)
 * @param strOrgSheetName The original name of the sheet
 * @param strNewSheetName The new name to assign to the sheet
 *
 * @return mapResults 		The results showing Passed, method details.
 *
 * @author pkanaris
 * Created: 06/24/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelRenameSheet(wb, strOrgSheetName, strNewSheetName) {
	//TODO update to new standards and test
	const mapResults = {};
	let boolPassed = true;
	let intSheetID;
	let strMethodDetails = '';

	// Original: wb.setSheetName(intSheetID, strNewSheetName)
	// In exceljs, you would get the sheet and change its name directly:
	// const sheet = wb.getWorksheet(strOrgSheetName); // or by id
	// if (sheet) { sheet.name = strNewSheetName; }

	intSheetID = excelReturnSheetID(wb, strOrgSheetName); // Get original sheet ID

	if (intSheetID !== -1) {
		if (wb && typeof wb.setSheetName === 'function') { // Apache POI style
			wb.setSheetName(intSheetID, strNewSheetName);
			strMethodDetails = `Renamed sheet '${strOrgSheetName}' to '${strNewSheetName}'.`;
		} else if (wb && wb.worksheets && wb.worksheets[intSheetID]) { // exceljs style (if you get it by index)
			wb.worksheets[intSheetID].name = strNewSheetName;
			strMethodDetails = `Renamed sheet '${strOrgSheetName}' to '${strNewSheetName}'.`;
		} else {
			boolPassed = false;
			strMethodDetails = `FAILED!!! Workbook or worksheet API doesn't support renaming for sheet '${strOrgSheetName}'.`;
		}
	} else {
		boolPassed = false;
		strMethodDetails = `FAILED!!! Original sheet '${strOrgSheetName}' not found to rename.`;
	}

	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * -------------------------------------  excelGetRowCount  -----------------------------------
 * Return the number of rows that are shown as containing values.
 * NOTE: cells with spaces or used cells that have no data can affect the row count. Delete all empty rows.
 *
 * @param assgSheet The assigned sheet (placeholder for SheetType)
 *
 * @return mapResults 		The results showing Passed, method details, and the intRowCount.
 *
 * @author pkanaris
 * Created: 06/24/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetRowCount(assgSheet) {
	const mapResults = {};
	let boolPassed = true;
	let strMethodDetails = '';
	// Original: String strSheetName = assgSheet.sheetName
	let strSheetName = assgSheet ? assgSheet.sheetName : 'Unknown Sheet'; // Placeholder for SheetType.sheetName
	// Original: int intRowCnt = assgSheet.getLastRowNum()
	let intRowCnt = assgSheet ? assgSheet.getLastRowNum() : 0; // Placeholder for SheetType.getLastRowNum()
	// In exceljs, intRowCnt might be assgSheet.rowCount or assgSheet.actualRowCount

	if (intRowCnt > 0) {
		strMethodDetails = 'The Sheet: "' + strSheetName + '" contains: ' + intRowCnt + ' row(s).';
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! The Sheet: "' + strSheetName + '" contains: ' + intRowCnt + ' row(s).';
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.RowCount = intRowCnt;
	return mapResults;
}

/**
 * -------------------------------------  excelGetRowAndColCount  -----------------------------------
 * Return the number of rows that are shown as containing values and the number of columns in the header row.
 * NOTE: cells with spaces or used cells that have no data can affect the row count. Delete all empty rows.
 *
 * @param assgSheet The assigned sheet (placeholder for SheetType)
 *
 * @return mapResults 		The results showing Passed, method details, intRowCount and intColCount.
 *
 * @author pkanaris
 * Created: 06/24/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetRowAndColCount(assgSheet) {
	const mapResults = {};
	let boolPassed = true;
	let strMethodDetails = '';
	// Original: String strSheetName = assgSheet.sheetName
	let strSheetName = assgSheet ? assgSheet.sheetName : 'Unknown Sheet'; // Placeholder for SheetType.sheetName
	// Original: int intColCnt = assgSheet.getRow(0).getLastCellNum()
	// Original: int intRowCnt = assgSheet.getLastRowNum()

	let intColCnt = 0;
	let intRowCnt = 0;

	if (assgSheet) {
		intRowCnt = excelGetRowCount(assgSheet).RowCount; // Reuse excelGetRowCount
		const headerRow = assgSheet.getRow(0); // Placeholder for SheetType.getRow(0)
		if (headerRow) {
			// Placeholder for RowType.getLastCellNum()
			intColCnt = headerRow.getLastCellNum(); // In exceljs, this might be row.cellCount
		}
	}


	if (intRowCnt > 0) {
		strMethodDetails = "The Sheet: '" + strSheetName + "' contains: " + intRowCnt + " row(s) and " +
			intColCnt + " column(s).";
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! The Sheet: "' + strSheetName + '" contains: ' + intRowCnt + ' row(s).';
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.RowCount = intRowCnt;
	mapResults.ColCount = intColCnt;
	return mapResults;
}

/**
 * -------------------------------------  excelReturnLoopStartEndRow  -----------------------------------
 * Returns the start and end row integers for the loop
 * @param intStartRow 		The start row number specified by the user
 * @param intEndRow			The end row number specified by the user
 * @param intInputRowCount	The number of rows in the input sheet
 *
 * @return mapResults 		The results showing Passed and method details.
 *
 * @author pkanaris
 * Created: 04/13/2022
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelReturnLoopStartEndRow(intStartRow, intEndRow, intInputRowCount) {
	const mapResults = {};
	let intLoopStart;
	let intLoopEnd;
	//Set the loop start
	if (intStartRow > intInputRowCount) {
		intLoopStart = intInputRowCount;
	}
	else {
		intLoopStart = intStartRow;
	}
	//Set the loop end
	if (intEndRow > intInputRowCount) {
		intLoopEnd = intInputRowCount;
	}
	else {
		intLoopEnd = intEndRow;
	}
	//Update the map
	mapResults.intLoopStart = intLoopStart;
	mapResults.intLoopEnd = intLoopEnd;
	return mapResults;
}

/**
 * -------------------------------------  excelGetHeaderColCount  -----------------------------------
 * Return the number of number of columns in the header
 * @param assgSheet The assigned sheet (placeholder for SheetType)
 *
 * @return mapResults 		The results showing Passed, method details, and HdrColCount.
 *
 * @author pkanaris
 * Created: 06/24/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetHeaderColCount(assgSheet) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	let boolPassed = true;
	let strMethodDetails = '';
	// Original: String strSheetName = assgSheet.sheetName
	let strSheetName = assgSheet ? assgSheet.sheetName : 'Unknown Sheet'; // Placeholder for SheetType.sheetName

	let intHeaderColCnt = 0;
	if (assgSheet) {
		const headerRow = assgSheet.getRow(0); // Placeholder for SheetType.getRow(0)
		if (headerRow) {
			// Placeholder for RowType.getLastCellNum()
			intHeaderColCnt = headerRow.getLastCellNum(); // In exceljs, this might be row.cellCount
		}
	}

	if (intHeaderColCnt > 0) {
		strMethodDetails = 'The Sheet: "' + strSheetName + '" header contains: ' + intHeaderColCnt + ' col(s).';
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! The Sheet: "' + strSheetName + '" header contains: ' + intHeaderColCnt + ' col(s).';
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.HdrColCount = intHeaderColCnt;
	return mapResults;
}

/**
 * -------------------------------------  excelGetHdrColNames  -----------------------------------
 * Return all the header column names in a delimited string
 * @param assgSheet The assigned sheet (placeholder for SheetType)
 *
 * @return mapResults 		The results showing Passed, method details, ColNames, and the ColValues.
 *
 * @author pkanaris
 * Created: 06/24/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetHdrColNames(assgSheet) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value");
	let boolPassed = true;
	let strMethodDetails = '';
	let strSheetName;
	let strColNames = '';
	let intHeaderColCnt = 0;
	let strColName;
	let strColValues = '';

	if (assgSheet) {
		strSheetName = assgSheet.sheetName; // Placeholder for SheetType.sheetName
		const headerRow = assgSheet.getRow(0); // Placeholder for SheetType.getRow(0)
		if (headerRow) {
			intHeaderColCnt = headerRow.getLastCellNum(); // Placeholder for RowType.getLastCellNum()
		}
	} else {
		strSheetName = 'Unknown Sheet';
	}


	if (intHeaderColCnt > 0) {
		strMethodDetails = 'The Sheet: "' + strSheetName + '" header contains: ' + intHeaderColCnt + ' col(s) see names below.' + gblLineFeed;
		//Build loop to return the values
		for (let loopIteration = 0; loopIteration < intHeaderColCnt; loopIteration++) {
			if (loopIteration > 0) {
				strColNames = strColNames + gblDelimiter;
				strMethodDetails = strMethodDetails + gblLineFeed;
			}
			//Return each column Name
			//all excel values are stored as text to maintain full values that have leading zeros
			// Original: StrNums.JComm_HandleNoData(assgSheet.getRow(0).getCell(loopIteration).getStringCellValue())
			let cell = assgSheet.getRow(0).getCell(loopIteration); // Placeholder for CellType
			strColName = StrNums.JComm_HandleNoData(excelGetCellValue(cell)); // Use internal helper to get cell value
			if (loopIteration === 0) {
				strColNames = 'ColumID: ' + loopIteration + '; Name: ' + strColName;
				strColValues = strColName;
				strMethodDetails = 'ColumID: ' + loopIteration + '; Name: ' + strColName;
			}
			else {
				strColNames = strColNames + 'ColumID: ' + loopIteration + '; Name: ' + strColName;
				strColValues = strColValues + gblDelimiter + strColName;
				strMethodDetails = strMethodDetails + 'ColumID: ' + loopIteration + '; Name: ' + strColName;
			}
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! The Sheet: "' + strSheetName + '" header contains: ' + intHeaderColCnt + ' col(s).';
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.ColNames = strColNames;
	mapResults.ColValues = strColValues;
	return mapResults;
}

/**
 * Return all the header column names in a delimited string
 * @param assgSheet The assigned sheet (placeholder for SheetType)
 * @param strReporterOutput The reporter output format (CSV, HTML, XHTML, DataSource)
 * @return strColNames The delimited string containing the column names (text)
 * Created: 04/21/2022
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetHdrColNamesValuesOnly(assgSheet, strReporterOutput) {
	const gblLineFeed = GVars.GblLineFeed('Value');
	// const gblDelimiter = GVars.GblDelimiter('Value'); // Not used in this version
	let boolPassed = true;
	let strMethodDetails = '';
	let strSheetName;
	let strColNames = '';
	const mapResults = {};
	let intHeaderColCnt = 0;
	let strColName;

	if (assgSheet) {
		strSheetName = assgSheet.sheetName; // Placeholder for SheetType.sheetName
		const headerRow = assgSheet.getRow(0); // Placeholder for SheetType.getRow(0)
		if (headerRow) {
			intHeaderColCnt = headerRow.getLastCellNum(); // Placeholder for RowType.getLastCellNum()
		}
	} else {
		strSheetName = 'Unknown Sheet';
	}

	if (intHeaderColCnt > 0) {
		strMethodDetails = 'The Sheet: "' + strSheetName + '" header contains: ' + intHeaderColCnt + ' col(s) see names below.' + gblLineFeed;
		//Build loop to return the values
		for (let loopIteration = 0; loopIteration < intHeaderColCnt; loopIteration++) {
			//Return each column Name
			// Original: StrNums.JComm_HandleNoData(assgSheet.getRow(0).getCell(loopIteration).getStringCellValue())
			let cell = assgSheet.getRow(0).getCell(loopIteration);
			strColName = StrNums.JComm_HandleNoData(excelGetCellValue(cell)); // all excel values are stored as text to maintain full values that have leading zeros
			if (loopIteration === 0) {
				strColNames = strColName;
			} else {
				strColNames = strColNames + ';' + strColName;
			}
			strMethodDetails = strMethodDetails + 'ColumID: ' + loopIteration + '; Name: ' + strColName;
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! The Sheet: "' + strSheetName + '" header contains: ' + intHeaderColCnt + ' col(s).';
	}
	//Output the method pass/fail and message
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	//TODO create step to output the results to a reporter method
	return mapResults;
}

/**
 * -------------------------------------  excelGetColNameByIndex  -----------------------------------
 * Return the name of the column based on the index assigned.
 * @param assgSheet The assigned sheet (placeholder for SheetType)
 * @param intColIndex The column index number
 *
 * @return mapResults 		The results showing Passed, method details, and ColName.
 *
 * @author pkanaris
 * Created: 06/24/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetColNameByIndex(assgSheet, intColIndex) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value");
	let boolPassed = true;
	let strMethodDetails = '';
	let strSheetName;
	let intHeaderColCnt = 0;
	let strColName = '';

	if (assgSheet) {
		strSheetName = assgSheet.sheetName; // Placeholder for SheetType.sheetName
		const headerRow = assgSheet.getRow(0); // Placeholder for SheetType.getRow(0)
		if (headerRow) {
			intHeaderColCnt = headerRow.getLastCellNum(); // Placeholder for RowType.getLastCellNum()
		}
	} else {
		strSheetName = 'Unknown Sheet';
	}

	if (intHeaderColCnt >= intColIndex) {
		//Subtract 1 from the intColIndex since index is zero based.
		// Original: StrNums.JComm_HandleNoData(assgSheet.getRow(0).getCell(intColIndex -1).getStringCellValue())
		let cell = assgSheet.getRow(0).getCell(intColIndex - 1);
		strColName = StrNums.JComm_HandleNoData(excelGetCellValue(cell)); // all excel values are stored as text to maintain full values that have leading zeros
		strMethodDetails = 'The sheet: ' + strSheetName + ' column number: ' + intColIndex + ' name is: ' + strColName + '.';
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! The Sheet: "' + strSheetName + '" header contains: ' + intHeaderColCnt + ' col(s) which is less then the assigned column index of: ' + intColIndex + '!!!';
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.ColName = strColName;
	return mapResults;
}

/**
 * -------------------------------------  excelGetColIndexByColName  -----------------------------------
 * Return the column index for the assigned column name
 * @param assgSht The sheet object assigned (placeholder for SheetType)
 * @param strColName The name of the assigned column
 *
 * @return mapResults 		The results showing Passed, method details, and ColIndex
 *
 * @author pkanaris
 * Created: 06/24/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetColIndexByColName(assgSht, strColName) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value");
	let boolPassed = true;
	let strMethodDetails = '';
	let strSheetName = assgSht ? assgSht.sheetName : 'Unknown Sheet'; // Placeholder for SheetType.sheetName
	// Original: Row rowHeader = assgSht.getRow(0)
	let rowHeader = assgSht ? assgSht.getRow(0) : null; // Placeholder for RowType
	let intRowColCnt = 0;
	if (rowHeader) {
		intRowColCnt = rowHeader.getLastCellNum(); // Placeholder for RowType.getLastCellNum()
	}
	let intColId = -1;
	let strCurColValue = '';
	//Build loop to return the values
	for (let loopIteration = 0; loopIteration < intRowColCnt; loopIteration++) {
		//return the column value and check for match
		// Original: StrNums.JComm_HandleNoData(rowHeader.getCell(loopIteration).getStringCellValue())
		let cell = rowHeader ? rowHeader.getCell(loopIteration) : null;
		strCurColValue = StrNums.JComm_HandleNoData(excelGetCellValue(cell));
		if (strCurColValue === strColName) {
			strMethodDetails = 'The Header row in sheet: "' + strSheetName + '" contains the column name: "' + strColName + '" in column ' + loopIteration + '.';
			intColId = loopIteration;
			break;
		}
	}
	if (intColId === -1) {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! The Header row in sheet: "' + strSheetName + '" DOES NOT CONTAIN the column name: "' + strColName + '" in ' + intRowColCnt + ' COLUMNS!!!';
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.ColIndex = intColId;
	return mapResults;
}

/**
 * -------------------------------------  excelGetDataRow  -----------------------------------
 * Return the row from the assigned sheet
 * @param assgSheet The assigned sheet (placeholder for SheetType)
 * @param intRowNum The row number to return. NOTE starts with 1 since zero is the header.
 *
 * @return mapResults 		The results showing Passed, method details, and objRow
 *
 * @author pkanaris
 * Created: 06/24/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetRow(assgSheet, intRowNum) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value");
	let boolPassed = true;
	let strMethodDetails = '';
	let assgRow = null; // Placeholder for RowType
	let strSheetName = assgSheet ? assgSheet.sheetName : 'Unknown Sheet'; // Placeholder for SheetType.sheetName
	let intRowCnt = assgSheet ? assgSheet.getLastRowNum() : 0; // Placeholder for SheetType.getLastRowNum()

	if (intRowCnt >= intRowNum) {
		assgRow = assgSheet.getRow(intRowNum); // Placeholder for SheetType.getRow(intRowNum)
		if (assgRow === null) {
			boolPassed = false;
			strMethodDetails = 'FAILED!!! The Sheet: "' + strSheetName + "' row number: '" + intRowNum + "' DID NOT RETURN A ROW OBJECT!!!";
		}
		strMethodDetails = 'The Sheet: "' + strSheetName + '" contains: ' + intRowCnt + ' row(s) and returned data row: ' + intRowNum + '.';
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! The Sheet: "' + strSheetName + '" contains: ' + intRowCnt + ' row(s) which is NOT EQUAL OR GREATER than THE ASSIGNED ROW: ' + intRowNum + '.';
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.ObjRow = assgRow; // Placeholder for RowType
	return mapResults;
}

/**
 * -------------------------------------  excelGetRowColCount  -----------------------------------
 * Return the number of columns for the assigned row object
 * @param assgRow The row object assigned (placeholder for RowType)
 *
 * @return mapResults 		The results showing Passed, method details, and colRowCount
 *
 * @author pkanaris
 * Created: 06/24/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetRowColCount(assgRow) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value");
	let boolPassed = true;
	let strMethodDetails = '';
	// Original: String strSheetName = assgRow.getSheet().sheetName
	let strSheetName = assgRow && assgRow.getSheet ? assgRow.getSheet().sheetName : 'Unknown Sheet'; // Placeholder for RowType.getSheet().sheetName
	// Original: int intRowColCnt = assgRow.getLastCellNum()
	let intRowColCnt = assgRow ? assgRow.getLastCellNum() : 0; // Placeholder for RowType.getLastCellNum()
	// Original: int intRowId = assgRow.getRowNum()
	let intRowId = assgRow ? assgRow.getRowNum() : -1; // Placeholder for RowType.getRowNum()

	if (intRowColCnt > 0) {
		strMethodDetails = 'Row ' + intRowId + ' in sheet: "' + strSheetName + '" contains ' + intRowColCnt + ' col(s).';
	}
	else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! Row ' + intRowId + ' in sheet: "' + strSheetName + '" DOES NOT CONTAIN COLUMN(s).';
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.ColRowCount = intRowColCnt;
	return mapResults;
}

/**
 * -------------------------------------  excelGetCellFromRowByColName  -----------------------------------
 * Return the cell from the assgRow based on column name
 * @param assgRow The row object assigned (placeholder for RowType)
 * @param strColName The name of the column for the cell
 *
 * @return mapResults 		The results showing Passed, method details, and objCell
 *
 * @author pkanaris
 * Created: 06/24/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetCellFromRowByColName(assgRow, strColName) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value");
	let boolPassed = true;
	let strMethodDetails = '';
	// Original: String strSheetName = assgRow.getSheet().sheetName
	let strSheetName = assgRow && assgRow.getSheet ? assgRow.getSheet().sheetName : 'Unknown Sheet'; // Placeholder for RowType.getSheet().sheetName
	// Original: Row rwHeader = assgRow.getSheet().getRow(0)
	let rwHeader = assgRow && assgRow.getSheet ? assgRow.getSheet().getRow(0) : null; // Placeholder for RowType
	let intHeaderColCnt = rwHeader ? rwHeader.getLastCellNum() : 0; // Placeholder for RowType.getLastCellNum()
	let intRowColCnt = assgRow ? assgRow.getLastCellNum() : 0; // Placeholder for RowType.getLastCellNum()
	let intRowId = assgRow ? assgRow.getRowNum() : -1; // Placeholder for RowType.getRowNum()
	let intColumnIndex = -1;
	let assgCell = null; // Placeholder for CellType
	//Check if the header and row count match. If not we cannot make a selection
	if (intHeaderColCnt === intRowColCnt && intHeaderColCnt > 0) {
		const mapGetColIndex = excelGetColIndexByColName(assgRow.getSheet(), strColName); // Reuse existing function
		//Check the results and mark pass or fail
		const boolMethPassed = StrNums.JComm_StringToBoolean(mapGetColIndex.boolPassed);
		const strMethResults = mapGetColIndex.strMethodDetails; // Not directly used in this logic block
		intColumnIndex = mapGetColIndex.ColIndex;
		if (boolMethPassed === true && intColumnIndex !== -1) {
			//return the cell
			assgCell = assgRow.getCell(intColumnIndex); // Placeholder for RowType.getCell(intColIndex)
			strMethodDetails = 'Returned the cell for column: "' + strColName + '" columnIndex: ' + intColumnIndex + ' from sheet: "' + strSheetName + '" in row number: ' + intRowId + '.';
		} else {
			boolPassed = false;
			strMethodDetails = 'FAILED!!! The header for sheet "' + strSheetName + '" did not contain the column name of: "' + strColName + '" !!!';
		}
	} else {
		boolPassed = false;
		strMethodDetails = `FAILED!!! Header column count (${intHeaderColCnt}) does not match row column count (${intRowColCnt}) or header count is zero.`;
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.ObjCell = assgCell; // Placeholder for CellType
	return mapResults;
}

/**
 * -------------------------------------  excelGetCellValueFromRowByColName  -----------------------------------
 * Return the cell value from the assgRow based on column name
 * @param assgRow The row object assigned (placeholder for RowType)
 * @param strColName The name of the column for the cell
 * @return mapResults 		The results showing Passed, method details, and CellValue
 * @author pkanaris
 * Created: 06/24/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetCellValueFromRowByColName(assgRow, strColName) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value");
	let boolPassed = true;
	let strMethodDetails = '';
	let strSheetName = assgRow && assgRow.getSheet ? assgRow.getSheet().sheetName : 'Unknown Sheet';
	let rwHeader = assgRow && assgRow.getSheet ? assgRow.getSheet().getRow(0) : null;
	let intHeaderColCnt = rwHeader ? rwHeader.getLastCellNum() : 0;
	let intRowColCnt = assgRow ? assgRow.getLastCellNum() : 0;
	let intRowId = assgRow ? assgRow.getRowNum() : -1;
	let intColumnIndex = -1;
	let assgCell = null;
	let strCellValue;
	//Check if the header and row count match.
	if (intHeaderColCnt !== intRowColCnt) {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The Header column count of: " + intHeaderColCnt + " DOES NOT MATCH THE ASSIGNED ROW COLUMN COUNT of: " + intRowColCnt + "!!!";
	}
	else {
		if (intHeaderColCnt > 0) {
			const mapGetColIndex = excelGetColIndexByColName(assgRow.getSheet(), strColName);
			const boolMethPassed = StrNums.JComm_StringToBoolean(mapGetColIndex.boolPassed);
			const strMethResults = mapGetColIndex.strMethodDetails; // Not directly used in this block
			intColumnIndex = mapGetColIndex.ColIndex;
			if (boolMethPassed === true && intColumnIndex !== -1) {
				assgCell = assgRow.getCell(intColumnIndex);
				if (assgCell !== null) {
					strCellValue = excelGetCellValue(assgCell); // Use internal helper
					strMethodDetails = 'Returned the cell for column: "' + strColName + '" columnIndex: ' + intColumnIndex + ' from sheet: "' + strSheetName + '" in row number: '
						+ intRowId + ' with a value of: ' + strCellValue + '.';
				} else {
					boolPassed = false;
					strMethodDetails = 'FAILED!!! DID NOT Return the cell for column: "' + strColName + '" columnIndex: ' + intColumnIndex + ' from sheet: "' + strSheetName + '" in row number: ' + intRowId + '.';
				}
			} else {
				boolPassed = false;
				strMethodDetails = 'FAILED!!! The header for sheet "' + strSheetName + '" did not contain the column name of: "' + strColName + '" !!!';
			}
		}
		else {
			boolPassed = false;
			strMethodDetails = "FAILED!!! The HEADER COLUMN COUNT IS NOT GREATER THAN ZERO!!!";
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.CellValue = strCellValue;
	mapResults.intColIndex = intColumnIndex;
	return mapResults;
}

/**
 * -------------------------------------  excelGetCellValueFromRowByColNumber  -----------------------------------
 * Return the cell value from the assgRow based on column name
 * @param assgRow The row object assigned (placeholder for RowType)
 * @param intColNum The column number for the cell
 * @return mapResults 		The results showing Passed, method details, and CellValue
 * @author pkanaris
 * Created: 06/24/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetCellValueFromRowByColNumber(assgRow, intColNum) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value");
	let boolPassed = true;
	let strMethodDetails = '';
	let strSheetName = assgRow && assgRow.getSheet ? assgRow.getSheet().sheetName : 'Unknown Sheet';
	let rwHeader = assgRow && assgRow.getSheet ? assgRow.getSheet().getRow(0) : null;
	let intHeaderColCnt = rwHeader ? rwHeader.getLastCellNum() : 0;
	let intRowColCnt = assgRow ? assgRow.getLastCellNum() : 0;
	let intRowId = assgRow ? assgRow.getRowNum() : -1;
	let intColumnIndex = intColNum; // Directly using intColNum as the index
	let assgCell = null;
	let strCellValue;
	//Check if the header and row count match.
	if (intHeaderColCnt === intRowColCnt && intHeaderColCnt > 0) {
		assgCell = assgRow.getCell(intColNum); // Placeholder for RowType.getCell(intColNum)
		if (assgCell !== null) {
			strCellValue = excelGetCellValue(assgCell); // Use internal helper
			strMethodDetails = 'Returned the cell for columnIndex: ' + intColumnIndex + ' from sheet: "' + strSheetName + '" in row number: '
				+ intRowId + ' with a value of: ' + strCellValue + '.';
		} else {
			boolPassed = false;
			strMethodDetails = 'FAILED!!! DID NOT Return the cell for columnIndex: ' + intColumnIndex + ' from sheet: "' + strSheetName + '" in row number: ' + intRowId + '.';
		}
	} else {
		boolPassed = false;
		strMethodDetails = "FAILED!!! Header column count (" + intHeaderColCnt + ") does not match row column count (" + intRowColCnt + ") or header count is zero.";
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.CellValue = strCellValue;
	return mapResults;
}

/**
 * -------------------------------------  excelGetCellValueByRowNumColName  -----------------------------------
 * Return the cell value from the assgRow based on column name
 * @param assgSheet The assigned sheet (placeholder for SheetType)
 * @param intRowNum The row number
 * @param strColName The name of the column for the cell
 * @return mapResults 		The results showing Passed, method details, and CellValue
 * @author pkanaris
 * Created: 06/24/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetCellValueByRowNumColName(assgSheet, intRowNum, strColName) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value");
	let boolPassed = true;
	let strMethodDetails = '';
	let rowData = null; // Placeholder for RowType
	let strCellValue;
	let intColIndex;
	//Return the row
	const mapGetRow = excelGetRow(assgSheet, intRowNum); // Reuse existing function
	//Return the values from the map
	boolPassed = mapGetRow.boolPassed;
	strMethodDetails = mapGetRow.strMethodDetails;
	rowData = mapGetRow.ObjRow;
	if (boolPassed === false) {
		strCellValue = 'FAILED!!! NO ROW DATA found in sheet: ' + (assgSheet ? assgSheet.sheetName : 'Unknown Sheet') + ' for row: ' + intRowNum + '!!!';
	} else {
		//Return the cell
		//Get the colId
		let mapGetCellValue = {};
		if (rowData === null) {
			boolPassed = false; // Fix: original had `== false` here
			strMethodDetails = 'FAILED!!! NO ROW OBJECT PRESENT for: ' + (assgSheet ? assgSheet.sheetName : 'Unknown Sheet') + ' for row: ' + intRowNum + '!!!';
		}
		else {
			mapGetCellValue = excelGetCellValueFromRowByColName(rowData, strColName); // Reuse existing function
			//Check the results
			boolPassed = StrNums.JComm_StringToBoolean(mapGetCellValue.boolPassed);
			strMethodDetails = mapGetCellValue.strMethodDetails;
			strCellValue = mapGetCellValue.CellValue;
			intColIndex = mapGetCellValue.intColIndex;
		}
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.CellValue = strCellValue;
	mapResults.intColIndex = intColIndex;
	return mapResults;
}

/**
 * -------------------------------------  excelGetCellValueByRowNumColNum  -----------------------------------
 * Return the cell value from the assgRow based on column number
 * @param assgSheet The assigned sheet (placeholder for SheetType)
 * @param intRowNum The row number
 * @param intColNum The column number for the cell
 * @return mapResults 		The results showing Passed, method details, and CellValue
 * @author pkanaris
 * Created: 06/24/2021
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetCellValueByRowNumColNum(assgSheet, intRowNum, intColNum) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value");
	let boolPassed = true;
	let strMethodDetails = '';
	let rowData = null; // Placeholder for RowType
	let strCellValue;
	//Return the row
	const mapGetRow = excelGetRow(assgSheet, intRowNum); // Reuse existing function
	//Return the values from the map
	boolPassed = mapGetRow.boolPassed;
	strMethodDetails = mapGetRow.strMethodDetails;
	rowData = mapGetRow.ObjRow;
	if (boolPassed === false) {
		strCellValue = 'FAILED!!! NO ROW DATA found in sheet: ' + (assgSheet ? assgSheet.sheetName : 'Unknown Sheet') + ' for row: ' + intRowNum + '!!!';
	} else {
		//Return the cell
		//Get the colId
		const mapGetCellValue = excelGetCellValueFromRowByColNumber(rowData, intColNum); // Reuse existing function
		//Check the results
		boolPassed = StrNums.JComm_StringToBoolean(mapGetCellValue.boolPassed);
		strMethodDetails = mapGetCellValue.strMethodDetails;
		strCellValue = mapGetCellValue.CellValue;
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.CellValue = strCellValue;
	return mapResults;
}

/**
 * ------------------ excelGetCellValue -----------------------
	 * Return the string value from the assigned cell.
	 * @param assgCell The cell that contains the value (placeholder for CellType)
	 *
	 * @return strCellValue The cell value as a string
	 * @author pkanaris
	 * @author Created 11/09/2021
	 * @author Last Edited:
	 * @author Last Edited By:
	 * @author Edit Comments: (Include email, date and details)
 */
function ExcelData_excelGetCellValue(assgCell) {
	//TODO update the method to new requirements PG 06/24/2021
	const wb = TCObj.objWorkbook; // Assumes TCObj.objWorkbook is the Workbook object
	// Original: FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator()
	let evaluator = { evaluateInCell: () => ({ getCellType: () => 'STRING' }) }; // Placeholder
	if (wb && typeof wb.getCreationHelper === 'function' && typeof wb.getCreationHelper().createFormulaEvaluator === 'function') {
		// This is highly dependent on the chosen Excel library.
		// In exceljs, you might not explicitly deal with a 'FormulaEvaluator' object for simple cell value retrieval if it's handled internally.
		// You might use `assgCell.value` directly or `assgCell.text` for formatted values.
		evaluator = wb.getCreationHelper().createFormulaEvaluator();
	}

	let strActualResults = '';
	let boolMethodPassed = true;
	let strCellValue = GVars.GblNull("Value");
	const boolDoDebug = TSExecParams.getBoolDoDebug();

	if (!assgCell) {
		boolMethodPassed = false;
		strActualResults = 'FAILED!!! The assigned cell is NULL!!!';
		strCellValue = strActualResults;
		if (boolDoDebug === true) {
			console.log(strActualResults);
		}
		return strCellValue;
	}

	// Original: String strCellType = evaluator.evaluateInCell(assgCell).getCellType()
	// This part is very specific to POI. In a Node.js library, you'd check `cell.type` or similar.
	let cellType;
	// Placeholder for Excel library cell type detection
	// Example with exceljs:
	if (assgCell && assgCell.type) {
		// Mapping exceljs types to POI-like strings for switch consistency.
		switch (assgCell.type) {
			case 'number': cellType = 'NUMERIC'; break;
			case 'string': cellType = 'STRING'; break;
			case 'boolean': cellType = 'BOOLEAN'; break;
			case 'formula': cellType = 'FORMULA'; break;
			case 'null': cellType = 'BLANK'; break; // exceljs null type can be blank
			case 'error': cellType = 'ERROR'; break;
			default: cellType = 'UNKNOWN'; // Fallback
		}
	} else {
		// Fallback or assume simple string if cell object is not complex
		cellType = 'STRING'; // Default if no complex cell object or type is undefined
	}


	switch (cellType) {
		case 'NUMERIC':
			//Return the numeric value to string
			// Original: assgCell.getNumericCellValue().toString()
			strCellValue = assgCell.value !== null ? String(assgCell.value) : GVars.GblNull("Value"); // Placeholder
			break;
		case 'BOOLEAN':
			//Return the boolean value and convert to a string
			// Original: assgCell.getBooleanCellValue().toString()
			strCellValue = assgCell.value !== null ? String(assgCell.value) : GVars.GblNull("Value"); // Placeholder
			break;
		case 'FORMULA':
			//Return the value from the cell to a string
			// Original: DataFormatter formatter = new DataFormatter(); strCellValue = formatter.formatCellValue(assgCell, evaluator);
			// In exceljs, you might use cell.result for formula result, or cell.text for formatted value.
			strCellValue = assgCell.value !== null ? String(assgCell.value) : GVars.GblNull("Value"); // Placeholder
			break;
		case 'STRING':
			//Return the string value of the cell
			// Original: StrNums.JComm_HandleNoData(assgCell.getStringCellValue())
			strCellValue = StrNums.JComm_HandleNoData(assgCell.value !== null ? String(assgCell.value) : StrNums.JComm_HandleNoData(null)); // Placeholder
			break;
		case 'BLANK':
			//Fail since we expect a value and cannot take blank since we do not know what the tester requires. Should be set a keyvalue such as *Null*
			boolMethodPassed = false;
			strActualResults = 'FAILED!!! The cell type is BLANK!!!';
			strCellValue = strActualResults;
			break;
		case 'ERROR':
			//Fail since we expect a value and cannot take blank since we do not know what the tester requires. Should be set a keyvalue such as *Null*
			boolMethodPassed = false;
			strActualResults = 'FAILED!!! The cell type is ERROR!!!';
			strCellValue = strActualResults;
			break;
		default:
			boolMethodPassed = false;
			strActualResults = 'FAILED!!! Unknown cell type: ' + cellType + ' for cell with value: ' + assgCell.value;
			strCellValue = strActualResults;
			break;
	}
	if (boolMethodPassed === true) {
		strActualResults = 'The cell value is: "' + strCellValue + '".';
	}
	if (boolDoDebug === true) {
		console.log(strActualResults);
	}
	return strCellValue;
}

function ExcelData_excelGetCellByRowColNumber(assgSheet, intRowNum, intColNum) {
	// Placeholder for SheetType.getRow() and RowType.getCell()
	if (assgSheet) {
		const row = assgSheet.getRow(intRowNum);
		if (row) {
			return row.getCell(intColNum);
		}
	}
	return null;
}


/**
 * Creates the base Excel File and the first sheet
 * @param strFilePath The full path for the file.
 * @param strSheetName The sheet name to be assigned.
 * @param boolDeleteIfExist Delete the file if it exists? True/False
 * @param boolOutputDetails Output the method details? True/False
 * @param boolStopOnFailError Stop the test on step failure? True/False
 * @param strReporterOutput The reporter output format (CSV, HTML, XHTML, DataSource)`
 * @return boolMethodPassed The method passed? True/False
 * Created: 04/21/2022
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelCreateWorkbook(strFilePath, strSheetName, boolDeleteIfExist, boolOutputDetails, boolStopOnFailError, strReporterOutput) {
	// Declare the variables
	const mapResults = {};
	let strMethodDetails;
	let boolPassed = true;
	//Check if the file exist
	// Original: def tmpFile = new File(strFilePath)
	if (fs.existsSync(strFilePath) && boolDeleteIfExist === true) {
		try {
			fs.unlinkSync(strFilePath); // Synchronous delete
			// Original: DateTime.WaitSecs(1) - synchronous wait
			DateTime.WaitSecs(1); // Assuming DateTime module has synchronous WaitSecs
			if (fs.existsSync(strFilePath)) {
				boolPassed = false;
				strMethodDetails = 'FAILED!!! ATTEMPTED TO DELETE The file: ' + strFilePath + ' HOWEVER, THE FILE STILL EXIST!!!';
			} else {
				strMethodDetails = 'Deleted file: ' + strFilePath + ' prior to file creation.' + GVars.GblLineFeed('Value');
			}
		} catch (e) {
			boolPassed = false;
			strMethodDetails = `FAILED!!! Error during file deletion: ${e instanceof Error ? e.stack : String(e)}`;
		}
	} else if (fs.existsSync(strFilePath) && boolDeleteIfExist === false) {
		boolPassed = false;
		strMethodDetails = 'FAILED!!!The file: ' + strFilePath + ' ALREADY EXIST!!! NO DELETION WAS REQUESTED!!!';
	}

	if (!fs.existsSync(strFilePath) && boolPassed) { // Only proceed if deletion was successful or file didn't exist
		//Create the Workbook and assign the first sheet
		// Original: Workbook wb = new XSSFWorkbook()
		// Original: Sheet sht = wb.createSheet(strSheetName)
		let wb = undefined; // Placeholder for WorkbookType
		let sht = undefined; // Placeholder for SheetType
		// Example with exceljs:
		// wb = new exceljs.Workbook();
		// sht = wb.addWorksheet(strSheetName);
		try {
			// Need to create a new workbook and sheet using your chosen Excel library
			// Assuming a placeholder structure:
			const newWorkbook = new WorkbookConstructor(); // Replace WorkbookConstructor with actual class/factory
			wb = newWorkbook;
			sht = newWorkbook.createSheet(strSheetName); // Replace with actual method to create sheet

			// Original: File outFile = new File(strFilePath)
			// Original: outFile.getParentFile().mkdirs()
			const dirName = path.dirname(strFilePath);
			if (!fs.existsSync(dirName)) {
				fs.mkdirSync(dirName, { recursive: true }); // Synchronous create directory
			}

			// Original: FileOutputStream out = new FileOutputStream(outFile)
			// Original: wb.write(out)
			// Original: out.close()
			// In Node.js, often done by writing the buffer:
			// For exceljs:
			// const buffer = await wb.xlsx.writeBuffer(); // 'async' is NOT allowed for this exercise
			// Need synchronous write:
			// const buffer = wb.xlsx.writeBufferSync(); // Some libraries might offer sync buffer write
			// fs.writeFileSync(strFilePath, buffer);
			excelWriteWBDataAndClose(wb, strFilePath); // Use internal helper function that should map to fs.writeFileSync

			strMethodDetails = (strMethodDetails || '') + strFilePath + ' successfully written to disk.'; // Concatenate any prior details
		} catch (e) {
			boolPassed = false;
			strMethodDetails = (strMethodDetails || '') + 'FAILED!!! Creation of file: ' + strFilePath + 'has an Exception!!! SEE ERROR STACK TRACE: ' + (e instanceof Error ? e.stack : String(e));
		}
	} else if (!fs.existsSync(strFilePath) && !boolPassed) {
		// If file didn't exist initially but a prior step (like delete) failed,
		// boolPassed might already be false due to that. No need to attempt creation.
		strMethodDetails = strMethodDetails || `File ${strFilePath} does not exist, but previous operations failed.`;
	}

	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * Adds a sheet to an existing excel file
 * @param strFilePath The full path for the file.
 * @param strSheetName The sheet name to be assigned.
 * @return boolMethodPassed The method passed? True/False
 * Created: 04/21/2022
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelAddSheet(strFilePath, strSheetName, boolOutputDetails, boolStopOnFailError, strReporterOutput) {
	// Declare the variables
	const mapResults = {};
	let strMethodDetails;
	let boolPassed = true;
	let intNumStartSheets;
	let intNumEndSheets;
	let intSheetsAdded;

	let wb = undefined; // Placeholder for WorkbookType
	//Check if the file exist
	if (fs.existsSync(strFilePath)) {
		try {
			// Original: FileInputStream inputfile = new FileInputStream(new File(strFilePath))
			// Original: wb = WorkbookFactory.create(inputfile)
			const inputfileStream = excelOpenExcelFileInputStream(strFilePath); // Internal helper
			wb = excelOpenWBFromInputStream(inputfileStream); // Internal helper

			intNumStartSheets = excelGetSheetCount(wb, false, ''); // Reuse helper
			// Original: wb.createSheet(strSheetName)
			if (wb && typeof wb.createSheet === 'function') { // POI-style
				wb.createSheet(strSheetName);
			} else if (wb && typeof wb.addWorksheet === 'function') { // exceljs-style
				wb.addWorksheet(strSheetName);
			} else {
				strMethodDetails = 'FAILED!!! Excel library API does not support sheet creation.';
				boolPassed = false;
			}

			if (boolPassed) { // Only proceed if sheet creation was attempted without error
				intNumEndSheets = excelGetSheetCount(wb, false, ''); // Reuse helper
				intSheetsAdded = intNumEndSheets - intNumStartSheets;
				if (intSheetsAdded === 1) {
					strMethodDetails = 'Added the sheet: "' + strSheetName + '" to the file: "' + strFilePath + '".';
					//Save the workbook
					try {
						// Original: FileOutputStream outputfile =new FileOutputStream(new File(strFilePath)); wb.write(outputfile); outputfile.close()
						excelWriteWBDataAndClose(wb, strFilePath); // Internal helper
						strMethodDetails = strMethodDetails + strFilePath + ' successfully written to disk.';
					} catch (e) {
						boolPassed = false;
						strMethodDetails = strMethodDetails + 'FAILED!!! Creation of file: ' + strFilePath + 'has an Exception!!! SEE ERROR STACK TRACE: ' + (e instanceof Error ? e.stack : String(e));
					}
				} else {
					boolPassed = false;
					strMethodDetails = 'FAILED TO ADD the sheet: "' + strSheetName + '" to the file: "' + strFilePath + '"!!!' + GVars.GblLineFeed('Value');
					strMethodDetails = strMethodDetails + "Start sheet count is: " + intNumStartSheets + ' end count is: ' + intNumEndSheets + ' and total sheets added: ' + intSheetsAdded + '!!!';
				}
			}
		} catch (e) {
			boolPassed = false;
			strMethodDetails = 'FAILED!!! Error opening/creating workbook: ' + (e instanceof Error ? e.stack : String(e));
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!!The file: ' + strFilePath + ' DOES NOT EXIST';
	}
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/** excelSaveOutputFile
 * Saves the latest changes to the output file
 * @param wb The workbook object (placeholder for WorkbookType)
 *
 * @return mapSOFResults
 * @return boolPassed The save file passed? true/false
 * @return strMethodDetails The details for the save file
 *
 * @author pkanaris
 * @author Created: 12/29/2022
 * @author Last Edited
 * @author Last Edited By:
 * @author Edit Comments: (Include email, date and details)
 *
 */
function ExcelData_excelSaveOutputFile(wb) {
	const mapSOFResults = {};
	let boolPassed = true;
	let strMethodDetails;
	const strInputFilePath = TCObj.getStrTCInputFilePath(); // Assumes TCObj has this method
	if (strInputFilePath === null) {
		boolPassed = false;
		strMethodDetails = "FAILED!!! The StrTCInputFilePath IS NULL!!!!";
	}
	else {
		// Original: FileOutputStream outputfile =new FileOutputStream(new File(strInputFilePath))
		let outputfileStream = null; // Placeholder for FileOutputStreamType
		try {
			outputfileStream = excelOpenExcelFileOutputStream(strInputFilePath); // Internal helper
			// Original: wb.write(outputfile)
			// This is the core save operation. In exceljs: await wb.xlsx.write(outputStream);
			// Need synchronous if not using async.
			excelWriteWBDataToFile(wb, outputfileStream); // Internal helper
		} catch (e) {
			// Original: Exception e
			// If the first write failed, try to load from the original stream and write again
			let wbTemp = undefined; // Placeholder for WorkbookType
			// Original: FileInputStream fisTemp = TCObj.getObjExcelFileInputStream()
			const fisTemp = TCObj.getObjExcelFileInputStream(); // Assumes TCObj.getObjExcelFileInputStream returns FileInputStream
			if (fisTemp === null) {
				boolPassed = false;
				strMethodDetails = (strMethodDetails || '') + 'FAILED!!! Creation of file: ' + strInputFilePath + 'has an Exception!!! SEE ERROR STACK TRACE: ' + (e instanceof Error ? e.stack : String(e));
			} else {
				// Original: wbTemp = WorkbookFactory.create(fisTemp)
				wbTemp = excelOpenWBFromInputStream(fisTemp); // Internal helper
				try {
					// Original: wbTemp.write(outputfile)
					excelWriteWBDataToFile(wbTemp, outputfileStream); // Internal helper
				} catch (e2) {
					boolPassed = false;
					strMethodDetails = (strMethodDetails || '') + 'FAILED!!! Creation of file: ' + strInputFilePath + 'has an Exception!!! SEE ERROR STACK TRACE: ' + (e2 instanceof Error ? e2.stack : String(e2));
				}
			}
		} finally {
			if (outputfileStream) {
				excelCloseExcelFileOutputStream(outputfileStream); // Internal helper
			}
			if (boolPassed === true) {
				strMethodDetails = (strMethodDetails || '') + strInputFilePath + ' successfully written to disk.';
			}
		}
		outputfile = null; SeSSleep(100)
	}

	//Update the map
	mapSOFResults.boolPassed = boolPassed.toString();
	mapSOFResults.strMethodDetails = strMethodDetails;
	return mapSOFResults;
}

/** excelAddSheetToTCFile
 * Adds a sheet to the test case input file
 * @param strSheetName The sheet name to be assigned.
 *
 * @return mapResults:
 * @return boolPassed The method passed? True/False
 * @return strMethodDetails The details on the method processing
 * Created: 04/21/2022
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelAddSheetToWorkBook(strSheetName) {
	// Declare the variables
	const mapResults = {};
	let strMethodDetails;
	let boolPassed = true;
	let intNumStartSheets;
	let intNumEndSheets;
	let intSheetsAdded;
	let objSh; // Placeholder for SheetType

	//Return the test case wb object
	let wb = undefined; // Placeholder for WorkbookType
	wb = TCObj.getObjWorkbook(); // Assumes TCObj has this method
	const strFilePath = TCObj.getStrTCInputFilePath(); // Assumes TCObj has this method
	if (wb === null) {
		//Error
		boolPassed = false;
		strMethodDetails = 'FAILED!!!The Test Case EXCEL DOES NOT EXIST!!!';
	}
	else {
		let strTempSheetName = strSheetName;
		let boolSheetPresentPassed = true;
		let intTempSheetNum = 0;
		//TODO Check if the sheet exist and should we fail it or delete the sheet and recreate it?
		let mapGetSheetByName = {};
		let boolDoCheck = true;
		while (boolDoCheck === true) {
			mapGetSheetByName = excelGetSheetByName(wb, strTempSheetName); // Reuse internal helper
			boolSheetPresentPassed = StrNums.JComm_StringToBoolean(mapGetSheetByName.boolPassed);
			if (boolSheetPresentPassed === true) {
				intTempSheetNum++;
				if (intTempSheetNum < 10) {
					//Update the sheetName with _intTempSheetNum
					strTempSheetName = strSheetName + "_" + intTempSheetNum;
				}
				else {
					boolPassed = false;
					strMethodDetails = "FAILED!!! Attempted to ADD MORE THAN '9' INSTANCES of the sheet" + strSheetName + "!!!";
					boolDoCheck = false;
				}
			}
			else {
				boolDoCheck = false; //The sheet did not exist go ahead and create it.
			}
		}
		if (boolPassed === true) {
			intNumStartSheets = excelGetSheetCount(wb, false, ''); // Reuse helper
			strSheetName = strTempSheetName;
			// Original: wb.createSheet(strSheetName)
			if (wb && typeof wb.createSheet === 'function') { // POI-style
				wb.createSheet(strSheetName);
			} else if (wb && typeof wb.addWorksheet === 'function') { // exceljs-style
				wb.addWorksheet(strSheetName);
			} else {
				strMethodDetails = 'FAILED!!! Excel library API does not support sheet creation.';
				boolPassed = false;
			}

			if (boolPassed) {
				intNumEndSheets = excelGetSheetCount(wb, false, ''); // Reuse helper
				intSheetsAdded = intNumEndSheets - intNumStartSheets;
				if (intSheetsAdded === 1) {
					strMethodDetails = 'Added the sheet: "' + strSheetName + '" to the test case workbook.';
					//Save the workbook
					const mapSaveWb = excelSaveOutputFile(wb); // Reuse internal helper
					boolPassed = StrNums.JComm_StringToBoolean(mapSaveWb.boolPassed);
					if (boolPassed === true) {
						//Check for the sheet
						let mapGetSheetByNameAfterSave = {};
						mapGetSheetByNameAfterSave = excelGetSheetByName(wb, strSheetName); // Reuse helper
						boolSheetPresentPassed = StrNums.JComm_StringToBoolean(mapGetSheetByNameAfterSave.boolPassed);
						if (boolSheetPresentPassed === true) {
							//return the sheet
							TCObj.setObjSheetTCActiveOutputSheet(mapGetSheetByNameAfterSave.objWbSheet); // Assuming TCObj has this method
							TCObj.objWorkbook = wb;
						}
					} else {
						strMethodDetails = mapSaveWb.strMethodDetails;
					}
				} else {
					boolPassed = false;
					strMethodDetails = 'FAILED TO ADD the sheet: "' + strSheetName + '" to the file: "' + strFilePath + '"!!!' + GVars.GblLineFeed('Value');
					strMethodDetails = strMethodDetails + "Start sheet count is: " + intNumStartSheets + ' end count is: ' + intNumEndSheets + ' and total sheets added: ' + intSheetsAdded + '!!!';
				}
			}
		}
	}

	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/** excelCreateShtWithHdrCols
 * Adds a sheet to the test case file to include add column names
 * @param strShtName		The sheet name to be added
 * @param strShtHdrColNames	The column names to be added to the sheet
 *
 * @return mapResults:
 * @return boolPassed 		The method passed? True/False
 * @return strMethodDetails The details on the method processing
 * @return objSheet			The sheet object that was added.
 * Created: 12/20/2022
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelCreateShtWithHdrCols(strShtName, strShtHdrColNames) {
	const mapResults = {};
	let boolPassed = true;
	let strMethodDetails;
	//Create the output sheet
	const mapAddSheet = excelAddSheetToWorkBook(strShtName); // Reuse internal helper
	const boolSheetAddedPassed = StrNums.JComm_StringToBoolean(mapAddSheet.boolPassed);
	if (boolSheetAddedPassed === true) {
		//Add the column names
		const mapAddColNames = excelCreateHeaderCols(strShtHdrColNames); // Reuse internal helper
		const boolAddColNames = StrNums.JComm_StringToBoolean(mapAddColNames.boolPassed);
		if (boolAddColNames === true) {
			//Return the header elements
			const lstColNames = mapAddColNames.lstColNames;
			//Update the header column(s) to the assigned theme
			const strAssgTheme = 'OutputDataHdrStd';
			const mapUpdateHdrTheme = excelSetRowCellFormatByTheme(0, strAssgTheme); // Reuse internal helper
			boolPassed = StrNums.JComm_StringToBoolean(mapUpdateHdrTheme.boolPassed);
			strMethodDetails = mapUpdateHdrTheme.strMethodDetails;
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
	if (boolPassed === true) {
		strMethodDetails = "Added the output sheet '" + strShtName + "' with the column names of: " + strShtHdrColNames;
	}
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

//****************************** Excel Streams ***********************************//
/**
 * @param strFilePath The full path for the file.
 * @return FileInputStream or null. (Placeholder for Node.js equivalent)
 */
function ExcelData_excelOpenExcelFileInputStream(strFilePath) {
	//Check if the file exist
	if (fs.existsSync(strFilePath)) {
		//check if the file is an Excel file (might need mimetypes check in Node.js)
		//Open and return the excel file
		// Original: FileInputStream inputStream = new FileInputStream(new File(strFilePath))
		// In Node.js, for simple file operations, often reading the whole file at once is common.
		// For streams, you'd use fs.createReadStream.
		// However, POI APIs like WorkbookFactory.create(InputStream) often consume the whole stream.
		// So, we'll simulate returning a "stream-like" object or buffer.
		// For simplicity, let's treat it as a placeholder that will be resolved to a buffer or actual stream when excelOpenWBFromInputStream is implemented.
		// Placeholder returning a path, assuming the next function will read it.
		return strFilePath; // Effectively, the file path is the "stream" for synchronous operations here.
	} else {
		return null;
	}
}

function ExcelData_excelCloseExcelFileInputStream(inputStream) {
	// Original: inputStream.close()
	// If inputStream was simply a path, no explicit close needed.
	// If it's an actual Node.js ReadStream, you would.
	if (inputStream && typeof inputStream.close === 'function') {
		inputStream.close();
	}
}

function ExcelData_excelOpenExcelFileOutputStream(strFilePath) {
	//Check if the file exist (or parent directories should be created)
	const dirName = path.dirname(strFilePath);
	if (!fs.existsSync(dirName)) {
		fs.mkdirSync(dirName, { recursive: true });
	}
	// Original: FileOutputStream outputStream = new FileOutputStream(new File(strFilePath))
	// Similar to FileInputStream, this function will essentially pass the filepath.
	// A real Node.js implementation for streams would use fs.createWriteStream.
	return strFilePath; // Return the path for synchronous operations.
}

function ExcelData_excelCloseExcelFileOutputStream(outputStream) {
	// Original: outputStream.close()
	// If outputStream was simply a path, no explicit close needed after fs.writeFileSync.
	// If it's an actual Node.js WriteStream, you would.
	if (outputStream && typeof outputStream.close === 'function') {
		outputStream.close();
	}
}

/**
 * Creates a workbook from an input stream (simulated with file path/buffer)
 * @param inputStream The input stream (simulated here as a file path or buffer)
 * @return Workbook object (placeholder for WorkbookType)
 */
function ExcelData_excelOpenWBFromInputStream(inputStream) {
	// Original: Workbook assgWb = new XSSFWorkbook(); assgWb = WorkbookFactory.create(inputStream)
	let assgWb = undefined; // Placeholder for WorkbookType
	if (typeof inputStream === 'string') { // If inputStream is the file path
		const fileBuffer = fs.readFileSync(inputStream); // Read content
		// Example with exceljs:
		// const workbook = new ExcelJS.Workbook(); await workbook.xlsx.load(fileBuffer);
		// You'd need to adapt to your chosen lib's synchronous loading.
		// Placeholder for a synchronous Excel library function:
		assgWb = createWorkbookFromBuffer(fileBuffer); // Replace with actual library call
	} else {
		// Handle other types of streams/buffers if needed
		assgWb = createWorkbookFromBuffer(inputStream); // Assume inputStream is already a buffer
	}
	return assgWb;
}

/**
 * Placeholder for creating workbook from buffer. Replace with actual library.
 * @param buffer
 */
function ExcelData_createWorkbookFromBuffer(buffer) {
	// Example using 'exceljs' (if it had sync load from buffer)
	// const wb = new ExcelJS.Workbook();
	// wb.xlsx.load(buffer, { sync: true }); // Hypothetical synchronous load option
	// return wb;
	return {}; // Placeholder for now
}

/**
 * Open a workbook from an output stream (typically not done this way for *creation*)
 * This method seems to imply creation or loading from something writable.
 * POI WorkbookFactory.create can sometimes take a File directly, which might be implied.
 * Re-evaluating the original intent. It's unusual to *create* a workbook from an output stream directly.
 * It's more common to *write* a workbook *to* an output stream.
 * Given `WorkbookFactory.create(outputStream)`, it implies `outputStream` was used like an `InputStream`.
 * I'll assume it's attempting to load a workbook from a path that will be used for output later.
 * @param outputStream (simulated here as a file path)
 * @return Workbook object (placeholder for WorkbookType)
 */
function ExcelData_excelOpenWBFromOutputStream(outputStream) {
	// Original: Workbook assgWb = new XSSFWorkbook(); assgWb = WorkbookFactory.create(outputStream)
	let assgWb = undefined; // Placeholder for WorkbookType
	if (typeof outputStream === 'string') { // If outputStream is the file path
		// This implies loading from this path, as creating from an empty output stream won't work.
		// It's likely intended to load an existing file, which will then be written to.
		const fileBuffer = fs.readFileSync(outputStream);
		assgWb = createWorkbookFromBuffer(fileBuffer); // Reuse createWorkbookFromBuffer
	}
	return assgWb;
}

/**
 * Writes workbook data to a file path and closes the associated stream if applicable.
 * @param assgWb The workbook object (placeholder for WorkbookType)
 * @param strFilePath The full path for the file.
 * @return mapResults
 */
function ExcelData_excelWriteWBDataAndClose(assgWb, strFilePath) {
	// Declare the variables
	const mapResults = {};
	let strMethodDetails;
	let boolPassed = true;
	let outputStream = null;

	try {
		// Original: FileOutputStream outputStream = new FileOutputStream(new File(strFilePath))
		// This helper will handle the actual file write.
		excelWriteWBDataToFileSync(assgWb, strFilePath); // Internal helper to write sync
		strMethodDetails = `Workbook successfully written to ${strFilePath}.`;
	} catch (e) {
		boolPassed = false;
		strMethodDetails = (strMethodDetails || '') + 'FAILED!!! Writing of file: ' + strFilePath + ' has an Exception!!! SEE ERROR STACK TRACE: ' + (e instanceof Error ? e.stack : String(e));
	}

	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * Internal helper to write workbook data to a file synchronously.
 * Replace with your Excel library's specific synchronous save function.
 * @param assgWb The workbook object.
 * @param filePath The path to save to.
 */
function excelWriteWBDataToFileSync(assgWb, filePath) {
	// Example with exceljs:
	// const buffer = assgWb.xlsx.writeBuffer(); // If sync buffer write exists
	// fs.writeFileSync(filePath, buffer);
	// OR if direct file write is sync:
	// assgWb.xlsx.writeFile(filePath, { sync: true }); // Hypothetical sync option
	//
	// For now, it's a placeholder:
	fs.writeFileSync(filePath, 'dummy content'); // Placeholder
}

/**
 * Internal helper to write workbook data to an output stream (used by excelSaveOutputFile).
 * This will use the Node.js `fs.writeFileSync` as a direct output.
 * @param wb The workbook object.
 * @param outputStreamPath The pathway to the output stream (representing a filename).
 */
function excelWriteWBDataToFile(wb, outputStreamPath) {
	// In Node.js, `FileOutputStream` is generally replaced by `fs.writeFileSync` for synchronous saves,
	// or `fs.createWriteStream` for streaming.
	// The `wb.write(outputfile)` would translate to getting the workbook data (e.g., as a buffer)
	// and writing that buffer to the file.
	// Example with exceljs:
	// const buffer = wb.xlsx.writeBuffer(); // Synchronous buffer generation
	// fs.writeFileSync(outputStreamPath, buffer);
	//
	// Placeholder:
	fs.writeFileSync(outputStreamPath, 'dummy content'); // Placeholder content
}

/**
 * Open the File, Create the column names and Output the file. NOTE!!! We will overwrite the existing values
 * @param strColNames The column names to be assigned to the header.
 * @param strColNamesDelimiter The Delimiter that is used to seperate the values.
 * @return mapResults The method processing Passed? True/False
 * Created: 04/21/2022
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelCreateHeaderCols(strColNames) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value");
	let boolPassed = true;
	let strMethodDetails = '';
	let lstColName; // Placeholder for string array

	//Return the test case wb object
	let wb = undefined; // Placeholder for WorkbookType
	wb = TCObj.getObjWorkbook(); // Assumes TCObj has this method
	let objSh = TCObj.getObjSheetTCActiveOutputSheet(); // Assumes TCObj has this method
	let strSheetName;
	if (objSh && objSh.sheetName) {
		strSheetName = objSh.sheetName; // Placeholder for SheetType.sheetName
	} else {
		strSheetName = 'Unknown Sheet';
	}
	const strFilePath = TCObj.getStrTCInputFilePath(); // Assumes TCObj has this method

	if (objSh !== null) {
		//Check for the header row and create if not present
		// Original: Row rowHdr = objSh.getRow(0); if (rowHdr == null) { objSh.createRow(0); rowHdr = objSh.getRow(0); }
		let rowHdr = objSh.getRow(0); // Placeholder for SheetType.getRow(0)
		if (rowHdr === null) {
			rowHdr = objSh.createRow(0); // Placeholder for SheetType.createRow(0)
		}
		//Create the array for the column names and make sure the value is not empty
		// Original: List<String> lstColName = StrNums.JComm_SplitString(strColNames, gblDelimiter)
		lstColName = StrNums.JComm_SplitString(strColNames, gblDelimiter);
		let intLstCnt = lstColName ? lstColName.length : 0;
		let strAryColName = '';
		let strFinalColNames = ''; // This variable is not used in the original Groovy, though it's set.
		let intRowColCnt = 0; // Not used accurately in calculation in original code

		if (intLstCnt > 0) {
			let assgCell = null; // Placeholder for CellType
			//Build loop to return the values
			for (let loopIteration = 0; loopIteration < intLstCnt; loopIteration++) {
				assgCell = null;
				strAryColName = lstColName[loopIteration];
				assgCell = rowHdr.getCell(loopIteration); // Placeholder for RowType.getCell(loopIteration)
				if (assgCell === null) {
					//Create the cell
					assgCell = rowHdr.createCell(loopIteration); // Placeholder for RowType.createCell(loopIteration)
				}
				//Assign the value
				assgCell.setCellValue(strAryColName.trim()); // Placeholder for CellType.setCellValue()
				//TODO add style for columns so we can set the column always to text
				//TODO add style for the header
			}
			//Once the loop is done, count the columns and output all column names
			let intColCnt = rowHdr.lastCellNum; // Placeholder for RowType.lastCellNum
			if (intColCnt === intLstCnt) {
				strMethodDetails = 'Added all the ' + intColCnt + ' col(s) to the sheet: "' + strSheetName + '" in worksheet: "' + strFilePath + '".';
			} else {
				boolPassed = false;
				strMethodDetails = 'FAILED!!! The ' + intColCnt + ' col(s) does NOT MATCH the expected ' + intLstCnt + ' col(s) to be added to the sheet: "' + strSheetName + '" in worksheet: "' + strFilePath + '".';
			}
			if (intColCnt > 0) {
				//Return all the columns added
				// Original: strFinalColNames = this.excelGetHdrColNames(objSh)
				const mapGetHdrColNames = excelGetHdrColNames(objSh); // Reuse internal helper
				strFinalColNames = mapGetHdrColNames.ColNames; // Assuming it returns a Map result
				strMethodDetails = strMethodDetails + gblLineFeed + 'The column names are: ' + gblLineFeed + strFinalColNames;
			}
			if (intColCnt > 0) {
				//Save the excel file
				const mapSaveWb = excelSaveOutputFile(wb); // Reuse internal helper
				boolPassed = StrNums.JComm_StringToBoolean(mapSaveWb.boolPassed);
				if (boolPassed === true) {
					TCObj.objWorkbook = wb; // Assuming TCObj has this property
				} else {
					strMethodDetails = mapSaveWb.strMethodDetails;
				}
			}
		} else {
			boolPassed = false;
			strMethodDetails = 'FAILED!!! Unable to add columns. The strColNames did not contain any items.';
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! THe sheet: "' + strSheetName + '" was not found in the workbook from the file: "' + strFilePath + '"!!!';
	}
	//Output the method pass/fail and message
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.lstColNames = lstColName;
	return mapResults;
}

/**
 * Autofit the column width based on widest column of data plus 10%
 *
 * @return boolMethodPassed The method processing Passed? True/False
 * Created: 05/01/2022
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelAutofitCols() {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value"); // Not directly used in this func
	let boolPassed = true;
	let strMethodDetails = '';
	let lstColName; // Not used in this func

	//Return the test case wb object
	let wb = undefined; // Placeholder for WorkbookType
	wb = TCObj.getObjWorkbook(); // Assumes TCObj has this method
	let objSh = TCObj.getObjSheetTCActiveOutputSheet(); // Assumes TCObj has this method
	let strSheetName;
	if (objSh && objSh.sheetName) {
		strSheetName = objSh.sheetName; // Placeholder for SheetType.sheetName
	} else {
		strSheetName = 'Unknown Sheet';
	}
	const strFilePath = TCObj.getStrTCInputFilePath(); // Assumes TCObj has this method

	if (objSh !== null) {
		//Check for the header row and create if not present
		let rowHdr = objSh.getRow(0); // Placeholder for SheetType.getRow(0)
		if (rowHdr === null) {
			boolPassed = false;
			strMethodDetails = 'FAILED!!! THe sheet: "' + strSheetName + '" in the workbook from the file: "' + strFilePath + '" DOES NOT CONTAIN A HEADER!!!';
		}
		else {
			let cntCols = rowHdr.getLastCellNum(); // Placeholder for RowType.getLastCellNum()
			if (cntCols > 0) {
				let intColAutoWidth;
				//Loop through the columns and set the width
				for (let loopCols = 0; loopCols < cntCols; loopCols++) {
					// Original: objSh.autoSizeColumn(loopCols)
					// Original: intColAutoWidth = objSh.getColumnWidth(loopCols)
					if (objSh && typeof objSh.autoSizeColumn === 'function') { // Placeholder
						objSh.autoSizeColumn(loopCols);
						intColAutoWidth = objSh.getColumnWidth(loopCols);
					} else {
						// Fallback for libraries without autoSizeColumn, or manual calculation
						intColAutoWidth = 256 * 10; // A reasonable default (e.g., 10 chars wide)
					}

					let intNewColWidth = intColAutoWidth * 1.10; //Added 10% for pixel changes.
					if (intNewColWidth > 254 * 256) { //Check if width is illegal 225 Char by 256 is may
						intNewColWidth = 254 * 256;
					}
					objSh.setColumnWidth(loopCols, intNewColWidth); // Placeholder
				}
				//Go through all of the cells and wrap the text as needed.
				let intRowCnt = objSh.getLastRowNum(); // Placeholder for SheetType.getLastRowNum()
				if (intRowCnt > 0) {
					let assgRow = null; // Placeholder for RowType
					let assgCell = null; // Placeholder for CellType
					let style = null; // Placeholder for CellStyleType
					// Original: CellStyle style = wb.createCellStyle();
					if (wb && typeof wb.createCellStyle === 'function') {
						style = wb.createCellStyle(); // Placeholder
					}

					for (let intLoopRows = 1; intLoopRows <= intRowCnt; intLoopRows++) {
						assgRow = objSh.getRow(intLoopRows); // Placeholder
						if (assgRow) {
							for (let intLoopCols = 0; intLoopCols < cntCols; intLoopCols++) {
								assgCell = assgRow.getCell(intLoopCols); // Placeholder
								if (assgCell) {
									style = assgCell.getCellStyle(); // Placeholder
									if (style && typeof style.setWrapText === 'function') {
										style.setWrapText(true); // Placeholder
									}
									assgCell.setCellStyle(style); // Placeholder
								}
							}
						}
					}
				}
				//Save the excel file
				const mapSaveWb = excelSaveOutputFile(wb); // Reuse internal helper
				boolPassed = StrNums.JComm_StringToBoolean(mapSaveWb.boolPassed);
				if (boolPassed === true) {
					TCObj.objWorkbook = wb;
					TCObj.objSheetTCActiveOutputSheet = objSh;
					strMethodDetails = 'The excel file: "' + strFilePath + '" sheet "' + strSheetName + '" columns were set to autofit plus 10%.';
				}
				else {
					strMethodDetails = mapSaveWb.strMethodDetails;
				}
			}
			else {
				boolPassed = false;
				strMethodDetails = 'FAILED!!! THe sheet: "' + strSheetName + '" in the workbook from the file: "' + strFilePath + '" HEADER CONTAINS ZERO COLUMNS!!!';
			}
		}

	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! THe sheet: "' + strSheetName + '" was not found in the workbook from the file: "' + strFilePath + '"!!!';
	}
	//Output the method pass/fail and message
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/** excelSetCellValueByRowColNumber
 * Set the value for the assigned cell to include cell theme (Fill and font)
 * @param intRowNum 			The excel row to be updated.
 * @param intColNum 			The excel column where the cell is located to be updated.
 * @param strAssgValue 			The value assigned to the cell.
 * @param strAssgTheme			The cell theme indicating the fill and font information
 * @return mapResults 			The map containing the pass/fail, method details, the Excel Workbook and Sheet
 * Created: 05/01/2022
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelSetCellValueByRowColNumber(intRowNum, intColNum, strAssgValue, strAssgTheme) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value"); // Not used
	let boolPassed = true;
	let strMethodDetails = '';
	let boolValueSet;
	let assgRow = null; // Placeholder for RowType

	//Return the test case wb object
	let wb = undefined; // Placeholder for WorkbookType
	wb = TCObj.getObjWorkbook(); // Assumes TCObj has this method
	let objSh = TCObj.getObjSheetTCActiveOutputSheet(); // Assumes TCObj has this method
	let strSheetName;
	if (objSh && objSh.sheetName) {
		strSheetName = objSh.sheetName; // Placeholder for SheetType.sheetName
	} else {
		strSheetName = 'Unknown Sheet';
	}
	const strFilePath = TCObj.getStrTCInputFilePath(); // Assumes TCObj has this method

	if (wb !== null) {
		if (objSh !== null) {
			//Row
			let rowSelected = objSh.getRow(intRowNum); // Placeholder for SheetType.getRow
			if (rowSelected === null) {
				//Create the row so we can continue
				rowSelected = objSh.createRow(intRowNum); // Placeholder for SheetType.createRow
			}
			//Cell
			let assgCell = null; // Placeholder for CellType
			assgCell = rowSelected.getCell(intColNum); // Placeholder for RowType.getCell
			if (assgCell === null) {
				//Set wraptext (implicitly or via style later)
				//Set the cell value
				assgCell = rowSelected.createCell(intColNum); // Placeholder for RowType.createCell
			}
			// Original: assgCell = rowSelected.getCell(intColNum) // This line is redundant after creation logic
			assgCell.setCellValue(strAssgValue); // Placeholder for CellType.setCellValue
			boolValueSet = true;
			//Open the output file (implicitly saves the workbook when needed)
			if (boolValueSet === true) {
				//Save the excel file
				const mapSaveWb = excelSaveOutputFile(wb); // Reuse internal helper
				boolPassed = StrNums.JComm_StringToBoolean(mapSaveWb.boolPassed);
				if (boolPassed === true) {
					TCObj.objWorkbook = wb;
					TCObj.objSheetTCActiveOutputSheet = objSh;
					strMethodDetails = 'The excel file: "' + strFilePath + '" sheet "' + strSheetName + '" row: ' + intRowNum + ' column: ' + intColNum + ' was updated with the value of: "' + strAssgValue + '".';
				}
				else {
					strMethodDetails = mapSaveWb.strMethodDetails;
				}
			}
		}
		//Set the cell theme first
		const mapSetCellThemeResults = excelSetSaveCellTheme(intRowNum, intColNum, strAssgTheme); // Reuse internal helper
		const boolSetCellTheme = StrNums.JComm_StringToBoolean(mapSetCellThemeResults.boolPassed);
		if (boolSetCellTheme === false) {
			boolPassed = false;
			strMethodDetails = mapSetCellThemeResults.strMethodDetails;
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! Unable to open the workbook from the file: "' + strFilePath + '"!!!';
	}
	//Output the method pass/fail and message
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/** excelSetCellValueByRowNumberColName
 * Set the value for the assigned cell to include cell theme (Fill and font) using row number and column name
 * @param intRowNum 			The excel row to be updated.
 * @param strAssgColName 		The excel column name in which the cell is to be updated.
 * @param strAssgValue 			The value assigned to the cell.
 * @param strAssgTheme			The cell theme indicating the fill and font information
 * @return mapResults 			The map containing the pass/fail, method details, the Excel Workbook and Sheet
 * Created: 05/01/2022
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelSetCellValueByRowNumberColName(intRowNum, strAssgColName, strAssgValue, strAssgTheme) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value"); // Not used
	let boolPassed = true;
	let strMethodDetails = '';
	let boolValueSet;
	let assgRow = null; // Placeholder for RowType
	let intAssgColNum; // Determined based on column name

	//Return the test case wb object
	let wb = undefined; // Placeholder for WorkbookType
	wb = TCObj.getObjWorkbook(); // Assumes TCObj has this method
	let objSh = TCObj.getObjSheetTCActiveOutputSheet(); // Assumes TCObj has this method
	let strSheetName;
	if (objSh && objSh.sheetName) {
		strSheetName = objSh.sheetName; // Placeholder for SheetType.sheetName
	} else {
		strSheetName = 'Unknown Sheet';
	}
	const strFilePath = TCObj.getStrTCInputFilePath(); // Assumes TCObj has this method

	if (wb !== null) {
		if (objSh !== null) {
			//Row
			let rowSelected = objSh.getRow(intRowNum); // Placeholder for SheetType.getRow
			if (rowSelected === null) {
				//Create the row so we can continue
				rowSelected = objSh.createRow(intRowNum); // Placeholder for SheetType.createRow
			}
			//Cell
			let assgCell = null; // Placeholder for CellType
			//Return the column number
			const mapGetColNumber = excelGetColIndexByColName(objSh, strAssgColName); // Reuse internal helper
			intAssgColNum = mapGetColNumber.ColIndex;
			if (intAssgColNum > -1) {
				assgCell = rowSelected.getCell(intAssgColNum); // Placeholder for RowType.getCell
				if (assgCell === null) {
					//Set wraptext (implicitly or via style later)
					//Set the cell value
					assgCell = rowSelected.createCell(intAssgColNum); // Placeholder for RowType.createCell
				}
				 // Original: assgCell = rowSelected.getCell(intAssgColNum) // This line is redundant after creation logic
				assgCell.setCellValue(strAssgValue); // Placeholder for CellType.setCellValue
				boolValueSet = true;
				//Open the output file (implicitly saves the workbook when needed)
				if (boolValueSet === true) {
					//Save the excel file
					const mapSaveWb = excelSaveOutputFile(wb); // Reuse internal helper
					boolPassed = StrNums.JComm_StringToBoolean(mapSaveWb.boolPassed);
					if (boolPassed === true) {
						TCObj.objWorkbook = wb;
						TCObj.objSheetTCActiveOutputSheet = objSh;
						strMethodDetails = 'The excel file: "' + strFilePath + '" sheet "' + strSheetName + '" row: ' + intRowNum + ' column: ' + intAssgColNum + ' was updated with the value of: "' + strAssgValue + '".';
					}
					else {
						strMethodDetails = mapSaveWb.strMethodDetails;
					}
				}
				//Set the cell theme
				const mapSetCellThemeResults = excelSetSaveCellTheme(intRowNum, intAssgColNum, strAssgTheme); // Reuse internal helper
				const boolSetCellTheme = StrNums.JComm_StringToBoolean(mapSetCellThemeResults.boolPassed);
				if (boolSetCellTheme === false) {
					boolPassed = false;
					strMethodDetails = mapSetCellThemeResults.strMethodDetails;
				}
			}
			else {
				boolPassed = false;
				strMethodDetails = " FAILED!!! NO COLUMN FOUND for NAME '" + strAssgColName + "'!!!";
			}
		}
		else {
			boolPassed = false;
			strMethodDetails = 'FAILED!!! NO ACTIVE WORKSHEET was FOUND!!!';
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! Unable to open the workbook from the file: "' + strFilePath + '"!!!';
	}
	//Output the method pass/fail and message
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/** excelSetCellValueByRowColNumberCustTheme
 * Set the value for the assigned cell to include custom cell theme values(Fill and font)
 * @param mapCustomValues 		The values including the row, column, value and custom them variables assigned.
 * @return mapResults 			The map containing the pass/fail, method details, the Excel Workbook and Sheet
 * Created: 05/08/2022
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelSetCellValueByRowColNumberCustTheme(mapCustomValues) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value"); // Not used
	let boolPassed = true;
	let strMethodDetails = '';
	const intRowNum = mapCustomValues.ExcelRow;
	const intColNum = mapCustomValues.ExcelCol;
	const strAssgValue = mapCustomValues.CellValue;
	let boolValueSet;
	let assgRow = null; // Placeholder for RowType

	//Return the test case wb object
	let wb = undefined; // Placeholder for WorkbookType
	wb = TCObj.getObjWorkbook(); // Assumes TCObj has this method
	let objSh = TCObj.getObjSheetTCActiveOutputSheet(); // Assumes TCObj has this method
	let strSheetName;
	if (objSh && objSh.sheetName) {
		strSheetName = objSh.sheetName; // Placeholder for SheetType.sheetName
	} else {
		strSheetName = 'Unknown Sheet';
	}
	const strFilePath = TCObj.getStrTCInputFilePath(); // Assumes TCObj has this method
	if (wb !== null) {
		if (objSh !== null) {
			//Row
			let rowSelected = objSh.getRow(intRowNum); // getRow returns a Row object or null
			if (rowSelected === null) {
				//Create the row so we can continue
				rowSelected = objSh.createRow(intRowNum); // createRow returns a new Row
			}
			//Cell
			let assgCell = null; // Placeholder for CellType
			assgCell = rowSelected.getCell(intColNum); // getCell returns a Cell object or null
			if (assgCell === null) {
				//Set wraptext (implicitly or via style later)
				//Set the cell value
				assgCell = rowSelected.createCell(intColNum); // createCell returns a new Cell
			}
			// Original: assgCell = rowSelected.getCell(intColNum) // This line is redundant after creation
			assgCell.setCellValue(strAssgValue); // Set the cell value
			boolValueSet = true;
			//Open the output file (implicitly saves the workbook when needed)
			if (boolValueSet === true) {
				//Save the excel file
				const mapSaveWb = excelSaveOutputFile(wb); // Reuse internal helper
				boolPassed = StrNums.JComm_StringToBoolean(mapSaveWb.boolPassed);
				if (boolPassed === true) {
					TCObj.objWorkbook = wb;
					TCObj.objSheetTCActiveOutputSheet = objSh;
					strMethodDetails = 'The excel file: "' + strFilePath + '" sheet "' + strSheetName + '" row: ' + intRowNum + ' column: ' + intColNum + ' was updated with the value of: "' + strAssgValue + '".';
				}
				else {
					strMethodDetails = mapSaveWb.strMethodDetails;
				}
			}
		}
		//Set the cell theme first
		const mapSetCellThemeResults = excelSetSaveCellCustomTheme(mapCustomValues); // Reuse internal helper
		const boolSetCellTheme = StrNums.JComm_StringToBoolean(mapSetCellThemeResults.boolPassed);
		if (boolSetCellTheme === false) {
			boolPassed = false;
			strMethodDetails = mapSetCellThemeResults.strMethodDetails;
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! Unable to open the workbook from the file: "' + strFilePath + '"!!!';
	}
	//Output the method pass/fail and message
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/** excelSetCellValueByRowColNaneCustTheme
 * Set the value for the assigned cell to include custom cell theme values(Fill and font)
 * @param mapCustomValues 		The values including the row, column, value and custom them variables assigned.
 * @return mapResults 			The map containing the pass/fail, method details, the Excel Workbook and Sheet
 * Created: 05/08/2022
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelSetCellValueByRowColNaneCustTheme(mapCustomValues) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value"); // Not used
	let boolPassed = true;
	let strMethodDetails = '';
	const intRowNum = mapCustomValues.ExcelRow;
	const strColName = mapCustomValues.ExcelColName;
	let intColNum;
	const strAssgValue = mapCustomValues.CellValue;
	let boolValueSet;
	let assgRow = null; // Placeholder for RowType

	//Return the test case wb object
	let wb = undefined; // Placeholder for WorkbookType
	wb = TCObj.getObjWorkbook(); // Assumes TCObj has this method
	let objSh = TCObj.getObjSheetTCActiveOutputSheet(); // Assumes TCObj has this method
	let strSheetName;
	if (objSh && objSh.sheetName) {
		strSheetName = objSh.sheetName; // Placeholder for SheetType.sheetName
	} else {
		strSheetName = 'Unknown Sheet';
	}
	const strFilePath = TCObj.getStrTCInputFilePath(); // Assumes TCObj has this method
	if (wb !== null) {
		if (objSh !== null) {
			//Row
			let rowSelected = objSh.getRow(intRowNum); // getRow returns a Row object or null
			if (rowSelected === null) {
				//Create the row so we can continue
				rowSelected = objSh.createRow(intRowNum); // createRow returns a new Row
			}
			//Cell
			let assgCell = null; // Placeholder for CellType
			const mapGetColID = excelGetColIndexByColName(objSh, strColName); // Reuse internal helper
			intColNum = mapGetColID.ColIndex;
			//Add the columnnumber to the mapCustomValues which is need is save cell theme
			mapCustomValues.ExcelCol = intColNum; // Directly assign to mapCustomValues
			assgCell = rowSelected.getCell(intColNum); // getCell returns a Cell object or null
			if (assgCell === null) {
				//Set wraptext (implicitly or via style later)
				//Set the cell value
				assgCell = rowSelected.createCell(intColNum); // createCell returns a new Cell
			}
			// Original: assgCell = rowSelected.getCell(intColNum) // This line is redundant after creation
			assgCell.setCellValue(strAssgValue); // Set the cell value
			boolValueSet = true;
			//Open the output file (implicitly saves the workbook when needed)
			if (boolValueSet === true) {
				//Save the excel file
				const mapSaveWb = excelSaveOutputFile(wb); // Reuse internal helper
				boolPassed = StrNums.JComm_StringToBoolean(mapSaveWb.boolPassed);
				if (boolPassed === true) {
					TCObj.objWorkbook = wb;
					TCObj.objSheetTCActiveOutputSheet = objSh;
					strMethodDetails = 'The excel file: "' + strFilePath + '" sheet "' + strSheetName + '" row: ' + intRowNum + ' column: ' + intColNum + ' was updated with the value of: "' + strAssgValue + '".';
					// Original: Thread.sleep(100) - blocking call
					// sleep(100);
				}
				else {
					strMethodDetails = mapSaveWb.strMethodDetails;
				}
			}
		}
		//Set the cell theme first
		const mapSetCellThemeResults = excelSetSaveCellCustomTheme(mapCustomValues); // Reuse internal helper
		const boolSetCellTheme = StrNums.JComm_StringToBoolean(mapSetCellThemeResults.boolPassed);
		if (boolSetCellTheme === false) {
			boolPassed = false;
			strMethodDetails = mapSetCellThemeResults.strMethodDetails;
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! Unable to open the workbook from the file: "' + strFilePath + '"!!!';
	}
	//Output the method pass/fail and message
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * Sets the header column name for a cell in the first row.
 * @param assgWb The workbook object (placeholder for WorkbookType)
 * @param assgSheet The assigned sheet (placeholder for SheetType)
 * @param intColNum The column number to update.
 * @param strAssgValue The value to assign to the cell.
 * @param boolOutputDetails Output the method details? True/False (not used)
 * @param strReporterOutput The reporter output format (CSV, HTML, XHTML, DataSource) (not used)
 * @return mapResults
 */
function ExcelData_excelSetHeaderColName(assgWb, assgSheet, intColNum, strAssgValue, boolOutputDetails, strReporterOutput) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value"); // Not used
	let boolPassed = true;
	let strMethodDetails = '';
	let strSheetName = assgSheet ? assgSheet.sheetName : 'Unknown Sheet'; // Placeholder for SheetType.sheetName
	//Return the number of cells and determine if we need to create the cell
	let intColCnt;
	//First get the row and check if it exist
	let assgRow = assgSheet ? assgSheet.getRow(0) : null; // Header row is always row 0 (Placeholder for SheetType.getRow)
	if (assgRow === null) {
		assgRow = assgSheet.createRow(0); // Create row if it doesn't exist (Placeholder for SheetType.createRow)
	}
	let assgCell = assgRow ? assgRow.getCell(intColNum) : null; // Placeholder for RowType.getCell
	let strOldValue;
	let strNewValue;
	if (assgCell === null) {
		assgCell = assgRow.createCell(intColNum); // Placeholder for RowType.createCell
		//Check that the cell has been created
		intColCnt = assgRow.getLastCellNum(); // Placeholder for RowType.getLastCellNum()
		if (intColCnt < intColNum) { // This condition is likely intended to check if creation extended beyond the intended column
			boolPassed = false; // Incorrect check. `lastCellNum` is 1-based, `intColNum` is 0-based.
			strMethodDetails = 'FAILED!!! Attempted to add column number: ' + intColNum + ' to sheet "' + strSheetName + '" but failed final column count is: ' + intColCnt + ' which is less then the assigned column!!!';
		}
	}
	if (boolPassed === true) {
		//Update the value and continue
		assgCell = assgRow.getCell(intColNum); // Re-get the cell to ensure it's the one we'll set
		strOldValue = StrNums.JComm_HandleNoData(excelGetCellValue(assgCell)); // Use internal helper
		assgCell.setCellValue(strAssgValue); // Placeholder
		//Verify the value is set.
		strNewValue = StrNums.JComm_HandleNoData(excelGetCellValue(assgCell)); // Use internal helper
		if (strNewValue === strAssgValue) {
			strMethodDetails = 'Updated column number: ' + intColNum + ' to sheet "' + strSheetName + '" to the value of: "' + strAssgValue + '".';
		} else {
			boolPassed = false;
			strMethodDetails = 'FAILED!!! Attempted to set the column number: ' + intColNum + ' in sheet "' + strSheetName + '" to the value of: "' + strAssgValue + '" but failed!!! Current value is: "' + strNewValue + '"!!!';
		}
	}
	//Output the method pass/fail and message
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * Open the File, write/update the data and Output the file. NOTE!!! We will overwrite the existing values and can write to columns without a name.
 * @param strFilePath The full path for the file.
 * @param strSheetName The sheet name to be assigned.
 * @param strColValues The values to assign for the cells in the row
 * @param strColValueDelimiter
 * @param intRowNumber The row number to update
 * @return Map
 */
function ExcelData_excelSetRowData(strFilePath, strSheetName, strColValues, strColValueDelimiter, intRowNumber, boolOutputDetails, boolStopOnFailError, strReporterOutput) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value"); // Not used
	let boolPassed = true;
	let strMethodDetails = '';

	if (fs.existsSync(strFilePath)) {
		try {
			//Create a file inputstream
			const inputStream = excelOpenExcelFileInputStream(strFilePath); // Internal helper
			//Open the WB
			let assgWb = excelOpenWBFromInputStream(inputStream); // Internal helper
			if (assgWb !== null) {
				//Return the specified sheet
				const mapGetSheet = excelGetSheetByName(assgWb, strSheetName); // Reuse internal helper
				let assgWs = mapGetSheet.objWbSheet; // Placeholder for SheetType
				if (mapGetSheet.boolPassed && assgWs !== null) { // Check that GetSheet was successful
					//Check for the header row and create if not present
					let rowSelected = assgWs.getRow(intRowNumber); // Placeholder for SheetType.getRow
					if (rowSelected === null) {
						//Create the row if not present
						rowSelected = assgWs.createRow(intRowNumber); // Placeholder for SheetType.createRow
					}
					//Create the array for the column names and make sure the value is not empty
					// Original: List<String> lstColValues = new ArrayList<String>(Arrays.asList(strColValues.split(strColValueDelimiter)))
					const lstColValues = strColValues.split(strColValueDelimiter);
					const intLstCnt = lstColValues.length;
					let strAryColName = '';
					let strFinalColNames = ''; // Not used
					let intRowColCnt = 0; // Not used

					if (intLstCnt > 0) {
						let assgCell = null; // Placeholder for CellType
						//Create the column Style
						let createHelper = undefined; // Placeholder
						let custStyle = undefined; // Placeholder for CellStyleType
						// Original: CreationHelper createHelper = assgWb.getCreationHelper(); CellStyle custStyle = assgWb.createCellStyle();
						if (assgWb && typeof assgWb.getCreationHelper === 'function' && typeof assgWb.createCellStyle === 'function') {
							createHelper = assgWb.getCreationHelper(); // Placeholder
							custStyle = assgWb.createCellStyle(); // Placeholder
						}

						//Build loop to return the values
						for (let loopIteration = 0; loopIteration < intLstCnt; loopIteration++) {
							assgCell = null;
							strAryColName = lstColValues[loopIteration];
							assgCell = rowSelected.getCell(loopIteration); // Placeholder
							if (assgCell === null) {
								//Create the cell
								assgCell = rowSelected.createCell(loopIteration); // Placeholder
							}
							//TODO add style for the cell to include the text, borders and color
							if (custStyle) {
								if (typeof custStyle.setHidden === 'function') { custStyle.setHidden(false); } // Placeholder
								//Set the column style
								//assgWs.setDefaultColumnStyle(loopIteration, colStyle) // Placeholder
							}
							//Assign the value
							assgCell.setCellValue(strAryColName); // Placeholder
						}
						//Close the inputstream always
						excelCloseExcelFileInputStream(inputStream); // Internal helper
						//Open the outputstream
						const outputStream = excelOpenExcelFileOutputStream(strFilePath); // Internal helper
						//Write the WB
						excelWriteWBDataToFile(assgWb, outputStream); // Internal helper
						//Close the outputStream
						excelCloseExcelFileOutputStream(outputStream); // Internal helper
						strMethodDetails = 'Updated the data for row: ' + intRowNumber + ' for ' + intLstCnt + ' column(s) in data file: "' + strFilePath + '".';
					} else {
						boolPassed = false;
						strMethodDetails = 'FAILED!!! Unable to add data. The strColValues did not contain any items.';
					}
				} else {
					boolPassed = false;
					strMethodDetails = 'FAILED!!! THe sheet: "' + strSheetName + '" was not found in the workbook from the file: "' + strFilePath + '"!!!';
				}
			} else {
				boolPassed = false;
				strMethodDetails = 'FAILED!!! Unable to open the workbook from the file: "' + strFilePath + '"!!!';
			}
		} catch (e) {
			// Original: catch (FileNotFoundException e), catch (IOException e)
			boolPassed = false;
			// Use specific error type detection if needed for more granular messages
			strMethodDetails = 'FAILED!!! File or IO exception see stack trace: ' + (e instanceof Error ? e.stack : String(e));
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!!The file: ' + strFilePath + ' DOES NOT EXIST';
	}
	//Output the method pass/fail and message
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

//Add formating for Cell, Range, Should we add graphs or charts?
/**
 * Sets the column data format.
 * @param assgSheet The assigned sheet (placeholder for SheetType)
 * @param intColNum The column number.
 * @param strAssgValue The format string.
 * @param boolOutputDetails (not used)
 * @param strReporterOutput (not used)
 */
function ExcelData_excelSetColDataFormat(assgSheet, intColNum, strAssgValue, boolOutputDetails, strReporterOutput) {
	let styColumn = null; // Placeholder for CellStyleType
	// Original: CellStyle styColumn;
	// Original: styColumn.setDataFormat(BuiltinFormats.getBuiltinFormat(strAssgValue))
	// Original: assgSheet.setDefaultColumnStyle(intColNum, styColumn)
	// Depends on Excel library's API for styles and data formats.
	// In exceljs, you might set cell.numFmt or a column's style.
	// Placeholder logic:
	if (assgSheet && typeof assgSheet.getWorkbook === 'function') { // Assuming getWorkbook exists to create styles
		const wb = assgSheet.getWorkbook();
		if (wb && typeof wb.createCellStyle === 'function') {
			styColumn = wb.createCellStyle();
			// Placeholder: styColumn.dataFormat = BuiltinFormats.getBuiltinFormat(strAssgValue);
			// Placeholder: assgSheet.setDefaultColumnStyle(intColNum, styColumn);
		}
	}
}

/**
 * Sets and saves the column width.
 * @param strFilePath The full path for the file.
 * @param strSheetName The sheet name.
 * @param intColNumber The column number.
 * @param intColWidth The desired width.
 * @return Map
 */
function ExcelData_excelSetSaveColWidth(strFilePath, strSheetName, intColNumber, intColWidth, boolOutputDetails, boolStopOnFailError, strReporterOutput) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value"); // Not used
	let boolPassed = true;
	let strMethodDetails = '';

	if (fs.existsSync(strFilePath)) {
		try {
			//Create a file inputstream
			const inputStream = excelOpenExcelFileInputStream(strFilePath); // Internal helper
			//Open the WB
			let assgWb = excelOpenWBFromInputStream(inputStream); // Internal helper
			if (assgWb !== null) {
				//Return the specified sheet
				const mapGetSheet = excelGetSheetByName(assgWb, strSheetName); // Reuse internal helper
				let assgWs = mapGetSheet.objWbSheet; // Placeholder for SheetType
				if (mapGetSheet.boolPassed && assgWs !== null) {
					//Check for the header row and create if not present
					let rowHeader = assgWs.getRow(0); // Placeholder for SheetType.getRow
					if (rowHeader !== null) {
						let assgCell = rowHeader.getCell(intColNumber); // Placeholder for RowType.getCell
						if (assgCell !== null) {
							//Set the column width
							assgWs.setColumnWidth(intColNumber, intColWidth); // Placeholder for SheetType.setColumnWidth
							//Close the inputstream always
							excelCloseExcelFileInputStream(inputStream); // Internal helper
							//Open the outputstream
							const outputStream = excelOpenExcelFileOutputStream(strFilePath); // Internal helper
							//Write the WB
							excelWriteWBDataToFile(assgWb, outputStream); // Internal helper
							//Close the outputStream
							excelCloseExcelFileOutputStream(outputStream); // Internal helper
							strMethodDetails = 'Updated the data file: "' + strFilePath + '" sheet "' + strSheetName + '" column: ' + intColNumber + ' width.';
						} else {
							boolPassed = false;
							strMethodDetails = 'FAILED!!! the workbook: "' + strFilePath + '" sheet: "' + strSheetName + '" header row does NOT HAVE COLUMN: ' + intColNumber + '!!!';
						}
					} else {
						boolPassed = false;
						strMethodDetails = 'FAILED!!! the workbook: "' + strFilePath + '" sheet: "' + strSheetName + '" DOES NOT have a HEADER ROW!!!';
					}
				} else {
					boolPassed = false;
					strMethodDetails = 'FAILED!!! THe sheet: "' + strSheetName + '" was not found in the workbook from the file: "' + strFilePath + '"!!!';
				}
			} else {
				boolPassed = false;
				strMethodDetails = 'FAILED!!! Unable to open the workbook from the file: "' + strFilePath + '"!!!';
			}
		} catch (e) {
			boolPassed = false;
			// Use specific error type detection if needed for more granular messages
			strMethodDetails = 'FAILED!!! File or IO exception see stack trace: ' + (e instanceof Error ? e.stack : String(e));
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!!The file: ' + strFilePath + ' DOES NOT EXIST';
	}
	//Output the method pass/fail and message
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * Create the font object by setting the font name, color, bold, italics and underline
 * @param assgWb The assigned Workbook (placeholder for WorkbookType)
 * @param assgCellStyle The cell style to apply font to (placeholder for CellStyleType)
 * @param strFontName The font name specified
 * @param strFontColor The font color specfied (e.g., IndexedColors name as string)
 * @param intFontHtInPts The font height in points
 * @param boolIsBold Bold font?
 * @param boolIsItalic Italic font?
 * @return assgCellStyle The updated cell style
 * Created: MM/dd/YYYY
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelCreateCellFont(assgWb, assgCellStyle, strFontName, strFontColor, intFontHtInPts, boolIsBold, boolIsItalic) {
	const gblNull = GVars.GblNull("Value");
	const gblSkip = GVars.GblSkip("Value");
	const gblNA = GVars.GblNotApplicable("Value");
	let assgFont = undefined; // Placeholder for FontType

	// Original: Font assgFont = assgWb.createFont()
	if (assgWb && typeof assgWb.createFont === 'function') {
		assgFont = assgWb.createFont(); // Placeholder for WorkbookType.createFont
	}

	if (assgFont) { // Proceed only if font object was created
		// Original: CellStyle assgCellStyle = assgWb.createCellStyle() (removed from function)
		if (strFontName !== gblNull) {
			assgFont.setFontName(strFontName); // Placeholder for FontType.setFontName
		}
		if (strFontColor !== gblNull) {
			// Original: assgFont.setColor(IndexedColors.valueOf(strFontColor).getIndex())
			// Needs mapping of 'IndexedColors' string names to actual color values
			// Placeholder: assgFont.setColor(mapIndexedColor(strFontColor));
		}
		//Set the font height
		if (intFontHtInPts > 0) {
			assgFont.setFontHeightInPoints(intFontHtInPts); // Placeholder for FontType.setFontHeightInPoints
		}
		if (boolIsBold === true) {
			assgFont.setBold(true); // Placeholder for FontType.setBold
			//assgFont.setBoldweight(Font.BOLDWEIGHT_BOLD) Katalon 6.3.4 (old POI API)
		} else {
			assgFont.setBold(false);
			//assgFont.setBoldweight(Font.BOLDWEIGHT_NORMAL) Katalon 6.3.4 (old POI API)
		}
		//Italics is either set or not
		assgFont.setItalic(boolIsItalic); // Placeholder for FontType.setItalic

		if (assgCellStyle && typeof assgCellStyle.setFont === 'function') {
			assgCellStyle.setFont(assgFont); // Placeholder for CellStyleType.setFont
		}
	}
	return assgCellStyle;
}

/**--------------------------------  excelCreateCellFontRGBColor --------------------------------------------
 * Create the font object by setting the font name, color in RGB, bold, italics and underline
 * @param assgWb The assigned Workbook (placeholder for XSSFWorkbookType)
 * @param assgCellStyle The cell style to apply font to (placeholder for CellStyleType)
 * @param strFontName The font name specified
 * @param strRGBAColor The font color specified (e.g., "rgba(R,G,B,A)")
 * @param intFontHtInPts The font height in points
 * @param boolIsBold Bold font?
 * @param boolIsItalic Italic font?
 * @return assgCellStyle The updated cell style
 * Created: MM/dd/YYYY
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelCreateCellFontRGBColor(assgWb, assgCellStyle, strFontName, strRGBAColor, intFontHtInPts, boolIsBold, boolIsItalic) {
	const gblNull = GVars.GblNull("Value");
	const gblSkip = GVars.GblSkip("Value");
	const gblNA = GVars.GblNotApplicable("Value");
	let assgFont = undefined; // Placeholder for XSSFFontType (or generic FontType)

	// Original: XSSFFont assgFont = assgWb.createFont()
	if (assgWb && typeof assgWb.createFont === 'function') {
		assgFont = assgWb.createFont(); // Placeholder
	}

	if (assgFont) {
		if (strFontName !== gblNull) {
			assgFont.setFontName(strFontName); // Placeholder
		}
		// Parsing RGB from string: "rgba(R,G,B,A)" -> "(R,G,B,A)" -> [R,G,B,A]
		let strTempColor = StrNums.JComm_GetRightTextInString(strRGBAColor, "(");
		strTempColor = StrNums.JComm_GetLeftTextInString(strTempColor, ")");
		const lstRGBA = StrNums.JComm_SplitString(strTempColor, ','); // This splits into string array

		const intRed = StrNums.JComm_StringToInteger(lstRGBA[0]);
		const intGreen = StrNums.JComm_StringToInteger(lstRGBA[1]);
		const intBlue = StrNums.JComm_StringToInteger(lstRGBA[2]);
		// Original: IndexedColorMap colorMap = assgWb.getStylesSource().getIndexedColors()
		// Original: XSSFColor color = new XSSFColor(new java.awt.Color(intRed,intGreen,intBlue), colorMap)
		// Original: assgFont.setColor(color)
		// This is complex POI-specific color mapping. In a Node.js library, you might set font color directly with RGB.
		// Placeholder:
		// const colorObject = createColorObject(intRed, intGreen, intBlue); // Custom helper for color object
		// assgFont.setColor(colorObject);
		// Or in exceljs, it's often `color: { argb: 'FF' + R.toString(16) + G.toString(16) + B.toString(16) }` or similar.

		//Set the font height
		if (intFontHtInPts > 0) {
			assgFont.setFontHeightInPoints(intFontHtInPts); // Placeholder
		}
		if (boolIsBold === true) {
			assgFont.setBold(true); // Placeholder
		} else {
			assgFont.setBold(false);
		}
		//Italics is either set or not
		assgFont.setItalic(boolIsItalic); // Placeholder

		if (assgCellStyle && typeof assgCellStyle.setFont === 'function') {
			assgCellStyle.setFont(assgFont); // Placeholder
		}
	}
	return assgCellStyle;
}

/**
 * Creates the alignment style
 * @param assgCellStyle The cell style assigned to the current cell (placeholder for CellStyleType)
 * @param strHorzAlign The Horizontal alignment to apply. Left, Right, Center, Justify
 * @param strVertAlign The Vertical alignment to apply. Top, Bottom, Center, Justify
 * @return assgCellStyle The alignment style created
 * Created: MM/dd/YYYY
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelCreateCellAllignment(assgCellStyle, strHorzAlign, strVertAlign) {
	const gblNull = GVars.GblNull("Value");
	const gblSkip = GVars.GblSkip("Value");
	if (!assgCellStyle) return assgCellStyle; // Ensure CellStyle object exists before trying to modify it

	//Set the horizontal alignment
	if (strHorzAlign !== gblNull && strHorzAlign !== gblSkip && typeof assgCellStyle.setAlignment === 'function') { // Placeholder
		switch (strHorzAlign) {
			case 'Left':
				// assgCellStyle.setAlignment(HorizontalAlignment.LEFT) // POI-style enum
				// Placeholder mapping to a string:
				assgCellStyle.setAlignment('left');
				break;
			case 'Right':
				assgCellStyle.setAlignment('right');
				break;
			case 'Center':
				assgCellStyle.setAlignment('center');
				break;
			case 'Justify':
				assgCellStyle.setAlignment('justify');
				break;
			default:
				assgCellStyle.setAlignment('general');
		}
	}
	//Set the vertical alignment
	if (strVertAlign !== gblNull && strVertAlign !== gblSkip && typeof assgCellStyle.setVerticalAlignment === 'function') { // Placeholder
		switch (strVertAlign) {
			case 'Top':
				assgCellStyle.setVerticalAlignment('top');
				break;
			case 'Bottom':
				assgCellStyle.setVerticalAlignment('bottom');
				break;
			case 'Center':
				assgCellStyle.setVerticalAlignment('center');
				break;
			case 'Justify':
				assgCellStyle.setVerticalAlignment('justify');
				break;
		}
	}
	return assgCellStyle;
}

/**
 * Set the border style and color
 * @param assgCellStyle The cell style assigned to the current cell (placeholder for CellStyleType)
 * @param strLeftStyle The Left border style (e.g., "THICK")
 * @param strLeftColor The Left border color (e.g., "BLACK")
 * ... (similar params for Top, Right, Bottom)
 * @return assgCellStyle The alignment style created
 * Created: MM/dd/YYYY
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelCreateCellBorders(assgCellStyle, strLeftStyle, strLeftColor, strTopStyle, strTopColor, strRightStyle, strRightColor, strBottomStyle, strBottomColor) {
	const gblNull = GVars.GblNull("Value");
	const gblSkip = GVars.GblSkip("Value");
	if (!assgCellStyle) return assgCellStyle;

	//Left border
	if (strLeftStyle !== gblNull && strLeftStyle !== gblSkip && typeof assgCellStyle.setBorderLeft === 'function') {
		// Original: assgCellStyle.setBorderLeft(BorderStyle.valueOf(strLeftStyle))
		assgCellStyle.setBorderLeft(strLeftStyle.toLowerCase()); // Placeholder: map string to enum or direct string
	}
	if (strLeftColor !== gblNull && strLeftColor !== gblSkip && typeof assgCellStyle.setLeftBorderColor === 'function') {
		// Original: assgCellStyle.setLeftBorderColor(IndexedColors.valueOf(strLeftColor).getIndex())
		// Placeholder: assgCellStyle.setLeftBorderColor(mapIndexedColor(strLeftColor));
	}
	//Top border
	if (strTopStyle !== gblNull && strTopStyle !== gblSkip && typeof assgCellStyle.setBorderTop === 'function') {
		assgCellStyle.setBorderTop(strTopStyle.toLowerCase()); // Placeholder
	}
	if (strTopColor !== gblNull && strTopColor !== gblSkip && typeof assgCellStyle.setTopBorderColor === 'function') {
		// Placeholder: assgCellStyle.setTopBorderColor(mapIndexedColor(strTopColor));
	}
	//Right border
	if (strRightStyle !== gblNull && strRightStyle !== gblSkip && typeof assgCellStyle.setBorderRight === 'function') {
		assgCellStyle.setBorderRight(strRightStyle.toLowerCase()); // Placeholder
	}
	if (strRightColor !== gblNull && strRightColor !== gblSkip && typeof assgCellStyle.setRightBorderColor === 'function') {
		// Placeholder: assgCellStyle.setRightBorderColor(mapIndexedColor(strRightColor));
	}
	//Bottom border
	if (strBottomStyle !== gblNull && strBottomStyle !== gblSkip && typeof assgCellStyle.setBorderBottom === 'function') {
		assgCellStyle.setBorderBottom(strBottomStyle.toLowerCase()); // Placeholder
	}
	if (strBottomColor !== gblNull && strBottomColor !== gblSkip && typeof assgCellStyle.setBottomBorderColor === 'function') {
		// Placeholder: assgCellStyle.setBottomBorderColor(mapIndexedColor(strBottomColor));
	}
	return assgCellStyle;
}

/**
 * Create the cell fill for background color (RGB) and fill pattern
 * @param WB The workbook that contains the cell used to get the indexed colors (placeholder for XSSFWorkbookType)
 * @param assgCellStyle The cell style assigned to the current cell (placeholder for CellStyleType)
 * @param intRed The red setting for the RGB color
 * @param intGreen The green setting for the RGB color
 * @param intBlue The blue setting for the RGB color
 * @param strFillPattern The fill pattern to be used for the cell (e.g., "SOLID_FOREGROUND")
 * @return assgCellStyle The updated cell style
 * Created: 04/28/2022
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelCreateCellFill(WB, assgCellStyle, intRed, intGreen, intBlue, strFillPattern) {
	const gblNull = GVars.GblNull("Value");
	const gblSkip = GVars.GblSkip("Value");
	if (!assgCellStyle) return assgCellStyle;

	// Original: IndexedColorMap colorMap = WB.getStylesSource().getIndexedColors()
	// Original: XSSFColor color = new XSSFColor(new java.awt.Color(intRed,intGreen,intBlue), colorMap)
	// Original: assgCellStyle.setFillForegroundColor(color)
	// This requires detailed mapping to your Excel library's color and fill APIs.
	// Placeholder:
	// const colorObject = { red: intRed, green: intGreen, blue: intBlue }; // Or suitable color object
	// if (typeof assgCellStyle.setFillForegroundColor === 'function') {
	//	 assgCellStyle.setFillForegroundColor(colorObject);
	// }

	if (strFillPattern !== gblNull && strFillPattern !== gblSkip && typeof assgCellStyle.setFillPattern === 'function') {
		// Original: assgCellStyle.setFillPattern(FillPatternType.valueOf(strFillPattern))
		assgCellStyle.setFillPattern(strFillPattern.toLowerCase().replace(/_/g, '')); // Placeholder
	}
	return assgCellStyle;
}

/**
 * Create the XSSF color (RGB)
 * @param WB The workbook that contains the cell used to get the indexed colors (placeholder for XSSFWorkbookType)
 * @param intRed The red setting for the RGB color
 * @param intGreen The green setting for the RGB color
 * @param intBlue The blue setting for the RGB color
 * @return color The XSSF Color object (placeholder for XSSFColorType)
 * Created: 005/05/2022
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelCreateXSSFColor(WB, intRed, intGreen, intBlue) {
	const gblNull = GVars.GblNull("Value");
	const gblSkip = GVars.GblSkip("Value");

	// Original: IndexedColorMap colorMap = WB.getStylesSource().getIndexedColors()
	// Original: XSSFColor color = new XSSFColor(new java.awt.Color(intRed,intGreen,intBlue), colorMap)
	// Similar to `excelCreateCellFill`, this is a deep POI integration.
	// In Node.js, you'd define a color object or string based on your library's API.
	// Placeholder returning a simple object:
	return { red: intRed, green: intGreen, blue: intBlue }; // Placeholder
}

/**
 * Set the cell formating to the specified theme for the row
 * @param intRowNumber The row number assigned (0-based)
 * @param strRowTheme The theme to apply (e.g., "OutputDataHdrStd")
 * @return Map
 */
function ExcelData_excelSetRowCellFormatByTheme(intRowNumber, strRowTheme) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value"); // Not used
	let boolPassed = true;
	let strMethodDetails = '';
	let intColCnt;

	//Return the test case wb object
	let wb = undefined; // Placeholder for WorkbookType
	wb = TCObj.getObjWorkbook(); // Assumes TCObj has this method
	let objSh = TCObj.getObjSheetTCActiveOutputSheet(); // Assumes TCObj has this method
	let strSheetName;
	if (objSh && objSh.sheetName) {
		strSheetName = objSh.sheetName; // Placeholder for SheetType.sheetName
	} else {
		strSheetName = 'Unknown Sheet';
	}
	const strFilePath = TCObj.getStrTCInputFilePath(); // Assumes TCObj has this method

	if (objSh !== null) {
		//Check for the header row
		// Original: int intInputRowCnt = objSh.getLastRowNum() // Seems unused
		let assgRow = objSh.getRow(intRowNumber); // Placeholder for SheetType.getRow
		if (assgRow === null) {
			boolPassed = false;
			strMethodDetails = 'FAILED!!! THe sheet: "' + strSheetName + '" in the workbook from the file: "' + strFilePath + '" DOES NOT HAVE A ROW OBJECT for ROW NUMBER: ' + intRowNumber + '!!!';
		}
		else {
			//Return the number of cells
			intColCnt = assgRow.lastCellNum; // Placeholder for RowType.lastCellNum
			if (intColCnt > 0) {
				let boolCellThemeSet; // Not used
				let strTmpResults = ''; // Used to concatenate failures
				//Format the cells
				let mapSetCellThemeResults = {};
				let boolSetCellTheme;
				for (let loopIteration = 0; loopIteration < intColCnt; loopIteration++) {
					mapSetCellThemeResults = excelSetSaveCellTheme(intRowNumber, loopIteration, strRowTheme); // Reuse internal helper
					boolSetCellTheme = StrNums.JComm_StringToBoolean(mapSetCellThemeResults.boolPassed);
					if (boolSetCellTheme === false) {
						boolPassed = false;
						strTmpResults = strTmpResults + 'FAILED!!! Unable to set the cell theme for row: ' + intRowNumber + ' and column: ' + loopIteration + '!!!' + gblLineFeed;
					}
				}
				if (boolPassed === true) {
					strMethodDetails = (strMethodDetails || '') + gblLineFeed + 'Applied the theme: "' + strRowTheme + '" to row: ' + intRowNumber + ' for the ' + intColCnt + ' column(s).';
				} else {
					strMethodDetails = 'FAILED!!! UNABLE TO COMPLETE SETTING THE THEME ASSIGNED. See details below: ' + gblLineFeed + strTmpResults;
				}
			} else {
				boolPassed = false;
				strMethodDetails = 'FAILED!!! THe sheet: "' + strSheetName + '" in the workbook from the file: "' + strFilePath + '" Row Number: ' + intRowNumber + ' DOES NOT HAVE ANY CELLS to FORMAT!!!';
			}
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! THe sheet: "' + strSheetName + '" was not found in the workbook from the file: "' + strFilePath + '"!!!';
	}
	//Output the method pass/fail and message
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * Set the specified cell to the assigned cell style class and save the output file to the theme
 * @param intRowNum The row number. Zero based which is the header row
 * @param intColNum The column number assigned
 * @param strCellTheme The theme to use for the cell.
 * @return Map
 */
function ExcelData_excelSetSaveCellTheme(intRowNum, intColNum, strCellTheme) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value"); // Not used
	let boolPassed = true;
	let strMethodDetails = '';
	let intRowCnt;
	let intColCnt;
	let boolSetTheme = false;

	//Return the test case wb object
	let wb = undefined; // Placeholder for WorkbookType (e.g., XSSFWorkbook equivalent)
	wb = TCObj.getObjWorkbook(); // Assumes TCObj has this method
	let objSh = TCObj.getObjSheetTCActiveOutputSheet(); // Assumes TCObj has this method
	let strSheetName;
	if (objSh && objSh.sheetName) {
		strSheetName = objSh.sheetName; // Placeholder for SheetType.sheetName
	} else {
		strSheetName = 'Unknown Sheet';
	}
	const strFilePath = TCObj.getStrTCInputFilePath(); // Assumes TCObj has this method

	if (objSh !== null) {
		//Check if the row exist and select
		intRowCnt = objSh.getLastRowNum(); // Placeholder for SheetType.getLastRowNum()
		if (intRowCnt >= intRowNum) {
			let assgRow = objSh.getRow(intRowNum); // Placeholder for RowType.getRow()
			//Check if the column exist and select
			intColCnt = assgRow ? assgRow.lastCellNum : -1; // Placeholder for RowType.lastCellNum
			if (intColCnt >= intColNum) {
				let assgCell = assgRow.getCell(intColNum); // Placeholder for RowType.getCell()
				//TODO Should we check if the cell is null
				let strCellValue = excelGetCellValue(assgCell); // Using internal helper to get string value
				let assgCellStyle = undefined; // Placeholder for CellStyleType
				// Original: CellStyle assgCellStyle = wb.createCellStyle()
				if (wb && typeof wb.createCellStyle === 'function') {
					assgCellStyle = wb.createCellStyle(); // Placeholder
				}

				if (assgCellStyle && assgCell) { // Ensure style and cell objects are valid before applying
					//Select the cellstyle to set
					switch (strCellTheme) {
						case 'OutputDataHdrStd':
							//Set the cell alignment
							assgCellStyle = excelCreateCellAllignment(assgCellStyle, "Center", "Bottom");
							//Set the cell font
							// Original: (this.excelCreateCellFont(wb, assgCellStyle, "Time New Roman", "WHITE", 12, true, true)) - uses wb, not XSSFWorkbook
							assgCellStyle = excelCreateCellFont(wb, assgCellStyle, "Time New Roman", "WHITE", 12, true, true);
							//Set the borders Style, Color, Each Border
							assgCellStyle = excelCreateCellBorders(assgCellStyle, "THICK", "BLACK", "THICK", "BLACK", "THICK", "BLACK", "THICK", "BLACK");
							//Set the fill
							//Medium Blue 70%
							let intRed = 102;
							let intGreen = 181;
							let intBlue = 255;
							// Original: (this.excelCreateCellFill(wb, assgCellStyle, intRed, intGreen, intBlue, "SOLID_FOREGROUND")) - uses wb, not XSSFWorkbook
							assgCellStyle = excelCreateCellFill(wb, assgCellStyle, intRed, intGreen, intBlue, "SOLID_FOREGROUND");
							assgCell.setCellStyle(assgCellStyle); // apply final style to cell
							boolSetTheme = true;
							break;
						case 'OutputDataLeftJustified':
							assgCellStyle = excelCreateCellAllignment(assgCellStyle, "Left", "Bottom");
							assgCellStyle = excelCreateCellFont(wb, assgCellStyle, "Time New Roman", "BLACK", 12, false, false);
							assgCellStyle = excelCreateCellBorders(assgCellStyle, "MEDIUM", "BLACK", "MEDIUM", "BLACK", "MEDIUM", "BLACK", "MEDIUM", "BLACK");
							intRed = 255; intGreen = 255; intBlue = 255;
							assgCellStyle = excelCreateCellFill(wb, assgCellStyle, intRed, intGreen, intBlue, "NO_FILL");
							assgCell.setCellStyle(assgCellStyle);
							boolSetTheme = true;
							break;
						case 'OutputDataCentered':
							assgCellStyle = excelCreateCellAllignment(assgCellStyle, "Center", "Bottom");
							assgCellStyle = excelCreateCellFont(wb, assgCellStyle, "Time New Roman", "BLACK", 12, false, false);
							assgCellStyle = excelCreateCellBorders(assgCellStyle, "MEDIUM", "BLACK", "MEDIUM", "BLACK", "MEDIUM", "BLACK", "MEDIUM", "BLACK");
							intRed = 255; intGreen = 255; intBlue = 255;
							assgCellStyle = excelCreateCellFill(wb, assgCellStyle, intRed, intGreen, intBlue, "NO_FILL");
							assgCell.setCellStyle(assgCellStyle);
							boolSetTheme = true;
							break;
						case 'VerifyDataHdrStd':
							assgCellStyle = excelCreateCellAllignment(assgCellStyle, "Center", "Bottom");
							assgCellStyle = excelCreateCellFont(wb, assgCellStyle, "Time New Roman", "BLACK", 12, true, true);
							assgCellStyle = excelCreateCellBorders(assgCellStyle, "THICK", "BLACK", "THICK", "BLACK", "THICK", "BLACK", "THICK", "BLACK");
							intRed = 255; intGreen = 179; intBlue = 102;
							assgCellStyle = excelCreateCellFill(wb, assgCellStyle, intRed, intGreen, intBlue, "SOLID_FOREGROUND");
							assgCell.setCellStyle(assgCellStyle);
							boolSetTheme = true;
							break;
						case 'TestRunHdrStd':
							assgCellStyle = excelCreateCellAllignment(assgCellStyle, "Center", "Bottom");
							assgCellStyle = excelCreateCellFont(wb, assgCellStyle, "Time New Roman", "WHITE", 12, true, true);
							assgCellStyle = excelCreateCellBorders(assgCellStyle, "THICK", "BLACK", "THICK", "BLACK", "THICK", "BLACK", "THICK", "BLACK");
							intRed = 92; intGreen = 214; intBlue = 92;
							assgCellStyle = excelCreateCellFill(wb, assgCellStyle, intRed, intGreen, intBlue, "SOLID_FOREGROUND");
							assgCell.setCellStyle(assgCellStyle);
							boolSetTheme = true;
							break;
						case 'TestRunDataStd':
							assgCellStyle = excelCreateCellAllignment(assgCellStyle, "Center", "Bottom");
							assgCellStyle = excelCreateCellFont(wb, assgCellStyle, "Time New Roman", "BLACK", 12, false, false);
							assgCellStyle = excelCreateCellBorders(assgCellStyle, "MEDIUM", "BLACK", "MEDIUM", "BLACK", "MEDIUM", "BLACK", "MEDIUM", "BLACK");
							intRed = 255; intGreen = 255; intBlue = 255;
							assgCellStyle = excelCreateCellFill(wb, assgCellStyle, intRed, intGreen, intBlue, "NO_FILL");
							assgCell.setCellStyle(assgCellStyle);
							boolSetTheme = true;
							break;
						case 'TestRunDataSkipIgnore':
							assgCellStyle = excelCreateCellAllignment(assgCellStyle, "Center", "Bottom");
							assgCellStyle = excelCreateCellFont(wb, assgCellStyle, "Time New Roman", "BLACK", 12, false, false);
							assgCellStyle = excelCreateCellBorders(assgCellStyle, "MEDIUM", "BLACK", "MEDIUM", "BLACK", "MEDIUM", "BLACK", "MEDIUM", "BLACK");
							intRed = 206; intGreen = 167; intBlue = 128;
							assgCellStyle = excelCreateCellFill(wb, assgCellStyle, intRed, intGreen, intBlue, "NO_FILL");
							assgCell.setCellStyle(assgCellStyle);
							boolSetTheme = true;
							break;
						case 'TestRunPassStd':
							assgCellStyle = excelCreateCellAllignment(assgCellStyle, "Center", "Bottom");
							assgCellStyle = excelCreateCellFont(wb, assgCellStyle, "Time New Roman", "WHITE", 12, true, false);
							assgCellStyle = excelCreateCellBorders(assgCellStyle, "MEDIUM", "BLACK", "MEDIUM", "BLACK", "MEDIUM", "BLACK", "MEDIUM", "BLACK");
							intRed = 0; intGreen = 163; intBlue = 0;
							assgCellStyle = excelCreateCellFill(wb, assgCellStyle, intRed, intGreen, intBlue, "SOLID_FOREGROUND");
							assgCell.setCellStyle(assgCellStyle);
							boolSetTheme = true;
							break;
						case 'TestRunFailStd':
							assgCellStyle = excelCreateCellAllignment(assgCellStyle, "Left", "Bottom");
							assgCellStyle = excelCreateCellFont(wb, assgCellStyle, "Time New Roman", "BLACK", 12, true, false);
							assgCellStyle = excelCreateCellBorders(assgCellStyle, "MEDIUM", "BLACK", "MEDIUM", "BLACK", "MEDIUM", "BLACK", "MEDIUM", "BLACK");
							intRed = 255; intGreen = 0; intBlue = 0;
							assgCellStyle = excelCreateCellFill(wb, assgCellStyle, intRed, intGreen, intBlue, "SOLID_FOREGROUND");
							assgCell.setCellStyle(assgCellStyle);
							boolSetTheme = true;
							break;
						case 'TestRunWarnStd':
							assgCellStyle = excelCreateCellAllignment(assgCellStyle, "Center", "Bottom");
							assgCellStyle = excelCreateCellFont(wb, assgCellStyle, "Time New Roman", "BLACK", 12, true, false);
							assgCellStyle = excelCreateCellBorders(assgCellStyle, "MEDIUM", "BLACK", "MEDIUM", "BLACK", "MEDIUM", "BLACK", "MEDIUM", "BLACK");
							intRed = 255; intGreen = 219; intBlue = 77;
							assgCellStyle = excelCreateCellFill(wb, assgCellStyle, intRed, intGreen, intBlue, "SOLID_FOREGROUND");
							assgCell.setCellStyle(assgCellStyle);
							boolSetTheme = true;
							break;
					}
					if (boolSetTheme === true) {
						//Save the excel file
						const mapSaveWb = excelSaveOutputFile(wb); // Reuse internal helper
						boolPassed = StrNums.JComm_StringToBoolean(mapSaveWb.boolPassed);
						if (boolPassed === true) {
							TCObj.objWorkbook = wb;
							TCObj.objSheetTCActiveOutputSheet = objSh;
							strMethodDetails = 'The excel file: "' + strFilePath + '" sheet "' + strSheetName + '" row: ' + intRowNum + ' column: ' + intColNum + ' was updated with the theme: "' + strCellTheme + '".';
						}
						else {
							strMethodDetails = mapSaveWb.strMethodDetails;
						}
					}
				} else {
					boolPassed = false;
					strMethodDetails = 'FAILED!!! Cell or CellStyle objects were invalid or not created.';
				}
			} else {
				boolPassed = false;
				strMethodDetails = 'FAILED!!! The excel file: "' + strFilePath + '" sheet "' + strSheetName + '" row: ' + intRowNum + ' contains ' + intColCnt + ' column(s) which is less then the assigned column of: ' + intColNum + '!!!.';
			}
		} else {
			boolPassed = false;
			strMethodDetails = 'FAILED!!! The excel file: "' + strFilePath + '" sheet "' + strSheetName + '" contains ' + intRowCnt + ' row(s) which is less then the assigned row of: ' + intRowNum + '!!!';
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! THe sheet: "' + strSheetName + '" was not found in the workbook from the file: "' + strFilePath + '"!!!';
	}
	//Output the method pass/fail and message
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}

/**
 * Set the specified cell to the assigned cell style class (custom theme)
 * @param mapCellTheme Object containing theme properties (ExcelRow, ExcelCol, RGBAColor, RGBABkColor, Alignment)
 * @return Map
 */
function ExcelData_excelSetSaveCellCustomTheme(mapCellTheme) {
	const mapResults = {};
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblDelimiter = GVars.GblDelimiter("Value"); // Not used
	const intRowNum = mapCellTheme.ExcelRow;
	const intColNum = mapCellTheme.ExcelCol;
	const strRGBAColor = mapCellTheme.RGBAColor;
	const strRGBABkColor = mapCellTheme.RGBABkColor;
	const strAlignment = mapCellTheme.Alignment;
	let boolPassed = true;
	let strMethodDetails = '';
	let intRowCnt;
	let intColCnt;
	let boolSetTheme = false; // Original had boolSetTheme, but it's never set to true logically in this variant.

	//Return the test case wb object
	let wb = undefined; // Placeholder for WorkbookType
	wb = TCObj.getObjWorkbook(); // Assumes TCObj has this method
	let objSh = TCObj.getObjSheetTCActiveOutputSheet(); // Assumes TCObj has this method
	let strSheetName;
	if (objSh && objSh.sheetName) {
		strSheetName = objSh.sheetName; // Placeholder for SheetType.sheetName
	} else {
		strSheetName = 'Unknown Sheet';
	}
	const strFilePath = TCObj.getStrTCInputFilePath(); // Assumes TCObj has this method

	if (objSh !== null) {
		//Check if the row exist and select
		intRowCnt = objSh.getLastRowNum(); // Placeholder for SheetType.getLastRowNum()
		if (intRowCnt >= intRowNum) {
			let assgRow = objSh.getRow(intRowNum); // Placeholder for RowType.getRow()
			//Check if the column exist and select
			intColCnt = assgRow ? assgRow.lastCellNum : -1; // Placeholder for RowType.lastCellNum
			if (intColCnt >= intColNum) {
				let assgCell = assgRow.getCell(intColNum); // Placeholder for RowType.getCell()
				let assgCellStyle = undefined; // Placeholder for CellStyleType
				// Original: CellStyle assgCellStyle = wb.createCellStyle()
				if (wb && typeof wb.createCellStyle === 'function') {
					assgCellStyle = wb.createCellStyle(); // Placeholder
				}

				if (assgCellStyle && assgCell) {
					//Set the cell alignment
					assgCellStyle = excelCreateCellAllignment(assgCellStyle, strAlignment, "Bottom"); // Reuse internal helper
					//Set the cell font to the forefront color using custom font
					// Original: (this.excelCreateCellFontRGBColor(wb, assgCellStyle, "Time New Roman", strRGBAColor, 12, true, true))
					assgCellStyle = excelCreateCellFontRGBColor(wb, assgCellStyle, "Time New Roman", strRGBAColor, 12, true, true);
					//Set the borders Style, Color, Each Border
					assgCellStyle = excelCreateCellBorders(assgCellStyle, "THICK", "BLACK", "THICK", "BLACK", "THICK", "BLACK", "THICK", "BLACK");
					//Set the fill
					//Update the values based on the background color as the fill. NOTE will work with just comma separated RGB as well
					let strTempColor = StrNums.JComm_GetRightTextInString(strRGBABkColor, "(");
					strTempColor = StrNums.JComm_GetLeftTextInString(strTempColor, ")");
					const lstRGBA = StrNums.JComm_SplitString(strTempColor, ',');
					const intRed = StrNums.JComm_StringToInteger(lstRGBA[0]);
					const intGreen = StrNums.JComm_StringToInteger(lstRGBA[1]);
					const intBlue = StrNums.JComm_StringToInteger(lstRGBA[2]);
					// Original:(this.excelCreateCellFill(wb, assgCellStyle, intRed, intGreen, intBlue, "SOLID_FOREGROUND"))
					assgCellStyle = excelCreateCellFill(wb, assgCellStyle, intRed, intGreen, intBlue, "SOLID_FOREGROUND");
					assgCell.setCellStyle(assgCellStyle); // apply final style to cell

					//Save the excel file
					const mapSaveWb = excelSaveOutputFile(wb); // Reuse internal helper
					boolPassed = StrNums.JComm_StringToBoolean(mapSaveWb.boolPassed);
					if (boolPassed === true) {
						TCObj.objWorkbook = wb;
						TCObj.objSheetTCActiveOutputSheet = objSh;
						strMethodDetails = 'The excel file: "' + strFilePath + '" sheet "' + strSheetName + '" row: ' + intRowNum + ' column: ' + intColNum + ' was updated with the custom theme".';
					}
					else {
						strMethodDetails = mapSaveWb.strMethodDetails;
					}
				} else {
					boolPassed = false;
					strMethodDetails = 'FAILED!!! Cell or CellStyle objects were invalid or not created.';
				}
			} else {
				boolPassed = false;
				strMethodDetails = 'FAILED!!! The excel file: "' + strFilePath + '" sheet "' + strSheetName + '" row: ' + intRowNum + ' contains ' + intColCnt + ' column(s) which is less then the assigned column of: ' + intColNum + '!!!.';
			}
		} else {
			boolPassed = false;
			strMethodDetails = 'FAILED!!! The excel file: "' + strFilePath + '" sheet "' + strSheetName + '" contains ' + intRowCnt + ' row(s) which is less then the assigned row of: ' + intRowNum + '!!!';
		}
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!! THe sheet: "' + strSheetName + '" was not found in the workbook from the file: "' + strFilePath + '"!!!';
	}
	//Output the method pass/fail and message
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	return mapResults;
}
	
/**
 * ------------------------------------------excelMapSheetRowValues------------------------------------------
 * Open the assigned Spreadsheet and using the column names and row/cell values create a map contain the ColName/Value pair
 * @params assgSheet The input sheet object (placeholder for SheetType)
 * @params intInputRowNumber The row number for the current test execution.
 * @return mapResults which contains the pass/fail, details, mapExcelData
 * Created: 10/11/2022
 * Author: PGKanaris
 * Last Edited:
 * Last Edited By:
 * Edit Comments: (Include email, date and details)
 */
function ExcelData_excelMapSheetRowValues(assgSheet, intInputRowNumber) {
	//Define module variables
	const mapResults = {};
	let strMethodDetails; // Will be assigned later
	let boolPassed = true;
	const mapExcelData = {}; // Initialize with empty object
	//Globals
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblNull = GVars.GblNull("Value");
	const gblSkip = GVars.GblSkip("Value");
	const boolDoDebug = TSExecParams.getBoolDoDebug();
	//Create excel variables
	// let sht = null; // Original: sht is declared but assignassgSheet is used
	// let intRowCnt, intColCnt; // Not used
	// let strCellValue; // Not used
	// let objCell; // Not used

	const mapGetColNames = excelGetHdrColNames(assgSheet); // Reuse internal helper
	const boolGetColNamesPassed = StrNums.JComm_StringToBoolean(mapGetColNames.boolPassed);
	const strGetColNameResults = mapGetColNames.strMethodDetails;
	const strGetColValues = mapGetColNames.ColValues;
	if (boolGetColNamesPassed === true) {
		//Return the values to a map
		const arryValues = strGetColValues.split('|'); // Split by literal '|'
		const lenArray = arryValues.length;
		//Create a loop and assign values
		let strTempColName;
		// let boolGetCellValuePassed; // Not used
		// let strGetCellValueResults; // Not used
		let strTempInputValue;
		let mapGetCellValue = {};
		for (let loopArray = 0; loopArray < lenArray; loopArray++) {
			mapGetCellValue = {}; // Re-initialize for each iteration
			//Return the colName
			strTempColName = arryValues[loopArray];
			//Get the input value from the row for the column
			mapGetCellValue = excelGetCellValueByRowNumColName(assgSheet, intInputRowNumber, strTempColName); // Reuse internal helper
			// boolGetCellValuePassed = StrNums.JComm_StringToBoolean(mapGetCellValue.boolPassed); // Not used
			// strGetCellValueResults = mapGetCellValue.strMethodDetails; // Not used
			strTempInputValue = mapGetCellValue.CellValue;
			if (boolDoDebug === true) {
				console.log("Column: " + strTempColName + " CellValue: " + strTempInputValue);
			}
			//Add to the map
			mapExcelData[strTempColName] = strTempInputValue; // Using bracket notation for dynamic keys
		}
		strMethodDetails = "Output " + lenArray + " column(s) of values for row: " + intInputRowNumber + ".";
	}
	else {
		strMethodDetails = strGetColNameResults;
		boolPassed = false; // If column names not retrieved, it's a failure
	}
	//Output the method pass/fail and message
	//Update the map
	mapResults.boolPassed = boolPassed.toString();
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.mapExcelData = mapExcelData;
	return mapResults;
}