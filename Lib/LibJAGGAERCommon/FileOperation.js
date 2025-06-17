// Describe the library purpose
/**
 * The library contains File operation methods
 */
SeSGlobalObject("FileOperation");

// Add Imports
// External Node.js built-in modules for file operations
//import fs from 'fs';	// For file system operations (e.g., fs.existsSync, fs.readFileSync)

// Add Jaggaer Libs (assuming these will also be converted to TypeScript modules)
//import * as GVars from './GlobalVariables'; // Represents resources.common.GlobalVariables

/**
 * Re-implementation of FileOperation class as a collection of functions.
 * Public static methods are converted to exported functions.
 */

function FileOperation_base64Encryption(strFileFullPath) {
	//Set global values
	const gblNull = GVars.GblNull("Value");
	const gblUndefined = GVars.GblUndefined("Value");
	const gblLineFeed = GVars.GblLineFeed("Value");
	const gblHTMLLineFeed = GVars.GblHTMLLineFeed("Value");

	//Create the result Map and variables
	// Groovy's Map mapResults = [:] maps directly to a JavaScript plain object {}
	const mapResults = {};
	let boolPassed = true;
	let strMethodDetails; // Will be assigned later
	let strFileEncData;   // Will be assigned later

	//Verify the file exist
	//Check if the file exist
	// Equivalent to new File(strFileFullPath) and tmpFile.exists()
	if (fs.existsSync(strFileFullPath)) {
		//Convert the file using base 64
		// Equivalent to byte[] fileContent = FileUtils.readFileToByteArray(new File(strFileFullPath));
		// Equivalent to strFileEncData = Base64.getEncoder().encodeToString(fileContent);

		// Read the file content as a Buffer
		const fileContentBuffer = fs.readFileSync(strFileFullPath);
		// Encode the Buffer to a Base64 string
		strFileEncData = fileContentBuffer.toString('base64');
	} else {
		boolPassed = false;
		strMethodDetails = 'FAILED!!!The file: ' + strFileFullPath + ' DOES NOT EXIST';
	}

	//Update the map
	// Using direct assignment for plain JavaScript objects instead of 'put' method
	mapResults.boolMethodPassed = boolPassed.toString(); // Convert boolean to string as in original
	mapResults.strMethodDetails = strMethodDetails;
	mapResults.strFileData = strFileEncData;

	return mapResults;
}