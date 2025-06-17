// Describe the library purpose
/**
 * The library contains system test environment methods
 */

SeSGlobalObject("TestEnvironment");

// Add Imports
// Note: org.apache.commons.lang3.exception.ExceptionUtils is not directly used in this snippet,
// but would represent a general utility to handle exceptions in TypeScript if needed elsewhere.

// Add Jaggaer Libs (assuming these will also be converted to TypeScript modules)
// They are kept here as they are imported in the original Groovy, even if not directly used
// in the provided method.
//import * as TCObj from './TestObjects'; // Represents resources.common.TestObjects
//import * as GVars from './GlobalVariables'; // Represents resources.common.GlobalVariables
//import * as TSExecParams from './TestCaseExecParams'; // Represents resources.common.TestCaseExecParams
//import * as StrNums from './StringsAndNumbers'; // Represents resources.common.StringsAndNumbers
// Note: 'resources.common.TestEnvironment' is the current module, so it's not re-imported.

/**
 * Re-implementation of TestEnvironment class as a collection of functions.
 * Public static methods are converted to exported functions.
 */

function TestEnvironment_getHostOS() {
	//Return the test system OS
	let strOS;
	// Original: String strSystemOS = System.getProperty('os.name')
	// In Node.js, process.platform gives the OS platform identifier.
	// 'darwin' for macOS, 'win32' for Windows.
	const systemPlatform = process.platform;

	//Generate the path for the specific system OS
	//Check if Mac or Window system based on SystemOS contains
	if (systemPlatform === 'darwin') {
		strOS = 'Mac';
	}
	if (systemPlatform === 'win32') {
		strOS = 'Windows';
	}
	// If the platform is neither Mac nor Windows, strOS will remain undefined,
	// mimicking the behavior of the original Groovy if none of the 'if' conditions were met.
	return strOS;
}