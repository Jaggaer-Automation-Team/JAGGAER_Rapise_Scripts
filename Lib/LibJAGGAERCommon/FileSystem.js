// Describe the library purpose
/**
 * The library contains File system methods
 */
SeSGlobalObject("FileSystem");

// Add Imports
// External Node.js built-in modules for file system operations
//import path from 'path'; // For path manipulation (e.g., path.sep, path.join, path.resolve)
//import fs from 'fs';    // For file system operations (e.g., fs.existsSync, fs.unlinkSync, fs.readdirSync, fs.statSync, fs.rmdirSync, fs.rmSync)

// Add Jaggaer Libs (assuming these will also be converted to TypeScript modules)
// We'll import them as modules containing functions/exports.
//import * as TCObj from './TestObjects'; // Represents resources.common.TestObjects
//import * as GVars from './GlobalVariables'; // Represents resources.common.GlobalVariables
//import * as TSExecParams from './TestCaseExecParams'; // Represents resources.common.TestCaseExecParams
//import * as StrNums from './StringsAndNumbers'; // Represents resources.common.StringsAndNumbers
//import * as TSEnv from './TestEnvironment'; // Represents resources.common.TestEnvironment

/**
 * Re-implementation of FileSystem class as a collection of functions.
 * Public static methods are converted to exported functions.
 */

function FileSystem_FileSystem_filePathConversion(strAssgPath) {
    const strFileSep = path.sep; // Equivalent to File.pathSeparator
    const strHostOS = TSEnv.getHostOS();
    let strOutputPath = '';
    const intStrLocation = StrNums.JComm_TextLocationInString(strAssgPath, ':'); // Assuming JComm_TextLocationInString exists in StrNums

    switch (strHostOS) {
        case 'Mac':
            //Process to change the path to Mac style
            strOutputPath = strAssgPath.replace(/\\/g, "/"); // Use /g for global replacement
            //remove C: if found
            if (intStrLocation >= 0) {
                strOutputPath = StrNums.JComm_GetRightTextInString(strOutputPath, ':'); // Assuming JComm_GetRightTextInString exists in StrNums
            }
            break;
        case 'Windows':
            //Process to change the path to Windows style
            strOutputPath = strAssgPath.replace(/\//g, "\\"); // Use /g for global replacement
            break;
    }
    return strOutputPath;
}

function FileSystem_fileSetTempDriveLetter() {
    // Original: File file = new File("C:" + GVars.GBLTCInputDirRoot("Value"))
    const cDrivePath = "C:" + GVars.GBLTCInputDirRoot("Value");
    let fileExist;

    // Equivalent to file.exists()
    fileExist = fs.existsSync(cDrivePath);
    if (fileExist === false) {
        const dDrivePath = "D:" + GVars.GBLTCInputDirRoot("Value");
        fileExist = fs.existsSync(dDrivePath);
        if (fileExist === true) {
            TCObj.setStrDriveLetter("D:");
        }
        else {
            TCObj.setStrDriveLetter("FAILED!!! DID NOT FIND C: " + GVars.GBLTCInputDirRoot("Value") + " or D: " + GVars.GBLTCInputDirRoot("Value") + " CREATE \\temp on your assigned drive!!!");
        }
    }
}

function FileSystem_fileSetDownloadDirPath() {
    const gblNull = GVars.GblNull("Value");
    let strDownloadDirPath, strDefaultDownloadDirStructure, strDefaultDownloadStructureSplitter;
    //Return the assigned directory structure and splitter
    //Return the drive letter
    //Create the path
    //Convert path to OS specific string
    //Return the absolute path
    return strDownloadDirPath;
}

//fileCreateExecutionDirectory
function FileSystem_fileCreateExecutionDirectory(boolIsInput, boolCreateTCDir) {
    const gblNull = GVars.GblNull("Value");
    //return the test suite and test case values
    // Original: String strDirectoryPath = this.filePathConversion(GVars.GBLOutputDirRoot("Value"))
    // Direct call within the same module (renamed from FileSystem class methods to functions)
    const strDirectoryPath = filePathConversion(GVars.GBLOutputDirRoot("Value"));
    let rootPath;
    //Return the Jira IDs
    const strTestCase = TSExecParams.strTCJiraID;
    const strTestSuite = TSExecParams.strTSJiraID;
    /*
    if ( strTestSuite != gblNull) {
        if ( boolIsInput == true) {
            strDirectoryPath = this.getFolderAbsolutePath(GVars.GBLTSInputDirRoot("Value")) + File.separator + this.filePathConversion(strTestSuite)
        } else {
            strDirectoryPath = strDirectoryPath + File.separator + DataStrings.dsGetRightText(gblTestSuite, '/')
        }
        if ( gblTestCase != gblNull && boolCreateTCDir == true) {
            //Added the test case directory so we can store the test cases snapshots as well as the test case specific input files.
            //Should we have a switch to create testcase subdirectory
            strDirectoryPath = strDirectoryPath + File.separator + DataStrings.dsGetRightText(strTestCase, '/')
        }
    }else if ( gblTestCase != gblNull) {
        if ( boolIsInput == true) {
            strDirectoryPath = this.getAbsolutePath(GlobalVariable.strGBLTCInputDirRoot) + File.separator + this.filePathConversion(gblTestCase)
        } else {
            strDirectoryPath = strDirectoryPath + File.separator + DataStrings.dsGetRightText(strTestCase, '/')
        }
    }
    //Add the file separator to the end
    strDirectoryPath = strDirectoryPath + File.separator
    if (gblTestSuite == gblNull && gblTestCase == gblNull) {
        strDirectoryPath = gblNull
    } else if (boolIsInput == false) {
        //Check host os
        if (TestEnvironment.testEnvGetHostOS() == 'Windows') {
            strDirectoryPath = 'C:' + strDirectoryPath
        }
    }
    */
    return strDirectoryPath;
}

/**
 * ------------------ getAbsolutePath -----------------------
    * Return the folders absolute path
    * @param strFolder the assigned folder
    *
    * @return strAbsPath the absolute path for the folder
    *
    * Created 06/22/2021
    * Author PGKanaris
    * Last Edited:
    * Last Edited By:
    * Edit Comments: (Include email, date and details)
 */
function FileSystem_getFolderAbsolutePath(strFolder) {
    const strAbsPath = GVars.GblNull("Value");
    // Equivalent to new File(strFolder).exists() and .absolutePath
    if (fs.existsSync(strFolder)) {
        strAbsPath = path.resolve(strFolder); // Get the absolute path using Node.js path module
    }
    return strAbsPath;
}

function FileSystem_deleteFiles(strFilePathList) {
    // Split by literal '|'. In JS, split with regex takes regex syntax.
    // If strFilePathList.split('\\|') was meant to split by a literal pipe, then we use '|'.
    // If it was meant to split by `\|` (backslash followed by pipe), then it would be /\\\|/
    const arryValues = strFilePathList.split('|');
    const lenArray = arryValues.length;
    //Create a loop and delete each file
    for (let loopArray = 0; loopArray < lenArray; loopArray++) {
        //Delete the files
        // Equivalent to FileUtils.getFile(arryValues[loopArray]) and FileUtils.forceDelete(tmpFile)
        try {
            // fs.unlinkSync deletes a file
            fs.unlinkSync(arryValues[loopArray]);
        } catch (error) {
            // Log the error if deletion fails. ExceptionUtils.getStackTrace replacement needs a logging library.
            // For simplicity, using console.error for now.
            console.error(`Error deleting file ${arryValues[loopArray]}:`, error);
        }
    }
}

function FileSystem_deleteDirectoryFiles(strDirPath) {
    // Equivalent to new File(strDirPath) and file.isDirectory()
    if (fs.existsSync(strDirPath) && fs.statSync(strDirPath).isDirectory()) {
        // Equivalent to file.list()
        const childFiles = fs.readdirSync(strDirPath); // Returns an array of file/directory names

        if (childFiles === null || childFiles.length === 0) {
            // Directory is empty or failed to read (readdirSync throws on failure)
            // Equivalent to FileUtils.forceDelete(file) for an empty directory
            try {
                fs.rmdirSync(strDirPath); // fs.rmdirSync removes an empty directory
            } catch (error) {
                console.error(`Error removing empty directory ${strDirPath}:`, error);
            }
        }
        else {
            // Directory contains files that must be deleted first
            for (const childFilePath of childFiles) {
                // recursive delete the files
                // Join path correctly to ensure OS-specific pathing
                const fullPath = path.join(strDirPath, childFilePath);
                deleteDirectoryFiles(fullPath); // Recursive call
            }
            // After deleting all children, the directory itself can be deleted.
            // This behavior mimics the manual recursive deletion logic in the original Groovy.
            try {
                fs.rmdirSync(strDirPath); // Now remove the (now empty) directory
            } catch (error) {
                console.error(`Error removing directory ${strDirPath} after emptying:`, error);
            }
        }
    }
    else {
        // Is a file and can be deleted
        // Equivalent to FileUtils.forceDelete(file) for a file
        try {
            fs.unlinkSync(strDirPath);
        } catch (error) {
            console.error(`Error deleting file ${strDirPath}:`, error);
        }
    }
}

//TODO add these methods
/*
 * getUserDownloadDirectory
 * getDirectoryFileCount
 * getLatestFileAdded
 * unzipFile
 */