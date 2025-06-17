const fs = require('fs');
const path = require('path');
//import * as HtmlWebpackPlugin from 'html-webpack-plugin';
//import * as webpack from 'webpack';
import { enUS } from 'date-fns/locale';
const moment = require('moment'); // Import moment.js for date/time manipulation
const { format } = require('date-fns');
import { GlobalVariables } from "./GlobalVariables"
import { TCObj } from './TestObjects';

class ExtentLogging {
    static ExtRptCreateOutput(strTCReportName: string, strTCAssignee: string, strTCJiraID: string): any {

		let rptExtent: any
		let sparksReporter: any
		let logger: any
	
        let strTempTCReportName = strTCReportName.replace("|", " ");
        let boolPassed: boolean = true;
        let strMethodDetails: string = "";
        let strExtRptFolder =  "test-output" //TestCaseExecParams.strTestExtentRptFolder//TSExecParams.getStrTestExtentRptFolder();
		 // Create folder if not exisits
		 fs.mkdirSync(strExtRptFolder, { recursive: true });
        let mapResults: Map<string, string> = new Map();
	
		 //TODO Implement
		 //const pluginHtml = new HtmlWebpackPlugin({
         // template: path.resolve('./src/index.html'), // load a custom template
        // });
	
	
       
		 const  report ="//TEST";
			TCObj.setRptExtent(report)
	    TCObj.setLogTestExtent(strTCReportName )
	  TCObj.strExtRptFilePath = strTCReportName
	   
	
        //TODO: implement report create
		 

       // mapResults.set('boolMethodPassed', boolPassed.toString());
       // mapResults.set('strMethodDetails', strMethodDetails);
		 return "done";
    }
	 static ExtRptTestStepResults (objExtRptMod: any, mapStepInput: any): any {
       //TODO: Implement the ExtRptTestStepResults method logic here
       console.log("Extent logging step result");
	   return {}
    }

	//TODO: Implmenet the function for ExtentReport .  
	static ExtRptDoSnapShot(objExtRptMod: any, mapStepInput: any) : any {
		return null
	}
}