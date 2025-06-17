// Put library code here
/** Original include
eval(File.IncludeOnce('/Lib/LibJAGGAERCommon/DateTime.js'));
eval(File.IncludeOnce('/Lib/LibJAGGAERCommon/ExcelData.js'));
eval(File.IncludeOnce('/Lib/LibJAGGAERCommon/FileOperation.js'));
eval(File.IncludeOnce('/Lib/LibJAGGAERCommon/FileSystem.js'));
eval(File.IncludeOnce('/Lib/LibJAGGAERCommon/GlobalVariables.js'));
eval(File.IncludeOnce('/Lib/LibJAGGAERCommon/StringsAndNumbers.js'));
eval(File.IncludeOnce('/Lib/LibJAGGAERCommon/TestCaseExecParams.js'));
eval(File.IncludeOnce('/Lib/LibJAGGAERCommon/TestEnvironment.js'));
eval(File.IncludeOnce('/Lib/LibJAGGAERCommon/TestExecReporting.js'));
*/
eval(g_helper.IncludeOnce('/Lib/LibJAGGAERCommon/DateTime.js'));
eval(g_helper.IncludeOnce('/Lib/LibJAGGAERCommon/ExcelData.js'));
eval(g_helper.IncludeOnce('/Lib/LibJAGGAERCommon/FileOperation.js'));
eval(g_helper.IncludeOnce('/Lib/LibJAGGAERCommon/FileSystem.js'));
eval(g_helper.IncludeOnce('/Lib/LibJAGGAERCommon/GlobalVariables.js'));
eval(g_helper.IncludeOnce('/Lib/LibJAGGAERCommon/StringsAndNumbers.js'));
eval(g_helper.IncludeOnce('/Lib/LibJAGGAERCommon/TestCaseExecParams.js'));
eval(g_helper.IncludeOnce('/Lib/LibJAGGAERCommon/TestEnvironment.js'));
eval(g_helper.IncludeOnce('/Lib/LibJAGGAERCommon/TestExecReporting.js'));

/** ---------------------------- RVL Functions ------------------------------*/
/*
	* RvlInfo
	* Dump all parameters passed into this function from RVL as ParamName: ParamValue pairs.
	*
	* return None
	* Created 10/15/2020
	* Author PGKanaris
	* Last Edited:
	* Last Edited By:
	* Edit Comments: (Include email, date and details)
	* https://www.inflectra.com/Support/KnowledgeBase/KB570.aspx
 */
function RvlInfo(/**string*/message)
{
	var msg = message;

	for (var lp in RVL.LastParams) {
		if (RVL.LastParams.hasOwnProperty(lp) && lp != "message") {
			msg+="<br/><b>"+lp+"</b>: "+RVL.LastParams[lp];
		}
	}
	
	Tester.Message(msg);
}


/**
 	* RvlVars
 	* Dump all local variables as VarName: Value and plus any additional parameters passed into this function.
 	*
	* return None
	* Created 10/15/2020
	* Author PGKanaris
	* Last Edited:
	* Last Edited By:
	* Edit Comments: (Include email, date and details)
	* https://www.inflectra.com/Support/KnowledgeBase/KB570.aspx
 */
function RvlVars(/**string*/message)
{
	var stack = RVL._current_rvl_execution_stack;
	var msg = message;
	
	var contextVars = stack.LocalContext.ContextVars;
	for (var lp in contextVars) {
		if(typeof contextVars[lp] == 'function') continue;
		msg+="<br/><b>"+lp+"</b>: "+contextVars[lp];
	}

	var contextVars = stack.GlobalContext.ContextVars;
	for (var lp in contextVars) {
		if(typeof contextVars[lp] == 'function') continue;
		msg+="<br/><b>"+lp+"</b>: "+contextVars[lp];
	}
	
	RvlInfo(msg);
}
