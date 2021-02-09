Dim App
Dim qtResultOpt

Set App = CreateObject("QuickTest.Application")
Set qtResultOpt=CreateObject("Quicktest.runresultsoptions")
qtResultOpt.ResultsLocation="C:\temp\results\"
Set qtAutoExportResultsOpts = app.Options.Run.AutoExportReportConfig
qtAutoExportResultsOpts.AutoExportResults = True ' Instruct UFT to automatically export the run results at the end of each run session
qtAutoExportResultsOpts.StepDetailsReport = True ' Instruct UFT to automatically export the step details part of the run results at the end of each run session
qtAutoExportResultsOpts.DataTableReport = True ' Instruct UFT to automatically export the data table part of the run results at the end of each run session
qtAutoExportResultsOpts.ScreenRecorderReport = True ' Instruct UFT to automatically export the screen recorder part of the run results at the end of each run session
qtAutoExportResultsOpts.ExportLocation = "C:\temp\HTML\" 'Instruct UFT to automatically export the run results to the Desktop at the end of each run session
qtAutoExportResultsOpts.ExportForFailedRunsOnly = false ' Instruct UFT to automatically export run results only for failed runs	
'
 @@ hightlight id_;_65696_;_script infofile_;_ZIP::ssf1.xml_;_
Window("Calculator").WinButton("Button").Click @@ hightlight id_;_459988_;_script infofile_;_ZIP::ssf10.xml_;_
Window("Calculator").WinButton("Button_2").Click @@ hightlight id_;_656070_;_script infofile_;_ZIP::ssf11.xml_;_
Window("Calculator").WinButton("Button").Click @@ hightlight id_;_1508056_;_script infofile_;_ZIP::ssf12.xml_;_
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf8.xml_;_
Window("Calculator").WinButton("Button").Click @@ hightlight id_;_1508056_;_script infofile_;_ZIP::ssf13.xml_;_
Window("Calculator").WinButton("Button_3").Click @@ hightlight id_;_1900972_;_script infofile_;_ZIP::ssf14.xml_;_

Reporter.ReportEvent micPass, "Reporting options set from UFT test","Details"

