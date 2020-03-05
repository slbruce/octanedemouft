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
Window("Calculator").WinButton("Button").Click @@ hightlight id_;_721876_;_script infofile_;_ZIP::ssf2.xml_;_
Window("Calculator").WinButton("Button_2").Click @@ hightlight id_;_1115006_;_script infofile_;_ZIP::ssf3.xml_;_
Window("Calculator").WinButton("Button").Click @@ hightlight id_;_721876_;_script infofile_;_ZIP::ssf4.xml_;_
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf8.xml_;_
Window("Calculator").WinButton("Button").Click @@ hightlight id_;_721876_;_script infofile_;_ZIP::ssf5.xml_;_
Window("Calculator").WinButton("Button_3").Click @@ hightlight id_;_656300_;_script infofile_;_ZIP::ssf6.xml_;_

Reporter.ReportEvent micPass, "Reporting options set from UFT test","Details"

