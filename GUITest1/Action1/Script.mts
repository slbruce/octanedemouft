SystemUtil.Run "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe","","C:\Program Files (x86)\Google\Chrome\Application",""
Browser("Browser").Navigate "https://www.google.com/search?q=websiote&rlz=1C1CHBF_enIE794IE794&oq=websiote&aqs=chrome..69i57j0l7.31652j0j7&sourceid=chrome&ie=UTF-8"
Reporter.ReportEvent micDone, "Test Completed", "Test has finished"
