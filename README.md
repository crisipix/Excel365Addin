# Excel365Addin
Excel 365 Addin for Excel

First Install the Developer tools first
1. Install the Office Developer Tools for Visual Studio 2013 or 2015, if you haven’t already. 
These instructions also assume you have Office 2013 or later installed on your computer. 
2. Then This absolutely needs to be installed otherwise you get a prompt saying the necessary program hasn't been created. 
Visual Studio 2010 Tools for Office Runtime https://www.microsoft.com/en-us/download/confirmation.aspx?id=48217

How to: Install the Visual Studio Tools for Office Runtime Redistributable
https://msdn.microsoft.com/en-us/library/ms178739.aspx


	1. Install the Office Developer Tools for Visual Studio 2013 or 2015, if you haven’t already. 
    These instructions also assume you have Office 2013 or later installed on your computer. 
	2. Open Visual Studio and go to File > New > Project. Under Office/SharePoint, choose Office Add-in (or App for Office) 
    and then choose OK. 
	3. In the app creation wizard, choose Task Pane, choose Next, and then uncheck all the options except Excel. 
        Press F5 or the green Start button to launch the add-in. The add-in will be hosted locally on IIS, 
        and Excel will be opened with the add-in loaded. 



	Visual Studio builds the project and does the following:
	1.Creates a copy of the XML manifest file and adds it to ProjectName\Output directory. The host application consumes this copy when you start Visual Studio and debug the add-in.
	2.Creates a set of registry entries on your computer that enable the add-in to appear in the host application.
	3.Builds the web application project, and then deploys it to the local IIS web server (http://localhost). 
	
	Next, Visual Studio does the following:
	1.Modifies the SourceLocation element of the XML manifest file by replacing the ~remoteAppUrl token with the fully qualified address of the start page (for example, http://localhost/MyAgave.html).
	2.Starts the web application project in IIS Express.
	3.Opens the host application. 