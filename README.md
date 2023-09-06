# PrintNote - A OneNote add-in

#### NOTE: PrintNote only works properly when OneNote is run as administrator!

### First Time Setup:
- By default, Microsoft Print to PDF does not support custom paper sizes which prevents PrintNote from working properly. To enable custom paper sizes, run the script `CustomSize.exe` as administrator; this script is located in `CustomSize.zip`. The C# code for this script is available [here](CustomSize/Program.cs). Note that you may need to install the latest version of the Microsoft .NET Runtime to run `CustomSize.exe` successfully

### Install:
- To install PrintNote, run the `.msi` file corresponding to your [Office version](https://support.microsoft.com/en-us/office/about-office-what-version-of-office-am-i-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-us&rs=en-us&ad=us).
    - For Office 64-bit: `PrintNote.msi`
    - For Office 32-bit: `PrintNoteX86.msi`

### Usage:
- Before printing a page, press the *Set paper size* button under the *PrintNote* tab
![Image of above](images/read1.png)
- When printing the page, use the paper size named *PrintNote* for a clean one page document
![Image of above](images/read2.png)

### Future Plans:
- Find an alternative solution without the requirement of administrative permissions
- Add support for exporting groups of pages, sections and notebooks