# JiraToExcel
Importer of data from jira into excel. Single file excel template.

VBA script only, for easy "deployment". Easily configureable. Queries Jira directly, parses resulting XML (implements IVBSAXContentHandler). Fast.

###Setup (see image below):
* Download / copy to drive then open file in excel. Be sure to enable macros (click options in "Security Warning" bar) 
* Open worksheet/tab "Con": Enter (a) URL to your server, (b) username and (c) password.
* Open sample sheet "Sheet2", go to cell F1 (named jql_1). Enter your projectname or any other jira JQL you want to use.
* Optional: Review columns/fields in row 3 (and edit headers in row 4) as needed.
* Click Do Import button.

###Some points to code:
* Has tests - support for running integration tests if you want to develop further. You need to get test XML yourself (easy).
* Robust and pretty fast implementation for a VBA, uses/implements Microsoft.XML SAX parser
* Fairly OO code. Pro: Easily expandable to do new things. Con: Takes a little time to grok code design.

**If you've found this and run into a snag or have questions, raise issue or message me.**
###Instruction image
![instruction image](https://raw.githubusercontent.com/chipbite/JiraToExcel/master/JiraExcelImport_Instruction.png)
