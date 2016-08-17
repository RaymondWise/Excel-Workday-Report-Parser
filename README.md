#Excel-Workday-Report-Parser

MIT License<br>
Copyright 2016 Raymond Wise @ [Github Source](https://github.com/RaymondWise/Excel-Weekly-Meal-Plan-Shopping-List-Creator) 

[![Code Review](http://www.zomis.net/codereview/shield/?qid=125698)](http://codereview.stackexchange.com/q/125698/75587)

2nd round [![Code Review](http://www.zomis.net/codereview/shield/?qid=125698)](http://codereview.stackexchange.com/q/137854/75587)

Language: VBA

[Workday](workday.com) ERP's generated reports can be exported to excel as `.xlsx` files. Due to the configuration of the tables within Workday, multiple lines of information are generated in the same cell.

For instance, you have data structures like this -

Security Group | Group Members
----------------|--------------
Local Admin | Adam <br> Bob <br> Alice
Power User |Jane 
DBA | Becky
Network Admin | John <br> Henrich <br> Ishmael
Standard | Eve
InfoSec|

When this report is exported from Workday, it retains its table format regardless of the number of characters per cell. 

Ignoring that it's possible to exceed the character limit of a cell, the issue becomes cells with multiple data points separated by non-visible Line Feed (LF) characters. 

If you want to try to Text-To-Columns it on the LF you will overwrite data to the right and end up with LF cells if you don't treat concecutive delimiters as one. 

You could also do a Find&Replace on the LF, but that doesn't help with the parsing - it just makes it easier to see the terrible data structure. 

`ParseWorkdayColumnVertically()` code should yeild you this structure - 

Security Group | Group Members
---------------|-----------------
Local Admin| Adam
Local Admin| Bob
Local Admin| Alice
Power User| Jane
DBA | Becky
Network Admin| John
Network Admin| Henrich
Network Admin| Ishmael
Standard| Eve
InfoSec|

Now you can easily manipulate the data for analysis!

There is also `ParseWorkdayColumnHorizontally` that will parse your data like `Text To Columns` - 


Security Group | Member 1 | Member 2 | Member 3
----------------|---------|----------|---------
Local Admin | Adam | Bob | Alice
Power User |Jane 
DBA | Becky
Network Admin | John | Henrich | Ishmael
Standard | Eve
InfoSec|

This project is an attempt to document how I've tried to overcome the limitations of exporting certain Workday reports into excel. I've found myself writing and rewriting code that's unmanagable, but the aim here is for it to be easily maintained. It is free for your use under the [MIT license](https://opensource.org/licenses/MIT).

##How to use this
Within excel you must open the Visual Basic Editor (VBE) to import the `.bas` file. On Windows you can open the VBE by pressing <kbd>Alt</kbd> + <kbd>F11</kbd>. Once open, you can import the file by pressing <kbd>Ctrl</kbd> + <kbd>M</kbd>, browse to the downloaded file `WorkdayLineFeedMgr.bas` and click Open. Once imported, on the left side in the Project Explorer you will have a module called *WorkdayLineFeedMgr* that contains this code.

Once imported, on Windows in excel you can press <kbd>Alt</kbd> + <kbd>F8</kbd> to run a macro. You can select either *ParseIntoColumns* or *ParseIntoRows*.

###Looking to contribute?
Please feel free to make a fork or raise an issue on my fork. I'm sure some data sets are more complicated than others and modifications may be needed.


The MIT License (MIT)
Copyright (c) <2016> <Raymond W Wise>

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
