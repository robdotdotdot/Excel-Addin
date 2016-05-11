# Excel-Data-Transformation

Visual Basic remains the most dreaded developer language[1], but it's still extremely useful for the many that use Excel Spreadsheets.

This repo provides a set of user-defined functions and stored procedures for the transformation, manipulation, and management of data. Enjoy!

Current list of UDFs:
* TextTransform - Cleans a string
* RegExTester - Test a string for a regex pattern


### Example Usage
In each example, the paramters of UDFs are provided as a reference. Please see code comments for details of each parameter

##### TextTransform
'''
Function TextTransform(t, Optional IgnoreCase, _
    Optional IgnoreSpace = False, Optional IgnoreSymbol = False, _
    Optional IgnoreNumber = False, Optional IgnoreQuote = False)
    ...
    ...
End Function
'''
![Sample usage image for TextTransform](/img/TextTransform.jpg)

##### RegExTester
'''
Function RegExTester(pattern As String, setence As String, _
    Optional outputTrueFalse = False, Optional MatchAll = False, _
    Optional delimiter = ",")
    ...
    ...
End Function
'''
![Sample usage image for RegExTester](/img/RegExTester.jpg)


[1]: http://stackoverflow.com/research/developer-survey-2016#technology-most-loved-dreaded-and-wanted