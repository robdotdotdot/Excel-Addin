# Excel-Data-Transformation

Visual Basic remains the most dreaded developer language<sup>[1], but it's still extremely useful for the many that use Excel Spreadsheets. A future translation of the code to Google App Script is not currently in the works, but could be a future development.

This repo is distrubuted under the MIT License and provides a set of user-defined functions and stored procedures for the transformation, manipulation, and management of data. Enjoy! :punch:

Current list of UDFs:
|Function|Description|
|:---|:---|
|TextTransform|Cleans a string|
|RegExTester|Test a string for a regex pattern|
|LOOKUPALL|Modified vlookup that can return matches to a single cell|

Current list of stored procedures:
|Procedure|Description|
|:---|:---|
|rearrangeColumns|Rearranges columnar data to your specification|

### Example Usage
In each example, an abbrevaited version of the of UDFs are provided as a reference. Please see code comments for details of each parameter.

##### TextTransform
```
TextTransform(t [, IgnoreCase, IgnoreSpace = False, IgnoreSymbol = False,
	IgnoreNumber = False, IgnoreQuote = False])
```
![Sample usage image for TextTransform](/img/TextTransform.jpg)

##### RegExTester
```
RegExTester(pattern, setence [, outputTrueFalse = False, MatchAll = False,
	delimiter = ","])
```
![Sample usage image for RegExTester](/img/RegExTester.jpg)

##### LOOKUPALL
```
LOOKUPALL(lookup_value, table_array, col_index_num [, delim = ","])
```
![Sample usage image for LOOKUPALL](/img/LOOKUPALL.jpg)

[1]: http://stackoverflow.com/research/developer-survey-2016#technology-most-loved-dreaded-and-wanted