# Excel-Data-Transformation

Visual Basic remains the most dreaded developer language<sup>[1]</sup>, but it's still extremely useful for the many that use Excel Spreadsheets. A future translation of the code to Google App Script is not currently in the works, but could be a future development.

This repo is distrubuted under the MIT License and provides a set of user-defined functions and stored procedures for the transformation, manipulation, and management of data. Enjoy! :punch:

Current list of UDFs

| Function | Description |
| :--- | :--- |
| `TextTransform` | Cleans a string based on IgnoreSetting(s)|
| `RegExTester` | Test a string for a regex pattern |
| `LOOKUPALL` | Returns all values matching the lookup value to a single cell |
| `LOOKUPLIST` | Returns the first match for every lookup value to a single cell |
| `arrLOOKUPLIST` | Modifed LOOKUPLIST as an array formula; returns each value to it's own cell |

Current list of stored procedures

| Procedure| Description |
| :--- |:--- |
| `RearrangeColumns` | Rearranges columnar data to your specification |
| `CustomFilterBySelection` | Rearranges columnar data to your specification |

### Example Usage
In each example, an abbrevaited version of the of UDF is provided as a reference. Please see the code comments for parameter details.

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

##### LOOKUPLIST
```
LOOKUPLIST(lookup_value, table_array, col_index_num [, lookup_as_num = False,
    hlook = False, input_delim = ",", output_delim = ","])
```
![Sample usage image for LOOKUPLIST](/img/LOOKUPLIST.jpg)

##### RearrangeColumns
```
Call RearrangeColumns()
```
![Sample usage image for rearrangeColumns](/img/rearrangeColumns.gif)

##### CustomFilterBySelection
```
Call CustomFilterBySelection()
```
![Sample usage image for CustomFilterBySelection](/img/CustomFilterBySelection.gif)

[1]: http://stackoverflow.com/research/developer-survey-2016#technology-most-loved-dreaded-and-wanted