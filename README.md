# codelist-compare
This tool creates a comparison spreadsheet comparing two or more versions of a code list.

# Example

The following is an example with four different versions of a code list.

## Input

These are the four inputs.  Each is a tab 
in an XML Spreadsheet 2003 format Excel 
file.  The tabs are in the following order.

### The Code - v1
| Code | Description |
|--|--|
| a | able |
| b | baker |

### The Code - v2
| Code | Description |
|--|--|
| a | able |
| c | charley |
| g | gamma |

### The Code - v3
| Code | Description |
|--|--|
| a | able |
| b | bastte |
| g | gamma |

### The Code - v4
| Code | Description |
|--|--|
| a | actor |
| g | gelding |
| z | zebra |

## Output

Here is the output of the comparison 
as an XML Spreadsheet 2003 formatted 
Excel spreadsheet.

| The Code - v1 | The Code - v2 | The Code - v3 | The Code - v4 | Description | The Code - v2 Alternate Description | The Code - v3 Alternate Description | The Code - v4 Alternate Description |
|--|--|--|--|--|--|--|--|
| a | a | a | a | able |  |  | actor |
| b |  | b |  | baker |  | bastte |  |
|  | c |  |  | charley |  |  |  |
|  | g | g | g | gamma |  |  | gelding |
|  |  |  | z | zebra |  |  |  |

## License

codelist-compare is released under the [MIT License](LICENSE). 
