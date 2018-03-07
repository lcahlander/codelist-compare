(:
 Copyright (c) 2018. Loren Cahlander 
 
:)
xquery version "3.0";
(:~
    This takes an Excel Spreadsheet in XML Spreadsheet 2003 format that has
    n spreadsheets where there are two columns and they are code and description.
    Each spreadsheet has a title row.
    
    A comparison spreadsheet between the n number of sets of codes is generated.
    
    Output Column Name                                     Description
    code[n]                                         - If the code is in the nth code list, 
                                                      then this column is populated with the code value
    Description                                     - The first instance of the description of the code 
                                                      in the following order: first spreadsheet, second spreadsheet
    Atlernate Description (if different from first) - If the code is in the nth (skipping 1) code list and the 
                                                      description is different from the Description column, 
                                                      then populate with the second code description.

    The input to the process is an n tab spreadsheet and the output is the n + 1 tab spreadsheet with the comparison as the first spreadsheet.
 :)
 
declare namespace map = "http://www.w3.org/2005/xpath-functions/map";
declare namespace ss="urn:schemas-microsoft-com:office:Spreadsheet";
declare namespace o="urn:schemas-microsoft-com:office:office";
declare namespace x="urn:schemas-microsoft-com:office:excel";

import module namespace functx = "http://www.functx.com" at "http://www.xqueryfunctions.com/xq/functx-1.0-doc-2007-01.xq";

(:~ The name of the code versions that is being compared. :)
declare variable $code-title as xs:string external := 'The Code';

(:~ 
    How are the codes sorted, 
    'alpha' - plain ascending sort on the code (default) and 
    'numeric-alpha' where it is sorted by first the numeric value 
    in the code and second being the non-numeric characters in the code.  
:)
declare variable $order-type as xs:string external := 'alpha';

(:~
 :)
declare function local:sort-order1($code as xs:string) 
{
    switch ($order-type)
    case 'alpha' 
    return $code
    
    case 'numeric-alpha'
    return
      let $int-string := fn:replace($code, '[^0-9]', '')
      let $int-value := if (fn:string-length($int-string) eq 0) then 0 else xs:integer($int-string)
      return $int-value
    
    default return $code
};

(:~
 :)
declare function local:sort-order2($code as xs:string) 
{
    switch ($order-type)
    case 'alpha' 
    return $code
    
    case 'numeric-alpha'
    return
      fn:replace($code, '[0-9]', '')
    
    default return $code
};


(:~
 : Generate a spreadsheet cell with Top alignment and text wrap set to on.
 :
 : @param $text the text to be in the cell
 : @return a ss:Cell element with the text as the content of the cell.
 :)
declare function local:ss-cell($type as xs:string, $text as xs:string) as element(ss:Cell)
{
    element { 'ss:Cell' } { 
        attribute { 'ss:StyleID' } { "s62" },
        element { 'ss:Data' } { 
            attribute { 'ss:Type' } { $type }, 
            $text 
        } 
    }
};

let $sheets := .//*:Workbook/*:Worksheet

let $codes-map   := map:merge(
                        for $sheet at $index in $sheets
                        return map:entry($index, fn:subsequence($sheet/*:Table/*:Row, 2))
                    )  
                    
(: Get the sorted list of unique code values from the two lists of code values. :)
let $codes := for $code in fn:distinct-values(
                            map:for-each($codes-map, 
                                            function($k, $v){
                                            $v/*:Cell[1]/*:Data/text()
                                            }))
              let $order1 := local:sort-order1($code)
              let $order2 := local:sort-order2($code)
              order by $order1, $order2
              return $code

let $tab :=
    element { 'ss:Worksheet' } {
        attribute { 'ss:Name' } { fn:concat($code-title, ' - Comparison') },
        element { 'ss:Table' } {
            for $sheet in $sheets
            return element { 'ss:Column' } { attribute { 'ss:Width' } { '60' } },
            for $sheet in $sheets
            return element { 'ss:Column' } { attribute { 'ss:Width' } { '360' } },
            element { 'ss:Row' } {
                for $sheet in $sheets
                return
                local:ss-cell('String', $sheet/@*:Name),
                local:ss-cell('String', 'Description'),
                    for $sheet at $index in $sheets
                    let $rows := map:get($codes-map, $index)
                    return
                        if ($index = 1)
                        then ()
                        else local:ss-cell('String', fn:concat($sheet/@*:Name, ' Alternate Description (if different from first)'))
            },                    

            for $code in $codes
            let $descriptions := fn:distinct-values(
                                    map:for-each(
                                        $codes-map, 
                                        function($k, $v){
                                            if ($v[*:Cell[1]/*:Data = $code]) then fn:normalize-space($v[*:Cell[1]/*:Data = $code]/*:Cell[2]/*:Data/text()) else ()
                                        }
                                    )
                                 )
            let $first-description := fn:subsequence($descriptions, 1, 1)
            return
                element { 'ss:Row' } {
                    map:for-each(
                        $codes-map, 
                        function($k, $v){
                            local:ss-cell('String', if ($v[*:Cell[1]/*:Data = $code]) then $code else '')
                            
                        }
                    ),
                    local:ss-cell('String', $first-description),
                    for $sheet at $index in $sheets
                    let $cells := map:get($codes-map, $index)
                    let $cell := $cells[*:Cell[1]/*:Data = $code]
                    let $cell-description := if (fn:exists($cell)) then fn:normalize-space($cell/*:Cell[2]/*:Data/text()) else ()
                    return
                        if ($index = 1)
                        then ()
                        else local:ss-cell('String', if (fn:exists($cell) and ($first-description ne $cell-description)) then $cell-description else ''),
                    ()
                }                    
        }
    }

let $spreadsheet :=
    document {
        processing-instruction { 'mso-application' } { 'progid="Excel.Sheet"' },
        element { 'ss:Workbook' } {
            element { 'o:DocumentProperties' } {
                element { 'o:Title' } { $code-title },
                element { 'o:Created' } { fn:current-dateTime() },
                element { 'o:Company' } { 'Example.com' }
            },
            element { 'ss:Styles' } {
                element { 'ss:Style' } {
                    attribute { 'ss:ID' } { "s62" },
                    element { 'ss:Alignment' } {
                        attribute { 'ss:Vertical' } { "Top" },
                        attribute { 'ss:WrapText' } { "1" }
                    }
                }
            },
            $tab,
            for $sheet in $sheets
            return functx:remove-attributes-deep($sheet, 'ss:StyleID')
        }
    }

return $spreadsheet
