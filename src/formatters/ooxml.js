/*global CSSLint*/
CSSLint.addFormatter({
    //format information
    id: "ooxml",
    name: "Open Office XML format", 

    /** 
     * Return content to be printed before all file results.
     * @return {String} to prepend before all results
     */
    startFormat: function() {
        return '<?xml version="1.0" encoding="UTF-8" ?>' +
                       '<?mso-application progid="Excel.Sheet"?>' +
                       '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:x2="http://schemas.microsoft.com/office/excel/2003/xml" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:html="http://www.w3.org/TR/REC-html40" xmlns:c="urn:schemas-microsoft-com:office:component:spreadsheet">' +
                       '<OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">' +
                       '<Colors>' +
                       '<Color>' +
                       '<Index>3</Index>' +
                       '<RGB>#c0c0c0</RGB>' +
                       '</Color>' +
                       '<Color>' +
                       '<Index>4</Index>' +
                       '<RGB>#ff0000</RGB>' +
                       '</Color>' +
                       '</Colors>' +
                       '</OfficeDocumentSettings>' +
                       '<ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">' +
                       '<WindowHeight>9000</WindowHeight>' +
                       '<WindowWidth>13860</WindowWidth>' +
                       '<WindowTopX>240</WindowTopX>' +
                       '<WindowTopY>75</WindowTopY>' +
                       '<ProtectStructure>False</ProtectStructure>' +
                       '<ProtectWindows>False</ProtectWindows>' +
                       '</ExcelWorkbook>' +
                       '<Styles>' +
                       '<Style ss:ID="Default" ss:Name="Default">' +
                       '<Font ss:FontName="Verdana" />' +
                       '</Style>' +
                       '<Style ss:ID="Result" ss:Name="Result">' +
                       '<Font ss:Bold="1" ss:FontName="Verdana" ss:Italic="1" ss:Underline="Single" />' +
                       '</Style>' +
                       '<Style ss:ID="Result2" ss:Name="Result2">' +
                       '<Font ss:Bold="1" ss:FontName="Verdana" ss:Italic="1" ss:Underline="Single" />' +
                       '<NumberFormat ss:Format="Currency" />' +
                       '</Style>' +
                       '<Style ss:ID="Heading" ss:Name="Heading">' +
                       '<Font ss:Bold="1" ss:FontName="Verdana" ss:Italic="1" ss:Size="16" />' +
                       '</Style>' +
                       '<Style ss:ID="Heading1" ss:Name="Heading1">' +
                       '<Alignment ss:Rotate="90" />' +
                       '<Font ss:Bold="1" ss:FontName="Verdana" ss:Italic="1" ss:Size="16" />' +
                       '</Style>' +
                       '<Style ss:ID="co1" />' +
                       '<Style ss:ID="co2" />' +
                       '<Style ss:ID="ta1" />' +
                       '<Style ss:ID="ta2" />' +
                       '<Style ss:ID="ta3" />' +
                       '<Style ss:ID="ce1">' +
                       '<Font ss:Bold="1" ss:FontName="Verdana" ss:Size="10" />' +
                       '</Style>' +
                       '<Style ss:ID="ce2">' +
                       '<NumberFormat ss:Format="[$$-409]#,##0.00"/>' +
                       '</Style>  <Style ss:ID="backed">' +
                       '<Font ss:FontName="Verdana" x:Family="Swiss" ss:Bold="1"/>' +
                       '<Interior ss:Color="#969696" ss:Pattern="Solid"/>' +
                       '</Style>' +
                       '<Style ss:ID="ce3">' +
                       '<Font ss:Bold="1" ss:FontName="Verdana" ss:Size="15" />' +
                       '</Style>' +
                       '<Style ss:ID="ce4">' +
                       '<Font ss:Bold="1" ss:FontName="Verdana" ss:Color="red" ss:Size="10" />' +
                       '</Style>' +
                       '<Style ss:ID="ce5">' +
                       '<Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>' +
                       '<Font ss:Bold="1" ss:FontName="Verdana"/>' +
                       '<Interior ss:Color="#91C489" ss:Pattern="Solid"/>' +
                       '</Style>' +
                       '<Style ss:ID="ce6">' +
                       '<Font ss:Bold="1" ss:FontName="Verdana"/>' +
                       '<Interior ss:Color="#91C489" ss:Pattern="Solid"/>' +
                       '</Style>' +
                       '<Style ss:ID="ce7">' +
                       '<Font ss:FontName="Verdana" ss:Bold="1" ss:Color="#FFFFFF" ss:Size="10"/>' +
                       '<Interior ss:Color="#E52828" ss:Pattern="Solid"/>' +
                       '</Style>' +
                       '<Style ss:ID="ce15">' +
                       '<Font ss:FontName="Verdana"/>' +
                       '<NumberFormat ss:Format="Percent"/>' +
                       '</Style>' +
                       '<Style ss:ID="ce9">' +
                       '<Font ss:Bold="1" ss:FontName="Verdana"/>' +
                       '<Interior ss:Color="#91C489" ss:Pattern="Solid"/>' +
                       '<NumberFormat ss:Format="Fixed"/>' +
                       '</Style>' +
                       '<Style ss:ID="ce50">' +
                       '<Font ss:FontName="Verdana" ss:Bold="1" ss:Size="10"/>' +
                       '<Interior ss:Color="#959595" ss:Pattern="Solid"/>' +
                       '</Style>' +
                       '<Style ss:ID="ce8">' +
                       '<Font ss:FontName="Verdana"/>' +
                       '<Interior ss:Color="#91C489"/>' +
                       '<NumberFormat ss:Format="Fixed"/>' +
                       '</Style>' +
                       '<Style ss:ID="ce55">' +
                       '<Font ss:FontName="Verdana" ss:Bold="1"/>' +
                       '<Interior ss:Color="#91C489" ss:Pattern="Solid"/>' +
                       '<Alignment ss:Horizontal="Left" ss:Vertical="Top"/>' +
                       '</Style>' +
                       '<Style ss:ID="s120">'+
                       '<Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="14" ss:Color="#FFFFFF"/>'+
                       '<Interior ss:Color="#FF0000" ss:Pattern="Solid"/>'+
                       '</Style>'+
                       '<Style ss:ID="s119">'+
                       '<Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="14" ss:Color="#FFFFFF"/>'+
                       '<Interior ss:Color="#FF0000" ss:Pattern="Solid"/>'+
                       '</Style>'+
                       '<Style ss:ID="s99">'+
                       '<Interior ss:Color="#F4F9B1" ss:Pattern="Solid"/>'+
                       '</Style>'+
                       '<Style ss:ID="s98">'+
                       '<Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="12"/>'+
                       '<Interior ss:Color="#F4F9B1" ss:Pattern="Solid"/>'+
                       '</Style>'+
                       '</Styles>' +
                       '<ss:Worksheet ss:Name="CSSLintResult"><Table ss:StyleID="ta1"><Column ss:AutoFitWidth="1" ss:Width="150" ss:StyleID="Default" /><Column ss:AutoFitWidth="1" ss:Width="150" ss:StyleID="Default" ss:Span="254" />';
    },

    /**
     * Return content to be printed after all file results.
     * @return {String} to append after all results
     */
    endFormat: function() {
        return '</Table><x:WorksheetOptions /></ss:Worksheet>' + '</Workbook>';
    },

    /**
     * Given CSS Lint results for a file, return output for this format.
     * @param results {Object} with error and warning messages
     * @param filename {String} relative file path
     * @param options {Object} (Optional) specifies special handling of output
     * @return {String} output for results
     */
    formatResults: function(results, filename, options) {


        // private utility functions

        function addRow(index,rowdata,sId) {
             //return '<Row ss:Index="' + index + '">' + rowdata + '</Row>';
             return '<Row ss:StyleID="'+sId+'">' + rowdata + '</Row>';
        }

        function addCell(cellData, sId) {
            return  '<Cell ss:StyleID="'+sId+'">' + cellData + '</Cell>';
        }

        function addRow1(index,rowdata) {
             //return '<Row ss:Index="' + index + '">' + rowdata + '</Row>';
             return '<Row>' + rowdata + '</Row>';
        }

        function addCell1(cellData) {
            return  '<Cell>' + cellData + '</Cell>';
        }

        function addData (ssType,data) {
            return  '<Data ss:Type="' + ssType + '">' + data + '</Data>';
        }
            
         var escapeSpecialCharacters = function(str) {
            /*if (!str || str.constructor !== String) {
                return "";
            }*/
            //return str.replace(/&/g, "&amp;").replace(/\"/g, "&quot;").replace(/'/g,"&apos;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); 

            return "<![CDATA[" + str + "]]>";
        };

        var messages = results.messages,
            output = "";
        options = options || {};

        if (messages.length === 0) {
            return options.quiet ? "" : "\n\ncsslint: No errors in " + filename + ".";
        }

        //output = "\n\ncsslint: There are " + messages.length  +  " problems in " + filename + ".";
        var pos = filename.lastIndexOf("/"),
            shortFilename = filename;

        if (pos === -1){
            pos = filename.lastIndexOf("\\");       
        }
        if (pos > -1){
            shortFilename = filename.substring(pos+1);
        }

        output = addRow(1, addCell(addData("String","File Name"),'s98') +
                 addCell(addData("String","Error Type"),'s98') +
                 addCell(addData("String","Line No"),'s98') +
                 addCell(addData("String","Column No"),'s98') +
                 addCell(addData("String","Browsers"),'s98') +
                 addCell(addData("String","Evidence"),'s98') +
                 addCell(addData("String","Description"),'s98') +
                 addCell(addData("String","Rule Name"),'s98'),'s99');

        CSSLint.Util.forEach(messages, function (message, i) {
          
          /*
            Message Data Structure
            
            message = {
              error_type
            }
          */
            var fileName = addData("String",shortFilename);
            var error_type = addData("String",escapeSpecialCharacters(message.type));
            var line = addData("String",message.line);
            var col = addData("String",message.col);
            var browsers = addData("String",message.rule.browsers);
            var info = addData("",escapeSpecialCharacters(message.message));
            var evidence = addData("String",escapeSpecialCharacters(message.evidence));
            var desc = addData("String",escapeSpecialCharacters(message.rule.desc));
            var name = addData("String",escapeSpecialCharacters(message.rule.name));
            var col_styleID="s98";
            var row_styleID="s99";
            //var cellWrapper = addCell(error_type)+addCell(line_number)+addCell(col_number)+addCell(browsers)+addCell(text_message)+addCell(evidence)+addCell(desc)+addCell(name);
            
          if("error"==message.type){
            col_styleID="s120";
            row_styleID="s119";
          }

            var cellWrapper =  addCell(fileName,col_styleID)+ addCell(error_type,col_styleID)+addCell(line,col_styleID)+addCell(col,col_styleID)+addCell(browsers,col_styleID)+addCell(evidence,col_styleID)+addCell(desc,col_styleID)+addCell(name,col_styleID);
          //var cellWrapper = addCell(error_type)+addCell(line)+addCell(col)+addCell(browsers)+addCell(evidence)+addCell(desc)+addCell(name);
            //var cellWrapper = addCell(info);//Ashok please check. Not working as a string, generated file not readable by Ms Excel
            output = output + addRow((i+1),cellWrapper,row_styleID);
            //output = output + addRow((i+1),cellWrapper);
        });
    
        return output;
    }
});
