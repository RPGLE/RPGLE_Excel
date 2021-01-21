**FREE
//--------------------------------------------------------------------------------------------------
// MIT License
// Copyright (c) 2020 Ghost +
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
//--------------------------------------------------------------------------------------------------
// EXCEL : Excel Spreadsheet procedures
//         This service program contains procedures related to creating spreadsheets.
//         To create a new spreadsheet, the following process must be followed:
//             - Create Workbook (Excel_createWorkbook)
//               - Create Worksheet (Excel_createWorsheet)
//                 - Set Column Widths - not mandatory (Excel_setColumnWidths)
//                 - Create Row (Excel_createRow)
//                   - Write Cell (Excel_writeCell)
//                 - Close Row (Excel_closeRow)
//               - Close Worksheet (Excel_createWorksheet)
//             - Close Workbook (Excel_closeWorkbook)
//          It is important to close each row before adding any new rows, and close off the previous
//          worksheet before adding new worksheets.
//          The available styles are stored within the EXCEL_P copybook. To add new styles, create
//          a new constant and ensure the style is defined within the Excel_createWorkbook.
//
//          The workbook is created within the IFS. The workbook name should contain the entire
//          path for creating the file.
//--------------------------------------------------------------------------------------------------
// Developer Date       Modification Description
// Ghost+    2020/07/07 Initial Coding
//--------------------------------------------------------------------------------------------------
ctl-opt option(*srcstmt: *nodebugio) datfmt(*iso) nomain;

/include QRPGLESRC,EXCEL_P
/include QRPGLESRC,STDIO_H

//--------------------------------------------------------------------------------------------------
// Procedure  : Excel_createWorkbook
// Purpose    : Creates a new file in the IFS for generating a formatted Excel Workbook.
// Returns    : success => Defaults to off, will fail if no workbook exists.
// Parameter/s: workbookName => Constant reference to the name of the file/workbook.
//--------------------------------------------------------------------------------------------------
dcl-proc Excel_createWorkbook export;
  dcl-pi *n ind;
    workbookName char(45) const;
  end-pi;

  dcl-s success ind inz(*off);
  dcl-s fileName char(50) inz;
  dcl-s assignedFile like(pFILE) inz;

  // Create our unique txt file in IFS
  fileName = %trim(workbookName) + '.xml';
  assignedFile = fopen(%trim(fileName) : 'w,crln=N');
  if assignedFile <> *null;

    // Write xml header
    fputs('<?xml version="1.0" encoding="UTF-8"?><?mso-application progid="Excel.Sheet"?>'
        : assignedFile);
    fputs('<ss:Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" ' : assignedFile);
    fputs('xmlns:x="urn:schemas-microsoft-com:office:excel" ' : assignedFile);
    fputs('xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" ' : assignedFile);
    fputs('xmlns:html="https://www.w3.org/TR/html401/">' : assignedFile);

    // Default Styles - can be customized or more styles added later...
    fputs('<ss:Styles>' : assignedFile);

    // Header 1 Style
    fputs('<ss:Style ss:ID="H1">' : assignedFile);
    fputs('<ss:Font ss:FontName="Calibri" ' : assignedFile);
    fputs('ss:Size="14" ss:Color="#222222" ss:Bold="1"/>' : assignedFile);
    fputs('<ss:Interior ss:Color="#CCCCCC" ss:Pattern="Solid"/>' : assignedFile);
    fputs('</ss:Style>' : assignedFile);

    // Header 2 Style
    fputs('<ss:Style ss:ID="H2">' : assignedFile);
    fputs('<ss:Font ss:FontName="Calibri" ' : assignedFile);
    fputs('ss:Size="14" ss:Color="#222222" ss:Bold="1" />' : assignedFile);
    fputs('<ss:Interior ss:Color="#CCCCCC" ss:Pattern="Solid"/>' : assignedFile);
    fputs('<ss:Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>' : assignedFile);
    fputs('</ss:Style>' : assignedFile);

    // Detail 1 Style
    fputs('<ss:Style ss:ID="D1">' : assignedFile);
    fputs('<ss:Font ss:FontName="Calibri" ' : assignedFile);
    fputs('ss:Size="12" ss:Color="#000000"/>' : assignedFile);
    fputs('</ss:Style>' : assignedFile);

    // Detail 2 Style
    fputs('<ss:Style ss:ID="D2">' : assignedFile);
    fputs('<ss:Font ss:FontName="Calibri" ' : assignedFile);
    fputs('ss:Size="12" ss:Color="#000000"/>' : assignedFile);
    fputs('<ss:Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>' : assignedFile);
    fputs('</ss:Style>' : assignedFile);

    // Money 1 Style
    fputs('<ss:Style ss:ID="M1">' : assignedFile);
    fputs('<ss:Font ss:FontName="Calibri" ' : assignedFile);
    fputs('ss:Size="12" ss:Color="#000000"/>' : assignedFile);
    fputs('<ss:NumberFormat ss:Format="Currency"/>' : assignedFile);
    fputs('</ss:Style>' : assignedFile);

    // Close Styles XML
    fputs('</ss:Styles>' : assignedFile);

    // Close the file as we are done.
    fclose(assignedFile);
    success = *on;
  endif;

  return success;
end-proc;

//--------------------------------------------------------------------------------------------------
// Procedure  : Excel_closeWorkbook
// Purpose    : Closes off the Workbook so it is ready for use.
// Returns    : success => Defaults to off, will fail if no workbook exists.
// Parameter/s: workbookName => Constant reference to the name of the file/workbook.
//--------------------------------------------------------------------------------------------------
dcl-proc Excel_closeWorkbook export;
  dcl-pi *n ind;
    workbookName char(45) const;
  end-pi;

  dcl-s success ind inz(*off);
  dcl-s fileName char(50) inz;
  dcl-s assignedFile like(pFILE) inz;

  // Open the file in the IFS.
  fileName = %trim(workbookName) + '.xml';
  assignedFile = fopen(%trim(fileName) : 'a+');
  if assignedFile <> *null;

    // Close the final tag for the workbook.
    fputs('</ss:Workbook>' : assignedFile);

    // Close the file as we are done.
    fclose(assignedFile);
    success = *on;
  endif;

  return success;
end-proc;

//--------------------------------------------------------------------------------------------------
// Procedure  : Excel_createWorksheet
// Purpose    : Creates a new worksheet, with the name provided. Workbook must exist before creating
//              a new worksheet. The closeWorksheet procedure must be called to close off the
//              worksheet when all processing is completed.
// Returns    : success => Defaults to off, will fail if no workbook exists.
// Parameter/s: workbookName => Constant reference to the name of the file/workbook.
//            : worksheetName => Constant reference to the name of the worksheet
//--------------------------------------------------------------------------------------------------
dcl-proc Excel_createWorksheet export;
  dcl-pi *n ind;
    workbookName char(45) const;
    worksheetName char(30) const;
  end-pi;

  dcl-s success ind inz(*off);
  dcl-s fileName char(50) inz;
  dcl-s assignedFile like(pFILE) inz;

  // Open the existing file in the IFS...
  fileName = %trim(workbookName) + '.xml';
  assignedFile = fopen(%trim(fileName) : 'a+');
  if assignedFile <> *null;

    // Create and name our worksheet, and open the table.
    fputs('<ss:Worksheet ss:Name="' + %trim(worksheetName) + '">' : assignedFile);
    fputs('<ss:Table>' : assignedFile);

    // Close the file as we are done.
    fclose(assignedFile);
    success = *on;
  endif;

  return success;
end-proc;

//--------------------------------------------------------------------------------------------------
// Procedure  : Excel_closeWorksheet
// Purpose    : Closes off a worksheet to complete off all data that has been added. Workbook and
//              worksheet must exist before closing the worksheet. This procedure must be called
//              before attempting to close a workbook or adding new worksheets.
//                Eventually support will be added to allow multiple worksheets to be worked with
//                at once, but currently, the worksheet name is not used and this will just close
//                the most recently opened worksheet.
// Returns    : success => Defaults to off, will fail if no workbook exists.
// Parameter/s: workbookName => Constant reference to the name of the file/workbook.
//            : worksheetName => Constant reference to the name of the worksheet (Not yet used...)
//--------------------------------------------------------------------------------------------------
dcl-proc Excel_closeWorksheet export;
  dcl-pi *n ind;
    workbookName char(45) const;
    worksheetName char(30) const;
  end-pi;

  dcl-s success ind inz(*off);
  dcl-s fileName char(50) inz;
  dcl-s assignedFile like(pFILE) inz;

  // Open the existing file in the IFS...
  fileName = %trim(workbookName) + '.xml';
  assignedFile = fopen(%trim(fileName) : 'a+');
  if assignedFile <> *null;

    // Close off the table and worksheet...
    fputs('</ss:Table>' : assignedFile);
    fputs('</ss:Worksheet>' : assignedFile);

    // Close the file as we are done.
    fclose(assignedFile);
    success = *on;
  endif;

  return success;
end-proc;

//--------------------------------------------------------------------------------------------------
// Procedure  : Excel_createRow
// Purpose    : Opens a new row within the workbook in the most recently opened worksheet. (Support
//              to be added later to allow updates to previously opened worksheets). Workbook and
//              worksheet must exist before attempting to create a row.
//              Use this to create a row before adding cells with data. Once finished with the
//              row, use closeRow to end.
// Returns    : success => Defaults to off, will fail if no workbook exists.
// Parameter/s: workbookName => Constant reference to the name of the file/workbook.
//            : worksheetName => constant reference to the name of the worksheet (Not yet used...)
//--------------------------------------------------------------------------------------------------
dcl-proc Excel_createRow export;
  dcl-pi *n ind;
    workbookName char(45) const;
    worksheetName char(30) const;
  end-pi;

  dcl-s success ind inz(*off);
  dcl-s fileName char(50) inz;
  dcl-s assignedFile like(pFILE) inz;

  // Open the existing file in the IFS...
  fileName = %trim(workbookName) + '.xml';
  assignedFile = fopen(%trim(fileName) : 'a+');
  if assignedFile <> *null;

    // Open a new row...
    fputs('<ss:Row>' : assignedFile);

    // Close the file as we are done.
    fclose(assignedFile);
    success = *on;
  endif;

  return success;
end-proc;

//--------------------------------------------------------------------------------------------------
// Procedure  : Excel_closeRow
// Purpose    : Closes a new row within the workbook in the most recently opened worksheet. (Support
//              to be added later to allow updates to previously opened worksheets). Workbook,
//              worksheet and row must exist before attempting to create a row.
//              Use this to close a row after it has been created and when the cells have been
//              filled.
// Returns    : success => default to off, will be set on if processing able to execute.
// Parameter/s: workbookName => Name of the workbook to add the row to.
//              worksheetName => Not used currently - will allow multiple worksheets to be worked on
//--------------------------------------------------------------------------------------------------
dcl-proc Excel_closeRow export;
  dcl-pi *n ind;
    workbookName char(45) const;
    worksheetName char(30) const;
  end-pi;

  dcl-s success ind inz(*off);
  dcl-s fileName char(50) inz;
  dcl-s assignedFile like(pFILE) inz;

  // Open the existing file in the IFS...
  fileName = %trim(workbookName) + '.xml';
  assignedFile = fopen(%trim(fileName) : 'a+');
  if assignedFile <> *null;

    // Close the row
    fputs('</ss:Row>' : assignedFile);

    // Close the file as we are done.
    fclose(assignedFile);
    success = *on;
  endif;

  return success;
end-proc;

//--------------------------------------------------------------------------------------------------
// Procedure  : Excel_writeCell
// Purpose    : This procedure is used to write a single cell into an existing row. The workbook,
//              worksheet and a row must exist to add this to.
// Returns    : success => default to off, will be set on if processing able to execute.
// Parameter/s: workbookName => Name of the workbook to add the row to.
//              worksheetName => Not used currently - will allow multiple worksheets to be worked on
//              cellData => contains the actual detail to add to the cell.
//              cellType => (use the defined constants) to advise excel what the data type is, being
//                          either a string or a numeric.
//              cellStyle => The constant for the style to apply to the cell.
//              mergeAcross => Optional parm if the cell should merge with multiple cells.
//--------------------------------------------------------------------------------------------------
dcl-proc Excel_writeCell export;
  dcl-pi *n ind;
    workbookName char(45) const;
    worksheetName char(30) const;
    cellData char(200) const;
    cellType char(1) const;
    cellStyle char(2) const;
    mergeAcross zoned(3:0) const options(*nopass);
  end-pi;

  dcl-s success ind inz(*off);
  dcl-s fileName char(50) inz;
  dcl-s assignedFile like(pFILE) inz;
  dcl-s mergeAcrossCells char(3) inz;

  // If we had a request for the merge, then we'll test it's a numeric. If it isn't, leave with fail
  if %parms >= %parmnum(mergeAcross);
    mergeAcrossCells = %char(mergeAcross);
  endif;

  // Open the existing file...
  fileName = %trim(workbookName) + '.xml';
  assignedFile = fopen(%trim(fileName) : 'a+');
  if assignedFile <> *null;
    success = writeCell(assignedFile : cellData : cellType : cellStyle : mergeAcrossCells);
    // Close the file as we are done.
    fclose(assignedFile);
  endif;

  return success;
end-proc;

//--------------------------------------------------------------------------------------------------
// Procedure  : writeCell
// Purpose    : Local proc for the actual write of the cells.
// Returns    : success => default to off, will be set on if processing able to execute.
// Parameter/s: workbookName => Name of the workbook to add the row to.
//              worksheetName => Not used currently - will allow multiple worksheets to be worked on
//              cellData => contains the actual detail to add to the cell.
//              cellType => (use the defined constants) to advise excel what the data type is, being
//                          either a string or a numeric.
//              cellStyle => The constant for the style to apply to the cell.
//              mergeAcrossCells => blank if no merge required
//--------------------------------------------------------------------------------------------------
dcl-proc writeCell;
  dcl-pi *n ind;
    assignedFile like(pFile) const;
    cellData char(200) const;
    cellType char(1) const;
    cellStyle char(2) const;
    mergeAcrossCells char(3) const;
  end-pi;

  dcl-s success ind inz(*off);

  if assignedFile <> *null;
    // If the cell was requested to be merged, then we'll use a different cell definition with
    // the number of columns we'll be merging, as well as the style.
    if mergeAcrossCells <> *blanks;
      fputs('<ss:Cell ss:MergeAcross="' + %trim(mergeAcrossCells)
             + '" ss:StyleID="' + cellStyle + '">' : assignedFile);

    // Otherwise we'll just create the cell and add the style for the single cell.
    else;
      fputs('<ss:Cell ss:StyleID="' + cellStyle + '">' : assignedFile);
    endif;

    // Depending on the style requested, add the data into the correct type.
    if cellType = EXCEL_CELL_TYPE_STRING;
      fputs('<ss:Data ss:Type="String">' + %trimr(cellData) + '</ss:Data> ' : assignedFile);
    elseif cellType = EXCEL_CELL_TYPE_NUMBER;
      fputs('<ss:Data ss:Type="Number">' + %trimr(cellData) + '</ss:Data> ' : assignedFile);
    endif;

    // Close the cell off.
    fputs('</ss:Cell>' : assignedFile);
    success = *on;
  endif;

  return success;
end-proc;

//--------------------------------------------------------------------------------------------------
// Procedure  : Excel_setColumnWidths
// Purpose    : Set the width for the columns within a worksheet. The workbook and worksheet must
//              already exist and not be closed. The count must be provided for the number of
//              columns we will be setting the width for.
//                Note: Excel will only autofitwidth for Numeric and Date Cells.
// Returns    : success => default to off, will be set on if processing able to execute.
// Parameter/s: workbookName => Name of the workbook to add the row to.
//              worksheetName => Not used currently - will allow multiple worksheets to be worked on
//              columnsData => Data structure containing the count of the columns we will be
//                             updating as well as the values themselves, within the array.
//--------------------------------------------------------------------------------------------------
dcl-proc Excel_setColumnWidths export;
  dcl-pi *n ind;
    workbookName char(45) const;
    worksheetName char(30) const;
    columnsData likeDS(dt_columnsData) const;
  end-pi;

  dcl-s success ind inz(*off);
  dcl-s fileName char(50) inz;
  dcl-s assignedFile like(pFILE) inz;
  dcl-s i int(10:0) inz;

  // Open the existing file...
  fileName = %trim(workbookName) + '.xml';
  assignedFile = fopen(%trimr(fileName) : 'a+');
  if assignedFile <> *null;

    // As long as we have no more than the number of columns we want to support, we will write
    // each of the widths for the count provided, even if it is zero.
    if columnsData.count <= %elem(columnsData.widths);
      for i = 1 to columnsData.count;
        // Self closing cell containing the width specified.
        fputs('<ss:Column ss:Width="' + %trim(%char(columnsData.widths(i))) + '" />'
            : assignedFile);
      endfor;

      // Close the file as we are done.
      fclose(assignedFile);
      success = *on;
    endif;
  endif;

  return success;
end-proc;

//--------------------------------------------------------------------------------------------------
// Procedure  : Excel_deleteWorkbook
// Purpose    : Removes the workbook from the IFS.
// Returns    : success => default to off, will be set on if processing able to execute.
// Parameter/s: workbookName => Name of the workbook to delete.
//--------------------------------------------------------------------------------------------------
dcl-proc Excel_deleteWorkbook export;
  dcl-pi *n ind;
    workbookName char(45) const;
  end-pi;

  dcl-s success ind inz(*off);


  return success;
end-proc;

//--------------------------------------------------------------------------------------------------
// Procedure  : Excel_writeSimpleRow
// Purpose    : Writes a row with the cells added to it using simple processing.
// Returns    : success => default to off, will be set on if processing able to execute.
// Parameter/s: workbookName => Name of the workbook to add to.
//            : worksheetName => Name of the worksheet to add to.
//            : simpleRow => DS containing a DS array with the cell data to add to a row.
//--------------------------------------------------------------------------------------------------
dcl-proc Excel_writeSimpleRow export;
  dcl-pi *n ind;
    workbookName char(45) const;
    worksheetName char(30) const;
    simpleRow likeDS(dt_simpleRow) const;
  end-pi;

  dcl-s success ind inz(*off);
  dcl-s fileName char(50) inz;
  dcl-s assignedFile like(pFILE) inz;
  dcl-s mergeAcrossCells char(3) inz;
  dcl-s i int(10:0) inz;

  // Open the existing file...
  fileName = %trim(workbookName) + '.xml';
  assignedFile = fopen(%trim(fileName) : 'a+');
  if assignedFile <> *null;
    // We don't support merge here, so we'll just set to blank and send it to the local proc.
    mergeAcrossCells = *blanks;
    fputs('<ss:Row>' : assignedFile);

    // Loop over each DS array and add it as a cell to the row. If for any reason it fails, then
    // we will exit and report the error.
    for i = 1 to simpleRow.count;
      success = writeCell(assignedFile
                        : simpleRow.cells(i).cellData
                        : simpleRow.cells(i).cellType
                        : simpleRow.cells(i).cellStyle
                        : mergeAcrossCells);
      if not success;
        leave;
      endif;
    endfor;

    // Close the row. Then close the file as we are done.
    fputs('</ss:Row>' : assignedFile);
    fclose(assignedFile);
  endif;

  return success;
end-proc;


//--------------------------------------------------------------------------------------------------
// Procedure  : Excel_email
// Purpose    : Emails the workbook to the specified email. The workbook will be attached to the
//              email.
//              We have support for up to 20 recipients - including BCC and CC types. Please be sure
//              that the number of recipients is correctly updated in email.recipientCount.
//              Uses SNDSMTPEMM to send email. IBM documentation as follows:
//                 https://www.ibm.com/support/pages/sndsmtpemm-command-instructions
//              Please note: This does NOT delete the file from the IFS after sending.
// Returns    : success => default to off, will be set on if processing able to execute.
// Parameter/s: workbookName => Name of the workbook to add to.
//            : email => Contains all information for sending.
//--------------------------------------------------------------------------------------------------
dcl-proc Excel_email export;
  dcl-pi *n ind;
    workbookName char(45) const;
    email likeDS(dt_excelEmail) const;
  end-pi;

  dcl-pr EXECUTE extpgm('QCMDEXC');
    commandString char(10000) const;
    commandLength packed(15:5) const;
  end-pr;

  dcl-s success ind inz(*off);
  dcl-s fileName char(50) inz;
  dcl-s i int(10:0) inz;
  dcl-s recipientString char(2000) inz;
  dcl-s command char(10000) inz;
  dcl-s command2 char(1000) inz;
  dcl-s user char(10) inz;

  // Complete the file name and prepare our command. The workbook name will contain the path.
  fileName = %trim(workbookName) + '.xml';

  // Don't even try to send this unless we have at LEAST one recipient.
  if email.recipientCount > *zero;
    // Build the recipient string as RCP((X@Y.COM) (Y@Y.COM *BCC) (etc...))
    recipientString = 'RCP(';
    for i = 1 to email.recipientCount;
      recipientString = %trim(recipientString) + '(' + %trim(email.emailRecipient(i).emailAddress);
      if email.emailRecipient(i).type <> EXCEL_RECIPIENT_TYPE_PRIMARY
         and email.emailRecipient(i).type <> *blanks;
        recipientString = %trim(recipientString) + ' ' + %trim(email.emailRecipient(i).type);
      endif;
      recipientString = %trim(recipientString) + ')';
    endfor;
    recipientString = %trim(recipientString) + ')';

    command = 'SNDSMTPEMM ' + %trim(recipientString)
            + ' SUBJECT(''' + %trim(email.subject)
            + ''') NOTE(''' + %trim(email.content)
            + ''') ATTACH(('''
            + %trim(fileName)
            + ''' *XML *TXT)) CONTENT(';
    if email.contentType = *blanks;
      command = %trim(command) + ' ' + %trim(email.contentType) + ')';
    else;
      command = %trim(command) + ' ' + %trim(EXCEL_CONTENT_TYPE_PLAIN) + ')';
    endif;
    
    // Send the email...
    EXECUTE(%trim(command) : %len(%trim(command)));
    success = *on;
  endif;

  return success;
end-proc; 
