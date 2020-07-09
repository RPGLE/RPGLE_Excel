**FREE
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
// Ghost++   2020/07/07 Initial Coding
//--------------------------------------------------------------------------------------------------
ctl-opt option(*srcstmt: *nodebugio) datfmt(*iso) nomain;

/include QRPGLESRC,EXCEL_P
/include QRPGLESRC,SQLERR_P
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
    fputs('<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>' : assignedFile);
    fputs('<ss:Workbook ' : assignedFile);
    fputs('xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">' : assignedFile);

    // Default Styles - can be customized or more styles added later...
    fputs('<ss:Styles>' : assignedFile);

    // Header 1 Style
    fputs('<ss:Style ss:ID="H1">' : assignedFile);
    fputs('<ss:Font ss:FontName="Calibri" ' : assignedFile);
    fputs('ss:Size="14" ss:Color="#222222"/>' : assignedFile);
    fputs('<ss:Interior ss:Color="#CCFFCC" ss:Pattern="Solid"/>' : assignedFile);
    fputs('</ss:Style>' : assignedFile);

    // Detail 1 Style
    fputs('<ss:Style ss:ID="D1">' : assignedFile);
    fputs('<ss:Font ss:FontName="Calibri" ' : assignedFile);
    fputs('ss:Size="12" ss:Color="#000000"/>' : assignedFile);
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

    // Close the file as we are done.
    fclose(assignedFile);
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
