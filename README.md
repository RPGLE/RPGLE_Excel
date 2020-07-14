# RPGLE_Excel
Service program to generate Excel spreadsheets with xml through a simple Service Program.

<h2>Terminology</h2>
<ul>
  <li>Workbook – the entire Excel document is known as the workbook. The workbook contains the style definitions, and all worksheets.</li>
  <li>Worksheet – A single page of data (sometimes referred to as a “tab”) contained within the workbook. The worksheet is a container for column definitions and rows. </li>
  <li>Rows – A horizontal list of cells, contained within a worksheet. A row does not have any real information but instead is a container for cells that are filled with the data displayed in the worksheet.</li>
  <li>Columns – A Vertical list of cells, contained with a worksheet. With the way the service program uses columns, they are only utilized for column width definitions. All data is instead handled by cells within the rows.</li>
  <li>Cell – The actual data that is contained within a row. Cells hold style references, the data to be presented as well as the type of cell. Cells can also be merged to create longer cells.</li>
</ul>

