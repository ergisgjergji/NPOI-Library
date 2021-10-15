using Microsoft.Extensions.Logging;
using NPOI.SS.UserModel;
using Npoi_Library.Excel.Configurations;
using Npoi_Library.Excel.CustomAttributes;
using Npoi_Library.Excel.CustomAttributes.Configurations;
using Npoi_Library.Excel.Helpers;
using Npoi_Library.Excel.Styling;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using Sp = Spire.Xls;
using NPOI.XSSF.UserModel;

/*
    This packagage uses 2 libraries:
    - NPOI: https://github.com/dotnetcore/NPOI
        This package is used for working solely with Excel files: generating excel files and/or reading from them.
    - FreeSpire.XLS: https://www.e-iceblue.com/Introduce/free-xls-component.html
        This package is used for converting Excel files to other formats (PDF, Html etc.)
 */

namespace Npoi_Library.Excel.XlsxManager
{
    public enum ColorIndex
    {
        HeaderFontColor = 56,
        HeaderBgColor = 57,
        BodyFontColor = 58,
        BodyBgColor = 59,
        HighlightColor = 60
    }

    public class XlsxManager
    {
        private static Type type;
        private static XSSFWorkbook workbook;
        private static ISheet Sheet;
        private static IDataFormat format;
        private static XSSFCellStyle headerCellStyle;
        private static ICellStyle bodyCellStyle;
        private static ICellStyle emptyCellStyle;

        private readonly ILogger<XlsxManager> _logger;

        public XlsxManager(ILogger<XlsxManager> logger = null)
        {
            _logger = logger;
        }

        /// <summary>
        /// Generates an Excel file containing a single sheet.
        /// </summary>
        /// <param name="dataList"> A collection of type T. </param>
        /// <param name="options"> (optional) Configurable options, like: HeaderStyle, BodyStyle </param>
        public byte[] GenerateExcel<T>(IEnumerable<T> dataList, ExcelOptions options) where T : class
        {
            if (dataList == null)
                throw new ArgumentNullException(nameof(dataList));

            if (dataList.Count() == 0)
                _logger?.LogWarning("ExcelManager.GenerateExcel<T> | Data is empty.");

            try
            {
                InitializeDocument(options);

                type = typeof(T);
                PropertyInfo[] properties;
                if (type.GetInterfaces().Contains(typeof(IPositionable)))
                    properties = type.GetProperties().Where(p => !p.Name.Equals("PositionMap")).ToArray();
                else
                    properties = type.GetProperties();

                List<PropertyConfig> configList = new List<PropertyConfig>();

                if (properties.Count() > 0)
                {
                    configList = GetPropertyConfigurationList(properties, bodyCellStyle);

                    // Set ColumnPosition for the properties in which ColumnPosition is not defined
                    SortColumnPositions(configList);

                    // Sort the columns
                    configList.Sort((c1, c2) => c1.ColumnPosition.CompareTo(c2.ColumnPosition));

                    IRow HeaderRow = Sheet.CreateRow(0);
                    // Generate column headers
                    GenerateHeaders(configList, HeaderRow, headerCellStyle);
                }

                // Iterating the data
                #region
                int rowIndex = 1;
                foreach (var item in dataList)
                {
                    IRow CurrentRow = Sheet.CreateRow(rowIndex);

                    int propIndex = 0;
                    foreach (PropertyConfig config in configList)
                    {
                        ICell Cell = CurrentRow.CreateCell(propIndex);
                        object propValue = item.GetType().GetProperty(config.PropertyName).GetValue(item, null);

                        PrintCellValue(Cell, propValue);

                        if (propValue == null)
                            Cell.CellStyle = emptyCellStyle;
                        else
                            Cell.CellStyle = config.CellStyle;

                        propIndex++;
                    }
                    rowIndex++;
                }
                #endregion

                AutoSizeSheet(Sheet);

                // Save Excel file as byte[] 
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    workbook.Write(memoryStream);
                    return memoryStream.ToArray();
                }
            }
            catch (Exception e)
            {
                _logger?.LogError("Error at: ExcelManager.GenerateExcel<T> | " + e.Message);
                throw new ApplicationException("Error at: ExcelManager.<T> | " + e.Message);
            }
            finally
            {
                ReleaseMemory();
            }
        }

        /// <summary>
        /// Generates an Excel file from a DataTable, containing a single sheet.
        /// </summary>
        /// <param name="table"> DataTable. </param>
        /// <param name="options"> (optional) Configurable options, like: HeaderStyle, BodyStyle </param>
        public byte[] GenerateExcel(DataTable table, ExcelOptions options)
        {
            if (table == null)
                throw new ArgumentNullException(nameof(table));

            if (table.Columns.Count == 0 || table.Rows.Count == 0)
                _logger?.LogWarning("ExcelManager.GenerateExcel<DataTable> | Data is empty.");

            try
            {
                InitializeDocument(options);

                List<PropertyConfig> configList = GetDataColumnConfigurationList(table.Columns, bodyCellStyle);

                IRow HeaderRow = Sheet.CreateRow(0);
                // Generate column headers
                GenerateHeaders(configList, HeaderRow, headerCellStyle);

                // Iterating the data
                #region
                int rowIndex = 1;
                foreach (DataRow row in table.Rows)
                {
                    IRow CurrentRow = Sheet.CreateRow(rowIndex);

                    int propIndex = 0;
                    foreach (PropertyConfig config in configList)
                    {
                        ICell Cell = CurrentRow.CreateCell(propIndex);
                        object propValue = row[config.HeaderName];

                        PrintCellValue(Cell, propValue);

                        if (propValue == null)
                            Cell.CellStyle = emptyCellStyle;
                        else
                            Cell.CellStyle = config.CellStyle;

                        propIndex++;
                    }
                    rowIndex++;
                }
                #endregion

                AutoSizeSheet(Sheet);

                // Save Excel file as byte[] 
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    workbook.Write(memoryStream);
                    return memoryStream.ToArray();
                }
            }
            catch (Exception e)
            {
                _logger?.LogError("Error at: ExcelManager.GenerateExcel<DataTable> | " + e.Message);
                throw new ApplicationException("Error at: ExcelManager.GenerateExcel<DataTable> | " + e.Message);
            }
            finally
            {
                ReleaseMemory();
            }
        }

        /// <summary>
        /// Generates an Excel file based on a template.
        /// </summary>
        /// <typeparam name="T"> Implementation of IPositionable </typeparam>
        /// <param name="data"> An object of type T. </param>
        /// <param name="templateLocation"> Absolute path of the template file. </param>
        /// <param name="tempSheetName"> Sheet name of the template file. </param>
        public byte[] GenerateExcelFromTemplate<T>(T data, string templateLocation, string tempSheetName) where T : IPositionable
        {
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            if (string.IsNullOrWhiteSpace(templateLocation))
                throw new ArgumentNullException(nameof(templateLocation));

            if (!File.Exists(templateLocation))
                throw new ArgumentException($"Can't find the excel template on path: {templateLocation}");

            if (string.IsNullOrWhiteSpace(tempSheetName))
                throw new ArgumentNullException(nameof(tempSheetName));

            type = typeof(T);
            try
            {
                using (var file = new FileStream(templateLocation, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    workbook = new XSSFWorkbook(file);
                    if (workbook == null)
                        throw new ArgumentNullException(nameof(workbook));

                    Sheet = workbook.GetSheet(tempSheetName);
                    if (Sheet == null)
                        throw new ArgumentNullException(nameof(Sheet));

                    if (data.PositionMap == null || data.PositionMap.Count() == 0)
                        throw new ArgumentException("PositionMap not configured.");

                    List<string> keys = new List<string>(data.PositionMap.Keys);

                    foreach (var key in keys)
                    {
                        object propValue = data.GetType().GetProperty(key).GetValue(data, null);

                        int rowIndex = data.PositionMap[key].RowIndex - 1;
                        int colIndex = ExcelHelpers.ColLetterToNumber(data.PositionMap[key].ColumnLetter) - 1;

                        IRow currRow = Sheet.GetRow(rowIndex);
                        if (currRow == null)
                            currRow = Sheet.CreateRow(rowIndex);

                        ICell currCell = currRow.GetCell(colIndex);
                        if (currCell == null)
                            currCell = currRow.CreateCell(colIndex);

                        PrintCellValue(currCell, propValue);
                    }

                    WorkbookFinalProcessing();

                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        workbook.Write(memoryStream);
                        return memoryStream.ToArray();
                    }
                }
            }
            catch (Exception e)
            {
                _logger?.LogError("Error at: ExcelManager.GenerateExcelFromTemplate | " + e.Message);
                throw new ApplicationException("Error at: ExcelManager.GenerateExcelFromTemplate | " + e.Message);
            }
            finally
            {
                ReleaseMemory();
            }
        }

        /// <summary>
        /// Returns byte array of content of an Excel file, built over a template file (.xls)
        /// </summary>
        /// <typeparam name="T"> Implementation of IPositionable </typeparam>
        /// <param name="dataList"> A collection of type T. </param>
        /// <param name="templateLocation"> Absolute path of the template file. </param>
        /// <param name="tempSheetName"> Sheet name of the template file. </param>
        public byte[] GenerateExcelFromTemplate<T>(IEnumerable<T> dataList, string templateLocation, string tempSheetName) where T : IPositionable
        {
            if (dataList == null)
                throw new ArgumentNullException(nameof(dataList));

            if (string.IsNullOrWhiteSpace(templateLocation))
                throw new ArgumentNullException(nameof(templateLocation));

            if (!File.Exists(templateLocation))
                throw new ArgumentException($"Can't find the excel template on path: {templateLocation}");

            if (string.IsNullOrWhiteSpace(tempSheetName))
                throw new ArgumentNullException(nameof(tempSheetName));

            type = typeof(T);
            try
            {
                using (var file = new FileStream(templateLocation, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    workbook = new XSSFWorkbook(file);
                    if (workbook == null)
                        throw new ArgumentNullException(nameof(workbook));

                    Sheet = workbook.GetSheet(tempSheetName);
                    Sheet.DefaultColumnWidth = 50;
                    if (Sheet == null)
                        throw new ArgumentNullException(nameof(Sheet));

                    if (dataList.Count() == 0)
                        _logger?.LogWarning("ExcelManager.GenerateExcel | Data is empty.");

                    foreach (var item in dataList)
                    {
                        if (item.PositionMap == null || item.PositionMap.Count() == 0)
                            throw new ArgumentException($"PositionMap not configured for item: dataList[{dataList.ToList().FindIndex(i => i.Equals(item))}].");

                        List<string> keys = new List<string>(item.PositionMap.Keys);

                        foreach (var key in keys)
                        {
                            object propValue = item.GetType().GetProperty(key).GetValue(item, null);

                            int rowIndex = item.PositionMap[key].RowIndex - 1;
                            int colIndex = ExcelHelpers.ColLetterToNumber(item.PositionMap[key].ColumnLetter) - 1;

                            IRow currRow = Sheet.GetRow(rowIndex);
                            if (currRow == null)
                                currRow = Sheet.CreateRow(rowIndex);

                            ICell currCell = currRow.GetCell(colIndex);
                            if (currCell == null)
                                currCell = currRow.CreateCell(colIndex);

                            PrintCellValue(currCell, propValue);
                        }
                    }

                    WorkbookFinalProcessing();

                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        workbook.Write(memoryStream);
                        return memoryStream.ToArray();
                    }
                }
            }
            catch (Exception e)
            {
                _logger?.LogError("Error at: ExcelManager.GenerateExcelFromTemplate | " + e.Message);
                throw new ApplicationException("Error at: ExcelManager.GenerateExcelFromTemplate | " + e.Message);
            }
            finally
            {
                ReleaseMemory();
            }
        }

        /// <summary>
        /// Returns byte array of content of an Excel file, built over a template file (.xls)
        /// </summary>
        /// <typeparam name="T"> Implementation of IPositionable </typeparam>
        /// <param name="dataList"> A collection of type T. </param>
        /// <param name="templateLocation"> Absolute path of the template file. </param>
        /// <param name="tempSheetName"> Sheet name of the template file. </param>
        /// <param name="dataSection"> Specifies the section where the table of data lives. </param>
        public byte[] GenerateExcelFromTemplate<T>(IEnumerable<T> dataList, string templateLocation, string tempSheetName, ExcelTemplateDataSection dataSection) where T : IPositionable
        {
            if (dataList == null)
                throw new ArgumentNullException(nameof(dataList));

            if (string.IsNullOrWhiteSpace(templateLocation))
                throw new ArgumentNullException(nameof(templateLocation));

            if (!File.Exists(templateLocation))
                throw new ArgumentException($"Can't find the excel template on path: {templateLocation}");

            if (string.IsNullOrWhiteSpace(tempSheetName))
                throw new ArgumentNullException(nameof(tempSheetName));

            type = typeof(T);
            try
            {
                using (var file = new FileStream(templateLocation, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    workbook = new XSSFWorkbook(file);
                    if (workbook == null)
                        throw new ArgumentNullException(nameof(workbook));

                    Sheet = workbook.GetSheet(tempSheetName);
                    Sheet.DefaultColumnWidth = 50;
                    if (Sheet == null)
                        throw new ArgumentNullException(nameof(Sheet));

                    if (dataList.Count() == 0)
                        _logger?.LogWarning("ExcelManager.GenerateExcel | Data is empty.");

                    foreach (var item in dataList)
                    {
                        if (item.PositionMap == null || item.PositionMap.Count() == 0)
                            throw new ArgumentException($"PositionMap not configured for item: dataList[{dataList.ToList().FindIndex(i => i.Equals(item))}].");

                        List<string> keys = new List<string>(item.PositionMap.Keys);

                        foreach (var key in keys)
                        {
                            object propValue = item.GetType().GetProperty(key).GetValue(item, null);

                            int rowIndex = item.PositionMap[key].RowIndex - 1;
                            int colIndex = ExcelHelpers.ColLetterToNumber(item.PositionMap[key].ColumnLetter) - 1;

                            IRow currRow = Sheet.GetRow(rowIndex);
                            if (currRow == null)
                                currRow = Sheet.CreateRow(rowIndex);

                            ICell currCell = currRow.GetCell(colIndex);
                            if (currCell == null)
                                currCell = currRow.CreateCell(colIndex);

                            PrintCellValue(currCell, propValue);
                        }
                    }

                    // Deletes the empty rows
                    ClearEmptyRows(dataSection.x1, dataSection.x2, dataSection.y1, dataSection.y2);

                    WorkbookFinalProcessing();

                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        workbook.Write(memoryStream);
                        return memoryStream.ToArray();
                    }
                }
            }
            catch (Exception e)
            {
                _logger?.LogError("Error at: ExcelManager.GenerateExcelFromTemplate | " + e.Message);
                throw new ApplicationException("Error at: ExcelManager.GenerateExcelFromTemplate | " + e.Message);
            }
            finally
            {
                ReleaseMemory();
            }
        }

        /// <summary>
        /// Returns byte array of content of an Excel file, built over a template file (.xls)
        /// </summary>
        /// <typeparam name="T"> Implementation of IPositionable </typeparam>
        /// <param name="dataList"> A collection of type T. </param>
        /// <param name="templateLocation"> Absolute path of the template file. </param>
        /// <param name="tempSheetName"> Sheet name of the template file. </param>
        /// <param name="dataSection"> Specifies the section where the table of data lives. </param>
        /// <param name="footer"> A textbox will be drawn below the data, acting like a footer. </param>
        public byte[] GenerateExcelFromTemplate<T>(IEnumerable<T> dataList, string templateLocation, string tempSheetName, ExcelTemplateDataSection dataSection, string footer) where T : IPositionable
        {
            if (dataList == null)
                throw new ArgumentNullException(nameof(dataList));

            if (string.IsNullOrWhiteSpace(templateLocation))
                throw new ArgumentNullException(nameof(templateLocation));

            if (!File.Exists(templateLocation))
                throw new ArgumentException($"Can't find the excel template on path: {templateLocation}");

            if (string.IsNullOrWhiteSpace(tempSheetName))
                throw new ArgumentNullException(nameof(tempSheetName));

            type = typeof(T);
            try
            {
                using (var file = new FileStream(templateLocation, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    workbook = new XSSFWorkbook(file);
                    if (workbook == null)
                        throw new ArgumentNullException(nameof(workbook));

                    Sheet = workbook.GetSheet(tempSheetName);
                    Sheet.DefaultColumnWidth = 50;
                    if (Sheet == null)
                        throw new ArgumentNullException(nameof(Sheet));

                    if (dataList.Count() == 0)
                        _logger?.LogWarning("ExcelManager.GenerateExcel | Data is empty.");

                    foreach (var item in dataList)
                    {
                        if (item.PositionMap == null || item.PositionMap.Count() == 0)
                            throw new ArgumentException($"PositionMap not configured for item: dataList[{dataList.ToList().FindIndex(i => i.Equals(item))}].");

                        List<string> keys = new List<string>(item.PositionMap.Keys);

                        foreach (var key in keys)
                        {
                            object propValue = item.GetType().GetProperty(key).GetValue(item, null);

                            int rowIndex = item.PositionMap[key].RowIndex - 1;
                            int colIndex = ExcelHelpers.ColLetterToNumber(item.PositionMap[key].ColumnLetter) - 1;

                            IRow currRow = Sheet.GetRow(rowIndex);
                            if (currRow == null)
                                currRow = Sheet.CreateRow(rowIndex);

                            ICell currCell = currRow.GetCell(colIndex);
                            if (currCell == null)
                                currCell = currRow.CreateCell(colIndex);

                            PrintCellValue(currCell, propValue);
                        }
                    }

                    // Deletes the empty rows
                    ClearEmptyRows(dataSection.x1, dataSection.x2, dataSection.y1, dataSection.y2);

                    // Insert a footer in the end of the data
                    DrawFooter(footer);

                    WorkbookFinalProcessing();

                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        workbook.Write(memoryStream);
                        return memoryStream.ToArray();
                    }
                }
            }
            catch (Exception e)
            {
                _logger?.LogError("Error at: ExcelManager.GenerateExcelFromTemplate | " + e.Message);
                throw new ApplicationException("Error at: ExcelManager.GenerateExcelFromTemplate | " + e.Message);
            }
            finally
            {
                ReleaseMemory();
            }
        }

        /// <summary>
        /// Returns byte array of content of an Excel file, built over a template file (.xls)
        /// </summary>
        /// <typeparam name="T"> Implementation of IPositionable </typeparam>
        /// <param name="dataList"> A collection of type T. </param>
        /// <param name="templateLocation"> Absolute path of the template file. </param>
        /// <param name="tempSheetName"> Sheet name of the template file. </param>
        /// <param name="dataSection"> Specifies the section where the table of data lives. </param>
        /// <param name="note"> Inserts a note in a merged-cells section. </param>
        /// <param name="footer"> A textbox will be drawn below the data, acting like a footer. </param>
        public byte[] GenerateExcelFromTemplate<T>(IEnumerable<T> dataList, string templateLocation, string tempSheetName, ExcelTemplateDataSection dataSection, ExcelNote note, string footer) where T : IPositionable
        {
            if (dataList == null)
                throw new ArgumentNullException(nameof(dataList));

            if (string.IsNullOrWhiteSpace(templateLocation))
                throw new ArgumentNullException(nameof(templateLocation));

            if (!File.Exists(templateLocation))
                throw new ArgumentException($"Can't find the excel template on path: {templateLocation}");

            if (string.IsNullOrWhiteSpace(tempSheetName))
                throw new ArgumentNullException(nameof(tempSheetName));

            type = typeof(T);
            try
            {
                using (var file = new FileStream(templateLocation, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    workbook = new XSSFWorkbook(file);
                    if (workbook == null)
                        throw new ArgumentNullException(nameof(workbook));

                    Sheet = workbook.GetSheet(tempSheetName);
                    Sheet.DefaultColumnWidth = 50;
                    if (Sheet == null)
                        throw new ArgumentNullException(nameof(Sheet));

                    if (dataList.Count() == 0)
                        _logger?.LogWarning("ExcelManager.GenerateExcel | Data is empty.");

                    foreach (var item in dataList)
                    {
                        if (item.PositionMap == null || item.PositionMap.Count() == 0)
                            throw new ArgumentException($"PositionMap not configured for item: dataList[{dataList.ToList().FindIndex(i => i.Equals(item))}].");

                        List<string> keys = new List<string>(item.PositionMap.Keys);

                        foreach (var key in keys)
                        {
                            object propValue = item.GetType().GetProperty(key).GetValue(item, null);

                            int rowIndex = item.PositionMap[key].RowIndex - 1;
                            int colIndex = ExcelHelpers.ColLetterToNumber(item.PositionMap[key].ColumnLetter) - 1;

                            IRow currRow = Sheet.GetRow(rowIndex);
                            if (currRow == null)
                                currRow = Sheet.CreateRow(rowIndex);

                            ICell currCell = currRow.GetCell(colIndex);
                            if (currCell == null)
                                currCell = currRow.CreateCell(colIndex);

                            PrintCellValue(currCell, propValue);
                        }
                    }

                    // Must come before ClearEmptyRows
                    DrawNote(note);

                    ClearEmptyRows(dataSection.x1, dataSection.x2, dataSection.y1, dataSection.y2);

                    DrawFooter(footer);

                    WorkbookFinalProcessing();

                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        workbook.Write(memoryStream);
                        return memoryStream.ToArray();
                    }
                }
            }
            catch (Exception e)
            {
                _logger?.LogError("Error at: ExcelManager.GenerateExcelFromTemplate | " + e.Message);
                throw new ApplicationException("Error at: ExcelManager.GenerateExcelFromTemplate | " + e.Message);
            }
            finally
            {
                ReleaseMemory();
            }
        }

        public IEnumerable<T> ReadFromExcel<T>(string location, string sheet, int rowIndex) where T : class
        {
            if (location == null)
                throw new ArgumentNullException(nameof(location));

            if (string.IsNullOrWhiteSpace(sheet))
                throw new ArgumentNullException(nameof(sheet));

            try
            {
                using (var file = new FileStream(location, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    workbook = new XSSFWorkbook(file);
                    if (workbook == null)
                        throw new ArgumentNullException(nameof(workbook));

                    Sheet = workbook.GetSheet(sheet);
                    if (Sheet == null)
                        throw new ArgumentNullException(nameof(Sheet));

                    type = typeof(T);
                    PropertyInfo[] properties;
                    if (type.GetInterfaces().Contains(typeof(IPositionable)))
                        properties = type.GetProperties().Where(p => !p.Name.Equals("PositionMap")).ToArray();
                    else
                        properties = type.GetProperties();

                    if (!properties.Any())
                        return null;

                    if (!PropertiesValid(properties))
                        throw new Exception($"{type.Name}: not all properties implement attribute {nameof(ExcelConfig)}");

                    var configList = GetPropertyConfigurationList(properties);

                    var dataList = new List<T>();
                    while (true)
                    {
                        if (IsRowEmpty(rowIndex, configList))
                            break;

                        var row = Sheet.GetRow(rowIndex);
                        T instance = (T)Activator.CreateInstance(type);

                        foreach (var propConfig in configList)
                        {
                            var propType = type.GetProperty(propConfig.PropertyName).PropertyType;

                            var cell = row.GetCell(propConfig.ColumnPosition - 1); // 0-based indexing
                            if (cell != null && cell.CellType != CellType.Blank)
                            {
                                var value = Convert.ChangeType(ReadCellValue(cell), Nullable.GetUnderlyingType(propType) ?? propType);
                                type.GetProperty(propConfig.PropertyName).SetValue(instance, value, null);
                            }
                        }

                        dataList.Add(instance);
                        rowIndex++;
                    }

                    return dataList;
                }
            }
            catch (Exception e)
            {
                _logger?.LogError("Error at: ExcelManager.ReadFromExcel<T> | " + e.Message);
                throw new ApplicationException("Error at: ExcelManager.ReadFromExcel<T> | " + e.Message);
            }
            finally
            {
                ReleaseMemory();
            }
        }

        public IEnumerable<T> ReadFromExcel<T>(string location, string sheet) where T : class
        {
            return ReadFromExcel<T>(location, sheet, 0);
        }

        public byte[] ConvertToPdf(byte[] xlsxBuffer)
        {
            if (xlsxBuffer is null)
                throw new ArgumentNullException(nameof(xlsxBuffer));

            try
            {
                using (var xlsxStream = new MemoryStream(xlsxBuffer))
                using (var pdfStream = new MemoryStream())
                {
                    var workbook = new Sp.Workbook();
                    workbook.LoadFromStream(xlsxStream);

                    workbook.SaveToStream(pdfStream, Sp.FileFormat.PDF);
                    return pdfStream.ToArray();
                }
            }
            catch (Exception e)
            {
                _logger?.LogError("Error at: ExcelManager.ConvertToPdf | " + e.Message);
                throw new ApplicationException("Error at: ExcelManager.ConvertToPdf | " + e.Message);
            }
        }

        #region  ---------------------------------------------------- Private methods -----------------------------------------------------------

        /// <summary>
        /// Configures document summary information
        /// </summary>
        private void InitializeDocument(ExcelOptions options)
        {
            workbook = new XSSFWorkbook();
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            Sheet = workbook.CreateSheet();
            if (Sheet == null)
                throw new ArgumentNullException(nameof(Sheet));

            Sheet.DefaultColumnWidth = 50;

            format = workbook.CreateDataFormat();

            // Cell styles
            if (options != null)
            {
                headerCellStyle = (XSSFCellStyle)GetHeaderStyle(options.HeaderStyle);
                bodyCellStyle = GetBodyStyle(options.BodyStyle);
            }
            else
            {
                headerCellStyle = (XSSFCellStyle)GetDefaultHeaderStyle();
                bodyCellStyle = GetDefaultBodyStyle();
            }

            emptyCellStyle = GetEmptyStyle();

            //DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            //dsi.Company = "InfoWeb Group";
            //workbook.DocumentSummaryInformation = dsi;

            //SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            //si.Subject = "NPOI SDK Example";
            //workbook.SummaryInformation = si;
        }

        private void ClearEmptyRows(int x1, int x2, int y1, int y2)
        {
            if (x1 > x2) throw new ArgumentException($"{nameof(x1)} cannot be greater than {nameof(x2)}");
            if (y1 > y2) throw new ArgumentException($"{nameof(y1)} cannot be greater than {nameof(y2)}");

            var realStartRow = x1 - 1;
            var realEndRow = x2 - 1;
            var startCol = y1 - 1;
            var endCol = y2 - 1;

            if (realEndRow > Sheet.LastRowNum)
                realEndRow = Sheet.LastRowNum;

            int i = realStartRow;
            while (i < realEndRow)
            {
                var row = Sheet.GetRow(i);
                if (row != null)
                {
                    var isEmptyRow = true;
                    for (int j = startCol; j <= endCol; j++)
                    {
                        var cell = row.GetCell(j);
                        if (cell != null && cell.CellType != CellType.Blank)
                        {
                            isEmptyRow = false;
                            break;
                        }
                    }
                    if (isEmptyRow)
                    {
                        Sheet.RemoveRow(row);
                        Sheet.ShiftRows(i + 1, Sheet.LastRowNum, -1);

                        realEndRow--;
                    }
                    else
                        i++;
                }
                else
                    break;
            }
        }

        /// <summary>
        /// Prints a note in a merged-cells section (x1, x2, y1, y2)
        /// </summary>
        private void DrawNote(ExcelNote note)
        {
            int startRow = note.x1 - 1;
            int endRow = note.x2 - 1;
            int startCol = note.y1 - 1;
            int endCol = note.y2 - 1;

            var noteStartRow = Sheet.GetRow(startRow);
            if (noteStartRow == null)
                noteStartRow = Sheet.CreateRow(startRow);

            var noteStartCell = noteStartRow.GetCell(startCol);
            if (noteStartCell == null)
                noteStartCell = noteStartRow.CreateCell(startCol);

            var noteEndRow = Sheet.GetRow(endRow);
            if (noteEndRow == null)
                Sheet.CreateRow(endRow);

            var noteEndCell = noteStartRow.GetCell(endCol);
            if (noteEndCell == null)
                noteStartRow.CreateCell(endCol);

            noteStartCell.SetCellValue(new XSSFRichTextString(note.Text));
            Sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRow, endRow, startCol, endCol));
        }

        /// <summary>
        /// Draw a textbox in the end of the file. No need for coordinates, it automatically inserts the footer in the end.
        /// </summary>
        private void DrawFooter(string footer)
        {
            int lastRowIndex = Sheet.LastRowNum;
            // We insert an offset of 5 here, in case there is some static text or data in between
            int txtStartIndex = lastRowIndex + 2;

            // Font
            IFont footerFont = workbook.CreateFont();
            footerFont.FontHeightInPoints = 8;
            footerFont.FontName = "Times New Roman";
            footerFont.IsBold = false;
            footerFont.IsItalic = true;

            // Textbox
            XSSFDrawing patriarch = (XSSFDrawing)Sheet.CreateDrawingPatriarch();

            XSSFTextBox textbox1 = patriarch.CreateTextbox(new XSSFClientAnchor(0, 0, 0, 0, 0, txtStartIndex, 12, txtStartIndex + 5));
            textbox1.LineStyle = LineStyle.None;
            textbox1.TopInset = 5;
            textbox1.BottomInset = 5;
            textbox1.LeftInset = 5;
            textbox1.RightInset = 5;

            var innerText = new XSSFRichTextString(footer);
            innerText.ApplyFont(footerFont);
            textbox1.SetText(innerText);
            // End: Textbox
        }

        /// <summary>
        /// Auto-sorts the properties for which the ColumnPosition is 0
        /// </summary>
        private void SortColumnPositions(List<PropertyConfig> configList)
        {
            ValidateColumnPosition(configList);

            int length = configList.Count();
            for (int i = 0; i < length; i++)
            {
                if (configList.ElementAt(i).ColumnPosition == 0)
                {
                    for (int j = i + 1; j <= length; j++)
                    {
                        if (configList.Where(c => c.ColumnPosition == j).Count() > 0)
                            continue;
                        else
                        {
                            configList.ElementAt(i).ColumnPosition = j;
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Checks if there are properties with invalid/duplicate ColumnPosition
        /// </summary>
        /// <param name="configList"></param>
        private void ValidateColumnPosition(List<PropertyConfig> configList)
        {
            if (configList.Where(c => c.ColumnPosition < 0).Count() > 0)
            {
                // Means there are many properties with same column position
                string errorMessage = $"Object `{type}` cannot have field with negative ColumnPosition.";
                _logger?.LogError(errorMessage);
                throw new ApplicationException(errorMessage);
            }

            var duplicateOrder = configList.Where(c => c.ColumnPosition > 0);
            foreach (var config in duplicateOrder)
            {
                var cnt = duplicateOrder.Where(c => c.ColumnPosition == config.ColumnPosition).Count();
                if (cnt > 1)
                {
                    // Means there are many properties with same column position
                    string errorMessage = $"Object `{type}` cannot have more than 1 field with ColumnPosition = {config.ColumnPosition}";
                    _logger?.LogError(errorMessage);
                    throw new ApplicationException(errorMessage);
                }
            }
        }

        private List<PropertyConfig> GetPropertyConfigurationList(PropertyInfo[] properties)
        {
            try
            {
                List<PropertyConfig> configList = new List<PropertyConfig>();

                for (int i = 0; i < properties.Length; i++)
                {
                    Type propType = properties[i].GetType();

                    PropertyConfig config = new PropertyConfig();
                    config.PropertyName = properties[i].Name;

                    var attr = properties[i].GetCustomAttribute<ExcelConfig>();
                    if (attr != null)
                    {
                        config.ColumnPosition = attr.ColumnPosition;

                        if (attr.HeaderName != null)
                            config.HeaderName = attr.HeaderName;
                        else
                            config.HeaderName = properties[i].Name;

                        if (attr.DataFormat != null)
                            config.DataFormat = attr.DataFormat;
                        else
                            config.DataFormat = GetTypeDefaultFormat(properties[i].PropertyType);
                    }
                    else
                    {
                        config.HeaderName = config.PropertyName;
                        config.ColumnPosition = 0;
                        config.DataFormat = GetTypeDefaultFormat(properties[i].PropertyType);
                    }

                    configList.Add(config);
                }

                return configList;
            }
            catch (Exception e)
            {
                _logger?.LogError("Error at: ExcelManager.GetPropertyConfigurationList | " + e.Message);
                throw new ApplicationException("Error at: ExcelManager.GetPropertyConfigurationList | " + e.Message);
            }
        }

        private List<PropertyConfig> GetPropertyConfigurationList(PropertyInfo[] properties, ICellStyle style)
        {
            try
            {
                List<PropertyConfig> configList = new List<PropertyConfig>();

                for (int i = 0; i < properties.Length; i++)
                {
                    Type propType = properties[i].GetType();

                    ICellStyle cellStyle = workbook.CreateCellStyle();
                    cellStyle.SetFont(style.GetFont(workbook));
                    cellStyle.FillForegroundColor = style.FillForegroundColor;
                    cellStyle.FillPattern = style.FillPattern;
                    cellStyle.BorderTop = style.BorderTop;
                    cellStyle.BorderBottom = style.BorderTop;
                    cellStyle.BorderLeft = style.BorderTop;
                    cellStyle.BorderRight = style.BorderTop;

                    PropertyConfig config = new PropertyConfig();
                    config.PropertyName = properties[i].Name;

                    var attributes = properties[i].GetCustomAttributes(true).Where(a => a.GetType().Equals(typeof(ExcelConfig)));
                    if (attributes.Count() == 1)
                    {
                        ExcelConfig attr = attributes.Single() as ExcelConfig;

                        config.ColumnPosition = attr.ColumnPosition;

                        if (attr.HeaderName != null)
                            config.HeaderName = attr.HeaderName;
                        else
                            config.HeaderName = properties[i].Name;

                        if (attr.DataFormat != null)
                            config.DataFormat = attr.DataFormat;
                        else
                            config.DataFormat = GetTypeDefaultFormat(properties[i].PropertyType);
                    }
                    else
                    {
                        config.HeaderName = config.PropertyName;
                        config.ColumnPosition = 0;
                        config.DataFormat = GetTypeDefaultFormat(properties[i].PropertyType);
                    }

                    if (config.DataFormat != null)
                        cellStyle.DataFormat = format.GetFormat(config.DataFormat);

                    config.CellStyle = cellStyle;
                    configList.Add(config);
                }

                return configList;
            }
            catch (Exception e)
            {
                _logger?.LogError("Error at: ExcelManager.GetPropertyConfigurationList | " + e.Message);
                throw new ApplicationException("Error at: ExcelManager.GetPropertyConfigurationList | " + e.Message);
            }
        }

        private List<PropertyConfig> GetDataColumnConfigurationList(DataColumnCollection columns, ICellStyle style)
        {
            try
            {
                List<PropertyConfig> configList = new List<PropertyConfig>();

                int i = 0;
                foreach (DataColumn col in columns)
                {
                    ICellStyle cellStyle = workbook.CreateCellStyle();
                    cellStyle.SetFont(style.GetFont(workbook));
                    cellStyle.FillForegroundColor = style.FillForegroundColor;
                    cellStyle.FillPattern = style.FillPattern;
                    cellStyle.BorderTop = style.BorderTop;
                    cellStyle.BorderBottom = style.BorderTop;
                    cellStyle.BorderLeft = style.BorderTop;
                    cellStyle.BorderRight = style.BorderTop;

                    PropertyConfig config = new PropertyConfig();
                    config.PropertyName = col.ColumnName;
                    config.HeaderName = col.ColumnName;
                    config.ColumnPosition = i;

                    string defaultFormat = GetTypeDefaultFormat(col.DataType);
                    if (defaultFormat != null)
                        cellStyle.DataFormat = format.GetFormat(defaultFormat);

                    config.CellStyle = cellStyle;
                    configList.Add(config);

                    i++;
                }

                return configList;
            }
            catch (Exception e)
            {
                throw new ApplicationException("Error at: ExcelManager.GetDataColumnConfigurationList | " + e.Message);
            }
        }

        private bool PropertiesValid(PropertyInfo[] properties)
        {
            if (properties.Any(p => p.GetCustomAttribute<ExcelConfig>() is null))
                return false;
            return true;
        }

        private bool IsRowEmpty(int rowIndex, IList<PropertyConfig> configList)
        {
            var row = Sheet.GetRow(rowIndex);
            if (row is null)
                return true;

            foreach (var config in configList)
            {
                var cell = row.GetCell(config.ColumnPosition);
                if (cell != null && cell.CellType != CellType.Blank)
                    return false;
                else
                    continue;
            }
            return true;
        }

        private void GenerateHeaders(List<PropertyConfig> configList, IRow HeaderRow, XSSFCellStyle headerCellStyle)
        {
            int i = 0;
            foreach (var config in configList)
            {
                ICell Cell = HeaderRow.CreateCell(i);
                Cell.SetCellValue(config.HeaderName);
                Cell.CellStyle = headerCellStyle;

                i++;
            }
        }

        private void AutoSizeSheet(ISheet Sheet)
        {
            int lastCell = Sheet.GetRow(0).LastCellNum;
            for (int i = 0; i <= lastCell; i++)
            {
                Sheet.AutoSizeColumn(i);
            }
        }

        /// <summary>
        /// Sets the cell value dynamically by checking the type of value.
        /// </summary>
        /// <param name="Cell">The cell to which you are setting the value.</param>
        /// <param name="value">Types: int, float, bool, string, DateTime, Guid etc. (Also supports nullable types)</param>
        private void PrintCellValue(ICell Cell, object value)
        {
            try
            {
                if (value == null)
                    Cell.SetCellValue(string.Empty);

                else
                {
                    Type t = value.GetType();
                    // Text
                    if (t.Equals(typeof(string)) || t.Equals(typeof(Guid)) || t.Equals(typeof(Guid?)))
                    {
                        Cell.SetCellValue(value.ToString());
                    }
                    // Number
                    else if (t.Equals(typeof(double)) || t.Equals(typeof(double?)) || t.Equals(typeof(int)) | t.Equals(typeof(int?)) || t.Equals(typeof(decimal)) || t.Equals(typeof(decimal?)) || t.Equals(typeof(float)) || t.Equals(typeof(float?)) || t.Equals(typeof(long)) || t.Equals(typeof(long?)))
                    {
                        Cell.SetCellValue(Convert.ToInt64(value));
                    }
                    // DateTime
                    else if (t.Equals(typeof(DateTime)) || t.Equals(typeof(DateTime?)))
                    {
                        Cell.SetCellValue(Convert.ToDateTime(value));
                    }
                    // Bool
                    else if (t.Equals(typeof(bool)) || t.Equals(typeof(bool?)))
                    {
                        Cell.SetCellValue(Convert.ToBoolean(value));
                    }
                }
            }
            catch (Exception e)
            {
                _logger?.LogError("Error at: ExcelManager.PrintCellValue | " + e.Message);
                throw new ApplicationException("Error at: ExcelManager.PrintCellValue | " + e.Message);
            }
        }

        private object ReadCellValue(ICell cell)
        {
            try
            {
                if (cell == null)
                    return null;

                var cellType = cell.CellType;

                if (cellType == CellType.String)
                    return cell.StringCellValue;

                if (cellType == CellType.Boolean)
                    return cell.BooleanCellValue;

                if (DateUtil.IsCellDateFormatted(cell))
                    return cell.DateCellValue;

                if (cellType == CellType.Numeric)
                    return cell.NumericCellValue;

                return cell.StringCellValue;
            }
            catch (Exception e)
            {
                _logger?.LogError("Error at: ExcelManager.PrintCellValue | " + e.Message);
                throw new ApplicationException("Error at: ExcelManager.PrintCellValue | " + e.Message);
            }
        }

        private string GetTypeDefaultFormat(Type t)
        {
            if (t.Equals(typeof(string)) || t.Equals(typeof(Guid)) || t.Equals(typeof(Guid?)))
                return null;

            if (t.Equals(typeof(int)) | t.Equals(typeof(int?)))
                return null;

            if (t.Equals(typeof(double)) || t.Equals(typeof(double?)) || t.Equals(typeof(decimal)) || t.Equals(typeof(decimal?)) || t.Equals(typeof(float)) || t.Equals(typeof(float?)) || t.Equals(typeof(long)) || t.Equals(typeof(long?)))
                return "0.00";

            if (t.Equals(typeof(DateTime)) || t.Equals(typeof(DateTime?)))
                return "dd-mm-yyyy hh:mm:ss";

            if (t.Equals(typeof(bool)) || t.Equals(typeof(bool?)))
                return null;

            return null;
        }

        private void WorkbookFinalProcessing()
        {
            // Force excel to recalculate all the formula while open
            XSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
            // Ensure the current Sheet is selected and the scroll is on top
            workbook.SetActiveSheet(workbook.GetSheetIndex(Sheet));
            Sheet.ActiveCell = NPOI.SS.Util.CellAddress.A1;
        }

        private void ReleaseMemory()
        {
            type = null;
            workbook = null;
            Sheet = null;
            format = null;
            headerCellStyle = null;
            bodyCellStyle = null;
            emptyCellStyle = null;
        }

        #region Style configurators

        private ICellStyle GetHeaderStyle(HeaderStyle styleConfig)
        {
            ICellStyle headerCellStyle = workbook.CreateCellStyle();
            //HSSFPalette palette = workbook.GetStylesSource();

            // Font
            IFont myFont = workbook.CreateFont();
            myFont.FontHeightInPoints = styleConfig.FontSize;
            myFont.FontName = styleConfig.FontFamily;
            myFont.IsBold = styleConfig.IsBold;
            myFont.IsItalic = styleConfig.IsItalic;
            myFont.Underline = styleConfig.IsUnderlined ? FontUnderlineType.Single : FontUnderlineType.None;
            if (styleConfig.FontColor != null)
            {
                //short index = (short)ColorIndex.HeaderFontColor;
                //byte r = styleConfig.FontColor[0];
                //byte g = styleConfig.FontColor[1];
                //byte b = styleConfig.FontColor[2];

                //palette.SetColorAtIndex(index, r, g, b);
                //myFont.Color = palette.GetColor(index).Indexed;
                myFont.Color = new XSSFColor(styleConfig.FontColor).Indexed;
            }
            headerCellStyle.SetFont(myFont);

            // Border
            if (styleConfig.IsBordered)
            {
                headerCellStyle.BorderTop = BorderStyle.Thin;
                headerCellStyle.BorderBottom = BorderStyle.Thin;
                headerCellStyle.BorderLeft = BorderStyle.Thin;
                headerCellStyle.BorderRight = BorderStyle.Thin;
            }

            // Background color
            if (styleConfig.BackgroundColor != null)
            {
                //short index = (short)ColorIndex.HeaderBgColor;
                //byte r = styleConfig.BackgroundColor[0];
                //byte g = styleConfig.BackgroundColor[1];
                //byte b = styleConfig.BackgroundColor[2];

                //palette.SetColorAtIndex(index, r, g, b);

                //headerCellStyle.FillForegroundColor = palette.GetColor(index).Indexed;
                headerCellStyle.FillBackgroundColor = IndexedColors.Yellow.Index;
                headerCellStyle.FillPattern = FillPattern.SolidForeground;
            }

            return headerCellStyle;
        }

        private ICellStyle GetDefaultHeaderStyle()
        {
            var defaultColor = new byte[] { 187, 255, 184};

            ICellStyle headerCellStyle = workbook.CreateCellStyle();
            //HSSFPalette palette = workbook.GetCustomPalette();

            // Font
            IFont myFont = workbook.CreateFont();
            myFont.FontHeightInPoints = 11;
            myFont.FontName = "Courier New";
            myFont.IsBold = true;
            myFont.IsItalic = false;
            myFont.Underline = FontUnderlineType.None;

            headerCellStyle.SetFont(myFont);

            // Border
            headerCellStyle.BorderTop = BorderStyle.Thin;
            headerCellStyle.BorderBottom = BorderStyle.Thin;
            headerCellStyle.BorderLeft = BorderStyle.Thin;
            headerCellStyle.BorderRight = BorderStyle.Thin;

            // Background color
            //short index = (short)ColorIndex.HeaderBgColor;
            //palette.SetColorAtIndex(index, defaultColor[0], defaultColor[1], defaultColor[2]);

            //headerCellStyle.FillForegroundColor = palette.GetColor(index).Indexed;
            headerCellStyle.FillForegroundColor = new XSSFColor(defaultColor).Indexed;
            headerCellStyle.FillPattern = FillPattern.SolidForeground;

            return headerCellStyle;
        }

        private ICellStyle GetBodyStyle(BodyStyle styleConfig)
        {
            ICellStyle bodyCellStyle = workbook.CreateCellStyle();
            //HSSFPalette palette = workbook.GetCustomPalette();

            // Font
            IFont myFont = workbook.CreateFont();
            myFont.FontHeightInPoints = styleConfig.FontSize;
            myFont.FontName = styleConfig.FontFamily;
            myFont.IsBold = styleConfig.IsBold;
            myFont.IsItalic = styleConfig.IsItalic;
            myFont.Underline = styleConfig.IsUnderlined ? FontUnderlineType.Single : FontUnderlineType.None;
            if (styleConfig.FontColor != null)
            {
                //short index = (short)ColorIndex.BodyFontColor;
                //byte r = styleConfig.FontColor[0];
                //byte g = styleConfig.FontColor[1];
                //byte b = styleConfig.FontColor[2];

                //palette.SetColorAtIndex(index, r, g, b);
                //myFont.Color = palette.GetColor(index).Indexed;
                myFont.Color = new XSSFColor(styleConfig.FontColor).Indexed;
            }
            bodyCellStyle.SetFont(myFont);

            // Border
            if (styleConfig.IsBordered)
            {
                bodyCellStyle.BorderTop = BorderStyle.Thin;
                bodyCellStyle.BorderBottom = BorderStyle.Thin;
                bodyCellStyle.BorderLeft = BorderStyle.Thin;
                bodyCellStyle.BorderRight = BorderStyle.Thin;
            }

            // Background color
            if (styleConfig.BackgroundColor != null)
            {
                //short index = (short)ColorIndex.BodyBgColor;
                //byte r = styleConfig.BackgroundColor[0];
                //byte g = styleConfig.BackgroundColor[1];
                //byte b = styleConfig.BackgroundColor[2];

                //palette.SetColorAtIndex(index, r, g, b);

                //bodyCellStyle.FillForegroundColor = palette.GetColor(index).Indexed;
                bodyCellStyle.FillForegroundColor = new XSSFColor(styleConfig.BackgroundColor).Indexed;
                bodyCellStyle.FillPattern = FillPattern.SolidForeground;
            }

            return bodyCellStyle;
        }

        private ICellStyle GetDefaultBodyStyle()
        {
            ICellStyle bodyCellStyle = workbook.CreateCellStyle();
            //HSSFPalette palette = workbook.GetCustomPalette();

            // Font
            IFont myFont = workbook.CreateFont();
            myFont.FontHeightInPoints = 11;
            myFont.FontName = "Courier New";
            myFont.IsBold = false;
            myFont.IsItalic = false;
            myFont.Underline = FontUnderlineType.None;

            bodyCellStyle.SetFont(myFont);

            return bodyCellStyle;
        }

        private ICellStyle GetEmptyStyle()
        {
            var defaultColor = new byte[] { 255, 255, 150 };

            ICellStyle emptyCellStyle = workbook.CreateCellStyle();
            //HSSFPalette palette = workbook.GetCustomPalette();

            //short index = (short)ColorIndex.HighlightColor;
            //palette.SetColorAtIndex(index, defaultColor[0], defaultColor[1], defaultColor[2]);

            //emptyCellStyle.FillForegroundColor = palette.GetColor(index).Indexed;
            emptyCellStyle.FillForegroundColor = new XSSFColor(defaultColor).Indexed;
            emptyCellStyle.FillPattern = FillPattern.SolidForeground;

            return emptyCellStyle;
        }

        #endregion

        #endregion
    }
}
