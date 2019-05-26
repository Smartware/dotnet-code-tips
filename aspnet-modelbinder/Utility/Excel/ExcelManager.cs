using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace aspnet_modelbinder.Utility.Excel
{
    public class ExcelManager
    {
        /// <summary>
        /// Get or Set Filepath
        /// </summary>
        //public String FilePath { get; set; }
        private IWorkbook _workbook;
        private readonly Stream _stream;
        private readonly IList<ISheet> _readableSheets = new List<ISheet>();

        public ExcelManager(IWorkbook workbook)
        {
            _workbook = workbook;
        }

        public ExcelManager(Stream stream)
        {
            _stream = stream;
            InitializeWorkBook();
        }

        public ExcelManager(String filePath)
            : this(new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
        }

        public void StyleSheetRecord(IWorkbook workBook, IFont headerFont, ICellStyle headerCellStyle)
        {
            var sheet = workBook.GetSheetAt(0);

            if (sheet == null)
                return;

            var headerRow = sheet.GetRow(0);

            for (int i = 0; i < headerRow.Count(); i++)
            {
                sheet.AutoSizeColumn(i);
                headerRow.Cells[i].CellStyle = headerCellStyle;
                headerRow.Cells[i].CellStyle.SetFont(headerFont);
            }
        }

        public IWorkbook ExportListToXLSX<T>(List<T> dataList
            , params Expression<Func<T, object>>[] excelHeaderExtracts)
        {
            if (dataList == null)
            {
                throw new ArgumentNullException("Invalid argument.");
            }

            DataTable shadowRecords = new DataTable();

            foreach (var item in excelHeaderExtracts)
            {
                if (item.Body is MemberExpression)
                {
                    shadowRecords.Columns.Add(((MemberExpression)item.Body).Member.Name, item.Body.Type); ;
                }
                else
                {
                    Expression currentExpression = item.Body;
                    var name = GetNameFromExpression(item, currentExpression);
                    shadowRecords.Columns.Add(name, currentExpression.Type);
                }
            }

            var sheet = InitializeWorkSheet(excelHeaderExtracts);

            int rowIndex = 0;
            int cellIndex = -1;

            PropertyInfo[] propInfos = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (T item in dataList)
            {
                var row = (XSSFRow)sheet.CreateRow(++rowIndex);
                foreach (var t in excelHeaderExtracts)
                {
                    string name;

                    if (t.Body is MemberExpression)
                    {
                        var expression = (MemberExpression)t.Body;
                        name = expression.Member.Name;
                    }
                    else
                    {
                        Expression currentExpression = t.Body;
                        name = GetNameFromExpression(t, currentExpression);
                    }

                    var propertyRefl = propInfos.FirstOrDefault(prop => prop.Name.Equals(name));

                    if (propertyRefl != null)
                    {
                        var cell = (XSSFCell)row.CreateCell(++cellIndex);
                        var itemVal = propertyRefl.GetValue(item);

                        if (itemVal != null)
                            cell.SetCellValue(itemVal.ToString());
                        else
                            cell.SetCellValue("[NA]");
                    }
                }
                cellIndex = -1;
            }

            return _workbook;
        }

        private static String GetNameFromExpression<T>(Expression<Func<T, object>> item, Expression currentExpression)
        {
            while (true)
            {
                switch (currentExpression.NodeType)
                {
                    case ExpressionType.Parameter:
                        return ((ParameterExpression)currentExpression).Name;

                    case ExpressionType.MemberAccess:
                        return ((MemberExpression)currentExpression).Member.Name;

                    case ExpressionType.Call:
                        return ((MethodCallExpression)item.Body).Method.Name;
                    case ExpressionType.Convert:
                    case ExpressionType.ConvertChecked:
                        currentExpression = ((UnaryExpression)currentExpression).Operand;
                        break;
                    case ExpressionType.Invoke:
                        currentExpression = ((InvocationExpression)currentExpression).Expression;
                        break;
                    case ExpressionType.Not:
                        currentExpression = ((UnaryExpression)currentExpression).Operand;
                        break;
                    default:
                        return "[NA]";
                }
            }
        }

        private ISheet InitializeWorkSheet<T>(Expression<Func<T, object>>[] excelHeaderExtracts, String sheetName = "Sheet 1")
        {
            var sheet = (XSSFSheet)_workbook.CreateSheet(sheetName);

            //make a header row
            var headerRow = (XSSFRow)sheet.CreateRow(0);
            int columnIndex = -1;

            foreach (var item in excelHeaderExtracts)
            {
                string name;

                if (item.Body is MemberExpression)
                {
                    var expression = (MemberExpression)item.Body;
                    name = expression.Member.Name;
                    var headerCell = (XSSFCell)headerRow.CreateCell(++columnIndex);
                    headerCell.SetCellValue(name);
                }
                else
                {
                    Expression currentExpression = item.Body;
                    name = GetNameFromExpression(item, currentExpression);
                    var headerCell = (XSSFCell)headerRow.CreateCell(++columnIndex);
                    headerCell.SetCellValue(name);
                }
            }

            return sheet;
        }

        private DataTable InitializeDataTable<T>(Expression<Func<T, object>>[] excelHeaderExtracts)
        {
            DataTable table = new DataTable();

            foreach (var item in excelHeaderExtracts)
            {
                string name;

                if (item.Body is MemberExpression)
                {
                    var expression = (MemberExpression)item.Body;
                    name = expression.Member.Name;
                    table.Columns.Add(name, expression.Type);
                }
                else
                {
                    var op = ((UnaryExpression)item.Body).Operand;
                    var mem = ((MemberExpression)op).Member;
                    name = mem.Name;
                    table.Columns.Add(name, op.Type);
                }
            }

            return table;
        }

        public IEnumerable<T> ReadSheetAtToList<T>(Int32 sheetAtIndex = 0, Boolean skipHeader = true, params Expression<Func<T, object>>[] excelHeaderExtracts)
        {
            DataTable shadowRecords = new DataTable();
            foreach (var item in excelHeaderExtracts)
            {
                if (item.Body is MemberExpression)
                {
                    shadowRecords.Columns.Add(((MemberExpression)item.Body).Member.Name, item.Body.Type); ;
                }
                else if (item.Body is MemberInitExpression)
                {
                    var memExpr = (item.Body as MemberInitExpression);
                    foreach (var modelProperty in memExpr.Bindings)
                    {
                        var properInfo = modelProperty.Member as PropertyInfo;
                        shadowRecords.Columns.Add(modelProperty.Member.Name, properInfo.PropertyType);
                    }
                }
                else
                {
                    var op = ((UnaryExpression)item.Body).Operand;
                    var mem = ((MemberExpression)op).Member;
                    shadowRecords.Columns.Add(mem.Name, op.Type); ;
                }
            }

            CopyToDataTable(skipHeader, sheetAtIndex, shadowRecords);

            List<T> recordSet = new List<T>();

            Array.ForEach(shadowRecords.Select(), row =>
            {
                var instance = Activator.CreateInstance<T>();
                foreach (var item in excelHeaderExtracts)
                {
                    string name;

                    if (item.Body is MemberExpression)
                    {
                        var expression = (MemberExpression)item.Body;
                        name = expression.Member.Name;
                        var cell = row[expression.Member.Name].ToString();

                        var property = instance.GetType().GetProperties()
                        .FirstOrDefault(p => p.Name == expression.Member.Name);

                        if (property.PropertyType == typeof(int))
                        {
                            Int32.TryParse(cell, out int parseValue);
                            property.SetValue(instance, parseValue);
                        }
                        else
                        {
                            property.SetValue(instance, cell);
                        }
                    }
                    else if (item.Body is MemberInitExpression)
                    {
                        var memExpr = (item.Body as MemberInitExpression);
                        foreach (var modelProperty in memExpr.Bindings)
                        {
                            var properInfo = modelProperty.Member as PropertyInfo;
                            var property = instance.GetType().GetProperties().FirstOrDefault(p => p.Name == modelProperty.Member.Name);

                            var cell = row[modelProperty.Member.Name].ToString();

                            if (property.PropertyType == typeof(int))
                            {
                                Int32.TryParse(cell, out int parseValue);
                                property.SetValue(instance, parseValue);
                            }
                            else
                            {
                                property.SetValue(instance, cell);
                            }
                        }
                    }
                    else
                    {
                        var op = ((UnaryExpression)item.Body).Operand;
                        var mem = ((MemberExpression)op).Member;
                        var cell = row[mem.Name].ToString();
                        var property = instance.GetType().GetProperties().FirstOrDefault(p => p.Name == mem.Name);

                        if (property.PropertyType == typeof(int))
                        {
                            Int32.TryParse(cell, out int parseValue);
                            property.SetValue(instance, parseValue);
                        }
                        else
                        {
                            property.SetValue(instance, cell);
                        }
                    }

                }

                recordSet.Add(instance);
            });

            return recordSet;
        }

        private void CopyToDataTable(Boolean skipHeader, Int32 sheetIndex, DataTable shadowRecord)
        {
            GetSheets(sheetIndex);
            ISheet sheet = _readableSheets[sheetIndex];
            var rows = sheet.GetRowEnumerator();

            if (skipHeader)
                rows.MoveNext();

            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                DataRow dr = shadowRecord.NewRow();

                for (int i = 0; i < row.LastCellNum; i++)
                {
                    ICell cell = row.GetCell(i);

                    if (cell == null)
                    {
                        dr[i] = null;
                    }
                    else
                    {
                        dr[i] = cell.ToString();
                    }
                }

                shadowRecord.Rows.Add(dr);
            }
        }

        /// <summary>
        /// Read all sheet in an excel file to the class specified
        /// </summary>
        /// <param name="predicate">A function returning a value for an unspecified column property</param>
        /// <param name="sheetsIndexes">specify sheet indexes to read</param>
        /// <returns></returns>
        public Dictionary<String, IEnumerable<T>> ReadAllSheetsToList<T>(int[] sheetsIndexes, Boolean skipHeader = true, params Expression<Func<T, object>>[] excelHeaderExtracts)
        {
            Dictionary<String, IEnumerable<T>> groupList = new Dictionary<String, IEnumerable<T>>();
            foreach (var sheetIndex in sheetsIndexes)
            {
                var dataList = ReadSheetAtToList(sheetIndex, skipHeader, excelHeaderExtracts);
                string sheetName = _readableSheets[sheetIndex].SheetName;
                groupList[sheetName] = dataList;
            }

            return groupList;
        }

        private void InitializeWorkBook()
        {
            if (_stream != null)
                _workbook = WorkbookFactory.Create(_stream);
        }

        private void GetSheets(params int[] sheetsIndexes)
        {
            if (sheetsIndexes.Any())
            {
                for (int i = 0; i < sheetsIndexes.Length; i++)
                {
                    var sheet = _workbook.GetSheetAt(sheetsIndexes[i]);

                    if (_readableSheets.FirstOrDefault(t => t.SheetName == sheet.SheetName) != null)
                        continue;

                    _readableSheets.Add(sheet);
                }
            }
            else
            {
                for (int i = 0; i < sheetsIndexes.Length; i++)
                {
                    var sheet = _workbook.GetSheetAt(i);

                    if (_readableSheets.FirstOrDefault(t => t.SheetName == sheet.SheetName) != null)
                        continue;

                    _readableSheets.Add(sheet);
                }
            }
        }
    }
}