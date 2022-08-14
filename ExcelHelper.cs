using Excel = ClosedXML.Excel;//Стороннее решение для создания файлов xlsx

namespace TestApp
{
    public class ExcelHelper
    {
        Excel.XLWorkbook workbook;
        Excel.IXLWorksheet worksheet;

        public ExcelHelper()
        {
            //В конструкторе создается worksheet
            workbook = new Excel.XLWorkbook();
            worksheet = workbook.AddWorksheet("Результат выборки");
        }

        internal void WriteToCell(int row, int column, object data)
        {
            //Пишу данные в ячейки worksheet
            worksheet.Cell(row, column).SetValue(data);
        }

        internal void SortByColumns(string columns)
        {
            //Сортировка по столбцам
            worksheet.Sort(columns);
        }

        internal void SaveFile(string fileName)
        {
            var path = Path.Combine(Environment.CurrentDirectory, fileName);
            if (File.Exists(path)) File.Delete(path);
            //Пишу worksheet в файл
            workbook.SaveAs(path);
            workbook.Dispose();
        }
    }
}