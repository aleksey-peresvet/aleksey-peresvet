using Newtonsoft.Json;//Стороннее решение для работы с JSON

namespace TestApp
{
    internal class TestApp
    {
        static readonly HttpClient client = new HttpClient();
        static async Task Main()
        {
            string responseBody = "";
            try
            {
                //Пытаюсь получить тело ответа(JSON) функции удаленного api 
                responseBody = 
                    await client.GetStringAsync("https://api.kontur.ru/dc.contacts/v1/cus");
            }
            catch (HttpRequestException e)
            {
                Console.WriteLine("\nException Caught!");
                Console.WriteLine("Message :{0} ", e.Message);
            }
            //Конвертация JSON в объект С#
            var rootjson = JsonConvert.DeserializeObject<JsonObject.Rootobject>(responseBody);
            MySelection selection = new MySelection();
            IEnumerable<JsonObject.Cu> temp_list = Enumerable.Empty<JsonObject.Cu>();
            if (rootjson != null)
            {
                temp_list =
                from json in rootjson.cus
                where json.region == "18"
                select json;
            }
            foreach (var data in temp_list)
            {
                if (data.type == null) data.type = "";
                if (data.code == null) data.code = "";
                if (data.name == null) data.name = "";
                if (data.soun.inn == null) data.soun.inn = "";
                selection.AddNewRow(
                        data.type.ToString(),
                        data.code.ToString(),
                        data.name.ToString(),
                        data.soun.inn.ToString());
            }
            //Начинается работа с экспортом выборки в файл эксцель
            //Снова использую сторонее решение
            int row_indent = 6;
            ExcelHelper excelHelper = new ExcelHelper();
            for(int row=0; row<selection.RowCount(); row++)
            {
                for (int column = 0; column < 4; column++)
                {
                    //Пишу данные выборки в таблицу эксцель
                    excelHelper.WriteToCell(row + row_indent, column+1, 
                        selection.GetRow(row)[column]);
                }
            }
            //Шапка выборки
            excelHelper.SortByColumns("A,B");
            excelHelper.WriteToCell(1, 1, "Дата выполнения: " + DateTime.Now.ToString());
            excelHelper.WriteToCell(2, 1, "Количество контролирующих органов:");
            excelHelper.WriteToCell(3, 1, "Общее - " + selection.RowCount());
            for (int column = 0; column < selection.GetTypesCount(); column++)
            {
                excelHelper.WriteToCell(4, column + 1, selection.GetTypes()[column]);
            }
            excelHelper.WriteToCell(5, 1, "Тип");
            excelHelper.WriteToCell(5, 2, "Код");
            excelHelper.WriteToCell(5, 3, "Имя");
            excelHelper.WriteToCell(5, 4, "ИНН");
            //Сохраняю файл в папке с *.exe файлом проекта
            excelHelper.SaveFile("Result.xlsx");
        }
    }
}