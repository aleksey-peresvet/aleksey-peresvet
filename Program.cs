using Newtonsoft.Json;//Стороннее решение для работы с JSON
using System.Linq;

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
            List<List<string>> result_list = new();
            foreach (var json_block in rootjson.cus)
            {
                //Делаю выборку по региону
                if (json_block.region == "18")
                {
                    if (json_block.type == null) json_block.type = "";
                    if (json_block.code == null) json_block.code = "";
                    if (json_block.name == null) json_block.name = "";
                    if (json_block.soun.inn == null) json_block.soun.inn = "";
                    //Пишу выборку в двумерный список
                    List<string> temp_list = new() {
                        json_block.type.ToString(),
                        json_block.code.ToString(),
                        json_block.name.ToString(),
                        json_block.soun.inn.ToString() };
                    result_list.Add(temp_list);
                }
            }
            //Начинается работа с экспортом выборки в файл эксцель
            //Снова использую сторонее решение
            int row = 5, column = 0, count = 0;
            ExcelHelper excelHelper = new ExcelHelper();
            //Создаю словарь для записи количества типов органов
            Dictionary<string, int> types = new Dictionary<string, int>();

            foreach (var sub_list in result_list)
            {
                row++;
                foreach (var data in sub_list)
                {
                    column++;
                    if (column == 1)
                    {
                        if (types == null)
                        {
                            types.Add(data, 1);
                        } 
                        else
                        {
                            if (types.Keys.Contains(data))
                            {
                                count = types.GetValueOrDefault(data);
                                count++;
                                types.Remove(data);
                                types.Add(data, count);
                            }
                            else
                            {
                                types.Add(data, 1);
                            }
                        }
                        
                    }
                    //Пишу данные выборки в таблицу эксцель
                    excelHelper.WriteToCell(row, column, data);
                }
                column = 0;
            }
            //Сортирую словарь типов, чтобы привести его к последовательности
            //получившейся выборки
            types = 
                types.OrderBy(x => x.Key).ToDictionary(x => x.Key, x => x.Value);
            //Шапка выборки
            excelHelper.SortByColumns("A,B");
            excelHelper.WriteToCell(1, 1, "Дата выполнения: " + DateTime.Now.ToString());
            excelHelper.WriteToCell(2, 1, "Количество контролирующих органов:");
            excelHelper.WriteToCell(3, 1, "Общее - " + result_list.Count);
            string splitter = ", ", row_4 = "";
            for (int i = 0; i < types.Count; i++)
            {
                if (types.Count - i == 1) splitter = "";
                row_4 += types.Keys.ElementAt(i).ToString() + 
                    " - " + types.Values.ElementAt(i) + splitter;
            }
            excelHelper.WriteToCell(4, 1, row_4);
            excelHelper.WriteToCell(5, 1, "Тип");
            excelHelper.WriteToCell(5, 2, "Код");
            excelHelper.WriteToCell(5, 3, "Имя");
            excelHelper.WriteToCell(5, 4, "ИНН");
            //Сохраняю файл в папке с *.exe файлом проекта
            excelHelper.SaveFile("Result.xlsx");
        }
    }
}