namespace TestApp
{
    internal class MySelection
    {
        private List<string[]> rows;

        public MySelection()
        {
            rows = new List<string[]>();
        }
        public void AddNewRow(string _type, string _code,
            string _name, string _inn)
        {
            rows.Add(new[] { _type, _code, _name, _inn });
        }

        public List<string> GetColumn(int column_number)
        {
            List<string> result = new List<string>();
            foreach (var row in rows)
            {
                result.Add(row[column_number]);
            }
            return result;
        }

        public List<string> GetRow(int row_number)
        {
            return rows[row_number].ToList<string>();
        }

        public int RowCount()
        {
            return rows.Count;
        }

        public List<string> GetTypes()
        {
            var types =  GetColumn(0).Distinct().ToList<string>();
            types.Sort();
            var result = new List<string>();
            foreach(var type in types)
            {
                var type_count = from cell in GetColumn(0)
                                 where cell == type
                                 select cell;
                result.Add(type + " - " + type_count.Count().ToString());
            }
            return result;
        }

        public int GetTypesCount()
        {
            return GetColumn(0).Distinct().ToList<string>().Count;
        }
    }
}
