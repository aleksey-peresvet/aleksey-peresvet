namespace TestApp
{
    //Класс описывающий все поля нужного мне JSON файла
    //Этот класс не есть моё творение, его сгенерила VS
    internal class JsonObject
    {

        public class Rootobject
        {
            public Cu[] cus { get; set; }
            Rootobject()
            {
                cus = new Cu[0];
            }
        }

        public class Cu
        {
            public string code { get; set; }
            public string name { get; set; }
            public string region { get; set; }
            public string type { get; set; }
            public Soun soun { get; set; }

            public Cu()
            {
                this.soun = new Soun();
                code = string.Empty;
                name = string.Empty;
                type = string.Empty;
                region = string.Empty;
            }
        }

        public class Soun
        {
            public string inn { get; set; }

            public Soun()
            {
                inn = string.Empty;
            }
        }
    }
}
