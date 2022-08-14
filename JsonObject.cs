namespace TestApp
{
    //Класс описывающий все поля нужного мне JSON файла
    //Этот класс не есть моё творение, его сгенерила VS
    internal class JsonObject
    {

        public class Rootobject
        {
            public Cu[] cus { get; set; }
        }

        public class Cu
        {
            public string code { get; set; }
            public string name { get; set; }
            public string region { get; set; }
            public int timezone { get; set; }
            public string type { get; set; }
            public string[] keproviders { get; set; }
            public string flags { get; set; }
            public string inputwsversion { get; set; }
            public Soun soun { get; set; }
            public Certificates certificates { get; set; }
            public string comment { get; set; }
            public string[] uftkproviders { get; set; }
            public string[] newcodes { get; set; }
        }

        public class Soun
        {
            public string shortname { get; set; }
            public string fullname { get; set; }
            public DateTime validfrom { get; set; }
            public string inn { get; set; }
            public string kpp { get; set; }
            public string address { get; set; }
            public string phone { get; set; }
            public string email { get; set; }
            public string document { get; set; }
            public string documentnum { get; set; }
            public DateTime documentdate { get; set; }
            public DateTime validto { get; set; }
        }

        public class Certificates
        {
        }

    }
}
