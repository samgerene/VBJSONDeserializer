namespace JSONGenerator
{
    using ServiceStack;
    using System.IO;

    class Program
    {
        static void Main(string[] args)
        {            
            var nrofobjects = 2;

            var pg = new PersonGenerator(nrofobjects);
            var json = pg.Persons.ToJson();
            var location = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var filename = string.Format("test - {0}.json", nrofobjects);
            var filepath = Path.Combine(location, filename);
            File.WriteAllText(filepath, json);
        }
    }
}
