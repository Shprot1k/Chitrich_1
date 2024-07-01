using Chitrich_1.Models;

namespace Chitrich_1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            /*
            Console.Write("Name of the file to be read: ");
            //var peoples_0 = PeopleBL.Read(Console.ReadLine());
            var peoples_0 = PeopleBL.Read("generated_excel_data.xlsx");
            PeopleBL.Print(peoples_0);
            Console.WriteLine("===========================================================");
            var peoples_1 = PeopleBL.SortByAge(peoples_0);
            PeopleBL.Print(peoples_1);
            PeopleBL.Save(peoples_1);
            */
            var peoples_0 = PeopleBL.Read("generated_excel_data.xlsx");
            PeopleBL.Print(peoples_0);
        }
    }
}