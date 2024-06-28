using Aspose.Cells;
using Chitrich_1.Models;

namespace Chitrich_1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("");
            var list_0 = People.Read();
            People.Print(list_0);
            Console.WriteLine("===========================================================");
            var list_1 = People.SortByAge(list_0);
            People.Print(list_1);
            People.Save(list_1);
            People.TestSave(); //тестовий метод для перевірки зберігання
        }
    }
}
