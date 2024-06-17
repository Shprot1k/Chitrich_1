using Aspose.Cells;
using Chitrich_1.Models;

namespace Chitrich_1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("");
            var list = People.Read();
            People.Write(list);
        }
    }
}
