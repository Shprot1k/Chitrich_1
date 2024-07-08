using Chitrich_1.Models;
using Chitrich_1.Services;

namespace Chitrich_1.Models
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<People> p = new List<People>();
            List<Car> c = new List<Car>();
            int choise = 1;
            while (choise != 0)
            {
                Console.Write("Select an action\n"
                    + "1 - Read file\n"
                    + "2 - Sort\n"
                    + "3 - Print\n"
                    + "4 - Save\n"
                    + "5 - Clear console\n"
                    + "6 - Read Cars\n"
                    + "7 - Print Cars\n"
                    + "0 - Exit\n"
                    + "Your choise: ");
                choise = int.Parse(Console.ReadLine());

                switch (choise)
                {
                    case 1: // Read peoples
                        Console.WriteLine(@"C:\Users\taras\source\repos\Chitrich_1\Chitrich_1\Files\generated_excel_data.xlsx");
                        Console.Write("Name of the file to be read: ");
                        p = ExelServices.Read<People>(Console.ReadLine());
                        Console.WriteLine("File read\n");
                        break;
                    case 2: // Sort
                        p = SharedServices.Sort(p);
                        Console.WriteLine("List sorted\n");
                        break;
                    case 3: // Print
                        SharedServices.PrintPeoples(p);
                        Console.WriteLine();
                        break;
                    case 4: // Save
                        ExelServices.Save(p);
                        Console.WriteLine("File saved\n");
                        break;
                    case 5: // Clear console
                        Console.Clear();
                        break;
                    case 6: // Read cars 
                        Console.WriteLine(@"C:\Users\taras\source\repos\Chitrich_1\Chitrich_1\Files\CarData.xlsx");
                        Console.Write("Name of the file to be read: ");
                        c = ExelServices.Read<Car>(Console.ReadLine());
                        Console.WriteLine("File read\n");
                        break;
                    case 7: // Print cars
                        SharedServices.PrintCars(c);
                        Console.WriteLine();
                        break;
                    case 0: // Exit
                        break;
                    default:
                        break;
                }
            }
        }
    }
}