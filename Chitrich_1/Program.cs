using Chitrich_1.Models;

namespace Chitrich_1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<People> p = new List<People>();
            int choise = 1;
            while (choise != 0)
            {
                Console.Write("Select an action\n"
                    + "1 - Read file   generated_excel_data.xlsx\n"
                    + "2 - Sort\n"
                    + "3 - Print\n"
                    + "4 - Save\n"
                    + "5 - Clear console\n"
                    + "0 - Exit\n"
                    + "Your choise: ");
                choise = int.Parse(Console.ReadLine());

                switch (choise)
                {
                    case 1: // Read file
                        Console.Write("Name of the file to be read: ");
                        p = PeopleBL.Read(Console.ReadLine());
                        Console.WriteLine("File read\n");
                        break;
                    case 2: // Sort
                        p = PeopleBL.Sort(p);
                        Console.WriteLine("List sorted\n");
                        break;
                    case 3: // Print
                        PeopleBL.Print(p);
                        Console.WriteLine();
                        break;
                    case 4: // Save
                        PeopleBL.Save(p);
                        Console.WriteLine("File saved\n");
                        break;
                    case 5: // Clear console
                        Console.Clear();
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