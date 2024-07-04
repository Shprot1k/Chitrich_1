using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Chitrich_1.Models;


namespace Chitrich_1.Services
{
    internal class SharedServices
    {
        public static void Print(List<People> peoples)
        {
            Console.WriteLine();
            foreach (People people in peoples)
            {
                Console.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}", people.Id, people.Name, people.Age, people.Salary, people.Department);
            }
        }
        public static List<People> Sort(List<People> peoples)
        {
            Console.Write("\nSelect a sorting method:\n"
                + "1 - By Id\n"
                + "2 - By Name (dont work)\n"
                + "3 - By Age\n"
                + "4- By Salare\n"
                + "5 - By Department (dont work)\n"
                + "0 - Don't sort\n"
                + "Your choice: ");
            int sortOption = int.Parse(Console.ReadLine());
            switch (sortOption)
            {
                case 0: // Don't sort
                    return peoples;

                case 1: // Id
                    for (int i = 0; i < peoples.Count; i++)
                    {
                        for (int j = 0; j < peoples.Count; j++)
                        {
                            if (peoples[i].Id < peoples[j].Id)
                            {
                                var temp = peoples[i];
                                peoples[i] = peoples[j];
                                peoples[j] = temp;
                            }
                        }
                    }
                    return peoples;

                case 2: // By Name (dont work)
                    return peoples;

                case 3: // By Age
                    for (int i = 0; i < peoples.Count; i++)
                    {
                        for (int j = 0; j < peoples.Count; j++)
                        {
                            if (peoples[i].Age < peoples[j].Age)
                            {
                                var temp = peoples[i];
                                peoples[i] = peoples[j];
                                peoples[j] = temp;
                            }
                        }
                    }
                    return peoples;

                case 4: // By Salare
                    for (int i = 0; i < peoples.Count; i++)
                    {
                        for (int j = 0; j < peoples.Count; j++)
                        {
                            if (peoples[i].Salary < peoples[j].Salary)
                            {
                                var temp = peoples[i];
                                peoples[i] = peoples[j];
                                peoples[j] = temp;
                            }
                        }
                    }
                    return peoples;

                case 5: //By Department (dont work)
                    return peoples;
                default:
                    return peoples;


            }
        }
    }
}
