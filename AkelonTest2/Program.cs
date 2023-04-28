using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Security.Cryptography;

namespace AkelonTest2
{
    
    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine(" Выберите Excel-файл.");
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            ofd.Title = "Выберите документ Excel";
            if (ofd.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Вы не выбрали файл для открытия", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string xlFileName = ofd.FileName;

            WorkExcel we = new WorkExcel(xlFileName);
            Menu:
            Console.Write("\n Меню.\n" 
                + " 1. Показать все листы\n" 
                + " 2. Показать лист\n" 
                + " 3. Информация о клиентах по товару\n" 
                + " 4. Изменить контактное лицо\n" 
                + " 5. Золотой клиент\n" 
                + " 6. Выход\n" 
                + " Введите номер:");
            string choice = Console.ReadLine();
            Console.WriteLine();
            switch (choice)
            {
                case "1":
                    we.ShowTable();
                    goto Menu;
                case "2":
                    Console.Write("\n Введите название листа: ");
                    string nameSheet = Console.ReadLine();
                    try { we.ShowSheet(nameSheet); } catch { goto default; }
                    goto Menu;
                case "3":
                    Console.Write("\n Введите название продукта: ");
                    string nameProduct = Console.ReadLine();
                    try { we.InfoForSeller(nameProduct); } catch { goto default; }
                    goto Menu;
                case "4":
                    Console.Write("\n Введите название организации: ");
                    string nameCompany = Console.ReadLine();
                    Console.Write("\n Введите новое контактное лицо: ");
                    string contactPersonNew = Console.ReadLine();
                    try { we.ChangeClient(nameCompany, contactPersonNew); } catch { goto default; }
                    goto Menu;
                case "5": //ЗАТЕСТИТЬ
                    Console.Write("\n Если не хотите вводить год или месяц, просто пропустите ввод клавишей Enter." +
                        "\n Введите год: ");
                    string year = Console.ReadLine();
                    Console.Write("\n Введите месяц: ");
                    string month = Console.ReadLine();                    
                    try { we.GoldenClient(year, month); } catch { goto default; }
                    goto Menu;
                case "6":
                    Environment.Exit(0);
                    break;
                default:
                    Console.WriteLine("\n Неправильный ввод");
                    goto Menu;
            }
        }
    }   
}

