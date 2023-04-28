using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AkelonTest2
{
    public class WorkExcel
    {
        private XLWorkbook workBook;
        private String filepath;
        public WorkExcel(String filepath)
        {
            this.filepath = filepath;
            workBook = new XLWorkbook(filepath);
        }
        public void ShowTable()
        {
            foreach (IXLWorksheet workSheet in workBook.Worksheets)
            {
                bool flagEnd = false;
                Console.Write("\n==========" + workSheet.Name + " ==========\n");
                foreach (IXLRow row in workSheet.Rows())
                {
                    foreach (IXLCell cell in row.Cells())
                    {
                        flagEnd = cell.IsEmpty();
                        Console.Write(" " + cell.Value.ToString() + " ");
                    }
                    if (flagEnd) break;
                    Console.WriteLine();
                }
            }
            
        }
        public void ShowSheet(String nameSheet)
        {
            IXLWorksheet workSheet = workBook.Worksheet(nameSheet);
                
            Console.Write("\n==========" + workSheet.Name + "==========\n");
            foreach (IXLRow row in workSheet.Rows())
            {
                foreach (IXLCell cell in row.Cells())
                {
                    if (cell.IsEmpty())
                    {
                        return;
                    }
                    Console.Write(" " + cell.Value.ToString() + " ");
                }
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Поиск номера данной колонки
        /// </summary>
        /// <param name="nameSheet">Лист</param>
        /// <param name="nameColumn">Колонка для поиска</param>
        /// <returns>Номер колонки</returns>
        private int FindColumn(String nameSheet, String nameColumn)
        { 
            IXLWorksheet workSheet = workBook.Worksheet(nameSheet);
            foreach (IXLRow row in workSheet.Rows())
            {
                int indexCol = 1;   
                foreach (IXLCell cell in row.Cells())
                {
                    if (cell.Value.ToString() == nameColumn) 
                    {
                        return indexCol; 
                    }
                    indexCol++;
                }
                return -1;
            }
            return -1;
        }

        /// <summary>
        /// Поиск индекса по строке с данным значением
        /// </summary>
        /// <param name="nameSheet">Лист</param>
        /// <param name="nameColumn">Колонка для поиска</param>
        /// <param name="value">Значение элемента в строке</param>
        /// <returns>Индекс с данных значением</returns>
        private int[] FindIndex(String nameSheet, String nameColumn, String value)
        {
            IXLWorksheet workSheet = workBook.Worksheet(nameSheet);
            int indexColumn = FindColumn(nameSheet, nameColumn);
            int[] index = { 0, indexColumn };

            foreach (IXLRow row in workSheet.Rows())
            {
                index[0]++;
                if (row.Cell(indexColumn).Value.ToString() == value)
                {
                    return index;
                }
            }
            return new int[] {};
        }
        
        /// <summary>
        /// Поиск всех индексов с данным значением в столбце
        /// </summary>
        /// <param name="nameSheet">Лист</param>
        /// <param name="nameColumn">Колонка для поиска</param>
        /// <param name="value">Значение элемента для поиска в столбце</param>
        /// <returns>Все индексы с данным значением</returns>
        private List<int[]> FindIndices(String nameSheet, String nameColumn, String value)
        {
            IXLWorksheet workSheet = workBook.Worksheet(nameSheet);
            int indexColumn = FindColumn(nameSheet, nameColumn);
            List<int[]> values = new List<int[]>();

            int count = 1;
            foreach (IXLRow row in workSheet.Rows())
            {
                int[] val = { 0, indexColumn };
                if (row.Cell(indexColumn).Value.ToString() == value)
                {
                    val[0] = count;
                    values.Add(val);
                }
                count++;
            }
            return values;
        }
        
        /// <summary>
        /// Поиск значения по строке в необходимой колонке
        /// </summary>
        /// <param name="nameSheet">Лист</param>
        /// <param name="nameColumn">Колонка для поиска</param>
        /// <param name="index">Индекс элемента для поиска строки</param>
        /// <returns>Значение в указанной колонке</returns>
        private string SearchString(String nameSheet, String nameColumn, int[] index)
        {
            IXLWorksheet workSheet = workBook.Worksheet(nameSheet);
            int indexCol = FindColumn(nameSheet, nameColumn);
            return workSheet.Cell(index[0], indexCol).Value.ToString();
        }

        /// <summary>
        /// Вывод информации о клиентах по наименованию товара
        /// с указанием информации по количеству товара, цене и дате заказа.
        /// </summary>
        /// <param name="nameProduct">Ниаменование продукта</param>
        public void InfoForSeller(String nameProduct)
        {
            String nameSheetProduct = "Товары";
            int[] indexProduct = FindIndex(nameSheetProduct, "Наименование", nameProduct);
            string productCode = SearchString(nameSheetProduct, "Код товара", indexProduct);
            int priceProduct = int.Parse(SearchString(nameSheetProduct, "Цена товара за единицу", indexProduct));

            Console.Write("Наименование товара: " + nameProduct + "\nИнформация о клиентах: \n");

            String nameSheetRequest = "Заявки";
            String nameSheetClient = "Клиенты";
            List<int[]> indexProductInRequests = FindIndices(nameSheetRequest, "Код товара", productCode);
            for(int count = 0; count < indexProductInRequests.Count; ++count)
            {
                int numberOfProduct = int.Parse(SearchString(nameSheetRequest, "Требуемое количество", indexProductInRequests[count]));
                string requestDate = SearchString(nameSheetRequest, "Дата размещения", indexProductInRequests[count]);
                int priceAllRequest = numberOfProduct * priceProduct;
                string clientCode = SearchString(nameSheetRequest, "Код клиента", indexProductInRequests[count]);

                int[] indexClient = FindIndex(nameSheetClient, "Код клиента", clientCode);

                string nameCompany = SearchString(nameSheetClient, "Наименование организации", indexClient);
                string contactPerson = SearchString(nameSheetClient, "Контактное лицо (ФИО)", indexClient);
                string address = SearchString(nameSheetClient, "Адрес", indexClient);

                Console.Write("\nКомпания: " + nameCompany + "\nКонтактное лицо: " + contactPerson + "\nАдрес: " + address
                    + "\nТребуемое количество: " + numberOfProduct + "\tЦена: " + priceAllRequest
                    + "\nДата размещения заявки: " + requestDate + "\n");
            }
        }

        /// <summary>
        /// Изменение контактного лица 
        /// с указанием названия организации и нового контактного лица
        /// </summary>
        /// <param name="nameCompany"></param>
        /// <param name="contactPersonNew"></param>
        public void ChangeClient(String nameCompany, String contactPersonNew)
        {
            String nameSheetClient = "Клиенты";
            IXLWorksheet workSheet = workBook.Worksheet(nameSheetClient);

            int[] indexNameCompany = FindIndex(nameSheetClient, "Наименование организации", nameCompany);
            string contactPerson = SearchString(nameSheetClient, "Контактное лицо (ФИО)", indexNameCompany);
            int[] indexPerson = FindIndex(nameSheetClient, "Контактное лицо (ФИО)", contactPerson);

            workSheet.Cell(indexPerson[0], indexPerson[1]).Value = contactPersonNew;

            Console.WriteLine(" Название организации " + nameCompany 
                + "\n Контактное лицо  " + contactPerson 
                + "заменено на: " + workSheet.Cell(indexPerson[0], indexPerson[1]).Value.ToString());
            workBook.Save();
        }

        /// <summary>
        /// Определение золотого клиента за указанный год, месяц.
        /// </summary>
        /// <param name="year">Год</param>
        /// <param name="month">Месяц</param>
        public void GoldenClient(string year = "", string month = "")
        {
            string MY = month + "." + year;
            String nameSheetRequest = "Заявки";
            IXLWorksheet workSheet = workBook.Worksheet(nameSheetRequest);
            int indexColumn = FindColumn(nameSheetRequest, "Код клиента");
            Console.WriteLine("Золотой клиент");
            
            // находим коды одинаковых клиентов
            List<string> clients = new List<string>();
            foreach (IXLRow row in workSheet.Rows())
            {
                int client = 0;
                if (Int32.TryParse(row.Cell(indexColumn).Value.ToString(), out client))
                {
                    client = Convert.ToInt32(row.Cell(indexColumn).Value.ToString());
                    clients.Add(client.ToString());
                }
            }
            clients = new HashSet<string>(clients).ToList();
            // подсчет заказов
            string clientGold = "";
            int maxRequests = 0;
            foreach(var client in clients)
            {
                int sum = 0;
                List<int[]> requests = FindIndices(nameSheetRequest, "Код клиента", client.ToString());
                foreach (var request in requests)
                {
                    if(SearchString(nameSheetRequest, "Дата размещения", request).Contains(MY))
                    {
                        string numberOfProduct = SearchString(nameSheetRequest, "Требуемое количество", request);
                        sum += int.Parse(numberOfProduct);
                    }
                }
                if(maxRequests < sum)
                {
                    maxRequests = sum;
                    clientGold = client;
                }
            }
            Console.WriteLine(clientGold);
        }




        ~WorkExcel() 
        {
            //excelFile.Close();
            workBook.Dispose();
        }
    }




    
}
