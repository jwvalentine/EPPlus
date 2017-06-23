using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;



namespace EPPlus
{

    public class Account
    {
        public int ID { get; set; }
        public double Balance { get; set; }     
        public Account(int givenID, double givenBalance)
        {
            ID = givenID;
            Balance = givenBalance;
        }
    }

    class Program
    {
        static List<Account> accountList = new List<Account>();


        public static void Main(string[] args)
        {

            using (ExcelPackage myPackage = new ExcelPackage())
            {
                //Create a new Worksheet
                ExcelWorksheet mySheet = myPackage.Workbook.Worksheets.Add("Sheet 1");

                //Add data to the Account List | See fillAccounts() Constructor
                fillAccounts();

                //Add some text to cell A1 to test
                mySheet.Cells["A1"].LoadFromCollection(accountList, true);
                
                //Set path to the file
                string filePath = "C:\\test\\test1.xlsx";

                //Write file to disk
                FileInfo file = new FileInfo(filePath);
                myPackage.SaveAs(file);
            }
        }


        // fillAccounts Constructor
        public static void fillAccounts()
        {
            accountList.Add(new Account(1, 10.0));
            accountList.Add(new Account(2, 30.7));
            accountList.Add(new Account(3, 70.3));
            accountList.Add(new Account(4, 14.3));
            accountList.Add(new Account(5, 10.0));
            accountList.Add(new Account(6, 10.0));
            accountList.Add(new Account(7, 10.0));
            accountList.Add(new Account(8, 10.0));
            accountList.Add(new Account(9, 10.0));
        }
    }
}







