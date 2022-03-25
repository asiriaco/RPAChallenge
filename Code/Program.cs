using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace SeleniumRPA
{
    public class Excel
    {
        string path = " ";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        public Excel(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];

        }

        public string readCell(int i, int j)    
        {
            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2;
            else
                return "";


        }
        
        public string readNConvert(int i, int j)
        {
            if (ws.Cells[i, j].Value2 != null)
                return Convert.ToString(ws.Cells[i, j].Value2);
            else
                return "";
        }

   
    }
    public class Program
    {
        private static IWebDriver driver = new ChromeDriver();
        
        
        
        

        static void Main()
        {
            int nTuples = 11;
            string path = @"C:\Users\fiska\Downloads\challenge.xlsx";

            Excel table = new Excel(path, 1);
            string[] data = new string[7];
            int line = 2;

            driver.Navigate().GoToUrl("https://www.rpachallenge.com/");
            driver.FindElement(By.XPath("/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button")).Click();

            while (line <= nTuples)
            {
                for (int i = 1; i <= 6; i++)
                {
                    data[i - 1] = table.readCell(line, i);

                }
                data[6] = table.readNConvert(line, 7);
                
                Fill(data);
                line++;
            }


            

            
        }

        

        static void Fill( String[] data)
        {
            
            driver.FindElement(By.CssSelector("input[ng-reflect-name='labelFirstName']")).SendKeys(data[0]);
            driver.FindElement(By.CssSelector("input[ng-reflect-name='labelLastName']")).SendKeys(data[1]);
            driver.FindElement(By.CssSelector("input[ng-reflect-name='labelCompanyName']")).SendKeys(data[2]);
            driver.FindElement(By.CssSelector("input[ng-reflect-name='labelRole']")).SendKeys(data[3]);
            driver.FindElement(By.CssSelector("input[ng-reflect-name='labelAddress']")).SendKeys(data[4]); ;
            driver.FindElement(By.CssSelector("input[ng-reflect-name='labelEmail']")).SendKeys(data[5]);
            driver.FindElement(By.CssSelector("input[ng-reflect-name='labelPhone']")).SendKeys(data[6]);
            driver.FindElement(By.XPath("/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input")).Click();


        }
    }
}