using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Net.Mail;

namespace pricemonitor
{
    class Program
    {
        List<bontique> notificationlist = new List<bontique>();
        static void Main(string[] args)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(@"C:\Git\SSENSE_Price_checker\data.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkBook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            List<bontique> buylist = new List<bontique>();
            for (int i = 1; i <= rowCount; i++)
            {
                bontique tmp = new bontique() { row = i, url = xlRange.Cells[i, 1].Text, lowest= Convert.ToDouble(xlRange.Cells[i, 2].Text), highest= Convert.ToDouble(xlRange.Cells[i, 3].Text), last = Convert.ToDouble(xlRange.Cells[i, 4].Text) };
                buylist.Add(tmp);
            }
            IWebDriver driver = new ChromeDriver(@"C:\Git\SSENSE_Price_checker");

            foreach (bontique tmp1 in buylist)
            {
                /*
                using (WebClient client = new WebClient()) // WebClient class inherits IDisposable
                {
                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                    string htmlCode = client.DownloadString(tmp1.url);
                    int position = htmlCode.IndexOf("class=\"price\">");
                }
                */

                driver.Url = tmp1.url;
                IWebElement pricetag = driver.FindElement(By.XPath("//span[@class='price']"));
                string ooo = pricetag.GetAttribute("innerHTML").ToString();
                double price = Convert.ToDouble(ooo.Replace(" CAD", "").Replace("$",""));
                xlRange.Cells[tmp1.row, 4].Value = price;
                
                if (price < tmp1.lowest)
                {
                    xlRange.Cells[tmp1.row, 2].Value = price;
                    tmp1.lowest = price;
                    email(tmp1);
                }
                if (price > tmp1.highest)
                {
                    xlRange.Cells[tmp1.row, 3].Value = price;
                }
                
            }
            driver.Close();
            driver.Quit();
            xlWorkBook.Save();
            xlWorkBook.Close(0);    
            xlApp.Quit();
        }


        static void email(bontique a)
        {

            var clientstmp = new SmtpClient("smtp.gmail.com", 587)
            {
                Credentials = new NetworkCredential("viarailtest333@gmail.com", "innovation123"),
                EnableSsl = true
            };
            MailMessage mail = new MailMessage("viarailtest333@gmail.com", "williamdu@me.com", "Price Changed!", a.url + "        Now the price is " + a.lowest + "        highest price is " + a.highest);
            clientstmp.Send(mail);
        }
    }

    class bontique
    {
        public Int32 row { get; set;  }
        public String url { get; set; } //1
        public Double lowest { get; set; } //2
        public Double highest { get; set; } //3
        public Double last { get; set; } //4

    }
}
