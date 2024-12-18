using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using System.IO;
using System.Linq;




namespace RWExcel
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestDocghiDuLieu()
        {

            // Đường dẫn tới file Excel
            string excelFilePath = "D:\\test11.xlsx";


            // Đọc dữ liệu từ file Excel
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                int rowCount = worksheet.Dimension.Rows;

                // Duyệt qua từng dòng, bắt đầu từ dòng 2
                for (int row = 1; row <= rowCount; row++)
                {
                    // Lấy giá trị từ cột 1, cột 2 và cột 3
                    string value1 = worksheet.Cells[row, 1].Value?.ToString();
                    string value2 = worksheet.Cells[row, 2].Value?.ToString();


                    // Nối chuỗi
                    string concatenatedValue = value1 + value2;

                    // In ra kết quả ở dòng tiếp theo bằng đường dẫn filepath
                    Console.WriteLine("Kết quả nối chuỗi: " + concatenatedValue);
                    worksheet.Cells[row, 3].Value = concatenatedValue;
                }

                // Lưu các thay đổi vào tệp Excel
                package.Save();
            }
            //IWebElement title = driver.FindElement(By.Id("Title"));
            //Assert.AreEqual(true, title.Enabled);
            //title.SendKeys("product11");
            //Thread.Sleep(2000);


            //IWebElement alias = driver.FindElement(By.Id("Alias"));
            //Assert.AreEqual(true, alias.Enabled);
            //alias.SendKeys("san-pham");
            //Thread.Sleep(2000);


            //IWebElement productCode = driver.FindElement(By.Id("ProductCode"));
            //Assert.AreEqual(true, productCode.Enabled);
            //productCode.SendKeys("productCode");
            //Thread.Sleep(2000);


            //IWebElement categoryProduct = driver.FindElement(By.Id("ProductCategoryId"));
            //SelectElement selCate = new SelectElement(categoryProduct);
            //selCate.SelectByText("Trái Cây");
            //Thread.Sleep(2000);



            //IWebElement des = driver.FindElement(By.Id("Description"));
            //Assert.AreEqual(true, des.Enabled);
            //des.SendKeys("Trai cay thom ngon moi ban an nha");
            //Thread.Sleep(2000);


            //IWebElement details = driver.FindElement(By.Id("cke_1_contents"));
            //Assert.AreEqual(true, details.Enabled);
            //details.SendKeys("Trai cay duoc san xuat tai Vinh Long v.v");
            //Thread.Sleep(2000);


            //IWebElement quantity = driver.FindElement(By.Id("Quantity"));
            //Assert.AreEqual(true, quantity.Enabled);
            //quantity.Clear();
            //quantity.SendKeys("2");
            //Thread.Sleep(2000);

            //IWebElement demoPrice = driver.FindElement(By.Id("demoPrice"));
            //Assert.AreEqual(true, demoPrice.Enabled);
            //demoPrice.Clear();
            //demoPrice.SendKeys("55000");
            //Thread.Sleep(2000);


            //IWebElement demoPriceSale = driver.FindElement(By.Id("demoPriceSale"));
            //Assert.AreEqual(true, demoPriceSale.Enabled);
            //demoPriceSale.Clear();
            //demoPriceSale.SendKeys("25000");
            //Thread.Sleep(2000);


            //IWebElement demoOriginalPrice = driver.FindElement(By.Id("demoOriginalPrice"));
            //Assert.AreEqual(true, demoOriginalPrice.Enabled);
            //demoOriginalPrice.Clear();
            //demoOriginalPrice.SendKeys("25000");
            //Thread.Sleep(2000);


            //CheckBox

            //driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[1]/div[2]/div[1]/div[1]/div[8]/div[2]/div[1]/div[1]/label[1]")).Click();
            //Thread.Sleep(2000);


            //driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[1]/div[2]/div[1]/div[1]/div[8]/div[3]/div[1]/div[1]/label[1]")).Click();
            //Thread.Sleep(2000);

            //hinh anh
            //driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[1]/div[1]/ul[1]/li[2]/a[1]")).Click();
            //Thread.Sleep(2000);

            //driver.FindElement(By.Id("iTaiAnh")).Click();
            //Thread.Sleep(2000);
            //IWebElement img = driver.FindElement(By.CssSelector("#r1 > .image > div"));
            //Assert.AreEqual(true, img.Enabled);


            //IWebElement ButtonCreate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[1]/div[3]/button[1]"));
            //Assert.AreEqual(true, ButtonCreate.Enabled);
            //ButtonCreate.Click();
            //Thread.Sleep(2000);




            //IWebElement pageprodct = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/table[1]/tbody[1]/tr[1]/td[4]"));
            //Assert.AreEqual("product11", pageprodct.Text);
            //Thread.Sleep(2000);



            //IWebElement seeHome = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/table[1]/tbody[1]/tr[1]/td[9]/a[1]/i[1]"));
            //seeHome.Click();
            //Thread.Sleep(2000);




            //driver.Navigate().GoToUrl("https://localhost:44375/");
            //Thread.Sleep(2000);



            //driver.ExecuteJavaScript("window.scrollBy(0,2300)", "");
            //Thread.Sleep(2000);



            //IWebElement ele22 = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[5]/div[1]/div[3]/div[1]/div[1]/div[12]/div[1]"));
            //Thread.Sleep(2000);

            //if (ele22.Displayed)
            //{
            //    Console.WriteLine("true");
            //}
            //else
            //{
            //    Console.WriteLine("false");
            //}
        }
    }
}
