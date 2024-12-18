using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using OpenQA.Selenium.Chrome;
using System.Diagnostics;
using OpenQA.Selenium.Edge;
using System.Threading;
using OpenQA.Selenium;
using selenium;
using OpenQA.Selenium.DevTools.V120.BackgroundService;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Support.Extensions;
using System.Collections.Generic;
using System.Linq;
using OpenQA.Selenium.DevTools.V120.Runtime;
using KiemThuSanPham;


namespace selenium
{
    [TestClass]
    public class TestCaseQLSanPham
    {
        IWebDriver driver;
        private string baseURL = "https://localhost:44375/admin/account";
        //path excel
        string pathExcel = "C:\\Users/nhan2/Downloads/TestData2.xlsx";
        ExcelApiTest an;

        [TestInitialize]
        public void TestOpenAdminPage()
        {
            driver = new EdgeDriver();
            driver.Url = "https://localhost:44375/admin/account";
            driver.Navigate();
            driver.Manage().Window.Maximize();
            an = new ExcelApiTest(pathExcel);
            Thread.Sleep(2000);


        }

        //Them san pham voi day du thong tin + xoa san pham
        [TestMethod]
        public void TestCase_ID_Them()
        {
            Test_LoginAdmin_GUI();
            Test_CreateProducts();
        }

        //Tim kiem san pham
        [TestMethod]
        public void TestCase_ID_TimKiem()
        {
            Test_LoginAdmin_GUI();
            Test_SearchProducts();
        }

        //Them danh muc san pham
        [TestMethod]
        public void TestCase_ID_DanhMuc_Them()
        {
            Test_LoginAdmin_GUI();
            Test_CategoryProducts();
        }

        public void Test_LoginAdmin_GUI()
        {

            By UserName = By.Id("UserName");
            By Password = By.Id("Password");

            //lay data tu file excel
            string account = an.GetCellData("TcQLSanPham", "Tài khoản", 3);
            string password = an.GetCellData("TcQLSanPham", "Mật khẩu", 3);


            driver.FindElement(UserName).SendKeys(account);
            Thread.Sleep(1000);
            driver.FindElement(Password).SendKeys(password);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("RememberMe")).Click();
            Thread.Sleep(1000);


            driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/form[1]/div[3]/div[2]/button[1]")).Click();

        }

        public void Test_CreateProducts()
        {
            driver.FindElement(By.CssSelector(".right.fas.fa-angle-left")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.CssSelector("a[href='/admin/products']")).Click();
            Thread.Sleep(2000);


            for (int i = 3; i <= 21; i++)
            {
                IWebElement ele1 = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[1]/div[1]/a[1]"));
                Assert.AreEqual(true, ele1.Enabled);
                ele1.Click();
                Assert.IsTrue(driver.Url.Contains("https://localhost:44375/admin/products/add"), "Loi khong vao duoc trang them san pham");
                //Kiem tra cac text va id
                By Title = By.Id("Title");
                By Alias = By.Id("Alias");
                By ProductCode = By.Id("ProductCode");
                By ProductCategoryId = By.Id("ProductCategoryId");
                By Description = By.Id("Description");
                By Quantity = By.Id("Quantity");
                By demoPrice = By.Id("demoPrice");
                By demoPriceSale = By.Id("demoPriceSale");
                By demoOriginalPrice = By.Id("demoOriginalPrice");
                By image = By.Id("txtImageUrl");



                string title = an.GetCellData("TcQLSanPham", "Tên sản phẩm", i);
                string alias = an.GetCellData("TcQLSanPham", "Alias", i);
                string productCode = an.GetCellData("TcQLSanPham", "SKU", i);
                string productCategoryId = an.GetCellData("TcQLSanPham", "Danh mục", i);
                string description = an.GetCellData("TcQLSanPham", "Mô tả", i);
                string quantity = an.GetCellData("TcQLSanPham", "Số lượng", i);
                string DemoPrice = an.GetCellData("TcQLSanPham", "Giá", i);
                string DemoPriceSale = an.GetCellData("TcQLSanPham", "Giá khuyến mãi", i);
                string DemoOriginalPrice = an.GetCellData("TcQLSanPham", "Giá nhập", i);
                string Image = an.GetCellData("TcQLSanPham", "Ảnh", i);

                var cbHot = an.GetCellData("TcQLSanPham", "Hot", i);
                var cbNoiBat = an.GetCellData("TcQLSanPham", "Nổi bật", i);
                var cbKhuyenMai = an.GetCellData("TcQLSanPham", "Khuyến mãi", i);


                //Ten san pham
                if (title != null)
                {
                    driver.FindElement(Title).SendKeys(title);
                    Thread.Sleep(100);
                }
                
                //Alias
                if (alias != null)
                {
                    driver.FindElement(Alias).SendKeys(alias);
                    Thread.Sleep(100);
                }
               
                //SKU
                if (productCode != null)
                {
                    driver.FindElement(ProductCode).SendKeys(productCode);
                    Thread.Sleep(100);
                }
                

                //Danh mục
                if (productCategoryId != null)
                {
                    driver.FindElement(ProductCategoryId).SendKeys(productCategoryId);
                    Thread.Sleep(100);
                }
                

                //Mô tả
                if (description != null)
                {
                    driver.FindElement(Description).SendKeys(description);
                    Thread.Sleep(100);
                }
               

                driver.ExecuteJavaScript("window.scrollBy(0,600)", "");

                //Số lượng
                if (quantity != null)
                {
                    driver.FindElement(Quantity).Clear();
                    driver.FindElement(Quantity).SendKeys(quantity);
                    Thread.Sleep(100);
                }
                

                //Giá
                if (DemoPrice != null)
                {
                    driver.FindElement(demoPrice).Clear();
                    driver.FindElement(demoPrice).SendKeys(DemoPrice);
                    Thread.Sleep(100);
                }
               

                //Giá Khuyến Mãi
                if (DemoPriceSale != null)
                {
                    driver.FindElement(demoPriceSale).Clear();
                    driver.FindElement(demoPriceSale).SendKeys(DemoPriceSale);
                    Thread.Sleep(100);
                }
                

                //Giá nhập
                if (DemoOriginalPrice != null)
                {
                    driver.FindElement(demoOriginalPrice).Clear();
                    driver.FindElement(demoOriginalPrice).SendKeys(DemoOriginalPrice);
                    Thread.Sleep(1000);
                }
                

               


                //Hiển thị giá
                if (cbHot == "R")
                    driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[1]/div[2]/div[1]/div[1]/div[8]/div[2]/div[1]/div[1]/label[1]")).Click();
                if (cbNoiBat == "R")
                    driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[1]/div[2]/div[1]/div[1]/div[8]/div[3]/div[1]/div[1]/label[1]")).Click();
                if (cbKhuyenMai == "R")
                    driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[1]/div[2]/div[1]/div[1]/div[8]/div[4]/div[1]/div[1]/label[1]")).Click();

                //Tab hinh anh
                driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[1]/div[1]/ul[1]/li[2]/a[1]")).Click();
                //Ảnh
                if (Image != null)
                {
                    driver.FindElement(image).SendKeys(Image);
                    //btn add image 
                    driver.FindElement(By.Id("iThemSanPham")).Click();
                    Thread.Sleep(1000);
                } 


                //Tab thong tin
                driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[1]/div[1]/ul[1]/li[1]/a[1]")).Click();

                driver.ExecuteJavaScript("window.scrollBy(0,600)", ""); 


                //Button create
                IWebElement ButtonCreate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[1]/div[3]/button[1]"));
                Assert.AreEqual(true, ButtonCreate.Enabled);
                ButtonCreate.Click();
                Thread.Sleep(2000);

                


                    try
                    {
                        //0. tim thay loi cua gia
                        IWebElement errorPrice = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[2]/div[2]/div[1]/div[1]/div[7]/div[2]/div[1]/span[1]"));
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", errorPrice);

                        string actualRS;
                        actualRS = "ERROR";
                        an.SetCellData("TcQLSanPham", 21, i, actualRS);
                        string anw = an.GetCellData("TcQLSanPham", "Expected result", i);
                        if (anw == actualRS)
                        {
                            an.SetCellData("TcQLSanPham", 23, i, "Passed");
                        }
                        else
                        {
                            an.SetCellData("TcQLSanPham", 23, i, "Failed");
                        }

                        driver.Url = "https://localhost:44375/Admin/Products";
                    }
                    catch (NoSuchElementException)
                    {
                        try
                        {
                            //1. tim thay loi trung san pham
                            IWebElement errorTittle = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[1]/div[2]/div[1]/div[1]/div[1]/span[1]"));
                            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", errorTittle);

                        string actualRS;
                            actualRS = "ERROR";
                            an.SetCellData("TcQLSanPham", 21, i, actualRS);
                            string anw = an.GetCellData("TcQLSanPham", "Expected result", i);
                            if (anw == actualRS)
                            {
                                an.SetCellData("TcQLSanPham", 23, i, "Passed");
                            }
                            else
                            {
                                an.SetCellData("TcQLSanPham", 23, i, "Failed");
                            }

                            driver.Url = "https://localhost:44375/Admin/Products";
                        }
                        catch (NoSuchElementException)
                        {
                            try
                            {
                                //2. tim thay loi danh muc
                                IWebElement errorCategory = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[2]/div[2]/div[1]/div[1]/div[4]/span[1]"));
                                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", errorCategory);
                                string actualRS;
                                actualRS = "ERROR";
                                an.SetCellData("TcQLSanPham", 21, i, actualRS);
                                string anw = an.GetCellData("TcQLSanPham", "Expected result", i);
                                if (anw == actualRS)
                                {
                                    an.SetCellData("TcQLSanPham", 23, i, "Passed");
                                }
                                else
                                {
                                    an.SetCellData("TcQLSanPham", 23, i, "Failed");
                                }

                                driver.Url = "https://localhost:44375/Admin/Products";
                            }
                            catch (NoSuchElementException)
                            {
                                try
                                {
                                    //3. tim thay loi so luong
                                    IWebElement errorTittle = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[2]/div[2]/div[1]/div[1]/div[7]/div[1]/div[1]/span[1]"));
                                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", errorTittle);
                                    string actualRS;
                                    actualRS = "ERROR";
                                    an.SetCellData("TcQLSanPham", 21, i, actualRS);
                                    string anw = an.GetCellData("TcQLSanPham", "Expected result", i);
                                    if (anw == actualRS)
                                    {
                                        an.SetCellData("TcQLSanPham", 23, i, "Passed");
                                    }
                                    else
                                    {
                                        an.SetCellData("TcQLSanPham", 23, i, "Failed");
                                    }

                                    driver.Url = "https://localhost:44375/Admin/Products";
                                }

                                catch (NoSuchElementException)
                                {
                                    try
                                    {
                                        //4. tim thay loi gia nhap
                                        IWebElement errorOriginPrice = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/div[1]/form[1]/div[2]/div[2]/div[1]/div[1]/div[7]/div[4]/div[1]/span[1]"));
                                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", errorOriginPrice);
                                        string actualRS;
                                        actualRS = "ERROR";
                                        an.SetCellData("TcQLSanPham", 21, i, actualRS);
                                        string anw = an.GetCellData("TcQLSanPham", "Expected result", i);
                                        if (anw == actualRS)
                                        {
                                            an.SetCellData("TcQLSanPham", 23, i, "Passed");
                                        }
                                        else
                                        {
                                            an.SetCellData("TcQLSanPham", 23, i, "Failed");
                                        }

                                        driver.Url = "https://localhost:44375/Admin/Products";
                                    }
                                    catch (NoSuchElementException)
                                    {
                                        IWebElement pageprodct = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/table[1]/tbody[1]/tr[1]/td[4]"));
                                        pageprodct.Click();

                                        //checkbox hien thi o trang chu
                                        var cbHome = an.GetCellData("TcQLSanPham", "Home", i);
                                        var cbSale = an.GetCellData("TcQLSanPham", "Sale", i);

                                        if (cbHome == "R")
                                        {
                                            driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/table[1]/tbody[1]/tr[1]/td[9]/a[1]/i[1]")).Click();
                                            Thread.Sleep(2000);
                                            //Duong dan driver moi dan toi trang chu
                                            driver.Url = "https://localhost:44375/";
                                            driver.Navigate();
                                            driver.Navigate().Refresh();
                                            Thread.Sleep(2000);

                                            //Kiem tra ten san pham o trang home
                                            IWebElement SPHome = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[5]/div[1]/div[3]/div[1]/div[1]/div[12]/div[1]"));
                                            Thread.Sleep(500);
                                            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", SPHome);
                                            Thread.Sleep(500);
                                            IWebElement tenHome = SPHome.FindElement(By.ClassName("product_name"));
                                            Thread.Sleep(500);

                                            
                                            //actualRS
                                            an.SetCellData("TcQLSanPham", 21, i, tenHome.Text.ToString());
                                            string anw = an.GetCellData("TcQLSanPham", "Expected result", i);
                                            if (anw == tenHome.Text)
                                            {
                                                an.SetCellData("TcQLSanPham", 23, i, "Passed");
                                            }
                                            else
                                            {
                                                an.SetCellData("TcQLSanPham", 23, i, "Failed");
                                            }
                                            driver.Url = "https://localhost:44375/Admin/Products";
                                            DeleteProduct();
                                    }
                                        else
                                        {
                                        driver.Url = "https://localhost:44375/";
                                        driver.Navigate();
                                        driver.Navigate().Refresh();
                                        Thread.Sleep(2000);
                                        IWebElement tenHome = null;
                                        try
                                        {
                                            tenHome = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[5]/div[1]/div[3]/div[1]/div[1]/div[12]/div[1]/div[4]/h6[1]/a[1]"));
                                            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", tenHome);
                                        }
                                        catch (NoSuchElementException)
                                        {
                                            // Nếu không tìm thấy phần tử, gán giá trị "NULL" cho tenHome
                                            tenHome = null;
                                        }
                                        // Lấy tên sản phẩm hoặc gán giá trị "NULL"
                                        string tenSanPham = tenHome != null ? tenHome.Text : "NULL";
                                        an.SetCellData("TcQLSanPham", 21, i, tenSanPham);
                                        string anw = an.GetCellData("TcQLSanPham", "Expected result", i);
                                        if (anw == tenSanPham)
                                        {
                                            an.SetCellData("TcQLSanPham", 23, i, "Passed");
                                        }
                                        else
                                        {
                                            an.SetCellData("TcQLSanPham", 23, i, "Failed");
                                        }
                                        driver.Url = "https://localhost:44375/Admin/Products";
                                        DeleteProduct();
                                    }
                                    if (cbSale == "R")
                                        driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/table[1]/tbody[1]/tr[1]/td[10]/a[1]/i[1]")).Click();
                                }
                                }
                            }
                        }
                    }
            }


            Thread.Sleep(2000);
        }

        public void Test_SearchProducts()
        {

            driver.FindElement(By.CssSelector(".right.fas.fa-angle-left")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.CssSelector("a[href='/admin/products']")).Click();
            Thread.Sleep(2000);
            for (int i = 3; i <= 7;  i++)
            {
                By TBSearch = By.Id("searchString");
                By btnSearch = By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/form[1]/div[1]/div[1]/div[1]/button[1]");
                string TextSearch = an.GetCellData("SearchProduct", "Tên sản phẩm", i);
                string ExceptedResult = an.GetCellData("SearchProduct", "Expected result", i);

                driver.FindElement(TBSearch).SendKeys(TextSearch);
                driver.FindElement(btnSearch).Click();

                Thread.Sleep(2000);

                IWebElement tbody = driver.FindElement(By.CssSelector("tbody"));

                var rows = tbody.FindElements(By.TagName("tr"));

                bool productFound = false;

                // Lặp qua từng hàng
                foreach (var row in rows)
                {
                    // Kiểm tra nội dung của hàng có chứa tên sản phẩm không
                    if (row.Text.Contains(TextSearch))
                    {
                        productFound = true;
                        string[] rowData = row.Text.Split(' ');
                        string productName = rowData[1] + " " + rowData[2];
                        Console.WriteLine("Tên sản phẩm được tìm thấy: " + productName);
                        an.SetCellData("SearchProduct", 6, i, "TRUE");
                        break;
                    }
                    else
                    {
                        an.SetCellData("SearchProduct", 6, i, "NULL");
                    }
                }
      

                string ActualResult = an.GetCellData("SearchProduct", "Actual Result", i);
                if (ActualResult == ExceptedResult)
                {
                    an.SetCellData("SearchProduct", 8, i, "Passed");

                }
                else
                {
                    an.SetCellData("SearchProduct", 8, i, "Failed");
                }

                driver.FindElement(TBSearch).Clear();
                driver.FindElement(btnSearch).Click();
            }

        }

        public void Test_CategoryProducts()
        {
            driver.FindElement(By.CssSelector(".right.fas.fa-angle-left")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.CssSelector("a[href='/admin/productcategory']")).Click();
            Thread.Sleep(2000);

            for (int i = 3; i <= 10; i++)
            {
                By BtnAddCategory = By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[1]/div[1]/a[1]");
                By Title = By.Id("Title");
                By Alias = By.Id("Alias");
                By Image = By.Id("txtImage");

                string title = an.GetCellData("CategoryProduct", "Tiêu đề", i);
                string alias = an.GetCellData("CategoryProduct", "Alias", i);
                string image = an.GetCellData("CategoryProduct", "Ảnh", i);

                driver.FindElement(BtnAddCategory).Click();

                //Ten danh muc
                if (title != null)
                {
                    driver.FindElement(Title).SendKeys(title);
                    Thread.Sleep(100);
                }

                //Alias danh muc
                if (alias != null)
                {
                    driver.FindElement(Alias).SendKeys(alias);
                    Thread.Sleep(100);
                }

                //Hinh anh danh muc
                if (image != null)
                {
                    driver.FindElement(Image).SendKeys(image);
                    Thread.Sleep(100);
                }

                IWebElement btnAdd = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/form[1]/div[1]/div[7]/button[1]"));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", btnAdd);
                Thread.Sleep(1000);
                btnAdd.Click();
                Thread.Sleep(1000);


                try
                {
                    //Loi tittle 
                    IWebElement errorTitleCategory = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/form[1]/div[2]/div[1]/span[1]"));
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", errorTitleCategory);
                    string actualRS;
                    actualRS = "ERROR";
                    an.SetCellData("CategoryProduct", 8, i, actualRS);
                    string anw = an.GetCellData("CategoryProduct", "Expected result", i);
                    if (anw == actualRS)
                    {
                        an.SetCellData("CategoryProduct", 10, i, "Passed");
                    }
                    else
                    {
                        an.SetCellData("CategoryProduct", 10, i, "Failed");
                    }

                    driver.Url = "https://localhost:44375/admin/productcategory";
                }
                catch (NoSuchElementException)
                {
                   try
                    {
                        //Loi Alias
                        IWebElement errorAliasCategory = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/div[1]/form[1]/div[2]/div[1]/span[1]"));
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", errorAliasCategory);
                        string actualRS;
                        actualRS = "ERROR";
                        an.SetCellData("CategoryProduct", 8, i, actualRS);
                        string anw = an.GetCellData("CategoryProduct", "Expected result", i);
                        if (anw == actualRS)
                        {
                            an.SetCellData("CategoryProduct", 10, i, "Passed");
                        }
                        else
                        {
                            an.SetCellData("CategoryProduct", 10, i, "Failed");
                        }

                        driver.Url = "https://localhost:44375/admin/productcategory";
                    }
                    catch (NoSuchElementException)
                    {
                        try
                        {
                            // Tìm phần tử chứa thông báo lỗi
                            IWebElement error = driver.FindElement(By.XPath("/html[1]/body[1]/span[1]/h1[1]"));

                            if (error.Text.Trim() == "Server Error in '/' Application.")
                            {
                                // Nếu nội dung khớp, đánh dấu là test đã pass
                                Console.WriteLine("Test passed!");
                                string actualRS;
                                actualRS = "ERROR";
                                an.SetCellData("CategoryProduct", 8, i, actualRS);
                                string anw = an.GetCellData("CategoryProduct", "Expected result", i);
                                if (anw == actualRS)
                                {
                                    an.SetCellData("CategoryProduct", 10, i, "Passed");
                                }
                                else
                                {
                                    an.SetCellData("CategoryProduct", 10, i, "Failed");
                                }

                                driver.Url = "https://localhost:44375/admin/productcategory";
                            }
                        }
                        catch (NoSuchElementException)
                        {
                            //Kiem tra danh muc
                            IWebElement tencategory = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/table[1]/tbody[1]/tr[6]/td[3]"));
                            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", tencategory);
                            tencategory.Click();

                            //actualRS
                            an.SetCellData("CategoryProduct", 8, i, tencategory.Text);
                            string anw = an.GetCellData("CategoryProduct", "Expected result", i);
                            if (anw == tencategory.Text)
                            {
                                an.SetCellData("CategoryProduct", 10, i, "Passed");
                            }
                            else
                            {
                                an.SetCellData("CategoryProduct", 10, i, "Failed");
                            }
                            driver.Url = "https://localhost:44375/admin/productcategory";
                            DeleteCategoryProduct();
                        }
                    }
                }

            }


        }

        //Delete Product
        public void DeleteProduct()
        {
            driver.Navigate().GoToUrl("https://localhost:44375/Admin/Products");
            driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/table[1]/tbody[1]/tr[1]/td[12]/a[2]")).Click();
            Thread.Sleep(3000);
            driver.SwitchTo().Alert().Accept();
            Thread.Sleep(3000);
            driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/table[1]/tbody[1]/tr[1]/td[12]/a[2]")).Click();
            driver.SwitchTo().Alert().Accept();
            driver.Navigate().Refresh();
            Thread.Sleep(3000);
        }

        //Delete Product
        public void DeleteCategoryProduct()
        {
            //btn <td> xóa
            IWebElement td = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/table[1]/tbody[1]/tr[6]/td[5]"));
            Thread.Sleep(1000);
            IWebElement tr = td.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/section[2]/div[1]/div[2]/table[1]/tbody[1]/tr[6]/td[5]/a[2]"));
            Thread.Sleep(1000);
            tr.Click();
            driver.SwitchTo().Alert().Accept();
            Thread.Sleep(1000);
        }


        [TestCleanup] 
        public void Cleanup()
        {
            driver.Close();
            driver.Quit();
        }
    }
}
