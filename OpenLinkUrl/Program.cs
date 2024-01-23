using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;


class Program
{
    static void Main()
    {
        try
        {
            // Set the EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Console.WriteLine("Enter the URL of the website:");
            string url = "https://authentication.liveperson.net/";

            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LP_Acount.xlsx");
            Console.WriteLine("Enter the path to the Excel file: " + path);
            string excelFilePath = path;

            if (File.Exists(excelFilePath))
            {
                OpenWebsiteAndLoginForAllRows(url, excelFilePath);
            }
            else
            {
                Console.WriteLine($"Excel file not found at path: {excelFilePath}");
            }


            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            Console.ReadKey();

        }
        
    }

    static void OpenWebsiteAndLoginForAllRows(string url, string excelFilePath)
    {
        // Create a new Excel package and add a worksheet
        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            // Check if the workbook already contains worksheets
            if (package.Workbook.Worksheets.Count == 0)
            {
                // If not, add a new worksheet
                package.Workbook.Worksheets.Add("Credentials");
            }

            // Access the first worksheet
            var worksheet = package.Workbook.Worksheets[0];

            // Rest of your code remains unchanged...


            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++) // Assuming the headers are in the first row
            {
                string siteNumber = worksheet.Cells[row, 1].Value?.ToString();
                string username = worksheet.Cells[row, 2].Value?.ToString();
                string password = worksheet.Cells[row, 3].Value?.ToString();

                if (string.IsNullOrEmpty(siteNumber) || string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
                {
                    Console.WriteLine($"Skipping row {row} - Credentials not found.");
                    continue;
                }

                Console.WriteLine($"Processing row {row} - SiteNumber: {siteNumber}, Username: {username}, Password: {password}");

                // Create a new ChromeDriver for each set of credentials
                var chromeOptions = new ChromeOptions();
                chromeOptions.AddArgument("--incognito");
                chromeOptions.AddArgument("--start-maximized");
                chromeOptions.AddArgument("no-sandbox");
                IWebDriver driver = new ChromeDriver(ChromeDriverService.CreateDefaultService(), chromeOptions,  TimeSpan.FromMinutes(3));

                driver.Navigate().GoToUrl(url);

                // Perform the login for each set of credentials
                PerformLogin(driver, siteNumber, username, password);

                // Optionally, add a delay between logins
                //Thread.Sleep(TimeSpan.FromSeconds(5));

            }
        }
    }

    static void PerformLogin(IWebDriver driver, string siteNumber, string username, string password)
    {
        try
        {
            // Continue with the login logic as before
            // (you can call the existing OpenWebsiteAndLogin method or include the login logic here)

            // Example:
            //Thread.Sleep(TimeSpan.FromSeconds(15));
            IWebElement siteNumberInput = driver.FindElement(By.XPath("//input[@id='siteNumber']"));
            siteNumberInput.SendKeys(siteNumber);

            IWebElement usernameInput = driver.FindElement(By.XPath("//input[@id='userName']"));
            usernameInput.SendKeys(username);

            IWebElement passwordInput = driver.FindElement(By.XPath("//input[@id='sitePass']"));
            passwordInput.SendKeys(password);

            IWebElement loginButton = driver.FindElement(By.XPath("//input[@name='loginButton']"));
            loginButton.Click();

            // Wait for a moment (you can adjust the time as needed)
            Thread.Sleep(TimeSpan.FromSeconds(50));

            // Additional steps after the initial login (you can adjust based on the actual page structure)
            IWebElement proxyUsernameInput = driver.FindElement(By.XPath("//input[@id='proxy-username']"));
            proxyUsernameInput.SendKeys(username);

            IWebElement proxyPasswordInput = driver.FindElement(By.XPath("//input[@id='password']"));
            proxyPasswordInput.SendKeys(password);

            IWebElement continueButton = driver.FindElement(By.XPath("//button[@class='submitButton']"));
            continueButton.Click();

            Console.WriteLine("Additional steps completed.");

            // Wait for a moment (you can adjust the time as needed)
            //Thread.Sleep(TimeSpan.FromSeconds(5));
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            
        }

    }
}
