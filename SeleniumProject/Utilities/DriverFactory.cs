using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace SeleniumProject.Utilities;

public static class DriverFactory
{
    public static IWebDriver CreateChromeDriver(bool headless = true)
    {
        var options = new ChromeOptions();
        options.AcceptInsecureCertificates = true;
        options.AddArgument("--window-size=1920,1080");
        options.AddArgument("--disable-gpu");
        options.AddArgument("--no-sandbox");

        if (headless)
        {
            options.AddArgument("--headless=new");
        }

        var driver = new ChromeDriver(options);
        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        return driver;
    }
}
