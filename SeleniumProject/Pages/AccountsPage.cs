using OpenQA.Selenium;

namespace SeleniumProject.Pages;

public sealed class AccountsPage
{
    private readonly IWebDriver _driver;
    private readonly By _openAccountLink = By.CssSelector("a[href*='OpenAccount']");

    public AccountsPage(IWebDriver driver)
    {
        _driver = driver;
    }

    public void Open(string baseUrl)
    {
        _driver.Navigate().GoToUrl($"{baseUrl}/Accounts");
    }

    public bool CanOpenNewAccount()
    {
        return _driver.FindElements(_openAccountLink).Count > 0;
    }
}
