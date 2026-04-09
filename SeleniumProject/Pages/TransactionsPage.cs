using OpenQA.Selenium;

namespace SeleniumProject.Pages;

public sealed class TransactionsPage
{
    private readonly IWebDriver _driver;
    private readonly By _findTransactionsInput = By.Name("transactionId");

    public TransactionsPage(IWebDriver driver)
    {
        _driver = driver;
    }

    public void Open(string baseUrl)
    {
        _driver.Navigate().GoToUrl($"{baseUrl}/Transactions/Find");
    }

    public bool HasFindTransactionsControl()
    {
        return _driver.FindElements(_findTransactionsInput).Count > 0;
    }
}
