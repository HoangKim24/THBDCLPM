using OpenQA.Selenium;

namespace SeleniumProject.Pages;

public sealed class TransferFundsPage
{
    private readonly IWebDriver _driver;

    private readonly By _amountInput = By.Name("amount");
    private readonly By _fromAccountSelect = By.Name("fromAccountId");
    private readonly By _toAccountSelect = By.Name("toAccountId");
    private readonly By _submitButton = By.CssSelector("button[type='submit'], input[type='submit']");
    private readonly By _successMessage = By.CssSelector(".alert-success, .text-success");

    public TransferFundsPage(IWebDriver driver)
    {
        _driver = driver;
    }

    public void Open(string baseUrl)
    {
        _driver.Navigate().GoToUrl($"{baseUrl}/Transfer");
    }

    public void Transfer(string amount)
    {
        var amountEl = _driver.FindElement(_amountInput);
        amountEl.Clear();
        amountEl.SendKeys(amount);

        _driver.FindElement(_submitButton).Click();
    }

    public bool HasTransferForm()
    {
        return _driver.FindElements(_amountInput).Count > 0
            && _driver.FindElements(_fromAccountSelect).Count > 0
            && _driver.FindElements(_toAccountSelect).Count > 0;
    }

    public bool HasSuccessMessage()
    {
        return _driver.FindElements(_successMessage).Count > 0;
    }
}
