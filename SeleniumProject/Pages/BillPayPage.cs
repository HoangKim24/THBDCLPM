using OpenQA.Selenium;

namespace SeleniumProject.Pages;

public sealed class BillPayPage
{
    private readonly IWebDriver _driver;
    private readonly By _payeeNameInput = By.Name("payeeName");
    private readonly By _amountInput = By.Name("amount");
    private readonly By _submitButton = By.CssSelector("button[type='submit'], input[type='submit']");

    public BillPayPage(IWebDriver driver)
    {
        _driver = driver;
    }

    public void Open(string baseUrl)
    {
        _driver.Navigate().GoToUrl($"{baseUrl}/BillPay");
    }

    public void PayBill(string payeeName, string amount)
    {
        _driver.FindElement(_payeeNameInput).SendKeys(payeeName);
        _driver.FindElement(_amountInput).SendKeys(amount);
        _driver.FindElement(_submitButton).Click();
    }

    public bool HasBillPayForm()
    {
        return _driver.FindElements(_payeeNameInput).Count > 0;
    }
}
