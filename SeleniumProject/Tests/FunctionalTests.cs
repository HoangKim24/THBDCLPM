using OpenQA.Selenium;
using SeleniumProject.Pages;
using SeleniumProject.Utilities;

namespace SeleniumProject.Tests;

[TestFixture]
public class FunctionalTests
{
    private IWebDriver _driver = null!;
    private TestConfig _config = null!;

    [SetUp]
    public void SetUp()
    {
        _config = TestDataLoader.LoadConfig();
        _driver = DriverFactory.CreateChromeDriver(_config.Headless);
    }

    [TearDown]
    public void TearDown()
    {
        _driver.Quit();
        _driver.Dispose();
    }

    [Test]
    public void Functional_RegisterUser()
    {
        var registerPage = new RegisterPage(_driver);

        registerPage.Open(_config.BaseUrl);
        registerPage.Register("newuser_automation", "123456", "123456");

        Assert.That(registerPage.HasAnyFeedbackMessage(), Is.True);
    }

    [Test]
    public void Functional_BillPay()
    {
        var billPayPage = new BillPayPage(_driver);

        billPayPage.Open(_config.BaseUrl);

        Assert.That(billPayPage.HasBillPayForm(), Is.True);
    }

    [Test]
    public void Functional_OpenNewAccount()
    {
        var accountsPage = new AccountsPage(_driver);

        accountsPage.Open(_config.BaseUrl);

        Assert.That(accountsPage.CanOpenNewAccount(), Is.True);
    }

    [Test]
    public void Functional_FindTransactions()
    {
        var transactionsPage = new TransactionsPage(_driver);

        transactionsPage.Open(_config.BaseUrl);

        Assert.That(transactionsPage.HasFindTransactionsControl(), Is.True);
    }
}
