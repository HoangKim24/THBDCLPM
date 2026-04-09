using OpenQA.Selenium;
using SeleniumProject.Pages;
using SeleniumProject.Utilities;

namespace SeleniumProject.Tests;

[TestFixture]
public class TransferTests
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
    public void Smoke_TransferFunds_SuccessFlow()
    {
        var loginPage = new LoginPage(_driver);
        var transferPage = new TransferFundsPage(_driver);

        loginPage.Open(_config.BaseUrl);
        loginPage.Login(_config.ValidUser.Username, _config.ValidUser.Password);

        transferPage.Open(_config.BaseUrl);

        Assert.That(transferPage.HasTransferForm(), Is.True);
    }
}
