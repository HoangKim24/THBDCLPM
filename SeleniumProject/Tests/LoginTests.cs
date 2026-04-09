using OpenQA.Selenium;
using SeleniumProject.Pages;
using SeleniumProject.Utilities;

namespace SeleniumProject.Tests;

[TestFixture]
public class LoginTests
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
    public void Smoke_LoginSuccess()
    {
        var loginPage = new LoginPage(_driver);

        loginPage.Open(_config.BaseUrl);
        loginPage.Login(_config.ValidUser.Username, _config.ValidUser.Password);

        Assert.That(loginPage.CurrentUrl(), Does.Contain("/Admin"));
    }

    [Test]
    public void Smoke_Logout_AfterLogin()
    {
        var loginPage = new LoginPage(_driver);

        loginPage.Open(_config.BaseUrl);
        loginPage.Login(_config.ValidUser.Username, _config.ValidUser.Password);
        _driver.Navigate().GoToUrl($"{_config.BaseUrl}/Admin/AdminAuth/Logout");

        Assert.That(_driver.Url, Does.Contain("/Admin/AdminAuth/Login"));
    }
}
