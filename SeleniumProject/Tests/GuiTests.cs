using OpenQA.Selenium;
using SeleniumProject.Pages;
using SeleniumProject.Utilities;

namespace SeleniumProject.Tests;

[TestFixture]
public class GuiTests
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
    public void GUI_LoginButton_Clickable()
    {
        var loginPage = new LoginPage(_driver);
        loginPage.Open(_config.BaseUrl);

        Assert.That(loginPage.IsLoginButtonClickable(), Is.True);
    }

    [Test]
    public void GUI_Textbox_Editable()
    {
        var loginPage = new LoginPage(_driver);
        loginPage.Open(_config.BaseUrl);

        Assert.Multiple(() =>
        {
            Assert.That(loginPage.IsUsernameTextboxEnabled(), Is.True);
            Assert.That(loginPage.IsPasswordTextboxEnabled(), Is.True);
        });
    }

    [Test]
    public void GUI_LinkNavigation_Correct()
    {
        var registerPage = new RegisterPage(_driver);
        registerPage.Open(_config.BaseUrl);

        Assert.That(_driver.Url, Does.Contain("/Register"));
    }

    [Test]
    public void GUI_ErrorMessage_Displayed_WhenInvalidLogin()
    {
        var loginPage = new LoginPage(_driver);
        loginPage.Open(_config.BaseUrl);
        loginPage.Login("", "123456");

        Assert.That(loginPage.HasErrorMessage(), Is.True);
    }
}
