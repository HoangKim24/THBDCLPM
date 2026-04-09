using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace SeleniumProject.Pages;

public sealed class LoginPage
{
    private readonly IWebDriver _driver;

    private readonly By _usernameInput = By.Name("username");
    private readonly By _passwordInput = By.Name("password");
    private readonly By _loginButton = By.CssSelector("button[type='submit'], input[type='submit']");
    private readonly By _errorMessage = By.CssSelector(".validation-summary-errors, .text-danger, .alert-danger");
    private readonly By _logoutLink = By.CssSelector("a[href*='Logout']");

    public LoginPage(IWebDriver driver)
    {
        _driver = driver;
    }

    public void Open(string baseUrl)
    {
        _driver.Navigate().GoToUrl($"{baseUrl}/Admin/AdminAuth/Login");
    }

    public void EnterUsername(string username)
    {
        var input = _driver.FindElement(_usernameInput);
        input.Clear();
        input.SendKeys(username);
    }

    public void EnterPassword(string password)
    {
        var input = _driver.FindElement(_passwordInput);
        input.Clear();
        input.SendKeys(password);
    }

    public void ClickLogin()
    {
        _driver.FindElement(_loginButton).Click();
    }

    public void Login(string username, string password)
    {
        EnterUsername(username);
        EnterPassword(password);
        ClickLogin();
    }

    public string CurrentUrl()
    {
        return _driver.Url;
    }

    public bool IsLoginButtonClickable()
    {
        return _driver.FindElement(_loginButton).Displayed && _driver.FindElement(_loginButton).Enabled;
    }

    public bool IsUsernameTextboxEnabled()
    {
        return _driver.FindElement(_usernameInput).Enabled;
    }

    public bool IsPasswordTextboxEnabled()
    {
        return _driver.FindElement(_passwordInput).Enabled;
    }

    public string ReadValidationMessage()
    {
        return _driver.FindElement(_errorMessage).Text;
    }

    public bool HasErrorMessage()
    {
        return _driver.FindElements(_errorMessage).Count > 0;
    }

    public bool CanSeeLogoutLink()
    {
        return _driver.FindElements(_logoutLink).Count > 0;
    }
}
