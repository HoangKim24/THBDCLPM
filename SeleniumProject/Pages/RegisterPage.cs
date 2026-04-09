using OpenQA.Selenium;

namespace SeleniumProject.Pages;

public sealed class RegisterPage
{
    private readonly IWebDriver _driver;

    private readonly By _usernameInput = By.Name("username");
    private readonly By _passwordInput = By.Name("password");
    private readonly By _confirmPasswordInput = By.Name("confirmPassword");
    private readonly By _submitButton = By.CssSelector("button[type='submit'], input[type='submit']");
    private readonly By _message = By.CssSelector(".alert, .text-danger, .text-success");

    public RegisterPage(IWebDriver driver)
    {
        _driver = driver;
    }

    public void Open(string baseUrl)
    {
        _driver.Navigate().GoToUrl($"{baseUrl}/Admin/AdminAuth/Register");
    }

    public void Register(string username, string password, string confirmPassword)
    {
        _driver.FindElement(_usernameInput).SendKeys(username);
        _driver.FindElement(_passwordInput).SendKeys(password);
        _driver.FindElement(_confirmPasswordInput).SendKeys(confirmPassword);
        _driver.FindElement(_submitButton).Click();
    }

    public bool HasAnyFeedbackMessage()
    {
        return _driver.FindElements(_message).Count > 0;
    }
}
