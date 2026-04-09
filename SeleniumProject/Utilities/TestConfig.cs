namespace SeleniumProject.Utilities;

public sealed class TestConfig
{
    public string BaseUrl { get; init; } = "https://localhost:7129";
    public bool Headless { get; init; } = true;
    public TestUser ValidUser { get; init; } = new();
}

public sealed class TestUser
{
    public string Username { get; init; } = "adminApproved";
    public string Password { get; init; } = "passwordCorrect";
}
