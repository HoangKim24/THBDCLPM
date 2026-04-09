using System.Text.Json;

namespace SeleniumProject.Utilities;

public static class TestDataLoader
{
    public static TestConfig LoadConfig()
    {
        var path = Path.Combine(TestContext.CurrentContext.TestDirectory, "TestData", "users.json");

        if (!File.Exists(path))
        {
            return new TestConfig();
        }

        var json = File.ReadAllText(path);
        var config = JsonSerializer.Deserialize<TestConfig>(json, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
        });

        return config ?? new TestConfig();
    }
}
