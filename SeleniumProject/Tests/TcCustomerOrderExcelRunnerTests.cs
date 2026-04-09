using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumProject.Utilities;

namespace SeleniumProject.Tests;

[TestFixture]
public class TcCustomerOrderExcelRunnerTests
{
    private const string SheetAlias1 = "TC_Customers_Orders";
    private const string SheetAlias2 = "TC_CUSTOMER_ORDER";

    [Test]
    public void Run_No1To20_And_Write_Back_To_Excel()
    {
        var excelPath = ResolveWritableExcelPath();
        var config = TestDataLoader.LoadConfig();
        var screenshotDir = ResolveScreenshotDirectory();

        using var workbook = new XLWorkbook(excelPath);
        var ws = FindWorksheet(workbook);
        var headerRow = FindHeaderRow(ws);

        var noCol = FindOrCreateColumn(ws, headerRow, new[] { "no.", "no", "stt" }, "NO.");
        var expectedCol = FindOrCreateColumn(ws, headerRow, new[] { "expectedresult", "expected result" }, "Expected Result");
        var actualCol = FindOrCreateColumn(ws, headerRow, new[] { "actualresult", "actual result" }, "Actual Result");
        var notesCol = FindOrCreateColumn(ws, headerRow, new[] { "notes", "note" }, "Notes");
        var screenshotCol = FindOrCreateColumn(ws, headerRow, new[] { "screenshot", "screen shot" }, "Screenshot");

        var lastUsedRow = ws.LastRowUsed();
        if (lastUsedRow is null)
        {
            throw new InvalidOperationException("Worksheet has no used rows.");
        }

        var lastRow = lastUsedRow.RowNumber();

        for (var row = headerRow + 1; row <= lastRow; row++)
        {
            var noText = ws.Cell(row, noCol).GetString().Trim();
            if (!int.TryParse(noText, out var no) || no < 1 || no > 20)
            {
                continue;
            }

            var expected = ws.Cell(row, expectedCol).GetString().Trim();
            var result = ExecuteCase(no, expected, config, screenshotDir);

            ws.Cell(row, actualCol).Value = result.Actual;
            ws.Cell(row, notesCol).Value = result.Passed ? "Passed" : "Falled";
            WriteScreenshotToCell(ws, row, screenshotCol, result);
        }

        workbook.Save();

        Assert.Pass($"Updated: {excelPath} | Sheet: {ws.Name} | Range: NO 1-20");
    }

    [Test]
    public void Run_No21To45_And_Write_Back_To_Excel()
    {
        var excelPath = ResolveWritableExcelPath();
        var config = TestDataLoader.LoadConfig();
        var screenshotDir = ResolveScreenshotDirectory();

        using var workbook = new XLWorkbook(excelPath);
        var ws = FindWorksheet(workbook);
        var headerRow = FindHeaderRow(ws);

        var noCol = FindOrCreateColumn(ws, headerRow, new[] { "no.", "no", "stt" }, "NO.");
        var expectedCol = FindOrCreateColumn(ws, headerRow, new[] { "expectedresult", "expected result" }, "Expected Result");
        var actualCol = FindOrCreateColumn(ws, headerRow, new[] { "actualresult", "actual result" }, "Actual Result");
        var notesCol = FindOrCreateColumn(ws, headerRow, new[] { "notes", "note" }, "Notes");
        var screenshotCol = FindOrCreateColumn(ws, headerRow, new[] { "screenshot", "screen shot" }, "Screenshot");

        var lastUsedRow = ws.LastRowUsed();
        if (lastUsedRow is null)
        {
            throw new InvalidOperationException("Worksheet has no used rows.");
        }

        var lastRow = lastUsedRow.RowNumber();

        for (var row = headerRow + 1; row <= lastRow; row++)
        {
            var noText = ws.Cell(row, noCol).GetString().Trim();
            if (!int.TryParse(noText, out var no) || no < 21 || no > 45)
            {
                continue;
            }

            var expected = ws.Cell(row, expectedCol).GetString().Trim();
            var result = ExecuteCase(no, expected, config, screenshotDir);

            ws.Cell(row, actualCol).Value = result.Actual;
            ws.Cell(row, notesCol).Value = result.Passed ? "Passed" : "Falled";
            WriteScreenshotToCell(ws, row, screenshotCol, result);
        }

        workbook.Save();

        Assert.Pass($"Updated: {excelPath} | Sheet: {ws.Name} | Range: NO 21-45");
    }

    [Test]
    public void Run_No46To70_And_Write_Back_To_Excel()
    {
        var excelPath = ResolveWritableExcelPath();
        var config = TestDataLoader.LoadConfig();
        var screenshotDir = ResolveScreenshotDirectory();

        using var workbook = new XLWorkbook(excelPath);
        var ws = FindWorksheet(workbook);
        var headerRow = FindHeaderRow(ws);

        var noCol = FindOrCreateColumn(ws, headerRow, new[] { "no.", "no", "stt" }, "NO.");
        var expectedCol = FindOrCreateColumn(ws, headerRow, new[] { "expectedresult", "expected result" }, "Expected Result");
        var actualCol = FindOrCreateColumn(ws, headerRow, new[] { "actualresult", "actual result" }, "Actual Result");
        var notesCol = FindOrCreateColumn(ws, headerRow, new[] { "notes", "note" }, "Notes");
        var screenshotCol = FindOrCreateColumn(ws, headerRow, new[] { "screenshot", "screen shot" }, "Screenshot");

        var lastUsedRow = ws.LastRowUsed();
        if (lastUsedRow is null)
        {
            throw new InvalidOperationException("Worksheet has no used rows.");
        }

        var lastRow = lastUsedRow.RowNumber();

        for (var row = headerRow + 1; row <= lastRow; row++)
        {
            var noText = ws.Cell(row, noCol).GetString().Trim();
            if (!int.TryParse(noText, out var no) || no < 46 || no > 70)
            {
                continue;
            }

            var expected = ws.Cell(row, expectedCol).GetString().Trim();
            var result = ExecuteCase(no, expected, config, screenshotDir);

            ws.Cell(row, actualCol).Value = result.Actual;
            ws.Cell(row, notesCol).Value = result.Passed ? "Passed" : "Falled";
            WriteScreenshotToCell(ws, row, screenshotCol, result);
        }

        workbook.Save();

        Assert.Pass($"Updated: {excelPath} | Sheet: {ws.Name} | Range: NO 46-70");
    }

    private static CaseResult ExecuteCase(int no, string expected, TestConfig config, string screenshotDir)
    {
        using var driver = DriverFactory.CreateChromeDriver(config.Headless);
        var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(12));

        try
        {
            var baseUrl = config.BaseUrl.TrimEnd('/');
            var loginUrl = $"{baseUrl}/Admin/AdminAuth/Login";

            var passed = no switch
            {
                1 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers?page=1"),
                2 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers?searchName=An"),
                3 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers?searchEmail=@gmail.com"),
                4 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers?searchPhone=09"),
                5 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers?membershipId=1"),
                6 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers?searchName=An&searchEmail=@gmail.com&searchPhone=09&membershipId=1"),
                7 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers"),
                8 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers"),
                9 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers"),
                10 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers"),
                11 => AuthenticatedPageCheckAndKeyword(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers/Details/999999", "notfound", "404", "không tìm thấy", "khong tim thay"),
                12 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers/Details/1"),
                13 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers/Details/1"),
                14 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers/Details/1"),
                15 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers/Details/1"),
                16 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers/Details/1"),
                17 => AuthenticatedPageCheckAndKeyword(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers/Edit/999999", "notfound", "404", "không tìm thấy", "khong tim thay"),
                18 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers/Edit/1"),
                19 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers/Edit/1"),
                20 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers/Edit/1"),
                21 => AuthenticatedPageCheckAndKeyword(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers/BlockCustomer?id=999999", "success", "false", "not found", "không tìm thấy"),
                22 => AuthenticatedPageCheckAndKeyword(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers/BlockCustomer?id=1", "success", "true"),
                23 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers/GetCustomerOrders?customerId=1"),
                24 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile"),
                25 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile?searchName=An"),
                26 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile?searchEmail=@gmail.com"),
                27 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Details/999999"),
                28 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Details/1"),
                29 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Edit/1"),
                30 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Edit/1"),
                31 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Edit/1"),
                32 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Edit/1"),
                33 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Edit/1"),
                34 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Edit/1"),
                35 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Edit/1"),
                36 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Edit/1"),
                37 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Edit/1"),
                38 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Edit/1"),
                39 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Delete/999999"),
                40 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Delete/1"),
                41 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Delete/2"),
                42 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders?page=1"),
                43 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders?searchName=An"),
                44 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders?searchName=@gmail.com"),
                45 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders?searchName=09"),
                46 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders?status=Đã hủy"),
                47 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders?fromDate=2026-01-01"),
                48 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders?toDate=2026-01-31"),
                49 => AuthenticatedPageCheckAndKeyword(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Details/999999", "notfound", "404", "không tìm thấy", "khong tim thay"),
                50 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Details/1"),
                51 => AuthenticatedPageCheckAndKeyword(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Edit/999999", "notfound", "404", "không tìm thấy", "khong tim thay"),
                52 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Edit/1"),
                53 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Edit/1"),
                54 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Edit/1"),
                55 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Edit/1"),
                56 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Edit/1"),
                57 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Edit/1"),
                58 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Edit/1"),
                59 => AuthenticatedPageCheckAndKeyword(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Cancel?id=999999", "success", "false", "not found", "không tìm thấy"),
                60 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Cancel?id=1"),
                61 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Cancel?id=2"),
                62 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Cancel?id=2"),
                63 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders"),
                64 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders/Edit/1"),
                65 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers?page=999"),
                66 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Edit/1"),
                67 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Edit/1"),
                68 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Profile/Edit/1"),
                69 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Customers/Details/1"),
                70 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Orders?searchName=An&status=Đã hủy&fromDate=2026-01-01&toDate=2026-01-31"),
                _ => false,
            };

            var strict = EvaluateExpectedStrict(driver, expected);
            passed = passed && strict.IsMatch;

            var screenshotName = $"NO{no:00}_{DateTime.Now:yyyyMMdd_HHmmss}.png";
            var screenshotFile = Path.Combine(screenshotDir, screenshotName);
            CaptureScreenshot(driver, screenshotFile);

            var actual = passed
                ? $"Observed behavior matched expected result for NO {no}."
                : $"Observed behavior did not match expected result for NO {no}.";

            if (!string.IsNullOrWhiteSpace(expected))
            {
                actual = $"Expected: {expected} | Actual: {actual}";
            }

            actual = $"{actual} | StrictCheck: {strict.Reason}";

            return new CaseResult(passed, actual, screenshotFile);
        }
        catch (Exception ex)
        {
            var screenshotName = $"NO{no:00}_ERROR_{DateTime.Now:yyyyMMdd_HHmmss}.png";
            var screenshotFile = Path.Combine(screenshotDir, screenshotName);
            CaptureScreenshot(driver, screenshotFile);

            return new CaseResult(false, $"Automation exception: {ex.Message}", screenshotFile);
        }
    }

    private static bool AuthenticatedPageCheck(IWebDriver d, WebDriverWait w, string loginUrl, TestConfig config, string targetUrl)
    {
        d.Navigate().GoToUrl(loginUrl);

        if (HasElement(d, By.CssSelector("input[name='username'],input[name='Username']"))
            && HasElement(d, By.CssSelector("input[name='password'],input[name='Password']")))
        {
            FillInput(d, new[] { "username", "Username" }, config.ValidUser.Username);
            FillInput(d, new[] { "password", "Password" }, config.ValidUser.Password);
            SubmitForm(d);

            try
            {
                w.Until(_ => !d.Url.Contains("/Admin/AdminAuth/Login", StringComparison.OrdinalIgnoreCase));
            }
            catch (WebDriverTimeoutException)
            {
                return false;
            }
        }

        d.Navigate().GoToUrl(targetUrl);

        if (d.Url.Contains("/Admin/AdminAuth/Login", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var source = d.PageSource;
        if (string.IsNullOrWhiteSpace(d.Url) || source.Length < 100)
        {
            return false;
        }

        if (source.Contains("error", StringComparison.OrdinalIgnoreCase)
            && source.Contains("exception", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return true;
    }

    private static bool AuthenticatedPageCheckAndKeyword(IWebDriver d, WebDriverWait w, string loginUrl, TestConfig config, string targetUrl, params string[] keywords)
    {
        if (!AuthenticatedPageCheck(d, w, loginUrl, config, targetUrl))
        {
            return false;
        }

        if (keywords.Length == 0)
        {
            return true;
        }

        var source = d.PageSource;
        return keywords.Any(k => source.Contains(k, StringComparison.OrdinalIgnoreCase));
    }

    private static bool HasElement(IWebDriver driver, By by)
    {
        return driver.FindElements(by).Count > 0;
    }

    private static void FillInput(IWebDriver driver, IEnumerable<string> names, string value)
    {
        IWebElement? element = null;

        foreach (var name in names)
        {
            element = driver.FindElements(By.Name(name)).FirstOrDefault()
                ?? driver.FindElements(By.Id(name)).FirstOrDefault()
                ?? driver.FindElements(By.CssSelector($"input[name='{name}']")).FirstOrDefault()
                ?? driver.FindElements(By.CssSelector($"input[id='{name}']")).FirstOrDefault();

            if (element is not null)
            {
                break;
            }
        }

        if (element is null)
        {
            throw new InvalidOperationException($"Input element not found. Candidates: {string.Join(", ", names)}");
        }

        element.Clear();
        element.SendKeys(value);
    }

    private static void SubmitForm(IWebDriver driver)
    {
        var submit = driver.FindElements(By.CssSelector("button[type='submit'],input[type='submit']")).FirstOrDefault();
        if (submit is null)
        {
            throw new InvalidOperationException("Submit button not found.");
        }

        submit.Click();
    }

    private static void CaptureScreenshot(IWebDriver driver, string filePath)
    {
        var folder = Path.GetDirectoryName(filePath)!;
        Directory.CreateDirectory(folder);

        if (driver is not ITakesScreenshot shotDriver)
        {
            return;
        }

        var shot = shotDriver.GetScreenshot();
        shot.SaveAsFile(filePath);
    }

    private static void WriteScreenshotToCell(IXLWorksheet ws, int row, int col, CaseResult result)
    {
        ws.Cell(row, col).Value = string.Empty;

        try
        {
            ws.Column(col).Width = Math.Max(ws.Column(col).Width, 34);
            ws.Row(row).Height = Math.Max(ws.Row(row).Height, 100);

            ws.AddPicture(result.ScreenshotAbsolutePath)
                .MoveTo(ws.Cell(row, col))
                .WithSize(240, 120);
        }
        catch
        {
            ws.Cell(row, col).Value = result.ScreenshotAbsolutePath;
        }
    }

    private static IXLWorksheet FindWorksheet(XLWorkbook workbook)
    {
        var target = workbook.Worksheets.FirstOrDefault(w =>
            string.Equals(w.Name, SheetAlias1, StringComparison.OrdinalIgnoreCase)
            || string.Equals(w.Name, SheetAlias2, StringComparison.OrdinalIgnoreCase));

        if (target is null)
        {
            throw new InvalidOperationException("Worksheet TC_Customers_Orders not found.");
        }

        return target;
    }

    private static int FindHeaderRow(IXLWorksheet ws)
    {
        for (var row = 1; row <= 30; row++)
        {
            for (var col = 1; col <= 25; col++)
            {
                var value = Normalize(ws.Cell(row, col).GetString());
                if (value is "no" or "no.")
                {
                    return row;
                }
            }
        }

        throw new InvalidOperationException("Could not locate header row.");
    }

    private static int FindOrCreateColumn(IXLWorksheet ws, int headerRow, IEnumerable<string> aliases, string createHeader)
    {
        var normalizedAliases = aliases.Select(Normalize).ToArray();
        var lastUsedCol = ws.LastColumnUsed();
        var lastCol = lastUsedCol?.ColumnNumber() ?? 1;

        for (var col = 1; col <= lastCol; col++)
        {
            var v = Normalize(ws.Cell(headerRow, col).GetString());
            if (normalizedAliases.Any(a => v.Equals(a, StringComparison.OrdinalIgnoreCase)))
            {
                return col;
            }
        }

        var newCol = lastCol + 1;
        ws.Cell(headerRow, newCol).Value = createHeader;
        return newCol;
    }

    private static string Normalize(string input)
    {
        return input
            .Trim()
            .ToLowerInvariant()
            .Replace(" ", string.Empty)
            .Replace("_", string.Empty);
    }

    private static StrictCheckResult EvaluateExpectedStrict(IWebDriver driver, string expected)
    {
        if (string.IsNullOrWhiteSpace(expected))
        {
            return new StrictCheckResult(true, "No expected text to validate.");
        }

        var tokens = ExtractExpectedTokens(expected);
        if (tokens.Count == 0)
        {
            return new StrictCheckResult(false, "Expected text has no verifiable token.");
        }

        var combined = $"{driver.Url}\n{driver.Title}\n{driver.PageSource}".ToLowerInvariant();
        var matched = tokens.Count(t => combined.Contains(t, StringComparison.Ordinal));
        var required = GetRequiredTokenMatches(tokens.Count);

        if (matched >= required)
        {
            return new StrictCheckResult(true, $"Matched {matched}/{tokens.Count} tokens (required {required}).");
        }

        return new StrictCheckResult(false, $"Matched {matched}/{tokens.Count} tokens (required {required}).");
    }

    private static List<string> ExtractExpectedTokens(string expected)
    {
        var separators = new[]
        {
            ' ', '\t', '\r', '\n', ',', '.', ';', ':', '|', '-', '_', '/', '\\',
            '(', ')', '[', ']', '{', '}', '?', '!', '"', '\''
        };

        var stopWords = new HashSet<string>(StringComparer.Ordinal)
        {
            "and", "the", "with", "from", "this", "that", "true", "false", "page", "data",
            "noi", "dung", "ket", "qua", "duoc", "hien", "thi", "cho", "khi", "nhap", "vao"
        };

        return expected
            .ToLowerInvariant()
            .Split(separators, StringSplitOptions.RemoveEmptyEntries)
            .Select(p => p.Trim())
            .Where(p => p.Length >= 4 && !stopWords.Contains(p))
            .Distinct(StringComparer.Ordinal)
            .Take(12)
            .ToList();
    }

    private static int GetRequiredTokenMatches(int tokenCount)
    {
        if (tokenCount <= 2)
        {
            return tokenCount;
        }

        if (tokenCount <= 4)
        {
            return 2;
        }

        if (tokenCount <= 7)
        {
            return 3;
        }

        return 4;
    }

    private static string ResolveExcelPath()
    {
        var candidate = Path.GetFullPath(Path.Combine(
            TestContext.CurrentContext.TestDirectory,
            "..",
            "..",
            "..",
            "..",
            "Nhom8 1.xlsx"));

        if (!File.Exists(candidate))
        {
            throw new FileNotFoundException("Excel file not found.", candidate);
        }

        return candidate;
    }

    private static string ResolveWritableExcelPath()
    {
        var sourcePath = ResolveExcelPath();

        try
        {
            using var fs = new FileStream(sourcePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return sourcePath;
        }
        catch (IOException)
        {
            var sourceDir = Path.GetDirectoryName(sourcePath)!;
            var copyPath = Path.Combine(sourceDir, "Nhom8 1_AUTOMATED.xlsx");
            File.Copy(sourcePath, copyPath, true);
            return copyPath;
        }
    }

    private static string ResolveScreenshotDirectory()
    {
        var screenshotDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "TC_CUSTOMERS_ORDERS");
        Directory.CreateDirectory(screenshotDir);
        return screenshotDir;
    }

    private sealed record CaseResult(bool Passed, string Actual, string ScreenshotAbsolutePath);
    private sealed record StrictCheckResult(bool IsMatch, string Reason);
}
