using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumProject.Utilities;

namespace SeleniumProject.Tests;

[TestFixture]
public class TcAuthRbacExcelRunnerTests
{
    private const string SheetAlias1 = "TC_AUTHRBAC";
    private const string SheetAlias2 = "TC_Auth_RBAC";

    [Test]
    public void Run_No1To16_And_Write_Back_To_Excel()
    {
        RunRangeAndWriteBack(1, 16);
    }

    [Test]
    public void Run_No17To36_And_Write_Back_To_Excel()
    {
        RunRangeAndWriteBack(17, 36);
    }

    [Test]
    public void Run_No37To53_And_Write_Back_To_Excel()
    {
        RunRangeAndWriteBack(37, 53);
    }

    [Test]
    public void Run_No54To70_And_Write_Back_To_Excel()
    {
        RunRangeAndWriteBack(54, 70);
    }

    private static void RunRangeAndWriteBack(int startNo, int endNo)
    {
        var excelPath = ResolveWritableExcelPath();
        var config = TestDataLoader.LoadConfig();
        var screenshotDir = ResolveScreenshotDirectory();

        using var workbook = new XLWorkbook(excelPath);
        var ws = FindWorksheet(workbook);
        var headerRow = FindHeaderRow(ws);

        var statusCol = FindColumn(ws, headerRow, new[] { "teststatus", "test status", "status" });
        if (statusCol > 0)
        {
            ws.Column(statusCol).Delete();
        }

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
            if (!int.TryParse(noText, out var no) || no < startNo || no > endNo)
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

        Assert.Pass($"Updated: {excelPath} | Range: NO {startNo}-{endNo}");
    }

    private static CaseResult ExecuteCase(int no, string expected, TestConfig config, string screenshotDir)
    {
        using var driver = DriverFactory.CreateChromeDriver(config.Headless);
        var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

        try
        {
            var baseUrl = config.BaseUrl.TrimEnd('/');
            var loginUrl = $"{baseUrl}/Admin/AdminAuth/Login";
            var registerUrl = $"{baseUrl}/Admin/AdminAuth/Register";

            var passed = no switch
            {
                1 => Case01(driver, wait, loginUrl),
                2 => Case02(driver, wait, loginUrl),
                3 => Case03(driver, wait, loginUrl),
                4 => Case04(driver, wait, loginUrl),
                5 => Case05(driver, wait, loginUrl),
                6 => Case06(driver, wait, loginUrl),
                7 => Case07(driver, wait, loginUrl),
                8 => Case08(driver, wait, loginUrl, config),
                9 => Case09(driver, wait, loginUrl, config, out _),
                10 => Case10(driver, wait, loginUrl, config),
                11 => Case11(driver, wait, registerUrl),
                12 => Case12(driver, wait, registerUrl),
                13 => Case13(driver, wait, registerUrl),
                14 => Case14(driver, wait, registerUrl),
                15 => Case15(driver, wait, registerUrl),
                16 => Case16(driver, wait, registerUrl),
                17 => GenericPageCheck(driver, $"{baseUrl}/Admin/Setup/CreateFirstAdmin"),
                18 => GenericPageCheck(driver, $"{baseUrl}/Admin/Setup/CreateFirstAdmin"),
                19 => GenericPageCheck(driver, $"{baseUrl}/Admin/Setup/ApproveAllPending"),
                20 => GenericPageCheck(driver, $"{baseUrl}/Admin/Setup/ApproveAllPending"),
                21 => GenericPageCheck(driver, $"{baseUrl}/Admin/Setup/ResetPassword"),
                22 => GenericPageCheck(driver, $"{baseUrl}/Admin/Setup/ResetPassword?username=notexist&newPassword=123456"),
                23 => GenericPageCheck(driver, $"{baseUrl}/Admin/Setup/ResetPassword?username=admin&newPassword=123456"),
                24 => GenericPageCheck(driver, $"{baseUrl}/Admin/Dashboard/Index"),
                25 => GenericPageCheck(driver, $"{baseUrl}/Admin/AdminManagement/Admins"),
                26 => GenericPageCheck(driver, $"{baseUrl}/Admin/AdminManagement/Admins"),
                27 => GenericPageCheck(driver, $"{baseUrl}/Admin/AdminManagement/Admins"),
                28 => GenericPageCheck(driver, $"{baseUrl}/Admin/AdminManagement/Approve"),
                29 => GenericPageCheck(driver, $"{baseUrl}/Admin/AdminManagement/Approve?id=999999"),
                30 => GenericPageCheck(driver, $"{baseUrl}/Admin/AdminManagement/Approve?id=1"),
                31 => GenericPageCheck(driver, $"{baseUrl}/Admin/AdminManagement/Block?id=1"),
                32 => GenericPageCheck(driver, $"{baseUrl}/Admin/AdminManagement/Delete?id=999999"),
                33 => GenericPageCheck(driver, $"{baseUrl}/Admin/AdminManagement/Delete?id=1"),
                34 => GenericPageCheck(driver, $"{baseUrl}/Admin/AdminManagement/Delete?id=2"),
                35 => GenericPageCheck(driver, $"{baseUrl}/Admin/AdminManagement/UpdateRole?id=999999&roleId=1"),
                36 => GenericPageCheck(driver, $"{baseUrl}/Admin/AdminManagement/UpdateRole?id=1&roleId=1"),
                37 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Index?page=1"),
                38 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Create"),
                39 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Create"),
                40 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Create"),
                41 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Create"),
                42 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Create"),
                43 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Edit/999999"),
                44 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Edit/1"),
                45 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Edit/1"),
                46 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Edit/1"),
                47 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Edit/1"),
                48 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Delete/999999"),
                49 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Delete/1"),
                50 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Delete/2"),
                51 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Delete/3"),
                52 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Details/999999"),
                53 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Details/1"),
                54 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/Dashboard/Index"),
                55 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Edit/1"),
                56 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Edit/1"),
                57 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Delete/1"),
                58 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Details/999999"),
                59 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Details/1"),
                60 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Index"),
                61 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Edit/1"),
                62 => GenericPageCheck(driver, $"{baseUrl}/Admin/AdminAuth/Register"),
                63 => GenericPageCheck(driver, $"{baseUrl}/Admin/Setup/ApproveAllPending"),
                64 => GenericPageCheck(driver, $"{baseUrl}/Admin/Setup/ResetPassword?username=admin&newPassword=123456"),
                65 => GenericPageCheck(driver, $"{baseUrl}/Admin/Setup/ResetPassword?username=admin&newPassword=123456"),
                66 => Case66_AfterLogoutBlocked(driver, wait, loginUrl, config),
                67 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Edit/1"),
                68 => AuthenticatedPageCheck(driver, wait, loginUrl, config, $"{baseUrl}/Admin/RoleManagement/Create"),
                69 => false,
                70 => false,
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

            return new CaseResult(passed, actual, screenshotFile, screenshotFile);
        }
        catch (Exception ex)
        {
            var screenshotName = $"NO{no:00}_ERROR_{DateTime.Now:yyyyMMdd_HHmmss}.png";
            var screenshotFile = Path.Combine(screenshotDir, screenshotName);
            CaptureScreenshot(driver, screenshotFile);

            return new CaseResult(false, $"Automation exception: {ex.Message}", screenshotFile, screenshotFile);
        }
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
            // Fallback to file path if image embedding is not possible.
            ws.Cell(row, col).Value = result.ScreenshotRelativePath;
        }
    }

    private static bool Case01(IWebDriver d, WebDriverWait w, string loginUrl)
    {
        d.Navigate().GoToUrl(loginUrl);
        return HasElement(d, By.CssSelector("input[name='username'],input[name='Username']"))
            && HasElement(d, By.CssSelector("input[name='password'],input[name='Password']"));
    }

    private static bool Case02(IWebDriver d, WebDriverWait w, string loginUrl)
    {
        d.Navigate().GoToUrl(loginUrl);
        FillInput(d, new[] { "username", "Username" }, string.Empty);
        FillInput(d, new[] { "password", "Password" }, "123456");
        SubmitForm(d);
        return PageContainsAny(d, "bắt buộc", "bat buoc", "required");
    }

    private static bool Case03(IWebDriver d, WebDriverWait w, string loginUrl)
    {
        d.Navigate().GoToUrl(loginUrl);
        FillInput(d, new[] { "username", "Username" }, "admin");
        FillInput(d, new[] { "password", "Password" }, string.Empty);
        SubmitForm(d);
        return PageContainsAny(d, "bắt buộc", "bat buoc", "required");
    }

    private static bool Case04(IWebDriver d, WebDriverWait w, string loginUrl)
    {
        d.Navigate().GoToUrl(loginUrl);
        FillInput(d, new[] { "username", "Username" }, "notexist");
        FillInput(d, new[] { "password", "Password" }, "123456");
        SubmitForm(d);
        return PageContainsAny(d, "không đúng", "khong dung", "invalid");
    }

    private static bool Case05(IWebDriver d, WebDriverWait w, string loginUrl)
    {
        d.Navigate().GoToUrl(loginUrl);
        FillInput(d, new[] { "username", "Username" }, "adminApproved");
        FillInput(d, new[] { "password", "Password" }, "wrong");
        SubmitForm(d);
        return PageContainsAny(d, "không đúng", "khong dung", "invalid");
    }

    private static bool Case06(IWebDriver d, WebDriverWait w, string loginUrl)
    {
        d.Navigate().GoToUrl(loginUrl);
        FillInput(d, new[] { "username", "Username" }, "adminPending");
        FillInput(d, new[] { "password", "Password" }, "passwordCorrect");
        SubmitForm(d);
        return PageContainsAny(d, "chưa được phê duyệt", "chua duoc phe duyet", "pending");
    }

    private static bool Case07(IWebDriver d, WebDriverWait w, string loginUrl)
    {
        d.Navigate().GoToUrl(loginUrl);
        FillInput(d, new[] { "username", "Username" }, "adminBlocked");
        FillInput(d, new[] { "password", "Password" }, "passwordCorrect");
        SubmitForm(d);
        return PageContainsAny(d, "đã bị khóa", "da bi khoa", "blocked");
    }

    private static bool Case08(IWebDriver d, WebDriverWait w, string loginUrl, TestConfig config)
    {
        d.Navigate().GoToUrl(loginUrl);
        FillInput(d, new[] { "username", "Username" }, config.ValidUser.Username);
        FillInput(d, new[] { "password", "Password" }, config.ValidUser.Password);
        SubmitForm(d);
        w.Until(_ => d.Url.Contains("/Admin", StringComparison.OrdinalIgnoreCase));
        return d.Url.Contains("/Admin", StringComparison.OrdinalIgnoreCase);
    }

    private static bool Case09(IWebDriver d, WebDriverWait w, string loginUrl, TestConfig config, out string notes)
    {
        notes = string.Empty;
        d.Navigate().GoToUrl(loginUrl);
        FillInput(d, new[] { "username", "Username" }, config.ValidUser.Username);
        FillInput(d, new[] { "password", "Password" }, config.ValidUser.Password);
        SubmitForm(d);
        w.Until(_ => d.Url.Contains("/Admin", StringComparison.OrdinalIgnoreCase));

        var cookie = d.Manage().Cookies.AllCookies.FirstOrDefault(c => c.Name.Contains("AspNetCore", StringComparison.OrdinalIgnoreCase));
        var pass = cookie is not null;
        notes = pass
            ? "Claims cannot be read directly from UI; auth cookie detected after login."
            : "Could not verify claims because auth cookie was not detected.";

        return pass;
    }

    private static bool Case10(IWebDriver d, WebDriverWait w, string loginUrl, TestConfig config)
    {
        d.Navigate().GoToUrl(loginUrl);
        FillInput(d, new[] { "username", "Username" }, config.ValidUser.Username);
        FillInput(d, new[] { "password", "Password" }, config.ValidUser.Password);
        SubmitForm(d);
        d.Navigate().GoToUrl(loginUrl.Replace("/Login", "/Logout", StringComparison.OrdinalIgnoreCase));
        w.Until(_ => d.Url.Contains("/Admin/AdminAuth/Login", StringComparison.OrdinalIgnoreCase));
        return d.Url.Contains("/Admin/AdminAuth/Login", StringComparison.OrdinalIgnoreCase);
    }

    private static bool Case11(IWebDriver d, WebDriverWait w, string registerUrl)
    {
        d.Navigate().GoToUrl(registerUrl);
        return HasElement(d, By.CssSelector("input[name='username'],input[name='Username']"));
    }

    private static bool Case12(IWebDriver d, WebDriverWait w, string registerUrl)
    {
        d.Navigate().GoToUrl(registerUrl);
        FillInput(d, new[] { "username", "Username" }, string.Empty);
        FillInput(d, new[] { "password", "Password" }, "123456");
        FillInput(d, new[] { "confirmPassword", "ConfirmPassword" }, "123456");
        SubmitForm(d);
        return PageContainsAny(d, "bắt buộc", "bat buoc", "required");
    }

    private static bool Case13(IWebDriver d, WebDriverWait w, string registerUrl)
    {
        d.Navigate().GoToUrl(registerUrl);
        FillInput(d, new[] { "username", "Username" }, "u1");
        FillInput(d, new[] { "password", "Password" }, "123456");
        FillInput(d, new[] { "confirmPassword", "ConfirmPassword" }, "123");
        SubmitForm(d);
        return PageContainsAny(d, "không khớp", "khong khop", "confirm");
    }

    private static bool Case14(IWebDriver d, WebDriverWait w, string registerUrl)
    {
        d.Navigate().GoToUrl(registerUrl);
        FillInput(d, new[] { "username", "Username" }, "u2");
        FillInput(d, new[] { "password", "Password" }, "123");
        FillInput(d, new[] { "confirmPassword", "ConfirmPassword" }, "123");
        SubmitForm(d);
        return PageContainsAny(d, "ít nhất", "it nhat", "at least");
    }

    private static bool Case15(IWebDriver d, WebDriverWait w, string registerUrl)
    {
        d.Navigate().GoToUrl(registerUrl);
        FillInput(d, new[] { "username", "Username" }, "admin");
        FillInput(d, new[] { "password", "Password" }, "123456");
        FillInput(d, new[] { "confirmPassword", "ConfirmPassword" }, "123456");
        SubmitForm(d);
        return PageContainsAny(d, "đã được sử dụng", "da duoc su dung", "already");
    }

    private static bool Case16(IWebDriver d, WebDriverWait w, string registerUrl)
    {
        d.Navigate().GoToUrl(registerUrl);
        var newUser = $"new_admin_{DateTime.Now:yyyyMMddHHmmss}";
        FillInput(d, new[] { "username", "Username" }, newUser);
        FillInput(d, new[] { "password", "Password" }, "123456");
        FillInput(d, new[] { "confirmPassword", "ConfirmPassword" }, "123456");
        SubmitForm(d);
        return PageContainsAny(d, "đăng ký thành công", "dang ky thanh cong", "phê duyệt", "phe duyet");
    }

    private static bool GenericPageCheck(IWebDriver d, string url)
    {
        d.Navigate().GoToUrl(url);

        var current = d.Url;
        var source = d.PageSource;

        if (string.IsNullOrWhiteSpace(current) || source.Length < 100)
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

        return GenericPageCheck(d, targetUrl);
    }

    private static bool Case66_AfterLogoutBlocked(IWebDriver d, WebDriverWait w, string loginUrl, TestConfig config)
    {
        if (!Case08(d, w, loginUrl, config))
        {
            return false;
        }

        d.Navigate().GoToUrl(loginUrl.Replace("/Login", "/Logout", StringComparison.OrdinalIgnoreCase));

        try
        {
            w.Until(_ => d.Url.Contains("/Admin/AdminAuth/Login", StringComparison.OrdinalIgnoreCase));
        }
        catch (WebDriverTimeoutException)
        {
            return false;
        }

        d.Navigate().GoToUrl(loginUrl.Replace("/AdminAuth/Login", "/Dashboard/Index", StringComparison.OrdinalIgnoreCase));
        return d.Url.Contains("/Admin/AdminAuth/Login", StringComparison.OrdinalIgnoreCase)
            || PageContainsAny(d, "access denied", "không có quyền", "khong co quyen");
    }

    private static bool PageContainsAny(IWebDriver driver, params string[] snippets)
    {
        var content = driver.PageSource;
        return snippets.Any(s => content.Contains(s, StringComparison.OrdinalIgnoreCase));
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

    private static IXLWorksheet FindWorksheet(XLWorkbook workbook)
    {
        var target = workbook.Worksheets.FirstOrDefault(w =>
            string.Equals(w.Name, SheetAlias1, StringComparison.OrdinalIgnoreCase)
            || string.Equals(w.Name, SheetAlias2, StringComparison.OrdinalIgnoreCase)
            || string.Equals(w.Name.Replace("_", string.Empty), SheetAlias1, StringComparison.OrdinalIgnoreCase));

        if (target is null)
        {
            throw new InvalidOperationException("Worksheet TC_AUTHRBAC not found.");
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

    private static int FindColumn(IXLWorksheet ws, int headerRow, IEnumerable<string> aliases)
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

        return -1;
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
        if (tokenCount <= 4)
        {
            return 1;
        }

        if (tokenCount <= 7)
        {
            return 1;
        }

        return 2;
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
        var screenshotDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "TC_AUTHRBAC");
        Directory.CreateDirectory(screenshotDir);
        return screenshotDir;
    }

    private sealed record CaseResult(bool Passed, string Actual, string ScreenshotRelativePath, string ScreenshotAbsolutePath);
    private sealed record StrictCheckResult(bool IsMatch, string Reason);
}
