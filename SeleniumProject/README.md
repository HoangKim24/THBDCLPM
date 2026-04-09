# SeleniumProject

Cau truc Page Object Model (POM):

- `Pages`: Moi page la 1 class, chua locator va action
- `Tests`: Test chi goi method tu Page
- `Utilities`: Khoi tao WebDriver va doc test data
- `TestData`: Du lieu test JSON

## Cac nhom test
- Smoke Test: Login thanh cong, Logout, Transfer funds
- GUI Test: Login button, textbox, dieu huong link, thong bao loi
- Functional Test: Register user, Bill pay, Open new account, Find transactions

## Chay test
```bash
dotnet test SeleniumProject.csproj
```

## Luu y
- Driver mac dinh: Chrome (headless)
- Base URL va tai khoan mau duoc cau hinh trong `TestData/users.json`
