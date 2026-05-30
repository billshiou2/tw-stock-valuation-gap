# Git 上傳流程

本流程採用「先人工檢視，再用 Git 指令上傳」的方式。建議先在 GitHub 網頁建立 repo，確認名稱、公開狀態與權限後，再把本機專案推上去。

## 上傳前檢查

先確認目前有哪些檔案會被 Git 看見：

```powershell
git status --short
```

確認目前所在專案與 Git 目標能合理對應：

```powershell
Get-Location
git remote -v
git status --short --branch
git log -1 --oneline
```

檢查時要比對：

- 目前資料夾名稱是否能合理對應到要上傳的專案。
- README 或專案說明是否能合理對應到要上傳的專案。
- remote URL 的 repo 名稱是否能合理對應到目前專案；若尚未設定 remote，則比對準備要設定的 repo URL。
- 目前 branch 與 upstream 是否是預期目標；若尚未設定 upstream，則確認準備推送的 branch。

資料夾名稱、專案說明、remote repo 名稱與 branch 不需要逐字相同，但要能合理對應到同一個專案。若出現無法合理對應的差異、看起來像不同專案、或與使用者指定目標不符，先停止上傳並向使用者確認，不要直接 `git push`。

確認本機機密與產生檔不會被提交：

```powershell
git check-ignore -v .env
git check-ignore -v output/test.txt
git check-ignore -v tmp/test.txt
```

如果 `output/test.txt` 或 `tmp/test.txt` 不存在，可以先略過；重點是確認 `.gitignore` 規則會保護 `.env`、`output/` 與 `tmp/`。

## 初始化本機 Git

```powershell
git init
git add .
git commit -m "Initial project structure"
git branch -M main
```

## 在 GitHub 建立 repo

到 GitHub 網頁建立新的 repo。

建議：

- 依專案需求選擇 private 或 public。
- 不要在 GitHub 預先加入 README、.gitignore 或 license，避免和本機檔案衝突。
- 建立後複製 repo URL。

## 連接遠端並上傳

```powershell
git remote add origin <你的 repo URL>
git push -u origin main
```

如果 `origin` 已經存在，先用 `git remote -v` 確認它能合理對應到這個專案的 GitHub repo。若看起來像不同專案，不要直接覆蓋 remote，也不要 push，先回報目前 remote 與預期 repo 的差異。

## 上傳後確認

```powershell
git remote -v
git status
```

最後到 GitHub 網頁確認檔案內容，特別檢查 `.env`、`output/`、`tmp/`、依賴資料夾與快取檔沒有被提交。

## 可選：GitHub CLI

GitHub CLI 是 GitHub 的命令列工具，指令名稱是 `gh`。它可以從 PowerShell 建立 repo、設定 remote 並 push。

本專案預設不依賴 GitHub CLI。若需要更自動化的流程，再另外加入 `gh` 相關指令。
