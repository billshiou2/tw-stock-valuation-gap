# Project

本專案採用簡潔的資料夾結構，將原始碼、文件、設定、輸出與暫存內容分開管理。

## 專案結構

- `AGENTS.md`：agent 必須遵守的工作規則。
- `AGENT_PROGRESS.md`：目前進度、重點討論與交接資訊。
- `.env.example`：本專案需要的環境變數名稱與格式。
- `config/`：非機密、可提交的設定檔。
- `docs/`：文件、需求、規格與說明。
- `output/`：程式或 agent 產生的輸出，預設不提交。
- `src/`：正式原始碼。
- `tmp/`：暫存檔，預設不提交。

## 環境變數

實際環境變數放在專案根目錄的 `.env`，不得提交到版本控制。

`.env.example` 用來說明需要的環境變數名稱與格式，不得填入真實密鑰、token 或密碼。

## Git 上傳

第一次上傳 GitHub 前，先參考 `docs/git-workflow.md`。建議先在 GitHub 網頁人工建立 repo，確認名稱、公開狀態與權限後，再用 Git 指令連接遠端並 push。
