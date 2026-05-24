# Agent Progress

## 目前目標

- 建立可直接使用的乾淨專案起始結構，並讓後續 agent 能接續目前進度。

## 目前進度

- 已保留 `AGENTS.md` 原本的套件版本安全規則。
- 已追加專案結構、機密管理與 agent 進度記錄規則。
- 已建立 `README.md`、`.gitignore`、`.env.example` 與基礎資料夾。
- 已新增 `.gitkeep` 以保留基礎空資料夾。
- 已新增 `docs/git-workflow.md`，記錄人工建立 GitHub repo 後的 Git 上傳流程與上傳前檢查。
- 已在 `README.md` 加入 Git 上傳流程入口。
- 使用者決定要將此專案上傳到 Git；已完成本機 Git 初始化與第一個 commit，目前等待 GitHub repo URL 以設定遠端並 push。

## 已決定事項

- 原本的套件版本安全規則不得刪除、修改或弱化。
- 不建立 `docs/DECISIONS.md`，交接重點集中記錄在 `AGENT_PROGRESS.md`。
- 不預設 Node、Python、React 或其他特定框架。
- 不使用「範本」、`template`、「請替換」等字眼描述本專案。
- `.env` 放在專案根目錄但不得提交，`.env.example` 可以提交。
- 新增、刪除或修改 `.env` 的環境變數時，必須同步更新 `.env.example`。
- `.env.example` 是環境變數名稱、格式與用途說明來源。
- `.env` 與 `.env.example` 的非敏感預設值應盡量保持一致；敏感或本機專用值可以不同。
- 任務收尾前必須檢查 `AGENT_PROGRESS.md`、環境變數同步、`README.md` 更新需求與測試執行情況。
- `config/` 放非機密、可提交的設定。
- `docs/` 放文件、需求、規格與說明。
- `output/` 放產生輸出，預設不提交。
- `src/` 放正式原始碼。
- `tmp/` 放暫存檔，預設不提交。
- 使用 `.gitkeep` 保留 `config/`、`docs/`、`src/`、`output/`、`tmp/` 的資料夾結構。
- GitHub repo 建議先由使用者在網頁人工建立，確認名稱、private/public 與權限後，再用 Git 指令上傳。
- GitHub CLI 不列入預設流程，只作為未來需要自動化時的可選方式。

## 歷程重點

- 討論後決定採用精簡結構，避免一開始加入過多管理文件。
- 討論後決定使用 `AGENT_PROGRESS.md` 記錄目前進度、重點討論、待確認事項與下一步。
- 討論後決定資料夾說明寫在 `AGENTS.md` 與 `README.md`，不在每個資料夾各放說明檔。
- 討論後決定修改 `.env` 時必須同步更新 `.env.example`。
- 參考 `poly-rewards4` 規則後，決定補強 `.env.example` 說明來源、非敏感預設值一致與任務收尾檢查規則。
- 討論後決定加入 `.gitkeep`，避免空資料夾在 Git 中消失。
- 討論後決定補上 Git 上傳流程，採用人工建立 GitHub repo 加本機 Git 指令 push 的保守流程。
- 討論後決定開始將目前專案上傳到 Git；已完成本機 Git 初始化、安全檢查與第一個 commit，接著等待 GitHub repo URL 連接遠端。

## 待確認事項

- 無。

## 下一步

- 依實際專案需求新增原始碼、設定或文件。
- 若新增套件，必須遵守 `AGENTS.md` 的套件版本安全規則。
- 新專案第一次上傳 GitHub 時，參考 `docs/git-workflow.md`。
- 若本次上傳流程中尚未設定遠端，請使用者在 GitHub 網頁建立 repo 後提供 repo URL，再執行 `git remote add origin <repo URL>` 與 `git push -u origin main`。
