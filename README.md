# hemonc-mm — 血液腫瘤科 M&M 工具

超長住院／Morbidity／Mortality 討論會的科務工具。**本 repo 只有程式與說明，沒有任何病人資料**；資料本體在院內 SharePoint 站台（清單：案例、對照表、根因分類、正式紀錄、改善行動、會議）。

| 頁面 | 用途 |
|---|---|
| `index.html` | 入口 |
| `mm_convert.html` | Word 討論會清單 → SharePoint 匯入檔（純前端，零外部套件，檔案不離開瀏覽器） |
| `mm_dashboard.html` | 載入 SharePoint 匯出 CSV：待討論案例、投影模式、根因 Pareto、月趨勢 |
| `handbook.html` | 研究醫師操作手冊 |

GitHub Pages：https://jnynlin.github.io/hemonc-mm/

## 相關（不在本 repo）
- SharePoint 清單建置腳本 `create_lists.js`、欄位規格 `List欄位規格.xlsx`、Python 版轉換器 `mm_import.py`、清單匯入腳本 `import_items.js`：院內保管。
- 轉換邏輯以 `mm_import.py` 為準；`mm_convert.html` 為逐函式移植，53 例回歸比對 0 差異（2026-11-14）。
- `theme.js`：主題切換（跟隨系統／深／淺），四頁共用；artifact 版不載入。
