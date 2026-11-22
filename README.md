# SRCD 博士研究｜閱讀理解 × 執行功能

此儲存庫是文獻證據網絡網站的正式來源，網站部署於儲存庫根網址：

<https://lin900212.github.io/srcd-research/>

## 維護流程

1. 先更新專案主目錄的 `SRCD_文獻總表.xlsx`，完成去重、文獻判讀及「三研究導覽」分頁更新。
2. 在專案根目錄執行 `python3 sites/github-srcd-research/scripts/build_srcd_site.py`。建置器讀取「三研究導覽」的書目、關係標籤、研究結果摘要及列號，更新本儲存庫 `public/index.html`，並同步更新本機圖表鏡像。
3. 新文獻的 DOI 加在 `site-src/doi-map.json`，以 `研究編號:文獻ID` 為鍵。工作表列號會依本機活頁簿自動讀取，摘要網址定位到相應 D 欄。
4. 檢查差異後，推送網站來源至 `main`。Pages workflow 只發布 `public/`；不發布 builder、DOI 對照檔或其他維護來源。
5. 本機及雲端工作簿、PDF、WoS 原始匯出與繁中摘要不放進儲存庫。網頁只保留可公開的書目、證據關係和研究結果摘要；Google 試算表仍限制存取。

## GitHub Pages 設定

此儲存庫目前為公開，GitHub Pages 使用自訂 GitHub Actions workflow 部署。workflow 將 `public/` 作為唯一網站產物；推送至 `main` 後自動部署至 `https://lin900212.github.io/srcd-research/`。
