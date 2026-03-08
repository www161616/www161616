# 央廚單機版 ERP 系統

## 專案概述
- 純 HTML/CSS/JS 單檔架構，無框架
- index.html 為主殼（sidebar + iframe 載入子頁面）
- 主題色：石板灰護眼主題（slate gray），與龍潭相同
- 列印抬頭：丸十水產股份有限公司

## Supabase
- URL: `https://xuykenpvonfqwsgbuvvj.supabase.co`
- REST: 加上 `/rest/v1` 才是 API endpoint
- KEY: JWT token（eyJhbG... 開頭）
- 變數名稱：`SUPABASE_URL`, `SUPABASE_KEY`, `H`
- RLS：anon 全開

## DB 欄位注意
- inventory 表：`current_stock`（龍潭用 qty）、`updated_at`（龍潭用 last_updated）
- inventory_log 表：`type`（龍潭用 log_type）、`balance`（龍潭用 after_qty）
- 注意表名也不同：央廚 `inventory_log` vs 龍潭 `inventory_logs`

## localStorage keys（ck_ prefix）
- `ck_calendarEvents` — 日曆記事
- `ck_calendarTodos` — 待辦事項
- `ck_paymentReminders` — 收款提醒
- 不要用無 prefix 的 key，會跟龍潭版衝突

## 與龍潭的差異摘要
| 項目 | 央廚 | 龍潭 |
|---|---|---|
| SB 變數 | SUPABASE_URL / SUPABASE_KEY / H | SB_URL / SB_KEY / HEADERS |
| inventory 庫存量 | current_stock | qty |
| inventory 更新時間 | updated_at | last_updated |
| 異動 log 表 | inventory_log | inventory_logs |
| log 類型欄位 | type | log_type |
| log 餘額欄位 | balance | after_qty |
| localStorage prefix | ck_ | lt_ |
