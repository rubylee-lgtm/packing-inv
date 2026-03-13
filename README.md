# Packing List / Commercial Invoice 產生工具

**版號：v1.1.0**  
**最終版本說明文件**

---

## 一、概述

本工具為**純前端**網頁應用，不需安裝後端或伺服器。人員上傳兩份 **CSV** 檔案（訂單查詢結果、產品目錄）後，於畫面上預覽訂單內容，再點擊按鈕即可在瀏覽器內產生並下載 **Excel (.xlsx)** 格式的：

- **Packing List**（裝箱單）
- **Commercial Invoice**（商業發票）

- **匯入**：僅接受 **CSV** 檔案（Catalog CSV + 訂單 CSV）  
- **產出**：**Excel (.xlsx)** 檔案（依模式輸出 `Packing_List_Output.xlsx` 或 `Commercial_Invoice.xlsx`）

---

## 二、使用方式

1. 以瀏覽器開啟專案內的 `templates/index.html`（雙擊或拖入瀏覽器即可）。
2. 於頁面上方下拉選單選擇模式：`Packing List` 或 `Commercial Invoice`。
3. 上傳「產品目錄 (Catalog)」CSV。
4. 上傳「訂單查詢結果」CSV，確認畫面上的預覽（最多顯示前 100 筆）。
5. 點擊「轉換並下載」，瀏覽器會自動下載對應模式的 `.xlsx` 檔案。

無需安裝 Python、Node 或任何伺服器，僅需可連網的瀏覽器（會載入 Bootstrap、PapaParse、SheetJS 的 CDN）。

---

## 三、專案結構

```
packing/
├── README.md                 # 本說明文件（v1.1.0）
└── templates/
    └── index.html            # 單一檔案：HTML + CSS + JavaScript
```

---

## 四、CSV 檔案規格

### 4.1 訂單查詢結果 CSV

- **必要欄位**（欄位名稱區分大小寫，以實際 CSV 表頭為準）：
  - `retailer_erp_customer_name`：買方名稱（對應 Excel 表頭 Buyer）
  - `retailer_address`：買方地址
  - `retailer_contact_phone`：買方電話
  - `retailer_contact_name`：買方聯絡人
  - `sku` 或 `Barcode`：品項條碼 / SKU（用於對照 Catalog 與計算毛重）
  - `quantity`：數量（PCS）
  - `pcs_per_carton`：每箱 PCS（若無 `carton` 時用來計算箱數）
  - `carton`：箱數（可選，若無則由 quantity / pcs_per_carton 計算）
  - `effective_date`：有效日期 / Best Before Date（建議 YYYY-MM-DD）
  - `measurement`：尺寸/Measurement
  - `Net Weight`：單件淨重（KG）
  - `quoted_price_currency`：商品報價（Invoice 模式使用）
- **DESCRIPTION（品名）**：產出時以 **Catalog 的 Product Name** 為單一來源；若 Catalog 無該 SKU，才 fallback 訂單的 `description` / `name`。

### 4.2 產品目錄 Catalog CSV

- 第一列可為標題列（例如「Juan Yuan Limited Product Catalog」），程式會自動掃描**下一列**起尋找表頭。
- **必要表頭**（表頭列中需包含下列關鍵字，大小寫不拘）：
  - **Barcode**：條碼 / SKU（與訂單的 sku 對應）
  - **G.W. (KGS)** 或 **G.W**：每箱毛重（KG）（Packing List 的 Gross Weight/Carton 來源）
  - **Product Name**：產品名稱（作為 Packing List / Invoice 的 DESCRIPTION 唯一來源）
- （選擇性）**Unit Price** / **Price**：若存在，會做為 Invoice 價格的優先來源
- 若表頭與資料列之間還有額外列，只要有一列同時出現 Barcode 與 G.W，該列會被視為表頭，其後列為資料。

---

## 五、產出 Excel 版面概要

### 5.1 表頭區（共通結構）

- 左側固定 Seller 區，右側為動態 Buyer 區，皆使用同一套結構，以空字串佔位達成左右並列：  
  - `Seller` 公司名（Packing：Juan Yuan Co., Ltd.；Invoice：Juan Yuan Limited）  
  - `Address : 6F., No. 19-9, Sanchong Rd., Nangang Dist., Taipei City 115601, Taiwan (R.O.C.)`  
  - `Tel :` 分別為 `886 2 77515350 #309`（Packing）或 `#216`（Invoice）  
  - `ATTN` / `Contact`：Invoice 為 `ATTN: Mary Wang`，Packing 為 `Contact:`（左），右側為訂單的聯絡人  
  - `Terms of Delivery: EXW Taoyuan Taiwan`（左右並列，Invoice 右側多一行 `Terms of Payment: 100% Prepaid T/T`）  


### 5.2 Packing List 明細區

- **雙層表頭**  
  - 第一列：`Packing No., Barcode, DESCRIPTION, Pcs per Carton, QUANTITY, Carton, Best Before Date, Measurement, Net Weight, Total Net Weight, Gross Weight/Carton, Total Gross Weight`  
  - 第二列：單位（PCS、(cm)、(KG) 等）  
- **明細邏輯**  
  - **DESCRIPTION**：Catalog 的 Product Name 為單一來源；Catalog 無資料時才 fallback 訂單的 description/name。  
  - **Best Before Date**：以 `effective_date` 為主，正規化為 `YYYY-MM-DD` 後尾端加一個空白字元，避免被 Excel 轉成日期序列號。  
  - **Carton**：若訂單未提供或為 0，則以 `quantity / pcs_per_carton` 計算。  
  - **Total Net Weight**：`Net Weight` × `QUANTITY`。  
  - **Gross Weight/Carton**：來自 Catalog 的 G.W. (KGS)。  
  - **Total Gross Weight**：`Carton` × `Gross Weight/Carton`。

### 5.3 Commercial Invoice 明細區

- **雙層表頭**  
  - 第一列：`No., Barcode, 品名, 箱入數, BBD, 商品報價, 下單數量, 金額`  
  - 第二列：`Description, Pcs / Carton, Days, Unit Price (USD), Order Quantity, Amount`  
- **明細邏輯**  
  - **No.**：由 1 開始的流水號。  
  - **Barcode**：訂單 `sku`（顯示原始值，內部比對會清洗為純英數）。  
  - **品名 / Description**：Catalog 的 Product Name；Catalog 無資料才 fallback 訂單的 description/name。  
  - **箱入數**：訂單 `pcs_per_carton`。  
  - **BBD / Days**：訂單 `bbd` 或 `effective_date`。  
  - **Unit Price (USD)**：訂單 `quoted_price_currency`，若 Catalog 有對應單價欄位則優先使用 Catalog。  
  - **Order Quantity**：訂單 `quantity`。  
  - **Amount**：`Unit Price × Order Quantity`。  
  - 明細結束後會額外加上一列 `TOTAL`，於 `Order Quantity` 與 `Amount` 欄位做加總。

---

## 六、技術與邏輯摘要

| 項目 | 說明 |
|------|------|
| 介面 | HTML5 + Bootstrap 5 (CDN) |
| CSV 解析 | PapaParse，訂單用 `header: true`，Catalog 用 `header: false` 掃表頭 |
| Excel 產出 | SheetJS (xlsx)，`XLSX.utils.book_new()` + `aoa_to_sheet` 組二維陣列後寫出 .xlsx |
| 執行順序 | 先解析 Catalog 建立 `catalogMap`（Barcode → `{ gw, name, unitPrice }`），再解析訂單，再依模式產出 Excel，確保毛重、品名與單價正確 |
| Barcode 對應 | Catalog 與訂單的 Barcode/SKU 皆會去除非英數字後再比對，以利格式差異時仍能對到 |
| 品名單一來源 | DESCRIPTION 優先使用 Catalog 的 Product Name，無則使用訂單的 description / name |

---

## 七、版號說明

- **v1.0.0**：初始穩定版  
  - 僅支援 Packing List，匯入 CSV、產出 Excel。  
  - Catalog 支援表頭列在第二列（第一列可為標題）。  
  - 買方資訊、明細欄位、毛重與品名對應如上所述。  

- **v1.1.0**（本版）：Packing / Invoice 雙模式版  
  - 新增 **Commercial Invoice** 模式（雙語雙層表頭）。  
  - 上半部 Seller/Buyer 資訊改為共用結構，依模式切換 Seller 名稱、分機與付款條件。  
  - 發票金額由 `quoted_price_currency × quantity` 自動計算並提供 TOTAL 加總。  
  - 維持 Catalog 為 Product Name 與毛重/單價的單一來源，Barcode 以純英數清洗後比對。

---

*文件依專案內 `templates/index.html` 最終版本撰寫。*
