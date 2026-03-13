# Packing List 產生工具

**版號：v1.0.0**  
**最終版本說明文件**

---

## 一、概述

本工具為**純前端**網頁應用，不需安裝後端或伺服器。人員上傳兩份 **CSV** 檔案（訂單查詢結果、產品目錄）後，於畫面上預覽訂單內容，再點擊按鈕即可在瀏覽器內產生並下載 **Excel (.xlsx)** 格式的 Packing List。

- **匯入**：僅接受 **CSV** 檔案  
- **產出**：**Excel (.xlsx)** 檔案（檔名：`Packing_List_Output.xlsx`）

---

## 二、使用方式

1. 以瀏覽器開啟專案內的 `templates/index.html`（雙擊或拖入瀏覽器即可）。
2. **步驟 1**：上傳「訂單查詢結果」CSV。
3. **步驟 2**：於頁面上確認訂單預覽（最多顯示前 100 筆）。
4. **步驟 3**：上傳「產品目錄 (Catalog)」CSV。
5. **步驟 4**：點擊「轉換並下載 Packing List」，瀏覽器會自動下載產出的 `.xlsx` 檔案。

無需安裝 Python、Node 或任何伺服器，僅需可連網的瀏覽器（會載入 Bootstrap、PapaParse、SheetJS 的 CDN）。

---

## 三、專案結構

```
packing/
├── README.md                 # 本說明文件（v1.0.0）
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
- **DESCRIPTION（品名）**：產出時以 **Catalog 的 Product Name** 為單一來源；若 Catalog 無該 SKU，才 fallback 訂單的 `description` / `name`。

### 4.2 產品目錄 Catalog CSV

- 第一列可為標題列（例如「Juan Yuan Limited Product Catalog」），程式會自動掃描**下一列**起尋找表頭。
- **必要表頭**（表頭列中需包含下列關鍵字，大小寫不拘）：
  - **Barcode**：條碼 / SKU（與訂單的 sku 對應）
  - **G.W. (KGS)** 或 **G.W**：每箱毛重（KG）
  - **Product Name**：產品名稱（作為 Packing List 的 DESCRIPTION 唯一來源）
- 若表頭與資料列之間還有額外列，只要有一列同時出現 Barcode 與 G.W，該列會被視為表頭，其後列為資料。

---

## 五、產出 Excel 版面概要

- **表頭區**  
  - 第 1 列：`PACKING LIST`（獨立一行）  
  - 第 2 列：左側 `Date:` + 當日日期，右側 E/F 欄 `NO：` + 單號（JY + 日期 8 碼）  
  - 第 3～6 列：左側 Seller 固定資訊（Juan Yuan Co., Ltd.、地址、電話、Contact），右側 E 欄起 Buyer 動態資訊（來自訂單第一筆的 retailer_erp_customer_name、retailer_address、retailer_contact_phone、retailer_contact_name）  
  - 其後：Container、Temprature request、Packing、Payment conditions、Delivery term 等固定選項列  

- **明細區**  
  - 雙層表頭：第一列為欄位名（Packing No., Barcode, DESCRIPTION, Pcs per Carton, QUANTITY, Carton, Best Before Date, Measurement, Net Weight, Total Net Weight, Gross Weight/Carton, Total Gross Weight），第二列為單位（PCS, (cm), (KG) 等）。  
  - 明細資料依序填入；**DESCRIPTION** 以 Catalog 的 **Product Name** 為準；**Gross Weight/Carton**、**Total Gross Weight** 由 Catalog 毛重與箱數計算。  
  - Best Before Date 以字串形式寫入並在尾端加空白，避免被 Excel 轉成日期序列號。

---

## 六、技術與邏輯摘要

| 項目 | 說明 |
|------|------|
| 介面 | HTML5 + Bootstrap 5 (CDN) |
| CSV 解析 | PapaParse，訂單用 `header: true`，Catalog 用 `header: false` 掃表頭 |
| Excel 產出 | SheetJS (xlsx)，`XLSX.utils.book_new()` + `aoa_to_sheet` 組二維陣列後寫出 .xlsx |
| 執行順序 | 先解析 Catalog 建立 `catalogMap`（Barcode → `{ gw, name }`），再解析訂單，再產出 Excel，確保毛重與品名正確 |
| Barcode 對應 | Catalog 與訂單的 Barcode/SKU 皆會去除非英數字後再比對，以利格式差異時仍能對到 |
| 品名單一來源 | DESCRIPTION 優先使用 Catalog 的 Product Name，無則使用訂單的 description / name |

---

## 七、版號說明

- **v1.0.0**（本版）：最終穩定版  
  - 僅支援匯入 CSV、產出 Excel。  
  - Catalog 支援表頭列在第二列（第一列可為標題）。  
  - 買方資訊、明細欄位、毛重與品名對應如上所述。  

---

*文件依專案內 `templates/index.html` 最終版本撰寫。*
