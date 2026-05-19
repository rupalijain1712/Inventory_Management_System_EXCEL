# 📦 Inventory & Supply Chain Management System with Reorder Alerts

An enterprise-grade Excel-based solution for real-time warehouse inventory tracking, supply chain movement monitoring, and automated reorder alert generation — powered by advanced Excel formulas and VBA Macros.

> ⚡ Fully Offline &nbsp;|&nbsp; 🔒 No External Tools Required &nbsp;|&nbsp; 📊 Dashboard-Ready &nbsp;|&nbsp; 🤖 Macro-Automated

---

## ✨ Features

### 🤖 Macro-Powered Automation
- Auto-fills **Product Name**, **Supplier**, and **Warehouse Location** the instant a Product ID is typed
- Eliminates manual data entry errors with programmatic lookups
- Runs seamlessly in the background without user intervention

### 🚨 Intelligent Reorder Alerts
- Automatically flags products that have fallen **below minimum stock thresholds**
- Uses `IF` and `SUMIF` logic to compute live stock levels in real time
- Displays clear **Yes / No** reorder status per product

### 📊 Interactive Dashboard
- Visual summary of inventory health across all categories
- At-a-glance KPIs for stock levels, turnover rates, and order status
- Built-in Pivot Table for dynamic slicing and filtering

### 🧹 Structured Data Cleaning Pipeline
- Dedicated **Cleaning Tasks** sheet documenting every transformation applied
- Handles negatives (`ABS()`), duplicates, blanks, spaces (`TRIM()`), and inconsistent casing (`PROPER()`)
- Raw Data preserved separately from cleaned output

### 📦 End-to-End Supply Chain Tracking
- Separate **Stock In** and **Stock Out** sheets for granular movement logging
- Products master sheet with full supplier and warehouse mapping
- Summary Analysis sheet with aggregated metrics

---

## 🗂️ Workbook Structure

| Sheet | Purpose |
|---|---|
| **Welcome Page** | Project title and entry point |
| **Content** | Table of contents and navigation guide |
| **Instructions** | Objective, usage guide, and function explanations |
| **Cleaning Tasks** | Data cleaning log and techniques applied |
| **Raw Data** | Original, unmodified source data |
| **Clean Data** | Processed and validated dataset |
| **Products** | Product master list with supplier and location info |
| **Stock In** | Incoming inventory / delivery records |
| **Stock Out** | Outgoing inventory / sales records |
| **Reorder Alerts** | Live alert table for below-threshold products |
| **Summary Analysis** | Aggregated KPIs and category-level insights |
| **Pivot Table** | Dynamic pivot for flexible data exploration |
| **Dashboard** | Visual overview of the entire inventory system |

---

## 🧠 How It Works

```
Product ID Entered
        ↓
VBA Macro Triggers
        ↓
Auto-fills Name, Supplier, Warehouse
        ↓
SUMIF Calculates Live Stock
        ↓
IF Function Checks Reorder Threshold
        ↓
Alert Flagged → Dashboard Updates
```

---

## ⚙️ Functions Used

| Function | Purpose |
|---|---|
| `IF` | Tests whether current stock is below minimum threshold; returns **Yes** or **No** reorder status |
| `SUMIF` | Calculates live stock by summing all incoming deliveries and subtracting outgoing sales per Product ID |
| `COUNTIF` | Counts how many unique products belong to each category |
| `MACRO (VBA)` | Auto-populates Product Name, Supplier, and Warehouse Location on Product ID entry |
| `ABS()` | Corrects negative stock quantity values in raw data |
| `TRIM()` | Removes extra spaces that caused VLOOKUP errors |
| `PROPER()` | Standardizes inconsistent text casing across columns |

---

## 🧹 Data Cleaning Process

The raw dataset underwent the following cleaning steps before analysis:

| Technique | Problem Addressed | Method |
|---|---|---|
| Remove Duplicates | Repeated product records | Data Ribbon → Remove Duplicates |
| Fix Negative Values | Negatives in `Stock_Quantity` column | `ABS()` function |
| Handle Missing Values | Blanks causing formula errors | Go To Special Feature |
| Trim Extra Spaces | Spaces breaking `VLOOKUP()` | `TRIM()` function |
| Standardize Text Case | Inconsistent casing across columns | `PROPER()` function |
| Fix Date Formats | Inconsistent date entries | Number Format → Date |
| Remove Irrelevant Data | Columns/rows not needed for analysis | Manual removal |

---

## 📋 Dataset Overview

- **Product Categories:** Fruits & Vegetables, Dairy, Grains & Pulses, Bakery, Beverages, Seafood, Oils & Fats
- **Fields Tracked:** Product ID, Supplier ID, Warehouse Location, Stock Quantity, Reorder Level, Reorder Quantity, Unit Price, Sales Volume, Inventory Turnover Rate, Dates (Received / Last Order / Expiration)
- **Stock Statuses:** Active, Backordered, Discontinued

---

## 🚀 Getting Started

**Requirements:**
- Microsoft Excel (2016 or later recommended)
- Macros must be **enabled** for auto-fill functionality to work

**Steps:**
1. Download the `.xlsm` file
2. Open in Microsoft Excel
3. When prompted, click **"Enable Macros"**
4. Navigate using the **Content** sheet
5. Enter or update Product IDs in the **Stock In** / **Stock Out** sheets — the macro will auto-fill related fields
6. Check the **Reorder Alerts** sheet for flagged items
7. View the **Dashboard** for a live summary

> ⚠️ If macros are disabled, the auto-fill feature will not work. Manual entry will be required for supplier and location fields.

---

## 🎯 Use Cases

- Warehouse inventory management
- Supply chain monitoring and reporting
- Procurement and reorder planning
- Business intelligence for FMCG / retail operations
- Academic projects on inventory systems and Excel modelling

---

## 🔮 Future Improvements

- Power BI integration for advanced visual analytics
- Automated email alerts for reorder-triggered items
- Multi-warehouse support with location-based filtering
- Barcode/QR scanner input compatibility
- Expiration date proximity alerts
- Supplier performance tracking sheet

---

## ⭐ Contributing

Contributions are welcome!

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Open a pull request
