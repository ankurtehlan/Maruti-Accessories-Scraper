# 🧾 Maruti Suzuki Genuine Accessories Scraper

This Python script scrapes part information and images from the [Maruti Suzuki Genuine Accessories - Grand Vitara](https://www.marutisuzuki.com/genuine-accessories/grand-vitara-accessories) website and saves the data into an Excel file with embedded product images.

---

## 📦 Features

- Scrapes:
  - ✅ Part Number
  - ✅ Part Name
  - ✅ MRP (Price)
  - ✅ Product Image
- Downloads and saves product images locally
- Embeds images into an Excel spreadsheet
- Supports pagination (up to 33 pages)
- Skips and logs missing or broken data without stopping

---

## 🛠 Requirements

**Python 3.7+** and **Google Chrome**

Install dependencies with:

```bash
pip install requests pandas beautifulsoup4 openpyxl playwright
playwright install
