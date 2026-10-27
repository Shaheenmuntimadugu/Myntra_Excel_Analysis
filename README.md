# Myntra_Excel_Analysis
End-to-end sales analysis of 5000 Myntra fashion products.

**Tools:** Advanced Excel 

## 📊 Dataset
- **File:** Myntra_Fashion_Sales_5000.xlsx
- **Rows:** 5000 transactions
- **Brands:** 10 (Zara, Nike, Adidas, Levis, H&M, Biba, Puma, Fabindia, Manyavar, Roadster)
- **Categories:** 8 (Jeans, Kurtas, Sarees, Shirts, Leggings, Shoes, Dresses, T-Shirts)
- **States:** 20+ Indian states

## 📑 7-Sheet Excel Model
1. **Raw Data:** Original 5000 rows
2. **Cleaned Sheet:** Data cleaning, duplicates removed
3. **XLOOKUP / VLOOKUP:** Practice - Syntax =XLOOKUP(what, search_column, return_column)
4. **PivotTable Dashboard:** Sales by Brand, Category, Gender
5. **Profit Calculation:** Discount_Amount, Cost_Price (60% of Original), Profit = Selling - Cost
6. **Profit Dashboard:** Brand-wise profit - Zara Top ₹5,02,267
7. **Formula Practice:** TEXT (Month-Year), YEAR, QUARTER, SUMIF, SUMIFS, IF

## 💰 Key Formulas
Discount_Amount = Original_Price * Discount_Pct
Cost_Price = Original_Price * 60%
Profit = (Selling_Price - Cost_Price) * Quantity
TEXT(Order_Date, "MMM-YYYY") -> Month-Year
SUMIF(Brand, "Zara", Profit) -> Brand-wise Profit

## 📈 Key Insights
- **Total Profit:** ₹41,15,425
- **Top Brand by Profit:** Zara ₹5,02,267
- **Discount Impact:** 50% discount = negative profit found

## 🎯 Skills
Advanced Excel: XLOOKUP, VLOOKUP, TEXT, YEAR, QUARTER, SUMIF, SUMIFS, IF, PivotTables, Slicers

---
**Author:** Shaheen M | github.com/Shaheenmuntimadugu
