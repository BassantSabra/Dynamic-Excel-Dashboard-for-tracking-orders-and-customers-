# 📊 Dynamic Excel Dashboard - README

This project guides you through building a **dynamic Excel dashboard** that connects customer and order data using formulas, PivotTables, and VBA. The dashboard enables filtering, auto-calculation, and visual insights for better decision-making.

---

# ✅ Step 1: Data Cleaning

Clean the **Customer** and **Order** sheets using Excel functions to standardize formats:

- `PROPER()` → Capitalizes the first letter of customer name
- `UPPER()`
- `CHOOSE()` → = CHOOSE(F2,"Speed Express","National Package","Inland Shipping")
- `TEXT()` → =TEXT(C2,"mmm") 
---
# 🔍 Step 2: Retrieve Customer Data with Formulas

Use the following functions to auto-fill customer info based on selection:

```excel
- IF(VLOOKUP($B$3, customer_info, 11, FALSE)=0, "--", VLOOKUP($B$3, customer_info, 11, FALSE))
- INDEX(customer_info[Address], MATCH($B$3, customer_info[Company Name], 0))

 # 🛠 Step 3: VBA Macro for Advanced Filter 






