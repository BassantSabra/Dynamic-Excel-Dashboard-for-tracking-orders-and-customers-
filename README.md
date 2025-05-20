# 📊 Dynamic Excel Dashboard - README

This project guides you through building a **dynamic Excel dashboard** that connects customer and order data using formulas, PivotTables, and VBA. The dashboard enables filtering, auto-calculation, and visual insights for better decision-making.

![Screenshot 2025-05-20 193524](https://github.com/user-attachments/assets/f252a600-da3f-4fe6-95a7-57c03608dcd0)

# ✅ Step 1: Data Cleaning

Clean the **Customer** and **Order** sheets using Excel functions to standardize formats:

- `PROPER()` → Capitalizes the first letter of the customer name
- `UPPER()`
- `CHOOSE()` → = CHOOSE(F2,"Speed Express","National Package","Inland Shipping")
- `TEXT()` → =TEXT(C2,"mmm") 
---
# 🔍 Step 2: Retrieve Customer Data with Formulas

Use the following functions to auto-fill customer info based on selection:

- IF(VLOOKUP($B$3, customer_info, 11, FALSE)=0, "--", VLOOKUP($B$3, customer_info, 11, FALSE))
- INDEX(customer_info[Address], MATCH($B$3, customer_info[Company Name], 0))

 # 🛠 Step 3: VBA Macro for Advanced Filter 

![Screenshot 2025-05-20 195310](https://github.com/user-attachments/assets/b0ff2a98-4014-45ba-aaf7-81e6e7e09782)


# 📈 Step 4: Calculate KPIs with SUBTOTAL()

-  SUBTOTAL(103, Order_Table[Order ID])      'Count
-  SUBTOTAL(101, Order_Table[Order Amount])  'Average

# 📊 Step 5: Create Pivot Table and Pivot Chart

# 🔄 Step 6: VBA Code to Link PivotChart with Dashboard

![Screenshot 2025-05-20 195218](https://github.com/user-attachments/assets/803f2fd0-cbbe-4628-8077-7f7756abe69f)


