# Espressheets-🍵

# Coffee Orders Data Analysis Using Excel ☕

In this project, I analyzed coffee order data using Excel.  
The goal was to extract insights to understand sales performance, identify popular products, and suggest improvements.

Each step is clearly documented as part of my data analysis learning journey, and to showcase the project on GitHub and LinkedIn.

---

## 📌 Acknowledgment

The data and initial project idea were inspired by [Mo Chen](https://youtu.be/m13o5aqeCbM?si=s9KFnvaMwjdmHFr7).  
Credit goes to the original creator for providing the dataset and concept.

---

## 🛠️ Step 1: Importing Customer Information

I started by linking the **Customer Name** from the `customers` table to the `orders` table.

Initially, I planned to use the `XLOOKUP` function — but I realized my Excel version doesn't support it.  
So instead, I used a combination of `INDEX` and `MATCH`:

```excel
=INDEX(customers!B:B, MATCH(C2, customers!A:A, 0))

````
This formula pulls the **Customer Name** based on the `Customer ID`.

---

## 📧 Step 2: Linking Emails and Other Info

I repeated the same approach to retrieve other customer details, such as **Email** and **Country**.

For example, to get the **Email**:

```excel
=INDEX(customers!$C:$C, MATCH($C2, customers!$A:$A, 0))
```

While working with email data, I noticed several missing or zero values.

To handle that, I added an `IF` condition to display a blank cell instead of a zero:

```excel
=IF(INDEX(customers!$C:$C, MATCH($C2, customers!$A:$A, 0)) = 0, "", INDEX(customers!$C:$C, MATCH($C2, customers!$A:$A, 0)))
```

---

## ☕ Step 3: Importing Product Info

In this step, I filled in the product-related columns:  
**Coffee Type**, **Roast Type**, **Size**, and **Unit Price**.

This part was a bit more challenging because it involved working with both **row and column matches**.

Here’s the formula I used to retrieve data dynamically based on the column headers:

```excel
=INDEX(products!$A$1:$G$49, MATCH(orders!$D2, products!$A$1:$A$49, 0), MATCH(orders!I$1, products!$A$1:$G$1, 0))
```

This formula checks the `Product ID` in column D and pulls the corresponding product detail based on the column name in row 1 (like Coffee Type, Roast Type, etc.).

---

### 💰 Sales Calculation

The last column in the `orders` table is **Sales**, which was simple to calculate:

```excel
=L2*E2
```

Where:
- `L2` is the Unit Price
- `E2` is the Quantity

---

## 🎨 Step 4: Data Formatting

Now that all the data is in place, I formatted the columns to make them cleaner and more insightful.

- **Date column**: Changed format to `dd/mmm/yyyy` so months appear as words (e.g., Jan, Feb, etc.). This will help in the analysis later.
  
- **Size column**: Since it represents kilograms, I used a custom number format to display values like `0.5 kg`:
  ```text
  0.0 "kg"
  ```

- **Price and Sales columns**: I formatted them as currency in dollars. No formula needed — just selected the `$` symbol from currency formatting.

---

### 📊 Step 5: Data Analysis and Visualization

This was my favorite part — turning raw data into meaningful insights!

#### ✅ Convert to Table

First, I converted the cleaned dataset into an official **Excel Table**.  
This makes it easier to reference and work with dynamic data ranges later on.

![Screenshot 2025-05-01 142801](https://github.com/user-attachments/assets/1b6611c5-dc5c-457b-8a82-3abab76f7b5a)

#### 📈 Create a Pivot Table

Then, I created a **Pivot Table** based on the main data table.

Here's how I structured it:
- **Rows**: Order Date  
- **Columns**: Coffee Type  
- **Values**: Sales

This setup helped me see how sales changed over time for each type of coffee.

#### 📉 Add a Chart

Next, I added a **Line Chart** to visualize the sales trend.  
Line charts are perfect for historical time-based data like this.
![Screenshot 2025-05-01 142750](https://github.com/user-attachments/assets/8be4698d-27c5-45bd-b05f-7c1b324a1316)

#### 🕒 Add a Timeline Filter

To make it more interactive, I added a **Timeline** for the **Order Date**.  
Now I can filter the data by month or specific time periods easily.
![Screenshot 2025-05-01 144235](https://github.com/user-attachments/assets/8c1907f0-ac52-4367-a8bf-79744e60c20b)


---

## 🎯 Step 6: Add Loyalty Card Column

Added a new column called **Loyalty Card** to the dataset, using this formula:

```excel
=INDEX(customers!$I$2:$I$1001, MATCH([@[Customer ID]], customers!$A$2:$A$1001, 0))
```

Refreshed the Pivot Table to include the new column.

---

## 🧩 Step 7: Add Slicers

To improve filtering and interactivity, I added **Slicers** to the Pivot Table for:

* Country
* Coffee Type
* Roast Type
* Size
* Loyalty Card

![Screenshot 2025-05-02 020305](https://github.com/user-attachments/assets/66d3e6a8-3907-435b-b0b4-69e4761eb7a9)


If a loyalty doesn’t exist, make sure to **Refresh** the Pivot Table.

---

## 📊 Step 8: Create Additional Charts

* **Country Sales** → Pie Chart
![Screenshot 2025-05-02 020948](https://github.com/user-attachments/assets/ac4867bd-35ae-4df6-9691-a0718e379915)

  
* **Roast Type Sales** → Pie Chart
![Screenshot 2025-05-02 021738](https://github.com/user-attachments/assets/4b74e1aa-7b83-4b9b-b808-0e0840071494)

  
* **Top 10 Customers** → Bar Chart (sorted by Sales descending)
![Screenshot 2025-05-02 021730](https://github.com/user-attachments/assets/e0140dd2-3253-478a-ad34-966941cdabb9)

  
---

### Step 6: Adding KPIs🆙🆙

Before wrapping up the project, I wanted to highlight some key performance indicators (KPIs) to give a clearer overview of business performance. These metrics are displayed alongside the dashboard to make insights easier to understand at a glance.

* **Total Sales**: I used the `SUM` function to calculate the total value of all sales based on the "Sales" column.

* **Total Orders**: To count the number of orders, I used the `COUNTA` function on the "Order ID" column.

* **Average Order Value**: This KPI shows the average value of each individual order. I calculated it with the following formula:

  ```
  =SUM(orders[Sales])/COUNTA(orders[Order ID])
  ```
![Screenshot 2025-05-02 142259](https://github.com/user-attachments/assets/943d8ed0-2330-49d2-8cdf-be33c1ddf630)

These indicators give a quick summary of how the business is performing and add a more professional touch to the final dashboard.

---

## 🧙 Step 9: Build the Dashboard ✨

I created a new sheet for the **Dashboard**, then moved all visuals (charts & slicers) into it.


![Screenshot 2025-05-02 024417](https://github.com/user-attachments/assets/3550cef1-8659-4fe7-90cd-9b78174bf76c)




To make slicers affect all visuals, I used:

> Slicer → **Options** → **Report Connections**

Finally, I improved the design with consistent colors and clean formatting.

![Screenshot 2025-05-02 153231](https://github.com/user-attachments/assets/daab9f0d-0125-40bf-99e2-b35b849f2f1c)


---

## ✅ Final Thoughts

This project was a big step in my journey toward learning **Data Analysis** and **Financial Analysis**.
I really enjoyed the process — especially building the dashboard and seeing everything come together.

### 📚 Key Excel Skills Gained:

* `INDEX`, `MATCH`, `IF`, and cell referencing
* Working with dynamic ranges and tables
* Pivot Tables and Slicers
* Data cleaning and formatting
* Creating interactive dashboards
* Visualizing data with Line, Pie, and Bar Charts

> I hope you enjoyed reading about my project as much as I enjoyed building it!

---
