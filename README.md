# i-just-analysed-coffee-

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
