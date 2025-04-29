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
````md
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

