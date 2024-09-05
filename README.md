# Analyzing Chocolate Sales with Excel

## Project Overview

In this project, I conducted a comprehensive analysis of chocolate sales data for a fictional chocolate company. The primary objective is to uncover key trends and insights that will help the company optimize its product offerings, improve customer satisfaction, and drive higher revenue. The project includes data cleaning, transformation, visualization, and the creation of an interactive dashboard in Excel.

## Business Problem

The chocolate industry is highly competitive, and understanding customer preferences and sales patterns is crucial for success. The company faces challenges such as:
- Identifying the most popular chocolate types and formats.
- Understanding how seasonal trends affect sales.
- Analyzing the factors that drive higher sales performance.

## Data Sources

The analysis is based on three core datasets:
1. **Orders Sheet**: Contains order-level data such as Order ID, Order Date, Customer ID, Product ID, and Quantity purchased.
2. **Products Sheet**: Contains product-level data such as Product ID, Chocolate Type, Form Type (Bars, Chips, etc.), Unit Price, and Profit.
3. **Customers Sheet**: Contains customer-level data such as Customer ID, Name, Email, Phone Number, Address, and Loyalty Card status.

## Steps Undertaken

1. **Data Cleaning and Preprocessing**
    - Integrated customer and product details into the orders sheet using XLOOKUP and INDEX MATCH functions.
    - Reformatted date, size, and price columns for consistency.
    - Converted raw data into structured Excel tables for better analysis.

2. **Data Analysis**
    - Created several Pivot Tables to explore key metrics:
        - **Total Sales Over Time**: Analyzed sales trends by month and chocolate type.
        - **Sales by Country**: Compared sales across different countries.
        - **Top 5 Customers**: Identified the top customers based on total sales.
    - Added slicers and timelines for interactive data filtering by form type, size, and loyalty card status.

3. **Data Visualization and Dashboard**
    - Developed an interactive dashboard combining multiple Pivot Charts and Slicers:
        - **Total Sales Over Time**: A line chart displaying sales trends by chocolate type over time.
        - **Sales by Country**: A bar chart highlighting the highest-grossing countries.
        - **Top 5 Customers**: A bar chart showing the top customers by sales.
    - Integrated slicers and a timeline to enable dynamic filtering for deeper data exploration.

## Tools Used

- **Microsoft Excel**: Primary tool for data preprocessing, analysis, and visualization.
- **Excel Functions**: XLOOKUP, INDEX MATCH, IF, Pivot Tables, Pivot Charts.
- **Interactive Elements**: Slicers, Timeline, Dashboard.

## Results and Insights

- **Dark Chocolate** consistently performs well, with sales spikes during specific months.
- The **United States** is the top-performing country, contributing the highest sales overall.
- The **Top 5 Customers** are frequent buyers, suggesting opportunities for targeted loyalty rewards.
  
## Conclusion

By analyzing the chocolate sales data, this project has uncovered valuable insights to help the company refine its sales strategy. The interactive dashboard provides an accessible and dynamic tool for exploring sales data across various dimensions, offering stakeholders the ability to make data-driven decisions.

## Future Enhancements

In future iterations, the project could be enhanced by:
- Incorporating predictive analysis for seasonal trends.
- Automating dashboard updates through VBA or macros.
- Refining the dashboard’s aesthetics for a more professional presentation.

## How to Use the Project

1. Download the provided Excel files (`Analyzing Chocolate Sales with Excel.xlsx` and `ChocolteOrdersData.xlsx`).
2. Open the main Excel file to view the integrated data and dashboard.
3. Use the slicers and timeline in the dashboard to filter data dynamically and explore key sales trends.

