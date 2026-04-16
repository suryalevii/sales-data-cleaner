"""Sales data cleaning, analysis, reporting, and export pipeline.

This script performs a full workflow on raw sales data:
1. Loads and profiles raw input data.
2. Cleans duplicates, missing values, and invalid records.
3. Validates data types and standardizes text fields.
4. Adds derived business columns and categories.
5. Runs grouped analyses for cities, products, and months.
6. Prints a human-readable management summary.
7. Saves detailed and aggregated Excel outputs.

Expected input file:
- raw_sales.xlsx

Generated output files:
- clean_sales_report.xlsx
- high_value_orders.xlsx
- city_summary.xlsx
- product_summary.xlsx
"""


"""Stage 1: Load source data and run an initial quality snapshot.

Includes schema preview, sample records, missing-value scan,
duplicate count, and quick categorical profiling.
"""
import pandas as pd

df=pd.read_excel("raw_sales.xlsx")
print("="*50)
print("RAW DATA LOADED")
rows=len(df)
columns=len(df.columns)
print(f"ROWS : {rows} || COLUMNS : {columns}")
print("="*50)
print()
print(df.head())
print()
print(df.tail(3))
print()
print(df.sample(3))
print()
print(df.info())
print()
print(df.dtypes)
print()
print("MISSING VALUES")
print(df.isnull().sum())
duplicate=df.duplicated().sum()
print("The number of duplicate rows found : ",duplicate)
print()
print(df.describe())
print()
print(df["City"].unique())
print()
print(df["Product"].unique())
print()
print(df["City"].value_counts())



"""Stage 2: Clean structurally invalid records.

Removes duplicate rows, drops records with missing product names,
fills missing prices with median, removes zero-quantity orders,
and resets the index for downstream consistency.
"""
print("="*50)
print("CLEANING STATE")
df.drop_duplicates(inplace=True)
rows_remaining_duplicates=len(df)
print("Rows remaining after removing the duplicates :",rows_remaining_duplicates)


print()


print(df.loc[df["Product"].isna()])
print()
df.dropna(subset=["Product"], inplace=True)
rows_remaining_products=len(df)
print("Rows remaining after removing the empty products : ",rows_remaining_products)


print()


print(df.loc[df["Price"].isna()])
print()
print("Missing Prices filled with median : ",df["Price"].median())
df.loc[df["Price"].isna(),"Price"]=df["Price"].median()

print()


print(df.loc[df["Quantity"]==0])
df.drop(df[df["Quantity"] == 0].index, inplace=True)
rows_remaining_quantity=len(df)
print("Row remaining after removing 0 quantity",rows_remaining_quantity)


print()


df.reset_index(drop=True,inplace=True)
print("Index reset is Done")
print()
print(df.isna().sum())
final_row=len(df)
print("="*50)
print("CLEANING COMPLETE")
print("FINAL NUMBER OF ROWS : ",final_row)
print("="*50)


"""Stage 3: Validate and standardize fields used for analysis.

Parses date values, drops invalid dates, derives calendar columns,
removes negative price/quantity records, and normalizes text labels.
"""
print("="*50)
print("VALIDATION STAGE")
print("="*50)
print()
print('The datatype of date : ',df["Date"].dtype)
print()
print("The unique date values before cleaning")
print(df["Date"].unique())
df["Date"]=pd.to_datetime(df["Date"],dayfirst=False,errors="coerce")
print()
print(df["Date"].isna().sum())
print()
df.drop(df.loc[df["Date"].isna()].index, inplace=True)
print(f"The date column cleaned. {len(df)} rows remain")


df["Month"]=df["Date"].dt.month_name()
df["Year"]=df["Date"].dt.year
df["Day"]=df["Date"].dt.day_name()


print(df.loc[df["Price"]<0].count())
print()
df.drop(df.loc[df["Price"]<0].index,inplace=True)
max_price=df["Price"].max()
min_price=df["Price"].min()
print(f"Price range: min={min_price}  max={max_price}")
print()


print(df.loc[df["Quantity"]<0].count())
df.drop(df.loc[df["Quantity"]<0].index,inplace=True)
max_quantity=df["Quantity"].max()
min_quantity=df["Quantity"].min()
print()
print(f"Quantity range:  Max={max_quantity}  Min={min_quantity}")
print()



df[["Customer", "Product", "City"]] = df[["Customer", "Product", "City"]].apply(
    lambda s: s.astype("string").str.strip()
)
df[["Customer","Product","City"]]=df[["Customer","Product","City"]].apply(
    lambda g:g.astype("string").str.title()
)
print(df["City"].unique())
print()


print("="*50)
print("VALIDATION COMPLETE")
print(f"Rows: {len(df)}  |  Columns:  {len(df.columns)}")
print("="*50)


"""Stage 4: Create business metrics and segmentation features.

Adds monetary totals, discount-adjusted amounts, and categorical
labels for pricing and quantity bands used in reporting.
"""

print("="*50)
print("ADDING CALCULATED COLUMNS")
print("="*50)
df["Total"]=df["Quantity"]*df["Price"]
print(df[["Product","Quantity","Price","Total"]].head())
def discount(total):
    """Return discount amount based on order total.

    Rules:
    - 10% discount for totals >= 1000
    - 5% discount for totals >= 500 and < 1000
    - No discount otherwise
    """
    if(total>=1000):
        return total*0.10
    elif(total>=500):
        return total*0.05
    else:
        return 0
df["Discount"]=df["Total"].apply(discount).round(2)
df["FinalAmount"]=(df["Total"]-df["Discount"]).round(2)

def price_category(price):
    """Classify a unit price into Budget, Standard, or Premium tiers."""
    if(price>=250):
        return "Premium"
    elif(price >=100):
        return "Standard"
    else:
        return "Budget"
df["PriceCategory"]=df["Price"].apply(price_category)

def quantity_category(quantity):
    """Classify order quantity into Small, Medium, or Bulk segments."""
    if(quantity>=15):
        return "Bulk"
    elif(quantity>=7):
        return "Medium"
    else:
        return "Small"
df["QuantityCategory"]=df["Quantity"].apply(quantity_category)

print("New columns added: Total, Discount, FinalAmount, PriceCategory, QuantityCategory")
total_before_discount=df["Total"].sum()
total_after_discount=df["FinalAmount"].sum()
total_discount=df["Discount"].sum()
print()
print(f"Total revenue (before discount): {total_before_discount}")
print(f"Total revenue (after discount): {total_after_discount}")
print(f"Total discount given : {total_discount}")

print()
print(df.head())
print()
print(df.columns.tolist())


"""Stage 5: Perform grouped analytics for decision support.

Generates ranked and segmented views by city, product, month,
and category, including top-customer summaries.
"""

print("="*50)
print("SORTING AND ANALYSIS")
print("="*50)
print()
df_sorted=df.sort_values(by='FinalAmount',ascending=False)
print(df_sorted[["OrderID",'Customer','Product','Quantity','Price','Total','FinalAmount']])
print()
print("CITY-WISE ANALYSIS")
print()
print(df.groupby('City')["FinalAmount"].sum())
print()
print(df.groupby("City")["Total"].mean())
print()
print(df.groupby("City")["OrderID"].count())
print()
print("PRODUCT WISE ANALYSIS")
print()
print(df.groupby("Product")["Quantity"].sum())
print()
print(df.groupby("Product")['FinalAmount'].sum())
print()
most_revenue_product=df.groupby("Product")['FinalAmount'].sum()
print("The product got the most revenue : ",most_revenue_product.idxmax())
print("The total amount revenued by the product is :",most_revenue_product.max())
most_sold_product=df.groupby("Product")["Quantity"].sum()
print("The product which is sold most is : ",most_sold_product.idxmax())
print()
print("The qunatity of the product sold is :",most_sold_product.max())
print()
print("MONTH-WISE ANALYSIS")
print()
print(df.groupby("Month")["FinalAmount"].sum())
print()
print(df.groupby("Month")["OrderID"].count())
print()
print("The Month with the Highest Revenue is :",df.groupby("Month")["FinalAmount"].sum().idxmax())
print()
print("CATEGORY-WISE ANALYSIS")
print()
print(df["PriceCategory"].value_counts())
print()
print(df["QuantityCategory"].value_counts())
print()
print(df.groupby("PriceCategory")["FinalAmount"].mean())
print()
print(df.groupby("QuantityCategory")["FinalAmount"].sum())
print()
print("The Top 3 customers are :")
top_customers = df.groupby("Customer")["FinalAmount"].sum().sort_values(ascending=False).head(3)
for i, (name, amount) in enumerate(top_customers.items(), 1):
    print(f"{i}. {name:<12} - Rs.{amount:.2f}")


"""Stage 6: Build and print an executive sales summary.

Outputs key KPIs such as order volume, customer/product breadth,
revenue metrics, top performers, and overall business status.
"""

print("="*50)
print("               SALES REPORT SUMMARY")
print("="*50)
print()
print("The total number of orders processed : ",df["OrderID"].count())
print()
unique_customers=df["Customer"].unique()
print("The total number of unique customers : ",len(unique_customers))
print()
unique_products=df["Product"].unique()
print("The total number of unique products : ",len(unique_products))
print()
unique_cities=df["City"].unique()
print("The number of unique cities : ",len(unique_cities))
print()
start_date = df["Date"].min().strftime("%Y-%m-%d")
end_date = df["Date"].max().strftime("%Y-%m-%d")
print(f"Date range : {start_date}   to   {end_date}")
print()
print("---REVENUE SUMMARY---")
print()
print("Gross Revenue (before discount):   Rs.",df["Total"].sum())
print()
print("Total Discount Given:              Rs.",df["Discount"].sum())
print()
print("Net Revenue (after discount):      Rs.",df["FinalAmount"].sum())
print()
print("Average Order Value:               Rs.",df["FinalAmount"].mean().round(2))
print()
print("Highest Single Order Value:        Rs.",df["FinalAmount"].max())
print()
print("Lowest Single Order Value:         Rs.",df["FinalAmount"].min())
print()
print("---BEST PERFORMERS---")
print()
top_revenue_product = df.groupby("Product")["FinalAmount"].sum()
print(f"Top Product by Revenue:    {top_revenue_product.idxmax():<15}    Rs.{top_revenue_product.max():.2f}")
top_quantity=df.groupby("Product")["Quantity"].sum()
print(f"Top Product by Quantity:   {top_quantity.idxmax():<15}      {top_quantity.max()} units")
print()
top_revenue_city=df.groupby("City")["FinalAmount"].sum()
print(f"Top City by Revenue:       {top_revenue_city.idxmax():<15}   Rs.{top_revenue_city.max():.2f}")
print()
top_revenue_month=df.groupby("Month")["FinalAmount"].sum()
print(f"Top Month by Revenue:      {top_revenue_month.idxmax():<15}      Rs.{top_revenue_month.max():.2f}")
final_top_customer=df.groupby("Customer")["FinalAmount"].sum()
print(f"Top Customer:              {final_top_customer.idxmax():<15}      Rs.{final_top_customer.max():.2f}")
print()
print("---CITY BREAKDOWN---")
city_breakdown_order=df.groupby("City")["OrderID"].count()
city_breakdown_revenue=df.groupby("City")["FinalAmount"].sum()
city_breakdown_avg=df.groupby("City")["FinalAmount"].mean()
print(f"{'City':<12}   {"Orders":<15}   {"Revenue":<12}   {"Avg Order":<12}")
for city in city_breakdown_order.index:
    print(f"{city:<12}   {city_breakdown_order[city]:<15}   {city_breakdown_revenue[city]:<12}   {city_breakdown_avg[city]:<12}")

discount_percentage=(df["Discount"].sum()/df["Total"].sum())/100
print()


net_revenue=df["FinalAmount"].sum()
if(net_revenue>=10000):
    print("Business Status: EXCELLENT")
elif(net_revenue>=5000):
    print("Business Status: GOOD")
else:
    print("Business Status: NEEDS IMPROVEMENT")
print()
print(f"Discount rate : {discount_percentage:.1f}%")
print()
print("="*50)
print("           REPORT GENERATED SUCCESSFULLY")
print("="*50)



"""Stage 7: Persist cleaned data products to Excel files.

Exports detailed transactions, filtered high-value orders,
and aggregate city/product summaries for external consumption.
"""
print("="*50)
print("SAVING OUTPUT FILES")
print("="*50)
final_report=df.sort_values(by="FinalAmount",ascending=False)
final_report.to_excel("clean_sales_report.xlsx",index=False)
print()
print(f"clean_sales_report.xlsx saved — {len(final_report)} rows")
print()
high_value_report=df.loc[df["FinalAmount"]>=500].copy()
high_value_report.to_excel("high_value_orders.xlsx",index=False)
print(f"high_value_orders.xlsx saved — {len(high_value_report)} rows")
print()
city_summary_order=final_report.groupby("City")["OrderID"].count()
city_summary_grossrevenue=final_report.groupby("City")["Total"].sum().round(2)
city_summary_netrevenue=final_report.groupby("City")["FinalAmount"].sum().round(2)
city_summary_avgordervalue=final_report.groupby("City")["FinalAmount"].mean().round(2)
city_summary_discount=final_report.groupby("City")["Discount"].sum().round(2)
city_summary=pd.DataFrame()
city_summary["Orders"]=city_summary_order
city_summary["GrossRevenue"]=city_summary_grossrevenue
city_summary["NetRevenue"]=city_summary_netrevenue
city_summary["AvgOrderValue"]=city_summary_avgordervalue
city_summary["TotalDiscount"]=city_summary_discount
city_summary.to_excel("city_summary.xlsx")
print("city_summary.xlsx saved")
product_summary=pd.DataFrame()
product_summary_quantity=final_report.groupby("Product")["Quantity"].sum()
product_summary_revenue=final_report.groupby("Product")["FinalAmount"].sum()
product_summary_avgprice=final_report.groupby("Product")["Price"].mean()
product_summary["TotalQuantitySold"]=product_summary_quantity
product_summary["TotalRevenue"]=product_summary_revenue
product_summary["AvgPrice"]=product_summary_avgprice
product_summary.sort_values(by="TotalRevenue",ascending=False,inplace=True)
product_summary.to_excel("product_summary.xlsx")
print()
print("product_summary.xlsx saved")
print()
print("="*50)
print("PROJECT 2 COMPLETE")
print("Files saved:")
print("  - clean_sales_report.xlsx")
print("  - high_value_orders.xlsx")
print("  - city_summary.xlsx")
print("  - product_summary.xlsx")
print("="*50)
