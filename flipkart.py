import pandas as pd
import requests
from bs4 import BeautifulSoup
Product_name = []
price = []
Description = []
Reviews = []

for i in range(2,12):
    url = f"https://www.flipkart.com/search?q=mobiles+under+50000+rupees&sid=tyy%2C4io&as=on&as-show=on&otracker=AS_QueryStore_OrganicAutoSuggest_1_19_na_na_na&otracker1=AS_QueryStore_OrganicAutoSuggest_1_19_na_na_na&as-pos=1&as-type=HISTORY&suggestionId=mobiles+under+50000+rupees%7CMobiles&requestId=61ff3ce5-7bb3-47ab-8d41-21b1552ae903&page="+str(i)

    r = requests.get(url)
        # print(r)

    soup = BeautifulSoup(r.text, "lxml")
    boxes = soup.find_all("div",class_ = "nZIRY7")
    print("Total boxes:",len(boxes))
    

    for i in boxes:
        names = i.find("div", class_="RG5Slk")

        if names:
            name = names.text
            Product_name.append(name)
        else:
            Product_name.append("N/A")


    for i in boxes:
        prices = i.find("div", class_="hZ3P6w")

        if prices:
            price.append(prices.text.strip())
        else:
            price.append("N/A")

    for i in boxes:
        desc = i.find("ul", class_="HwRTzP")

        if desc:
            Description.append(desc.get_text(" | ", strip=True))
        else:
            Description.append("N/A")

    for i in boxes:
        reviews = i.find("div", class_="MKiFS6")

        if reviews:
            Reviews.append(reviews.text.strip())
        else:
            Reviews.append("N/A")
        
print("Product Name:", len(Product_name))
print("Price:", len(price))
print("Description:", len(Description))
print("Reviews:", len(Reviews))


df = pd.DataFrame({"Product Name":Product_name, "Price":price, "Description":Description, "reviews":Reviews})
print(df)

with pd.ExcelWriter("Flipkart_data.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, index=False)

    sheet = writer.sheets["Sheet1"]

    sheet.column_dimensions["A"].width = 50
    sheet.column_dimensions["B"].width = 15
    sheet.column_dimensions["C"].width = 80
    sheet.column_dimensions["D"].width = 12

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    from openpyxl.styles import Alignment, Font

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(
                wrap_text=True,
                vertical="top"
            )


df.to_excel("Flipkart_data.xlsx", index=False)

print("Excel file created successfully!")








































    
        