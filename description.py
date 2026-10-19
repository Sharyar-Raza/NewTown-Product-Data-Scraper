from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

from bs4 import BeautifulSoup
import pandas as pd
import time
import re
import os



BASE = "https://www.newtownspares.com"

FILE = "Internal_Hard_Drive_Child_Data.xlsx"

LIVE_FILE = "Internal_Hard_Drive_Live.xlsx"



# =========================
# Chrome Setup
# =========================

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.page_load_strategy = "eager"


driver = webdriver.Chrome(
    service=Service(
        ChromeDriverManager().install()
    ),
    options=options
)


driver.set_page_load_timeout(60)





# =========================
# Categories
# =========================

categories = {

"Hard Drive Accessories":
"/hard-drive-accessories/",

"Storage Media":
"/storage-media/",

"Desktop Hard Drive":
"/desktop-hard-drive/",

"Laptop Hard Drive":
"/laptop-hard-drive/",

"Printer Hard Drive":
"/printer-hard-drive/",

"Server Hard Drive":
"/server-hard-drive/"

}






# =========================
# Load Existing Data
# =========================

data=[]


if os.path.exists(FILE):

    old=pd.read_excel(FILE)

    data=old.to_dict(
        "records"
    )

    print(
        "Old Data Loaded:",
        len(data)
    )



done=set()


for row in data:

    if "URL" in row:

        done.add(
            row["URL"]
        )







# =========================
# Save Excel
# =========================

def save_excel():

    df = pd.DataFrame(data)


    df.drop_duplicates(
        subset=["URL"],
        inplace=True
    )


    try:

        # Live Excel file (open this to watch progress)
        df.to_excel(
            LIVE_FILE,
            index=False
        )

        print(
            "Live Excel Updated:",
            len(df)
        )


    except PermissionError:

        print(
            "Live Excel open hai, next save me retry hoga"
        )


    # Safe backup
    df.to_csv(
        "Internal_Hard_Drive_Backup.csv",
        index=False,
        encoding="utf-8-sig"
    )








# =========================
# Product Scraper
# =========================

def scrape_product(url, category):


    if url in done:

        print("Skip:",url)

        return



    try:


        try:

            driver.get(url)

        except:

            driver.execute_script(
                "window.stop();"
            )



        time.sleep(4)



        soup=BeautifulSoup(
            driver.page_source,
            "html.parser"
        )


        text=soup.get_text(
            " ",
            strip=True
        )






        # =====================
        # Product Name
        # =====================

        name="-"


        h1=soup.find("h1")


        if h1:

            name=h1.get_text(
                " ",
                strip=True
            )







        # =====================
        # Brand
        # =====================

        brand="-"


        brand_match=re.search(
            r"Brand:\s*(.*?)\s*(SKU:|No reviews|$)",
            text
        )


        if brand_match:

            brand=brand_match.group(1).strip()







        # =====================
        # SKU
        # =====================

        sku="-"


        sku_match=re.search(
            r"SKU:\s*([A-Za-z0-9\-]+)",
            text
        )


        if sku_match:

            sku=sku_match.group(1)







        # =====================
        # PRICE EXACT XPATH
        # =====================

        old_price="-"

        current_price="-"



        try:

            price_element = driver.find_element(
                By.XPATH,
                "/html/body/div[8]/div[1]/div/div[2]/div/section[1]/div[1]/div[7]"
            )


            price_text = price_element.text.strip()


            print(
                "PRICE:",
                price_text
            )


            prices=re.findall(
                r"\$[\d,]+\.\d{2}",
                price_text
            )


            if len(prices)>=2:

                old_price=prices[0]

                current_price=prices[1]


            elif len(prices)==1:

                current_price=prices[0]

                old_price="-"



        except Exception as e:

            print(
                "Price Error:",
                e
            )









        # =====================
        # DESCRIPTION EXACT XPATH
        # =====================

        description="-"



        try:


            desc_element=driver.find_element(
                By.XPATH,
                '//*[@id="productDescription"]/div[2]'
            )


            description=desc_element.text.strip()



        except Exception as e:

            print(
                "Description Error:",
                e
            )









        # =====================
        # DETAIL
        # =====================

        detail="-"



        try:


            detail_heading=soup.find(
                string=re.compile(
                    "^Details$",
                    re.I
                )
            )


            if detail_heading:


                parent=detail_heading.find_parent()


                detail_text=parent.parent.get_text(
                    " | ",
                    strip=True
                )


                detail=re.sub(
                    r"\s+",
                    " ",
                    detail_text
                )



        except:

            pass







        data.append({

            "Product Name":
            name,


            "Brand":
            brand,


            "SKU":
            sku,


            "Listing Price":
            old_price,


            "Current Price":
            current_price,


            "Description":
            description,


            "Detail":
            detail,


            "Category Path":
            "Storage Devices > Internal Hard Drives > " + category,


            "URL":
            url

        })



        done.add(url)



        print(
            "Saved:",
            name,
            "|",
            current_price
        )



        save_excel()



    except Exception as e:


        print(
            "Product Error:",
            e
        )

        time.sleep(5)









# =========================
# Category Scraper
# =========================

def scrape_category(category,url):


    page=1



    while True:


        if page==1:

            page_url=BASE+url

        else:

            page_url=(
                BASE+
                url+
                f"?page={page}"
            )



        print(
            "\nCATEGORY:",
            category,
            "PAGE:",
            page
        )



        try:

            driver.get(
                page_url
            )

        except:

            driver.execute_script(
                "window.stop();"
            )



        time.sleep(4)



        soup=BeautifulSoup(
            driver.page_source,
            "html.parser"
        )



        cards=soup.select(
            "article.card"
        )



        print(
            "Products Found:",
            len(cards)
        )



        if len(cards)==0:

            break




        links=[]



        for card in cards:


            a=card.find(
                "a",
                href=True
            )


            if a:


                link=a["href"]


                if not link.startswith("http"):

                    link=BASE+link



                links.append(link)



        links=list(set(links))



        for link in links:


            scrape_product(
                link,
                category
            )



        page+=1







# =========================
# START
# =========================

for cat,url in categories.items():

    scrape_category(
        cat,
        url
    )





save_excel()



print("===================")
print("SCRAPING COMPLETE")
print(
    "TOTAL PRODUCTS:",
    len(data)
)
print("===================")



try:

    pd.DataFrame(data).to_excel(
        FILE,
        index=False
    )

    print("Final Excel Created")


except Exception as e:

    print("Final Excel Error:", e)


driver.quit()