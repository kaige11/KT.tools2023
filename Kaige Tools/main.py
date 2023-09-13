#packages
import pandas as pd
import openpyxl

# Replace 'your_file.xlsx' with the path to your Excel file
excel_file_path = 'Spendlove Budget.xlsx'

# List of sheet names you want to process (add or remove sheet names as needed)
sheet_names_to_process = [
    'July'  # Replace with your sheet names
]
fileName = input("enter new file name: ")
# Define a function to classify transactions based on their description
def classify_transaction(transaction):
    keywords = {
    'FOOD': [
        "domino's", "sq *indaba c", "fiesta mexic", "girl scout c", "subway", "tst* twigs b", "safeway #",
        "chipotle", "sq *indaba c", "tailwind psc", "tokyo joes g", "texas roadho", "paypal *kach", "conoco - sei",
        "sq *indaba c", "nayax vendin", "costco by in", "sq *indaba c", "pp*instacart", "walmart.com",
        "bobo's oat b", "shiki hibach", "paypal *thri", "tst* la bell", "paypal *uber", "paypal *uber", "poppis -",
        "sq *indaba c", "guys  q", "walgreens #", "panda expres", "fred meyer", "sq *indaba c", "tst* costa v",
        "royal mart", "paypal *fift", "tst* costa v", "paypal *uber", "paypal *uber", "chipotle", "safeway #",
        "safeway #", "sq *indaba c", "greek island", "mid columbia", "sq *popular", "yoke's fresh", "blue dolphin",
        "royalmart", "indabacoffee", "bonefishgril", "indabacoffee", "tomatilloaut", "circlek#", "til*tpblackr",
        "til*tpblackr", "starbuckssto", "popeyes", "shikihibachi", "indabacoffee", "greekislandc", "midcolumbiaw",
        "starbuckssto", "subzero", "supersupplem", "indabacoffee", "olivegardeni", "madeleinesca", "nayaxvending",
        "tacobell", "baskin#", "athleticgree", "auntieanne's", "chipotle", "tacobell", "circlek#", "nayaxvending",
        "doordash*cru"],
    'ENTERTAINMENT': [
        "apple.com/bi", "recreation.g", "fairchild ci", "prime video", "spotify usa", "spectrum",
        "iplay experi", "iplay experi", "spotify usa", "peter attia", "blue dolphin", "apple.com/bi",
        "apple.com/bi", "netflix.com", "audible*hmj", "nintendo", "paypal *goog", "all american",
        "apple.com/bi", "bestbuycom", "bestbuycom", "apple.com/bi", "sq *bonnie?s", "fairchild ci",
        "fairchild ci", "audible*wm", "microsoft*xb", "amazon prime", "u-haulsundow", "hulu",
        "k krates llc", "z place salo", "z place salo", "apple.com/bi", "ouraring inc", "chuck e chee",
        "apple.com/bi", "chuck e chee", "chuck e chee", "chuck e chee", "chuck e chee", "chuck e chee",
        "chuck e chee", "audible*dta", "sxm*siriusxm", "tnailsands", "tnailsands", "kenssporting", "tnailsands",
        "apple.com/bi", "apple.com/bi", "apple.com/bi", "apple.com/bi", "wixpayments*", "apple.com/bi",
        "apple.com/bi", "porscheclubo", "spaphuntakil", "apple.com/bi", "apple.com/bi", "thedailywire","nintendo","audible"],
    'HEALTH/FITNESS': [
        "TRIOS HEALTH KENNEWICK WA",
        "SIRI BRAZILIAN JIU JITSU 714-3253774 WA",
        "TRIOS HEALTH KENNEWICK WA",
        "FAMILY WELLNESS CENTER RICHLAND WA",
        "TRIOS HEALTH KENNEWICK WA",
        "THE G.O.A.T. SPORTS BA 970-3024377 CO",
        "TRIOS HEALTH KENNEWICK WA",
        "VTG*Couple & Family Insti 208-2075898 ID",
        "PELOTON* MEMBERSHIP HTTPSWWW.ONEP NY",
        "LUMIN PDF KUB0JMZBCO1P AUCKLAND",
        "TRIOS HEALTH KENNEWICK WA",
        "TRIOS HEALTH KENNEWICK WA",
        "ATOMIC DERMATOLOGY HTTPSWWW.ATOM WA",
        "ATOMIC DERMATOLOGY HTTPSWWW.ATOM WA",
        "TRIOS HEALTH KENNEWICK WA",
        "TRIOS HEALTH KENNEWICK WA",
        "VTG*Couple & Family Insti 208-2075898 ID",
        "TRIOS HEALTH KENNEWICK WA",
        "PAYPAL *BETTERME 35314369001",
        "EMPOWERED HEALTH HTTPSEMPOWERE WA",
        "TALKSPACE HTTPSWWW.TALK NY",
        "ATOMICDERMATOLOGY",
        "BH*REGAIN.US",
        "TRIOSHEALTH0000",
        "BIOLAYNETECHNOLOGIES",
        "PLANET Fit",
        "PLANET Fit","Jiu Jitsu"
    ],

     'AMAZON': [
        "AMZN"
    ],
    'AUTO': [
        "AUTOZONE #3727 KENNEWICK WA",
        "JAGUARLANDROVER","auto"
    ],
    'CAR WASH': [
        "Carwash",
        "Carwash",
        "MISTER CAR WASH #579 866-2543229 WA",
        "BLUE DOLPHIN CAR WASH KENNEWICK WA",
        "BLUE DOLPHIN CAR WASH KENNEWICK WA",
        "MCW0575-Kennewick CarWash",
        "MCW0575-Kennewick CarWash",
        "MCW0575-Kennewick CarWash",
        "BLUEDOLPHINCARWASH"
    ],
    'CLOTHING': [
        "Oak&Luna USA DE",
        "COLEHAAN.COM 800-488-2000 NH",
        "FAMOUS FOOTWEAR #2744 RICHLAND WA",
        "Etsy.com - RusticElegance 718-8557955 NY",
        "WALLAWALLACLOTHINGCO.-930553411192",
        "DICK'SSPORTINGGOODS1370",
        "TILLYS",
        "BUCKLE#2140000",
        "AMERICANEAGLEOUTFITTERS",
        "DICK'SSPORTINGGOODS1370",
        "SPSHADY RAYS",
        "tj maxx"
    ],
    'DEBT SERVICE': [
        "WEALTHFRONT",
        "PORSCHE FINANCIAPAYMENTS",
        "TRI CU",
        "WEALTHFRONT Payments",
        "FIRSTSTATEBANKACH",
        "HARBORSTONECREDLOAN",
        "MORTG",
        "Interest Expense",
        "Interest Expense"
    ],
    'DEPOSIT': [
        "deposit"],
    'GAS': [
        "CIRCLE K # 06034 KENNEWICK WA",
        "PAYPAL *MDLECLAIR13 402-935-7733 CA",
        "SHELL OIL 12669977006 KENNEWICK WA",
        "CIRCLEK#06030/CIRCLEK",
        "CIRCLEK#06044/CIRCLEK",
        "CIRCLEK",
        "FREDM FUEL#9286",
        "MAVERIK","Shell"
    ],
    'GROCERIES': [
        "TARGET",
        "WALGREENS",
        "Costco",
        "target"
    ],
    'HOME': [
        "ACE HDWE RICHLAND WA",
        "IN *CENTURY 21 TRI-CITIES 509-9470920 WA",
        "IN *CENTURY 21 TRI-CITIES 509-9470920 WA",
        "SQ *CENTURY 21 TRI-CITIES gosq.com WA",
        "LOWES #00249* KENNEWICK WA",
        "BENTON PUD SMARTHUB.BENT WA",
        "THE HOME DEPOT 4746 RICHLAND WA",
        "ATT*BILL PAYMENT 800-288-2020 TX",
        "WASTE MGMT - KENNEWICK KENNEWICK WA",
        "THE HOME DEPOT 4746 RICHLAND WA",
        "lowes",
        "LOWE'S",
        "HOMEGOODS#0776000000776",
        "THEHOMEDEPOT",
        "THEHOMEDEPOT",
        "THEHOMEDEPOT",
        "KENNEWICKRANCHANDHOME9489084706836"
    ],
    'INVESTMENT': [
        "FUNDRISEINCOMERE",
        "Fundrise Real"
    ],
    'OTHER': [
        "SUPRA RE 877-699-6787 FL",
        "ZOEK 888-7379635 CA",
        "FOUNTAIN GIFTS HTTPSWWW.FOUN NJ",
        "POCKETGUARD, INC. HTTPSSECURE.P CA",
        "Carwash",
        "SP S E E K E R HTTPSSEEKER.M CA",
        "76 - DALLAS RD 76 RICHLAND WA",
        "CBEGLLC WWW.CBEVENTGR WA",
        "GOVX INC 888-468-5511 CA",
        "gozoek.com 415-4499034 CA",
        "TST*PROOF00043582",
        "10799RIDGELINEDR12669977006",
        "CASHAPP*JENNIFERLINDBERG",
        "BLOSSOMFLOWERDELIVERY55811920000064",
        "ROGUE084870052089448",
        "TRIOSHEALTH0000",
        "UNION7609476490",
        "FAIRCHILDCINEMASSOUTHGA650000012492",
        "10799RIDGELINEDR12669977006",
        "TRIOSHEALTH0000",
        "GRAZELLC0000",
        "CASH",
        "CASH"
    ],
    'PET': [
        "PAYPAL *CHEWY INC 402-935-7733 FL",
        "LUCKYPUPPYGROOMING"
    ],
    'TRAVEL': [
        "RPSPASCOTRICITIESAP",
        "DENVERAIRPORTENTERPR",
        "HOMEWOODSUITESGREELEYHOMEWOODSUITE",
        "THEFINEST",
        "DELTAAIRLINES",
        "TAILWINDPSC545500001248145",
        "DELTAAIRLINES",
        "RPSPASCOTRICITIESAP",
        "SEI4174394174398"
    ],
    # Add more categories and keywords as needed
}

    for category, category_keywords in keywords.items():
        for keyword in category_keywords:
            if keyword in str(transaction).lower():  # Use str() to handle potential non-string values
                return category
    return 'Other'

# Load the Excel file and create a Pandas Excel writer
with pd.ExcelFile(excel_file_path) as xls:
    # Create an Excel writer
    output_file_path = f'{excel_file_path.split(".")[0]}_{fileName}.xlsx'
    writer = pd.ExcelWriter(output_file_path, engine='openpyxl')

    # Iterate through the sheets you want to process
    for sheet_name in sheet_names_to_process:
        # Read data from the current sheet into a DataFrame
        data = pd.read_excel(xls, sheet_name=sheet_name)

        # Apply the classification function to add a 'DESCRIPTION' column with classification
        data['DESCRIPTION'] = data['TRANSACTION'].apply(classify_transaction)

        # Save the updated DataFrame to the output Excel file
        data.to_excel(writer, sheet_name=sheet_name, index=False)

    # Save the Excel writer to the output file
    writer._save()

print(f"Classification complete. Results saved to {output_file_path}")

