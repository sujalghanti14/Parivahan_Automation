# Parivahan_Automation
This project automates the download and preprocessing of data from the Parivahan portal.

🔍 What does this code do?
When you run the Python script on your local machine, it will:
-Automatically open the Parivahan Portal
-Apply all the necessary filters
-Refresh and download the data
-Repeat this process in a loop to fetch datasets month-wise and year-wise, RTA-wise (for TN, MH, AP, as well as all other states).

📂 Combine.py - 
This script merges all the downloaded files into a single Excel file for analysis.
It also performs data cleaning by removing unnecessary entries such as null values and zeros — essentially serving as the preprocessing step before analysis.

⚙️ How does it work?
Automation (Data Download):
    -Built using Python and Selenium.
    -Each button or filter on the Parivahan portal has a unique element ID (visible through browser inspection).
    -The script detects these IDs, clicks the buttons, applies filters, and downloads the data.
    -This process runs in a loop to cover multiple states, months, and years.
Preprocessing (Data Cleaning & Structuring):
    -Implemented with Pandas, NumPy, and Regex (re).
    -Cleans the raw data by removing nulls, zeros, and unwanted records.
    -Structures the dataset into a format that’s ready for analysis.

🎯 Why is this project important?
 This project is especially valuable for professionals in the auto retail and automobile industry.
    ⏳ Saves time – A manual process that usually takes days can now be done in just minutes or hours.
    📊 Enables detailed market analysis – Downloads data RTA-wise, month-wise, and fuel-type-wise for in-depth insights.
    ⚡ Boosts efficiency – Provides clean, ready-to-use data for quick decision-making.
If you are not in the automobile industry, this project may not be directly relevant to you. But for industry professionals, it can be a game changer.