from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Setup
driver = webdriver.Chrome()
driver.get("https://vahan.parivahan.gov.in/vahan4dashboard/vahan/view/reportview.xhtml")
wait = WebDriverWait(driver, 30)

# Helper: wait until any loading overlay disappears
def wait_for_loading():
    try:
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "ui-blockui")))
    except:
        pass  # In case overlay is not present

# Helper: safe click download button
def safe_click_download():
    wait_for_loading()
    btn = wait.until(EC.element_to_be_clickable((By.ID, "groupingTable:xls")))
    btn.click()


State_dropdown_id = "j_idt39"
refresh_button_id = "j_idt71"
vehicle_category_refrash_id = "j_idt76"



# --- State list ---
states = [
   
   
   
   


   
    "Andaman & Nicobar Island(3)",
    "Andhra Pradesh(84)",
    "Arunachal Pradesh(29)",
    "Assam(33)",
    "Bihar(48)",
    "Chhattisgarh(31)",
    "Chandigarh(1)",
    "Delhi(16)",
    "UT of DNH and DD(3)",
    "Goa(13)",
                "Gujarat(37)",
                "Himachal Pradesh(96)",
                "Haryana(98)",
                "Jharkhand(25)",
                "Jammu and Kashmir(21)",
                "Karnataka(68)",
                "Kerala(87)",
                "Ladakh(3)",
                "Lakshadweep(6)",
                "Maharashtra(59)",
    "Meghalaya(15)",
    "Manipur(13)",
    "Madhya Pradesh(53)",
    "Mizoram(10)",
    "Nagaland(9)",
    "Odisha(39)",
    "Punjab(96)",
    "Rajasthan(59)",
    "Puducherry(8)",
    "Sikkim(9)",
                "Tamil Nadu(148)",
                "Uttarakhand(21)",
                "Tripura(9)",
                "Uttar Pradesh(77)",
                "West Bengal(59)"
]

# Select Y-Axis = Maker
dropdown_y_axis = wait.until(EC.element_to_be_clickable((By.ID, "yaxisVar_label")))
dropdown_y_axis.click()
time.sleep(1)
y_axis_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(., 'Maker')]")))
y_axis_option.click()
time.sleep(2)

# Select X-Axis = Fuel
dropdown_x_axis = wait.until(EC.element_to_be_clickable((By.ID, "xaxisVar_label")))
dropdown_x_axis.click()
time.sleep(1)
fuel_option = wait.until(EC.element_to_be_clickable((By.ID, "xaxisVar_3")))
fuel_option.click()
time.sleep(2)

# Select Year
dropdown_year = wait.until(EC.element_to_be_clickable((By.ID, "selectedYear_label")))
dropdown_year.click()
time.sleep(1)
year_option = wait.until(EC.element_to_be_clickable((By.ID, "selectedYear_3")))
year_option.click()
time.sleep(2)

# Open Vehicle Filter
vehicle_filter_button = wait.until(EC.element_to_be_clickable((By.ID, "filterLayout-toggler")))
vehicle_filter_button.click()
time.sleep(1)

# --- Month list ---
months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

# --- Loop through States ---
for state_name in states:
    print(f"Processing state: {state_name}")
    
    # Select State
    dropdown_state = wait.until(EC.element_to_be_clickable((By.ID, State_dropdown_id)))
    dropdown_state.click()
    time.sleep(1)

    state_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[contains(text(), '{state_name}')]")))
    state_option.click()
    time.sleep(2)

    # Click refresh after selecting axis
    refresh_button = wait.until(EC.element_to_be_clickable((By.ID, refresh_button_id)))
    refresh_button.click()
    wait_for_loading()


     # Select TWO WHEELER(NT)
    M2WN_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[normalize-space()='TWO WHEELER(NT)']")))
    M2WN_label.click()
    wait_for_loading()

    # Select TWO WHEELER(T)
    M2WT_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[normalize-space()='TWO WHEELER(T)']")))
    M2WT_label.click()
    wait_for_loading()

    # Loop through months
    for month in months:
        # Open month dropdown
        month_dropdown = wait.until(EC.element_to_be_clickable((By.ID, "groupingTable:selectMonth_label")))
        month_dropdown.click()
        time.sleep(1)

        # Select month
        Month_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[normalize-space()='{month}']")))
        Month_option.click()
        time.sleep(1)

        # Refresh for vehicle category
        refresh_button = wait.until(EC.element_to_be_clickable((By.ID, vehicle_category_refrash_id)))
        refresh_button.click()
        wait_for_loading()

        # Download file (fresh lookup every time)
        safe_click_download()

        # Print confirmation
        print(f"Excel file of {state_name} for {month} has been downloaded")

input("Press Enter to close...")
driver.quit()