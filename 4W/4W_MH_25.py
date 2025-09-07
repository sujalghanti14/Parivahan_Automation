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

# --- Select State ---
dropdown_state = wait.until(EC.element_to_be_clickable((By.ID, State_dropdown_id)))
dropdown_state.click()
time.sleep(1)

state_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Maharashtra(59)')]")))
state_option.click()
time.sleep(2)

# Open Vehicle Filter
vehicle_filter_button = wait.until(EC.element_to_be_clickable((By.ID, "filterLayout-toggler")))
vehicle_filter_button.click()
time.sleep(1)

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


# --- RTO list ---
rto_list = [









    "AKLUJ - MH45( 03-APR-2017 )",
    "AMBEJOGAI - MH44( 02-MAY-2017 )",
    "AMRAWATI - MH27( 21-JAN-2017 )",
    "BARAMATI - MH42( 10-MAR-2017 )",
    "BEED - MH23( 17-MAR-2017 )",
    "BHADGAON - MH54( 20-MAR-2024 )",
    "BHANDARA - MH36( 12-APR-2017 )",
    "BULDHANA - MH28( 07-NOV-2017 )",
    "CHALISGAON - MH52( 05-MAR-2024 )",
    "CHHATRAPATI SAMBHAJINAGAR - MH20( 20-OCT-2016 )",
    "Chiplun Chiplun Track - MH202( 04-DEC-2019 )",
    "DHARASHIV - MH25( 31-OCT-2017 )",
    "DHULE - MH18( 03-JAN-2017 )",
    "DY REGIONAL TRANSPORT OFFICE, HINGOLI - MH38( 15-JUL-2017 )",
    "DY RTO RATNAGIRI - MH8( 10-APR-2017 )",
    "GADCHIROLI - MH33( 18-APR-2017 )",
    "GONDHIA - MH35( 11-APR-2017 )",
    "ICHALKARANJI - MH51( 07-MAR-2024 )",
    "JALANA - MH21( 03-AUG-2017 )",
    "KALYAN - MH5( 11-MAY-2017 )",
    "KARAD - MH50( 20-MAR-2017 )",
    "KHAMGAON - MH56( 15-APR-2025 )",
    "KOLHAPUR - MH9( 02-MAR-2017 )",
    "MALEGAON - MH41( 23-AUG-2017 )",
    "MIRA BHAYANDAR - MH58( 07-MAY-2025 )",
    "MUMBAI (CENTRAL) - MH1( 15-DEC-2016 )",
    "MUMBAI (EAST) - MH3( 13-DEC-2016 )",
    "MUMBAI (WEST) - MH2( 21-APR-2017 )",
    "NAGPUR (EAST) - MH49( 17-APR-2017 )",
    "NAGPUR (RURAL) - MH40( 17-JAN-2017 )",
    "NAGPUR (U) - MH31( 18-JAN-2017 )",
    "NANDED - MH26( 12-JAN-2017 )",
    "NANDURBAR - MH39( 02-MAY-2017 )",
    "NASHIK - MH15( 05-JAN-2017 )",
    "PANVEL - MH46( 31-JAN-2017 )",
    "PARBHANI - MH22( 25-APR-2017 )",
    "PEN (RAIGAD) - MH6( 16-MAY-2017 )",
    "PHALTAN - MH53( 03-SEP-2024 )",
    "PUNE - MH12( 25-JAN-2017 )",
    "RTO AHEMEDNAGAR - MH16( 16-MAR-2017 )",
    "RTO AKOLA - MH30( 20-FEB-2017 )",
    "R.T.O.BORIVALI - MH47( 21-APR-2017 )",
    "RTO CHANDRAPUR - MH34( 25-APR-2017 )",
    "RTO JALGAON - MH19( 24-MAR-2017 )",
    "RTO LATUR - MH24( 15-MAR-2017 )",
    "RTO MH04-Mira Bhayander FitnessTrack - MH203( 01-MAY-2022 )",
    "RTO PIMPRI CHINCHWAD - MH14( 06-FEB-2017 )",
    "RTO SATARA - MH11( 04-MAR-2017 )",
    "RTO SOLAPUR - MH13( 05-APR-2017 )",
    "SANGLI - MH10( 03-MAR-2017 )",
    "SINDHUDURG(KUDAL) - MH7( 10-APR-2017 )",
    "SRIRAMPUR - MH17( 22-MAR-2017 )",
    "TC OFFICE - MH99( 06-JUN-2018 )",
    "THANE - MH4( 08-MAR-2017 )",
    "UDGIR - MH55( 28-AUG-2024 )",
    "VAIJAPUR - MH57( 06-JUN-2025 )",
    "VASAI - MH48( 08-JUN-2017 )",
    "VASHI (NEW MUMBAI) - MH43( 07-JUL-2016 )",
    "WARDHA - MH32( 06-APR-2017 )",
    "WASHIM - MH37( 11-APR-2017 )",
    "YAWATMAL - MH29( 07-JUL-2017 )"

]


# --- Month list ---
months = ["AUG"]

# --- Loop through RTOs ---
for rto_name in rto_list:
    # Click RTO dropdown
    dropdown_rta = wait.until(EC.element_to_be_clickable((By.ID, "selectedRto")))
    dropdown_rta.click()
    time.sleep(1)

    # Select current RTO
    rto_option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[contains(., '{rto_name}')]")))
    driver.execute_script("arguments[0].scrollIntoView(true);", rto_option)
    rto_option.click()
    time.sleep(2)

    # Click refresh after selecting axis
    refresh_button = wait.until(EC.element_to_be_clickable((By.ID, refresh_button_id)))
    refresh_button.click()
    wait_for_loading()

    # Select MOTOR CAR
    motor_car_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[normalize-space()='MOTOR CAR']")))
    motor_car_label.click()
    wait_for_loading()

    # Select Maxi Cabs
    motor_car_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[normalize-space()='MAXI CAB']")))
    motor_car_label.click()
    wait_for_loading()

    # Select Luxury Cars
    motor_car_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[normalize-space()='LUXURY CAB']")))
    motor_car_label.click()
    wait_for_loading()

    # Select MOTOR CAB
    motor_car_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[normalize-space()='MOTOR CAB']")))
    motor_car_label.click()
    wait_for_loading()

    # Select Vintner Cars
    motor_car_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[normalize-space()='VINTAGE MOTOR VEHICLE']")))
    motor_car_label.click()
    wait_for_loading()

    # Select Light Motor Vehicle
    motor_car_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[normalize-space()='LIGHT MOTOR VEHICLE']")))
    motor_car_label.click()
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
        print(f"Excel file of Maharashtra, {rto_name} for {month} has been downloaded")

input("Press Enter to close...")
driver.quit()
