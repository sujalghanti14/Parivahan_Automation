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


State_dropdown_id = "j_idt34"
refresh_button_id = "j_idt68"
vehicle_category_refrash_id = "j_idt77"



# --- Select State ---
dropdown_state = wait.until(EC.element_to_be_clickable((By.ID, State_dropdown_id)))
dropdown_state.click()
time.sleep(1)

state_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Andhra Pradesh(84)')]")))
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

   
   
#         "Adoni RTO - AP221( 06-MAY-2022 )",
#         "Amalapuram RTA - AP205( 25-APR-2022 )",
#         "Anakapalli RTA - AP131( 25-APR-2022 )",
#         "Anantapur RTA - AP2( 25-APR-2022 )",
#         "Atmakur-Kurnool MVI Office - AP321( 06-MAY-2022 )",
#         "Atmakur MVI Office - AP126( 11-MAY-2022 )",
#         "Badvel MVI Office - AP104( 06-MAY-2022 )",
#         "BAPATLA RTO OFFICE - AP207( 25-APR-2022 )",
#         "Bhimavaram RTA - AP137( 25-APR-2022 )",
#         "Chilakaluripeta MVI Office - AP307( 06-MAY-2022 )",
# "Chirala UO - AP127( 06-MAY-2022 )",
# "Chintoor - AP905( 09-JUN-2022 )",
# "Chittoor RTA - AP3( 25-APR-2022 )",
# "Cuddapah RTA - AP4( 25-APR-2022 )",
# "Darsi UO - AP427( 11-MAY-2022 )",
# "Dharamavaram unit office - AP602( 04-DEC-2024 )",
# "Dhone MVI Office - AP421( 06-MAY-2022 )",
# "Gudiwada RTA - AP116( 05-MAY-2022 )",
# "Gajuwaka RTA - AP231( 06-MAY-2022 )",
# "Gudur RTA - AP226( 06-MAY-2022 )",
#         "Guntakal UO - AP202( 06-MAY-2022 )",
#         "Guntur RTA - AP7( 25-APR-2022 )",
#         "Itchapuram MVI Office - AP130( 12-MAY-2022 )",
#         "Hindupur RTA - AP102( 25-APR-2022 )",
#         "Jaggayyapet UO - AP616( 05-MAY-2022 )",
#         "JANGAREDDYGUDEM RTA - AP237( 06-MAY-2022 )",
#         "Kalyandurg RTO office - AP702( 04-DEC-2024 )",
#         "Kandukur MVI Office - AP227( 11-MAY-2022 )",
#         "Kavali UO - AP326( 11-MAY-2022 )",
#         "Kovvuru UO - AP337( 12-MAY-2022 )",
# "Kurnool RTA - AP21( 25-APR-2022 )",
# "Macherla MVI Office - AP407( 06-MAY-2022 )",
# "Mandapeta UO - AP405( 06-MAY-2022 )",
# "Mangalagiri MVI Office - AP507( 06-MAY-2022 )",
# "Markapur UO - AP327( 12-MAY-2022 )",
# "Nandigama RTA - AP316( 05-MAY-2022 )",
# "Nagari MVI office - AP114( 30-NOV-2024 )",
# "Nandyal RTA - AP121( 25-APR-2022 )",
# "Narasaraopet RTA - AP107( 25-APR-2022 )",
# "Narsipatnam MVI Office - AP331( 06-MAY-2022 )",
#         "Nellore RTA - AP26( 25-APR-2022 )",
#         "Nuzvid UO - AP416( 06-MAY-2022 )",
#         "Paderu RTA - AP41( 09-JUN-2022 )",
#         "Palakole UO - AP637( 06-MAY-2022 )",
#         "PALAKONDA RTA - AP230( 12-MAY-2022 )",
#         "Palamaner MVI Office - AP303( 10-MAY-2022 )",
#         "Palasa MVI Office - AP330( 12-MAY-2022 )",
#         "PARVATHIPURAM RTO OFFICE - AP135( 25-APR-2022 )",
#         "Peddapuram MVI Office - AP505( 06-MAY-2022 )",
#         "Piduguralla UO - AP607( 06-MAY-2022 )",
# "Piler MVI Office - AP403( 12-MAY-2022 )",
# "Prakasam RTA - AP27( 25-APR-2022 )",
# "Proddutur RTA - AP204( 06-MAY-2022 )",
# "Pulivendula MVI Office - AP304( 06-MAY-2022 )",
# "Punganur UO - AP113( 09-JUN-2022 )",
# "Puttur MVI Office - AP503( 12-MAY-2022 )",
# "Rajahmundry RTA - AP105( 25-APR-2022 )",
# "Rajampet MVI Office - AP404( 12-MAY-2022 )",
# "Ramachandrapuram UO - AP605( 09-JUN-2022 )",
# "Ravulapalem UO - AP705( 06-MAY-2022 )",
#         "Rayachoti MVI Office - AP504( 25-APR-2022 )",
#         "REGIONAL TRANSPORT OFFICE RAMPACHODAVARAM - AP141( 04-APR-2023 )",
#         "RTA Eluru - AP37( 25-APR-2022 )",
#         "RTA Kakinada - AP5( 25-APR-2022 )",
#         "RTA MACHILIPATNAM - AP216( 25-APR-2022 )",
#         "RTO KADIRI - AP302( 11-MAY-2022 )",
#         "RTO MADANAPALLE - AP203( 12-MAY-2022 )",
#         "Salur MVI Office - AP235( 12-MAY-2022 )",
#         "Srikakulam RTA - AP30( 25-APR-2022 )",
#         "Srikalahasthi MVI Office - AP603( 12-MAY-2022 )",
# "Sullurpet UO - AP426( 11-MAY-2022 )",
# "Tadepalli Gudem UO - AP437( 06-MAY-2022 )",
# "Tadipatri UO - AP402( 06-MAY-2022 )",
# "Tanuku UO - AP537( 06-MAY-2022 )",
# "Tekkali MVI Office - AP430( 12-MAY-2022 )",
# "TENALI RTA - AP707( 06-MAY-2022 )",
# "Tirupati RTA - AP103( 25-APR-2022 )",
# "UNIT OFFICE KATHIPUDI - AP305( 06-MAY-2022 )",
# "UNIT OFFICE RAYADURG - AP502( 23-JUN-2022 )",
# "Vijayawada RTA - AP16( 17-FEB-2022 )",
#         "Vishakapatnam RTA - AP31( 25-APR-2022 )",
#         "Vizianagaram RTA - AP35( 25-APR-2022 )",
        "Vuyyuru UO - AP516( 05-MAY-2022 )"
]


# --- Month list ---
months = ["SEP"]

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

    # Select Light Motor Vehicle
    motor_car_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[normalize-space()='LIGHT MOTOR VEHICLE']")))
    motor_car_label.click()
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
        print(f"Excel file of Andhra Pradesh, {rto_name} for {month} has been downloaded")

input("Press Enter to close...")
driver.quit()
