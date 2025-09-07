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

state_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), 'Tamil Nadu(148)')]")))
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






    "ALANGUDI UO - TN641( 04-JUN-2018 )",
    "ALANGULAM UO - TN644( 05-MAR-2024 )",
    "AMBASAMUTHIRAM UO - TN611( 28-AUG-2018 )",
    "AMBATTUR RTO - TN612( 07-JUN-2018 )",
    "AMBUR UO - TN628( 11-JUL-2018 )",
    "ARAKKONAM UO - TN609( 06-JUL-2018 )",
    "ARANI RTO - TN516( 21-MAY-2018 )",
    "ARANTHANGI UO - TN592( 19-DEC-2017 )",
    "ARAVAKURICHI UO - TN632( 27-JUN-2018 )",
    "ARIYALUR RTO - TN61( 14-AUG-2018 )",
            "ARUPPUKOTTAI UO - TN622( 25-JUL-2018 )",
            "ATTUR RTO - TN591( 08-JAN-2018 )",
            "AVINASHI UO - TN580( 14-AUG-2018 )",
            "BATLAGUNDU UO - TN596( 20-JUL-2018 )",
            "BHAVANI UO - TN578( 30-JUL-2018 )",
            "CHENGALPATTU RTO - TN19( 07-JUN-2018 )",
            "CHENNAI (CENTRAL) RTO - TN1( 12-JUN-2018 )",
            "CHENNAI (EAST) RTO - TN4( 03-JUL-2018 )",
            "CHENNAI (NORTH-EAST) RTO - TN3( 03-JUL-2018 )",
            "CHENNAI (NORTH) RTO - TN5( 01-JUL-2018 )",
    "CHENNAI (SOUTH-EAST) RTO - TN6( 11-JUN-2018 )",
    "CHENNAI (SOUTH) RTO - TN7( 01-JUN-2018 )",
    "CHENNAI (SOUTH-WEST) RTO - TN10( 03-JUL-2018 )",
    "CHENNAI (WEST) RTO - TN9( 28-MAY-2018 )",
    "CHEYYAR UO - TN635( 07-APR-2017 )",
    "CHIDAMBARAM RTO - TN544( 19-JUN-2018 )",
    "COIMBATORE (CENTRAL) RTO - TN66( 17-JUL-2018 )",
    "COIMBATORE (NORTH) RTO - TN38( 17-JUL-2018 )",
    "COIMBATORE (SOUTH) RTO - TN37( 28-MAY-2018 )",
    "COIMBATORE (WEST) RTO - TN99( 17-JUL-2018 )",
            "CUDDALORE RTO - TN31( 15-JUN-2018 )",
            "DHARAPURAM RTO - TN594( 18-JUL-2018 )",
            "DHARMAPURI RTO - TN29( 03-AUG-2018 )",
            "DINDIGUL RTO - TN57( 19-JUN-2018 )",
            "ERODE RTO - TN33( 14-AUG-2018 )",
            "ERODE (WEST) RTO - TN86( 10-AUG-2018 )",
            "GINGEE UO - TN627( 14-JUN-2018 )",
            "GOPICHETTIPALAYAM RTO - TN36( 30-JUL-2018 )",
            "GUDALORE UO - TN582( 18-JUL-2018 )",
            "GUDIYATHAM UO - TN514( 06-JUL-2018 )",
    "GUMMIDIPOONDI UO - TN625( 07-JUN-2018 )",
    "HARUR UO - TN527( 03-AUG-2018 )",
    "HOSUR RTO - TN70( 06-JUL-2018 )",
    "ILLUPPUR UO - TN629( 24-MAY-2018 )",
    "KALLAKURICHI RTO - TN615( 15-JUN-2018 )",
    "KANCHEEPURAM RTO - TN21( 05-JUL-2018 )",
    "KANGEYAM UO - TN593( 17-AUG-2018 )",
    "KARAIKUDI UO - TN602( 21-AUG-2018 )",
    "KARUR RTO - TN47( 27-JUN-2018 )",
    "KOVILPATTI RTO - TN607( 24-AUG-2018 )",
            "KRISHNAGIRI RTO - TN24( 06-JUL-2018 )",
            "KULITHALI UO - TN585( 27-JUN-2018 )",
            "KUMARAPALAYAM RTO - TN638( 28-JAN-2018 )",
            "KUMBAKONAM RTO - TN68( 06-AUG-2018 )",
            "KUNDRATHUR RTO - TN85( 19-JUN-2018 )",
            "LALKUDI UO - TN631( 28-JUN-2018 )",
            "MADURAI (CENTRAL) RTO - TN64( 17-AUG-2018 )",
            "MADURAI (NORTH) RTO - TN59( 01-JUN-2018 )",
            "MADURAI (SOUTH) RTO - TN58( 06-JUL-2018 )",
            "MADURANTAGAM UO - TN508( 07-JUN-2018 )",
    "MANAPARAI UO - TN584( 30-MAY-2018 )",
    "MANMANGALAM UO - TN637( 25-JUN-2018 )",
    "MANNARGUDI UO - TN588( 09-AUG-2018 )",
    "MARTHANDAM RTO - TN75( 31-JUL-2018 )",
    "MAYILADUTHURAI RTO - TN589( 07-AUG-2018 )",
    "MEENAMBAKKAM RTO - TN22( 09-JUL-2018 )",
    "MELUR UO - TN600( 01-JUN-2018 )",
    "METTUPALAYAM RTO - TN40( 18-JUL-2018 )",
    "METTUR RTO - TN590( 30-AUG-2018 )",
    "MUSURI UO - TN621( 28-JUN-2018 )",
            "NAGAPATTINAM RTO - TN51( 07-AUG-2018 )",
            "NAGERCOIL RTO - TN74( 28-AUG-2018 )",
            "NAMAKKAL (NORTH) RTO - TN28( 30-AUG-2018 )",
            "NAMAKKAL (SOUTH) RTO - TN88( 30-AUG-2018 )",
            "NATHAM UO - TN640( 04-JUN-2018 )",
            "NEYVELI UO - TN562( 20-JUN-2018 )",
            "ODDANCHATRAM  UO - TN595( 24-JUL-2018 )",
            "OMALURE UO - TN535( 29-AUG-2018 )",
            "OOTY RTO - TN43( 18-JUL-2018 )",
            "PALACODE UO - TN616( 03-AUG-2018 )",
    "PALANI RTO - TN597( 24-JUL-2018 )",
    "PANRUTI UO - TN626( 19-JUN-2018 )",
    "PARAMAKUDI UO - TN603( 24-JUL-2018 )",
    "PARAMATHI VELLURE UO - TN517( 30-AUG-2018 )",
    "PATTUKOTTAI UNIT OFFICE - TN587( 06-AUG-2018 )",
    "PERAMBALUR RTO - TN46( 14-AUG-2018 )",
    "PERUNDURAI RTO - TN56( 14-AUG-2018 )",
    "POLLACHI RTO - TN41( 23-JUL-2018 )",
    "POONAMALLEE RTO - TN511( 11-JUN-2018 )",
    "PUDUKOTTAI RTO - TN55( 28-DEC-2017 )",
                "RAJAPALAYAM UO - TN643( 15-NOV-2023 )",
                "RAMANATHAPURAM RTO - TN65( 21-AUG-2018 )",
                "RANIPET RTO - TN73( 06-JUL-2018 )",
                "RASIPURAM UO - TN526( 30-AUG-2018 )",
                "REDHILLS RTO - TN18( 07-JUN-2018 )",
                "RTO CHENNAI (NORTH WEST) - TN2( 25-JUN-2018 )",
                "SALEM (EAST) RTO - TN54( 29-AUG-2018 )",
                "SALEM (SOUTH) RTO - TN90( 29-AUG-2018 )",
                "SALEM (WEST) RTO - TN30( 29-AUG-2018 )",
                "SANKAGIRI RTO - TN52( 30-AUG-2018 )",
    "SANKARANKOVIL RTO - TN610( 28-AUG-2018 )",
    "SATHYAMANGALAM UO - TN579( 30-JUL-2018 )",
    "SHOLINGANALLUR RTO - TN512( 19-JUN-2018 )",
    "SIRKALI UO - TN623( 07-AUG-2018 )",
    "SIVAGANGAI RTO - TN63( 21-AUG-2018 )",
    "SIVAKASI RTO - TN604( 21-AUG-2018 )",
    "SRIPERUMBUDUR RTO - TN614( 01-AUG-2018 )",
    "SRIRANGAM RTO - TN48( 28-JUN-2018 )",
    "SRIVILLIPUTHUR RTO - TN605( 17-AUG-2018 )",
    "STATE TRANSPORT AUTHORITY - TN999( 13-JUN-2018 )",
                "SULUR UO - TN620( 28-MAY-2018 )",
                "TAMBARAM RTO - TN513( 25-JUN-2018 )",
                "TENKASI RTO - TN76( 25-JUL-2018 )",
                "THANJAVUR RTO - TN49( 06-AUG-2018 )",
                "THENI RTO - TN60( 20-JUL-2018 )",
                "THIRUCHENDUR RTO - TN606( 28-AUG-2018 )",
                "THIRUKALUKUNTRAM UO - TN639( 04-JUN-2018 )",
                "THIRUMANGALAM  UO - TN598( 07-JUL-2018 )",
                "THIRUPATTUR RTO - TN624( 11-JUL-2018 )",
                "THIRUTHURAIPOONDI UO - TN630( 09-AUG-2018 )",
    "THIRUTTANI UO - TN634( 11-JUN-2018 )",
    "THOOTHUKUDI RTO - TN69( 24-AUG-2018 )",
    "THURAIYUR UO - TN586( 28-JUN-2018 )",
    "TINDIVANAM RTO - TN571( 14-JUN-2018 )",
    "TIRUCHENGODE RTO - TN34( 15-MAR-2018 )",
    "TIRUCHI(EAST) RTO - TN81( 06-JUL-2018 )",
    "TIRUCHI RTO - TN45( 30-MAY-2018 )",
    "TIRUNELVELI RTO - TN72( 25-JUL-2018 )",
    "TIRUPPUR (NORTH) RTO - TN39( 25-JUL-2018 )",
    "TIRUPPUR (SOUTH) RTO - TN42( 25-JUL-2018 )",   
                "TIRUVALLUR RTO - TN20( 11-JUN-2018 )",
                "TIRUVANNAMALAI RTO - TN25( 21-MAY-2018 )",
                "TIRUVARUR RTO - TN50( 09-AUG-2018 )",
                "TIRUVERANBUR UO - TN583( 09-JUL-2018 )",
                "UDUMALPET RTO - TN581( 17-AUG-2018 )",
                "ULUNDURPET RTO - TN577( 15-JUN-2018 )",
                "USILAMPATTI UO - TN636( 14-AUG-2018 )",
                "UTHAMAPALAYAM UO - TN601( 20-JUL-2018 )",
                "VADIPATTI UO - TN599( 05-JUN-2018 )",
                "VALAPPADI UO - TN613( 08-JAN-2018 )",
    "VALLIYUR UO - TN608( 25-JUL-2018 )",
    "VALPARAI UO - TN633( 24-JUL-2018 )",
    "VANIYAMBADI RTO - TN515( 11-JUL-2018 )",
    "VEDACHANDUR UO - TN617( 20-JUL-2018 )",
    "VELLORE RTO - TN23( 05-JUL-2018 )",
    "VILUPPURAM RTO - TN32( 14-JUN-2018 )",
    "VIRUDHACHALAM UO - TN553( 19-JUN-2018 )",
    "VIRUDHUNAGAR RTO - TN67( 25-JUL-2018 )"
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
        print(f"Excel file of Tamil Nadu, {rto_name} for {month} has been downloaded")

input("Press Enter to close...")
driver.quit()
