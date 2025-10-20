import pandas as pd

df = pd.read_csv(r"C:\Users\sujal\Downloads\2.csv")

df['todate'] = pd.to_datetime(df['todate'], dayfirst=True, errors='coerce')
df['Month'] = df['todate'].dt.strftime('%b %Y')


df = df.drop(['modelDesc' ,'colour', 'makeYear', 'seatCapacity','secondVehicle','tempRegistrationNumber',
               'category', 'fromdate', 'todate'], axis=1)

df = df[df['vehicleClass'].isin(['MOTOR CYCLE', 'MOTOR CAR', 'Motor Cab', 'Maxi Cab', 'Motor Cycle for Hire'])]


df = df.rename(columns={
    "makerName": "Maker Name",
    "OfficeCd": "RTA",
    "fuel": "Fuel Type",
    "vehicleClass": "Vehicle Type",
    "Month": "Month"
})

result = (
    df[["Maker Name", "RTA", "Fuel Type", "Vehicle Type", "Month"]]
    .groupby(["Maker Name", "RTA", "Fuel Type", "Vehicle Type", "Month"])
    .size()
    .reset_index(name="Count")
)

result.to_csv("output.csv", index=False)

print("✅ Transformation complete. File saved as Transformed_RTA_Data.xlsx")


