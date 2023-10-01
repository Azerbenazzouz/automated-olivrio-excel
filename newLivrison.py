# Description: This file contains the list of the governorates and their respective state and aria codes.
Governorate = [
    "Ariana",
    "Bizerte",
    "Jendouba",
    "Monastir",
    "Tunis",
    "Mannouba",
    "Gafsa",
    "Sfax",
    "Gabes",
    "Tataouine",
    "Medenine",
    "Kef",
    "Kebili",
    "Siliana",
    "Kairouan",
    "Zaghouan",
    "Ben Arous",
    "Sidi Bouzid",
    "Mahdia",
    "Tozeur",
    "Kasserine",
    "Sousse",
    "Nabeul",
    "Beja",
]
State = [
    2550,
    2551,
    2552,
    2553,
    2554,
    2555,
    2556,
    2557,
    2558,
    2559,
    2560,
    2561,
    2562,
    2563,
    2564,
    2565,
    2566,
    2567,
    2568,
    2569,
    2570,
    2571,
    2572,
    2573,
]
Aria = [
    2164,
    2165,
    2166,
    2167,
    2168,
    2169,
    2170,
    2171,
    2172,
    2173,
    2174,
    2175,
    2176,
    2177,
    2178,
    2179,
    2180,
    2181,
    2182,
    2183,
    2184,
    2185,
    2186,
    2187,
]
# ======================================================================================================================
# import the necessary libraries
import glob
import os
import pandas as pd
import datetime

# ======================================================================================================================
# get the last modified file
list_of_files = glob.glob("*.xlsx")  # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getctime)
# read the excel file
df = pd.read_excel(
    latest_file,
    dtype={
        "nom": str,
        "tel": str,
        "tel2": str,
        "email": str,
        "adr": str,
        "ville": str,
        "local": str,
        "cod": str,
        "desc": str,
        "prix": str,
    },
)
# df to dict
data = df.to_dict(orient="index")
data = list(data.values())

# ======================================================================================================================
# type	branch_id	shipping_date	client_phone	client_address	reciver_name	reciver_phone	reciver_address	from_country_id	to_country_id	from_state_id	to_state_id	from_area_id	to_area_id	payment_type	payment_method_id	attachments_before_shipping	package_id	description	qty	weight	length	width	height	amount_to_be_collected	order_id	delivery_time	collection_time
# new df without data
newDf = pd.DataFrame(
    columns={
        "type": int,
        "branch_id": int,
        "shipping_date": str,
        "client_phone": str,
        "client_address": str,
        "reciver_name": str,
        "reciver_phone": str,
        "reciver_address": str,
        "from_country_id": int,
        "to_country_id": int,
        "from_state_id": int,
        "to_state_id": int,
        "from_area_id": int,
        "to_area_id": int,
        "payment_type": int,
        "payment_method_id": int,
        "attachments_before_shipping": str,
        "package_id": int,
        "description": str,
        "qty": int,
        "weight": float,
        "length": float,
        "width": float,
        "height": float,
        "amount_to_be_collected": float,
        "order_id": str,
        "delivery_time": str,
        "collection_time": str,
    }
)


# ======================================================================================================================
def GovernorateId(governorate):
    for i in range(len(Governorate)):
        if Governorate[i] == governorate:
            return State[i], Aria[i]
    return 0, 0


# ======================================================================================================================
# fill the newDf
for i in range(len(data)):
    if i == 0:
        for j in range(2):
            newDf.loc[j] = [
                2,
                1,
                datetime.datetime.now().strftime("%d-%m-%Y"),
                "",
                19,
                data[i]["nom"],
                data[i]["tel"],
                data[i]["adr"],
                224,
                224,
                GovernorateId("Ben Arous")[0],
                GovernorateId(data[i]["ville"])[0],
                GovernorateId("Ben Arous")[1],
                GovernorateId(data[i]["ville"])[1],
                1,
                "cash_payment",
                "",
                1,
                data[i]["desc"],
                1,
                1,
                1,
                1,
                1,
                int(data[i]["prix"]) - 8,
                "",
                "",
                "",
            ]
    else:
        newDf.loc[i + 1] = [
            2,
            1,
            datetime.datetime.now().strftime("%d-%m-%Y"),
            "",
            19,
            data[i]["nom"],
            data[i]["tel"],
            data[i]["adr"],
            224,
            224,
            GovernorateId("Ben Arous")[0],
            GovernorateId(data[i]["ville"])[0],
            GovernorateId("Ben Arous")[1],
            GovernorateId(data[i]["ville"])[1],
            1,
            "cash_payment",
            "",
            1,
            data[i]["desc"],
            1,
            1,
            1,
            1,
            1,
            int(data[i]["prix"]) - 8,
            "",
            "",
            "",
        ]
# ======================================================================================================================
# save the newDf to a csv file
newDf.to_csv("newLivrison.csv", index=False)
# ======================================================================================================================

print(data[0]["nom"])
