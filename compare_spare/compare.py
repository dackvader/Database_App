import pandas as pd
import re
from collections import defaultdict

# 📌 กำหนดชื่อไฟล์ที่ใช้
file_inventory = "All Inventory.xlsx"
file_spare = "Spare part Serial Item.xlsx"
output_file = "Comparison Result.xlsx"

# ✅ **1. อ่านข้อมูลจากไฟล์ Excel**
print("📌 Reading data from Excel files...")
df_inventory = pd.read_excel(file_inventory)
df_spare = pd.read_excel(file_spare)

# ✅ **2. ลบช่องว่างจากชื่อคอลัมน์ (ป้องกันปัญหาชื่อผิด)**
df_inventory.columns = df_inventory.columns.str.strip()
df_spare.columns = df_spare.columns.str.strip()

# ✅ **3. ตรวจสอบว่าคอลัมน์ที่ต้องใช้มีอยู่จริง**
required_columns_inventory = ["PartNumber", "Quantity", "ProjectCode", "CustomerID", "ProjectSaleState", "ProjectMAState"]
required_columns_spare = ["PartNumber", "SpareStatus", "SpareStore"]

# เช็คว่าคอลัมน์มีครบหรือไม่
for col in required_columns_inventory:
    if col not in df_inventory.columns:
        raise ValueError(f"❌ Missing column '{col}' in All Inventory.xlsx")

for col in required_columns_spare:
    if col not in df_spare.columns:
        raise ValueError(f"❌ Missing column '{col}' in Spare part Serial Item.xlsx")

# ✅ **4. รายการที่ต้องกรองออก (Exclusions)**
exclusions = [
    "Freeproduct", "Man-day", "Outdoor/Indoor", "FiberOptic", "Eflex", "PowerNYY", "PowerVCT", 
    "C13-C14", "MODE-NXOS", "MODE-ACI-LEAF", "MODE-ACI-SPINE", "NETWORK-PNP-LIC", "NXK-ACC-KIT-1RU"
]

# ✅ **5. กรองข้อมูลเฉพาะที่มีสถานะ "Warranty" หรือ "Maintenance"**
df_inventory_filtered = df_inventory[
    ((df_inventory["ProjectSaleState"] == "Warranty") | (df_inventory["ProjectMAState"] == "Maintenance"))
    & (~df_inventory["PartNumber"].astype(str).str.contains("|".join(exclusions), case=False, na=False))
]

# ✅ **6. สร้าง Dictionary เก็บค่าต่าง ๆ**
inventory_dict = defaultdict(float)   # เก็บจำนวนรวมของแต่ละ Part
project_dict = defaultdict(set)       # เก็บรหัสโครงการ
customer_dict = defaultdict(set)      # เก็บรหัสลูกค้า
spare_dict = defaultdict(int)         # เก็บจำนวนอะไหล่
spare_status_dict = defaultdict(lambda: defaultdict(int))  # เก็บสถานะอะไหล่
spare_store_dict = defaultdict(lambda: defaultdict(int))   # เก็บที่เก็บอะไหล่

# ✅ **7. ประมวลผลข้อมูลจาก All Inventory**
print("📌 Processing All Inventory data...")
for _, row in df_inventory_filtered.iterrows():
    part_number = re.sub(r"[ =]", "", str(row["PartNumber"]))  # ลบช่องว่างและเครื่องหมาย =
    inventory_dict[part_number] += row["Quantity"]
    project_dict[part_number].add(str(row["ProjectCode"]))
    customer_dict[part_number].add(str(row["CustomerID"]))

# ✅ **8. ประมวลผลข้อมูลจาก Spare part Serial Item**
print("📌 Processing Spare part Serial Item data...")
for _, row in df_spare.iterrows():
    part_number = re.sub(r"[ =]", "", str(row["PartNumber"]))
    spare_dict[part_number] += 1
    spare_status_dict[part_number][row["SpareStatus"]] += 1

    store = row["SpareStore"]
    if store in ["AIT-HQ", "AIT-MT"]:
        spare_store_dict[part_number][store] += 1
    else:
        spare_store_dict[part_number]["Other"] += 1

# ✅ **9. คำนวณคำแนะนำในการสั่งซื้ออะไหล่**
print("📌 Calculating spare part recommendations...")
recommendations = []
for part_number in inventory_dict:
    total_qty = inventory_dict[part_number]
    total_spares = spare_dict[part_number]

    # 🛠️ กำหนดคำแนะนำ
    if total_spares > 0:
        spare_recommend = "No need to order"
    elif total_qty <= 10:
        spare_recommend = "1 Spare"
    elif total_qty <= 50:
        spare_recommend = "1-2 Spares"
    elif total_qty <= 100:
        spare_recommend = "2-3 Spares"
    elif total_qty <= 200:
        spare_recommend = "3-4 Spares"
    else:
        spare_recommend = "5+ Spares"

    recommendations.append([
        part_number, 
        total_qty, 
        total_spares,
        ", ".join(f"{status} ({count})" for status, count in spare_status_dict[part_number].items()) or "N/A",
        ", ".join(f"{store} ({count})" for store, count in spare_store_dict[part_number].items()) or "N/A",
        spare_recommend,
        ", ".join(customer_dict[part_number]) or "N/A",
        ", ".join(project_dict[part_number]) or "N/A"
    ])

# ✅ **10. สร้าง DataFrame สำหรับผลลัพธ์**
df_results = pd.DataFrame(recommendations, columns=[
    "Part Number", "Total QTY", "Spare Count", "Spare Status", "Spare Store", 
    "Spare Order Recommendation", "Customer ID", "Project Refer"
])

# ✅ **11. บันทึกผลลัพธ์ลงไฟล์ Excel**
print("📌 Saving results to Comparison Result.xlsx...")
with pd.ExcelWriter(output_file) as writer:
    df_results.to_excel(writer, sheet_name="Comparison Result", index=False)

print("✅ Process completed! Data saved in 'Comparison Result.xlsx'")
