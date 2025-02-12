import pyodbc
import pandas as pd

# ✅ ตรวจสอบ ODBC Driver ที่ติดตั้ง
drivers = pyodbc.drivers()
print("📌 Available ODBC Drivers:", drivers)

# ✅ ตรวจสอบว่ามี Driver สำหรับ SQL Server หรือไม่
odbc_driver = None
for driver in drivers:
    if "ODBC Driver 17 for SQL Server" in driver:
        odbc_driver = driver
        break

if not odbc_driver:
    raise ValueError("❌ No valid ODBC Driver for SQL Server found! Please install 'ODBC Driver 17 for SQL Server'.")

# ✅ กำหนดค่าการเชื่อมต่อ MSSQL
# server = "your_server"
# database = "your_database"
# username = "your_username"
# password = "your_password"

server = "10.0.0.58"
database = "AIT_Inventory"
username = "cms"
password = "cms372!"

# ✅ สร้าง Connection String ที่ถูกต้อง
conn_str = f"DRIVER={{{odbc_driver}}};SERVER={server};DATABASE={database};UID={username};PWD={password}"

# ✅ เชื่อมต่อกับ MSSQL
print("📌 Connecting to SQL Server...")
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()
print("✅ Connected successfully!")

# ✅ สร้าง SQL Query เพื่อดึงข้อมูลจากหลาย Table
sql_query = """
SELECT 
    inv.PartNo, inv.Qty, inv.ProjectCode, inv.CustomerCode, 
    inv.ProjectSaleState, inv.ProjectMAState, 
    spare.SpareStatus, spare.SpareStore
FROM dbo.inv_vw_Customer_Item inv
LEFT JOIN dbo.csm_vw_SparePart spare ON inv.PartNo = spare.PartNumber
"""

# ✅ ดึงข้อมูลจากฐานข้อมูล
df = pd.read_sql(sql_query, conn)

# ✅ บันทึกข้อมูลลงไฟล์ combine_data.xlsx
output_file = "combine_data.xlsx"
df.to_excel(output_file, index=False)

# ✅ ปิดการเชื่อมต่อ
cursor.close()
conn.close()
print(f"✅ Data successfully saved to {output_file}!")
