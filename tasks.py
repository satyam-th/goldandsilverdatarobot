from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Email.ImapSmtp import ImapSmtp
import sqlite3
import os
import json

browser = Selenium()


@task
def run_task():
    data = data_taker()
    for datas in data:
            print(datas)

    database(data)
    excel_file = create_excel_report()
    send_email_report(excel_file)
    print("\nAll tasks completed!")

#collecting data
def data_taker():
    browser.set_selenium_speed(0.0)
    browser.open_available_browser("https://www.fenegosida.org/")
    browser.maximize_browser_window()
    year = browser.find_element('xpath:/html/body/div[3]/div[1]/div[1]/div[3]').text
    month = browser.find_element('xpath:/html/body/div[3]/div[1]/div[1]/div[2]').text
    day = browser.find_element('xpath:/html/body/div[3]/div[1]/div[1]/div[1]').text
    tola_gold_price = browser.find_element('xpath:/html/body/div[3]/div[1]/div[2]/div/div[2]/div/div[1]/p/b').text
    tola_silver_price = browser.find_element('xpath:/html/body/div[3]/div[1]/div[2]/div/div[2]/div/div[3]/p/b').text
    print("day ", day, " month ", month, " year ", year)
    print("Tola Gold Price:", tola_gold_price, "and  silver ", tola_silver_price)
    browser.mouse_over('xpath:/html/body/div[3]/div[1]/div[2]/div/ul/li[1]')
    gram_gold_price = browser.find_element('xpath:/html/body/div[3]/div[1]/div[2]/div/div[1]/div/div[1]/p/b').text
    gram_silver_price = browser.find_element('xpath:/html/body/div[3]/div[1]/div[2]/div/div[1]/div/div[3]/p/b').text
    print("10 gram Gold Price:", gram_gold_price, "and  silver ", gram_silver_price)
    per_gram_gold_price = int(gram_gold_price) / 10 
    per_gram_Silver_price =int(gram_silver_price) / 10
    print("gram Gold Price:", per_gram_gold_price, "and  silver ", per_gram_Silver_price)
    data = {
        "year": year,
        "month": month,
        "day": day,
        "tola_gold": float(tola_gold_price),
        "tola_silver": float(tola_silver_price),
        "per_gram_gold": float(per_gram_gold_price),
        "per_gram_silver": float(per_gram_Silver_price)
        }
    browser.close_browser()
    return data


def database(data):
    """
    Store scraped data into SQLite database
    """
    # 1. Connect to database
    conn = sqlite3.connect('gold_silver_prices.db')
    cursor = conn.cursor()
    
    # 2. Create table if it doesn't exist
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS prices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            year TEXT,
            month TEXT,
            day TEXT,
            tola_gold REAL,
            tola_silver REAL,
            per_gram_gold REAL,
            per_gram_silver REAL,
            gold_change_pct REAL,
            silver_change_pct REAL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # ADDED: Check and add missing columns
    cursor.execute("PRAGMA table_info(prices)")
    columns = [column[1] for column in cursor.fetchall()]
    
    if 'gold_change_pct' not in columns:
        print(" Adding gold_change_pct column...")
        cursor.execute('ALTER TABLE prices ADD COLUMN gold_change_pct REAL')
    
    if 'silver_change_pct' not in columns:
        print(" Adding silver_change_pct column...")
        cursor.execute('ALTER TABLE prices ADD COLUMN silver_change_pct REAL')
    
    # Calculate percentage change
    gold_change, silver_change = calculate_percentage_change(cursor, data)
    
    # 3. Insert data
    cursor.execute('''
        INSERT INTO prices (year, month, day, tola_gold, tola_silver, 
                           per_gram_gold, per_gram_silver, gold_change_pct, silver_change_pct)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        data['year'],
        data['month'],
        data['day'],
        data['tola_gold'],
        data['tola_silver'],
        data['per_gram_gold'],
        data['per_gram_silver'],
        gold_change,
        silver_change
    ))
    
    # 4. Commit and close
    conn.commit()
    print(f"Data saved successfully! Record ID: {cursor.lastrowid}")
    conn.close()


def calculate_percentage_change(cursor, data):
    """
    Calculate percentage change from previous day
    """
    # Get previous day's data
    cursor.execute('''
        SELECT tola_gold, tola_silver
        FROM prices
        ORDER BY id DESC
        LIMIT 1
    ''')
    
    previous = cursor.fetchone()
    
    if previous:
        prev_gold = previous[0]
        prev_silver = previous[1]
        
        # Calculate percentage change
        gold_change = ((data['tola_gold'] - prev_gold) / prev_gold) * 100
        silver_change = ((data['tola_silver'] - prev_silver) / prev_silver) * 100  
        print(f"Gold Change: {gold_change:+.2f}%")
        print(f"Silver Change: {silver_change:+.2f}%")
        
        return round(gold_change, 2), round(silver_change, 2)
    else:
        print("ℹ First record - no percentage change")
        return None, None


def create_excel_report():
    """
    Generate Excel report with prices and percentage changes
    """
    # Connect to database
    conn = sqlite3.connect('gold_silver_prices.db')
    cursor = conn.cursor()
    
    # Get all records
    cursor.execute('''
        SELECT day, month, year, tola_gold, tola_silver, gold_change_pct, silver_change_pct
        FROM prices
        ORDER BY id DESC
        LIMIT 30
    ''')
    
    records = cursor.fetchall()
    conn.close()
    
    if not records:
        print(" No data available for Excel report")
        return None
    
    # Create Excel file
    excel = Files()
    
    # Prepare data for Excel
    excel_data = []
    
    # Add header row
    header = ["Date", "Gold (per tola)", "Silver (per tola)", "Gold Change (%)", "Silver Change (%)"]
    excel_data.append(header)
    
    # Add data rows
    for record in records:
        day = record[0]
        month = record[1]
        year = record[2]
        tola_gold = record[3]
        tola_silver = record[4]
        gold_change = record[5]
        silver_change = record[6]
        
        # Format date
        date_str = f"{day} {month} {year}"
        
        # Format prices
        gold_str = f"{tola_gold:,.0f}"
        silver_str = f"{tola_silver:,.0f}"
        
        # Format percentage changes
        if gold_change is not None:
            gold_change_str = f"{gold_change:+.2f}%"
        else:
            gold_change_str = "-"
        
        if silver_change is not None:
            silver_change_str = f"{silver_change:+.2f}%"
        else:
            silver_change_str = "-"
        
        # Add row
        row = [date_str, gold_str, silver_str, gold_change_str, silver_change_str]
        excel_data.append(row)
    
    # Create filename
    latest_day = records[0][0]
    latest_month = records[0][1]
    latest_year = records[0][2]
    filename = f"gold_silver_prices_{latest_day}{latest_month}{latest_year}.xlsx"
    
    # Create workbook and add data
    excel.create_workbook(filename)
    excel.append_rows_to_worksheet(excel_data, header=False)
    excel.save_workbook()
    excel.close_workbook()
    
    print(f"Excel report created: {filename}")
    return filename

def send_email_report(excel_file):
    if not excel_file:
        print("No Excel file to send")
        return

    try:
        secrets = load_email_credentials()  # load manually

        mail = ImapSmtp()
        mail.authorize_smtp(
            account=secrets["username"],
            password=secrets["password"],
            smtp_server=secrets["smtp_server"],
            smtp_port=secrets["smtp_port"]
        )

        recipients = [r.strip() for r in secrets.get("recipients", "").split(",") if r.strip()]

        mail.send_message(
            sender=secrets["username"],
            recipients=recipients,
            subject="Daily Gold & Silver Prices",
            body="This is a test email from Robocorp using App Password.",
            html=False,
            attachments=excel_file
        )

        print(f" Email sent successfully to: {', '.join(recipients)}")

    except Exception as e:
        print(f" Email sending failed: {e}")


def view_all_records():
    """
    View all database records
    """
    # Connect to database
    conn = sqlite3.connect('gold_silver_prices.db')
    cursor = conn.cursor()
    
    # Get all records
    cursor.execute('''
        SELECT day, month, year, tola_gold, tola_silver, gold_change_pct, silver_change_pct
        FROM prices
        ORDER BY id DESC
    ''')
    
    records = cursor.fetchall()
    
    # Print header
    print("\n" + "="*90)
    print(f"{'Date':<20} {'Gold/Tola':>15} {'Silver/Tola':>15} {'Gold %':>15} {'Silver %':>15}")
    print("="*90)
    
    # Print records
    for record in records:
        day = record[0]
        month = record[1]
        year = record[2]
        gold = record[3]
        silver = record[4]
        gold_change = record[5]
        silver_change = record[6]
        
        date_str = f"{day} {month} {year}"
        
        if gold_change is not None:
            gold_change_str = f"{gold_change:+.2f}%"
        else:
            gold_change_str = "-"
        
        if silver_change is not None:
            silver_change_str = f"{silver_change:+.2f}%"
        else:
            silver_change_str = "-"
        
        print(f"{date_str:<20} {gold:>15,.0f} {silver:>15,.0f} {gold_change_str:>15} {silver_change_str:>15}")
    
    print("="*90 + "\n")
    
    conn.close()



def load_email_credentials():
    vault_path = os.path.join(os.getcwd(), "vault.json")
    with open(vault_path, "r") as f:
        secrets = json.load(f)
    return secrets["email_credentials"]
