import os
import tempfile
import pandas as pd
import psutil
import time
import subprocess
import pyautogui
from flask import Flask, render_template
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import LineChart, Reference
from flask import jsonify
import random
app = Flask(__name__)
temp_files_info = []
cpu_usage = []
memory_total = []
memory_used = []
memory_percent = []
disk_total = []
disk_used = []
disk_percent = []


def get_temp_files_info():
    temp_files_info = []
    for file_name in os.listdir(tempfile.gettempdir()):
        file_path = os.path.join(tempfile.gettempdir(), file_name)
        if os.path.isfile(file_path):
            file_size = os.path.getsize(file_path)
            temp_files_info.append({"Name": file_name, "Size": file_size, "Location": file_path})
    return temp_files_info


def check_cpu_usage():
    cpu_usage_value = psutil.cpu_percent(interval=1)
    cpu_usage.append(cpu_usage_value)


def check_memory_usage():
    memory_info = psutil.virtual_memory()
    total = f"{memory_info.total / (1024 ** 3):.2f} GB"
    memory_total.append(total)

    used = f"{memory_info.used / (1024 ** 3):.2f} GB"
    memory_used.append(used)

    percent = memory_info.percent
    memory_percent.append(percent)


def check_disk_space():
    disk_usage = psutil.disk_usage('/')
    total = f"{disk_usage.total / (1024 ** 3):.2f} GB"
    disk_total.append(total)

    used = f"{disk_usage.used / (1024 ** 3):.2f} GB"
    disk_used.append(used)

    percent = disk_usage.percent
    disk_percent.append(percent)


def save_to_excel():
    df_main = pd.DataFrame({
        "CPU Usage": cpu_usage,
        "Memory Total": memory_total,
        "Memory Used": memory_used,
        "Memory Percent": memory_percent,
        "Disk Total": disk_total,
        "Disk Used": disk_used,
        "Disk Percent": disk_percent
    })
    df_temp = pd.DataFrame(get_temp_files_info())
    df_main.to_csv('main.csv', index=False)
    # Rest of the code for saving to Excel remains unchanged
    df_temp.to_csv('temporary_files.csv')
    df_main = pd.read_csv('main.csv')
    df_temp = pd.read_csv('temporary_files.csv')

    # Create a new Excel workbook
    wb = openpyxl.Workbook()

    # Create main data sheet and add data
    ws_main = wb.active
    ws_main.title = 'System_Health'
    for row in dataframe_to_rows(df_main, index=False, header=True):
        ws_main.append(row)

    # Create temporary files sheet and add data
    ws_temp = wb.create_sheet(title='Temporary_Files')
    for row in dataframe_to_rows(df_temp, index=False, header=True):
        ws_temp.append(row)

    # Create line chart in System_Health sheet
    chart_main = LineChart()
    data_main = Reference(ws_main, min_col=2, min_row=1, max_col=df_main.shape[1], max_row=df_main.shape[0] + 1)
    categories_main = Reference(ws_main, min_col=1, min_row=2, max_row=df_main.shape[0] + 1)
    chart_main.add_data(data_main, titles_from_data=True)
    chart_main.set_categories(categories_main)
    ws_main.add_chart(chart_main, 'D2')

    # Create line chart in Temporary_Files sheet
    chart_temp = LineChart()
    data_temp = Reference(ws_temp, min_col=2, min_row=1, max_col=df_temp.shape[1], max_row=df_temp.shape[0] + 1)
    categories_temp = Reference(ws_temp, min_col=1, min_row=2, max_row=df_temp.shape[0] + 1)
    chart_temp.add_data(data_temp, titles_from_data=True)
    chart_temp.set_categories(categories_temp)
    ws_temp.add_chart(chart_temp, 'D2')

    # Save the workbook
    wb.save('output.xlsx')

    subprocess.run(["C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE", 'output.xlsx'], check=True)

    time.sleep(5)

@app.route('/')
def home():
    return render_template('index.html', data=combine_data())


def combine_data():
    data = {
        "CPU Usage": cpu_usage,
        "Memory Total": memory_total,
        "Memory Used": memory_used,
        "Memory Percent": memory_percent,
        "Disk Total": disk_total,
        "Disk Used": disk_used,
        "Disk Percent": disk_percent
    }

    data["Temporary Files"] = get_temp_files_info()
    generated_cpu_value = random.uniform(25, 40)  # Adjust the range as needed
    data["Generated CPU Usage"] = f"{generated_cpu_value:.2f}%"
    generated_memory_value = random.uniform(25, 70)  # Adjust the range as needed
    data["Generated Memory Usage"] = f"{generated_memory_value:.2f}%"
    generated_disk_value = random.uniform(9, 60)  # Adjust the range as needed
    data["Generated Disk Usage"] = f"{generated_disk_value:.2f}%"
    return data


def run_excel_macro():
    excel_file_path = os.path.abspath('output.xlsx')

    subprocess.Popen(["start", "excel", excel_file_path], check=True)
    time.sleep(5)

    pyautogui.hotkey('Alt', 'F8')
    time.sleep(1)
    pyautogui.write('RefreshCharts')
    pyautogui.press('Enter')
   
     
@app.route('/chart-data')
def chart_data():
    # Sample chart data (replace with your actual data)
    chart_info = {
        'cpuChart': {
            'labels': [1, 2, 3, 4, 5],
            'values': [10, 20, 15, 25, 30],
            'label': 'CPU Usage',
            'borderColor': '#e74c3c',
            'backgroundColor': '#e74c3c',
        },
        'memoryChart': {
            'labels': [1, 2, 3, 4, 5],
            'values': [5, 10, 8, 12, 15],
            'label': 'Memory Usage',
            'borderColor': '#2ecc71',
            'backgroundColor': '#2ecc71',
        },
        'diskChart': {
            'labels': [1, 2, 3, 4, 5],
            'values': [15, 12, 18, 22, 20],
            'label': 'Disk Usage',
            'borderColor': '#3498db',
            'backgroundColor': '#3498db',
        },
    }

    return jsonify(chart_info)

def main():
    print("Checking System Health:")
    for _ in range(6):  # Run for 120 seconds
        check_cpu_usage()
        check_memory_usage()
        check_disk_space()

        get_temp_files_info()
        print("Done")

        time.sleep(1)  # Sleep for 1 second between each check
    print("System health checks completed")
    try:
      save_to_excel()
      print("save_to_excel completed")
    except Exception as e:
      print(f"Error in save_to_excel: {e}")
    time.sleep(5)  # Add a delay of 10 seconds (adjust as needed)

    try:
      run_excel_macro()
      print("run_excel_macro completed")
    except Exception as e:
      print(f"Error in run_excel_macro: {e}")

    print("Main function completed")


main()
print("Flask app running on port 7777")
app.run(debug=True, port=7777)









