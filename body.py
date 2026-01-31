import py_win_keyboard_layout

import os
import shutil

import time
import keyboard

import win32com.client
import win32gui

import xml.etree.ElementTree as ET
from xml.dom import minidom

import csv

def main(n_copy, path) -> None:
    #n_copy=3
    #path=r'C:\Users\Pavel\Desktop\elevate_in\TEST\K21'

    file = os.listdir(path)
    n_copy = int(n_copy)

    for i in range(2, n_copy + 1):
        if i < 10:
            shutil.copyfile(path + '\\' + file[0], path + '\\' + file[0][:-6] + str(i) + '.elvx')
        else:
            shutil.copyfile(path + '\\' + file[0], path + '\\' + file[0][:-7] + str(i) + '.elvx')
       
        file = os.listdir(path)
        modify_handling_capacity(path + '\\' + file[i-1], i)
    
    get_area(path)

    os.startfile(r'C:\Program Files (x86)\Elevate 9\Elevate.exe')
    time.sleep(3.0)
    py_win_keyboard_layout.change_foreground_window_keyboard_layout(0x04090409)
    keyboard.press_and_release('alt + a')
    keyboard.press_and_release('down arrow')
    keyboard.press_and_release('down arrow')
    keyboard.press_and_release('down arrow')
    keyboard.press_and_release('down arrow')
    keyboard.press_and_release('down arrow')
    keyboard.press_and_release('enter')
    time.sleep(0.5)
    keyboard.press_and_release('tab')
    keyboard.press_and_release('tab')
    keyboard.press_and_release('tab')
    time.sleep(0.5)
    keyboard.press_and_release('ctrl + v')
    keyboard.press_and_release('enter')

def print_repot(path) -> None:
        
    excel_app = win32com.client.Dispatch("Excel.Application")
    #excel_app.WindowState = -4137
    excel_app.Visible = True

    workbook = excel_app.Workbooks.Open(path + '\\' + 'batch_results.csv')
    win32gui.SetForegroundWindow(excel_app.Hwnd) 

    workbook.RefreshAll()
    excel_app.CalculateUntilAsyncQueriesDone()

    keyboard.press_and_release('alt + c')
    keyboard.press_and_release('left arrow')
    keyboard.press_and_release('tab')
    keyboard.press_and_release('enter')
    
    keyboard.press_and_release('left arrow')
    keyboard.press_and_release('enter')

def modify_handling_capacity(xml_file, new_capacity):

     # Parse the XML file
    tree = ET.parse(xml_file)
    root = tree.getroot()
    
    # Find and modify HandlingCapacity in PassengerData/Standard
    data_standard = root.find('.//PassengerData/Standard')
    if data_standard is not None:
        handling_capacity = data_standard.find('HandlingCapacity')
        if handling_capacity is not None:
            handling_capacity.text = str(new_capacity)
    
     # Find and modify TotalArrivalRate in PassengerData/Traffic
    data_traffic = root.find('.//PassengerData/Traffic')
    if data_traffic is not None:
        data_periods = data_traffic.findall('Period')
        for data_period in data_periods:
            if data_period.get('Id') == '0':
                data_period.set('TotalArrivalRate', str(new_capacity))

  # Save the modified XML with pretty formatting
    xml_str = ET.tostring(root, encoding='utf-8')
    pretty_xml = minidom.parseString(xml_str).toprettyxml(indent="  ")
    
    with open(xml_file, 'w', encoding='utf-8') as f:
        f.write(pretty_xml)

def get_area(path) -> None:

    xml_file = path + '\\' + os.listdir(path)[0]
    csv_file = os.path.join(path, 'floor_area.csv')

    tree = ET.parse(xml_file)
    root = tree.getroot()

    data = []

    # Find FloorArea in ElevatorData/Advanced/Configuration/Car
    data_cars = root.findall('.//ElevatorData/Advanced/Configuration/Car')
    if data_cars is not None:
        for data_car in data_cars:
            car_id = data_car.get('Id')
            floor_area = data_car.get('FloorAreaM2')

            data.append({
                'CarId' : car_id,
                'FloorAreaM2' : floor_area
            })

    # Write CSV
    if data:
        fieldnames = ['CarId', 'FloorAreaM2']

        with open(csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=';')
            writer.writeheader()
            writer.writerows(data)

if __name__ == '__main__':
    main()