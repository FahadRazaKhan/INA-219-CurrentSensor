# INA219 Current sensor program
# @author: FahadRaza
'''
This Program read values from AdaFruit INA219 current sensor using 'ina219' module.
It also stores Current, Voltage , and Power values in an XLSX File.
Sensor is connected to I2C pins of Raspberry Pi.

ina219 module can be installed on Raspbian/Linux system by

sudo pip install pi-ina219

To check the Insatlled version of pi-ina219

pip show pi-ina219

'''
from ina219 import INA219
import time
import xlsxwriter
from collections import deque



workbook = xlsxwriter.Workbook('SensorValues.xlsx',{'constant_memory': True})  # Creating XLSX File for Data Keeping 
worksheet = workbook.add_worksheet() # Generating worksheet

bold = workbook.add_format({'bold':True}) # Formating for Bold text

worksheet.write('A1', 'Time', bold)
worksheet.write('B1', 'Current (mA)', bold)
worksheet.write('C1', 'Voltage (v)', bold)
worksheet.write('D1', 'Power (mW)', bold)

row = 1 # Starting Row (0 indexed)
col = 0 # Starting Column (0 indexed) 


DataPoints = deque(maxlen=None) # Creating Array of datatype deque to store values

Shunt_OHMS = 0.1 # For this sensor is 0.1 ohm
Max_Expected_Amps = 0.3 # must be close to expected value


def CurrentRead():
    ina = INA219(Shunt_OHMS, Max_Expected_Amps)
    ina.configure(ina.RANGE_16V)
    #print('Bus Voltage: %.3f V' % ina.voltage())
    print('Bus Current: %.3f mA' % ina.current())
    #print('Power: %.3f mW' % ina.power())
    currentvalue = round(ina.current()) # Rounding off values to nearest integer
    voltagevalue = round(ina.voltage()) # //
    powervalue = round(ina.power()) # //
    timevalue = float('{0:.1f}'.format(time.time()-start)) # Elapsed time in Seconds with 1 decimal point floating number 

    DataPoints.append([timevalue, currentvalue, voltagevalue, powervalue]) # Updating DataPoints Array

def main():
    
    if __name__ == '__main__':
        CurrentRead()



try:
    
    print('Starting Current Sensor')
    print('Collecting Sensor Values...')
    start = time.time() # Start Time
    Iterations = 0

    while True:

        if Iterations >= 50: # Number of Iterations
            break
        
        main() # Calling main function

        time.sleep(0.5) # Reading value after half second
        Iterations += 1 # Updating Iteration number
        

    for Time, value1, value2, value3 in (DataPoints):   # Writing Data in XLSX file
        worksheet.write(row, col, Time)
        worksheet.write(row, col+1, value1)
        worksheet.write(row, col+2, value2)
        worksheet.write(row, col+3, value3)
        row += 1

    chart1 = workbook.add_chart({'type': 'line'}) # adding chart of type 'Line' for Current values
    chart2 = workbook.add_chart({'type': 'line'}) # Chart for Voltage
    chart3 = workbook.add_chart({'type': 'line'}) # Chart for Power

    n = len(DataPoints) # Total number of rows
    
    chart1.add_series({'name':['Sheet1',0,1],
                      'categories': ['Sheet1', 1,0,n,0],
                      'values': ['Sheet1', 1,1,n,1]
                      })
    chart2.add_series({'name':['Sheet1',0,2],
                      'categories': ['Sheet1', 1,0,n,0],
                      'values': ['Sheet1', 1,2,n,2]
                      })
    chart3.add_series({'name':['Sheet1',0,3],
                      'categories': ['Sheet1', 1,0,n,0],
                      'values': ['Sheet1', 1,3,n,3]
                      })
    
    chart1.set_title({'name': 'Current Chart'}) # Setting Title name
    chart1.set_x_axis({'name': 'Elapsed Time (s)'}) # Setting X-Axis name
    chart1.set_y_axis({'name': 'Value'}) # Setting Y-Axis name

    chart2.set_title({'name': 'Voltage Chart'})
    chart2.set_x_axis({'name': 'Elapsed Time (s)'})
    chart2.set_y_axis({'name': 'Value'})

    chart3.set_title({'name': 'Power Chart'})
    chart3.set_x_axis({'name': 'Elapsed Time (s)'})
    chart3.set_y_axis({'name': 'Value'})


    chart1.set_style(8) # Setting Chart Color
    chart2.set_style(5)
    chart2.set_style(9)

    worksheet.insert_chart('D2', chart1, {'x_offset': 25, 'y_offset': 10}) # Inserting Charts in the Worksheet
    worksheet.insert_chart('D2', chart2, {'x_offset': 25, 'y_offset': 10}) # //
    worksheet.insert_chart('D5', chart3, {'x_offset': 25, 'y_offset': 10}) # //


except:

    print('An Error Occured, Please try again!')


workbook.close() # Closing Workbook 
time.sleep(1)
print('Operation Complete')
