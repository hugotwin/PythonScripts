import tkinter as tk
import serial
import struct
import Adafruit_DHT
import time
import datetime
import RPi.GPIO as GPIO
from gpiozero import Button
import threading

# Initialize 'now' before entering the loop
running = False
pin_rele = None  # Initialize pin_rele as None
program_state = "A iniciar programa"  # Initial program state
valor2 = ""  # Initialize valor2 as an empty string
temp_amb = " A RECOLHER TEMPERATURA ..."
diferenca_temp = ""
button = Button(21)  # Create the Button object outside of the loop
n = 0
ligado = 0


def update_state_label():
    state_label.config(text=program_state)
    state_label2.config(text=valor2)

    state_label3.config(text=temp_amb)
    state_label4.config(text=diferenca_temp)
    window.after(1000, update_state_label)


def start_program():
    global running, pin_rele, ligado
    if ligado == 0:
        ligado = 1
        running = True
        program_thread = threading.Thread(target=run_program)
        program_thread.start()
        time.sleep(0.5)


def stop_program():
    global running, pin_rele, program_state, ligado
    ligado = 0
    running = False
    pin_rele = 4
    GPIO.setmode(GPIO.BCM)
    GPIO.setup(pin_rele, GPIO.OUT)
    time.sleep(0.1)
    GPIO.output(pin_rele, GPIO.HIGH)
    time.sleep(0.3)
    GPIO.output(pin_rele, GPIO.HIGH)
    GPIO.output(pin_rele, GPIO.HIGH)
    program_state = "SISTEMA PARADO"
    if pin_rele is not None:
        pin_rele = 4
        GPIO.setmode(GPIO.BCM)
        GPIO.setup(pin_rele, GPIO.OUT)
        GPIO.output(pin_rele, GPIO.HIGH)
        program_state = "SISTEMA PARADO"


def run_program():
    global now2, now, n, rele, program_state, valor2, temp_amb, diferenca_temp
    now = datetime.datetime.now()
    now2 = datetime.datetime.now()

    rele = 0

    while running:
        # Create a serial connection for the Calex sensor
        ser_calex = serial.Serial('/dev/ttyUSB0', 9600, timeout=0.2)

        # Define the Modbus RTU request parameters for the Calex sensor
        slave_id = 1
        function_code = 3
        starting_address = 11
        number_of_registers = 1
        crc = 0xFFFF

        # Read data from the Calex sensor continuously
        # Pack the Modbus RTU request into a bytes object
        request = struct.pack('>BBHH', slave_id, function_code, starting_address, number_of_registers)
        crc = crc & 0xFFFF
        for b in request:
            crc = crc ^ b
            for i in range(8):
                if crc & 0x0001:
                    crc = (crc >> 1) ^ 0xA001
                else:
                    crc = crc >> 1
        request += struct.pack('<H', crc)

        # Send the request to the sensor and read the response
        ser_calex.write(request)
        response = ser_calex.read(5 + 2 * number_of_registers)

        # Check the length of the response
        if len(response) < 5 + 2 * number_of_registers:
            print('Error: incomplete Modbus response')
        else:
            # Unpack the response and print the register value
            if response[1] != function_code:
                print('Error: incorrect function code')
            elif response[2] != 2 * number_of_registers:
                print('Error: incorrect number of bytes')
            else:
                registers = struct.unpack('>' + 'H' * number_of_registers, response[3:-2])
                valor = registers[0] / 10
                valor2 = (f' A TEMPERATURA DO GARNULADO: {valor:.1f}  ')  # Corrected variable name
                time.sleep(0.1)

        # Check if 60 seconds have passed to read data from the DHT11 sensor
        if (button.is_pressed and ((datetime.datetime.now() - now2)).total_seconds() > 300):
            if (datetime.datetime.now() - now).total_seconds() >= 10 or n == 0:
                global temperature

                # Read data from the DHT11 sensor
                humidity, temperature = Adafruit_DHT.read_retry(11, 14)
                temperature = temperature - 5
                now = datetime.datetime.now()
                diferenca = abs(valor - temperature)
                diferenca_temp = (f' O DIFERENCIAL DE  TEMPERATURA: {diferenca:.1f}')
                print(f' foi ao medidor de temperatura ambiente {diferenca:.1f}')
                n = 1
                time.sleep(0.1)
                temp_amb = (f' TEMPERATURA AMBIENTE: {temperature:.1f}  ')
            else:
                diferenca = abs(valor - temperature)
                diferenca_temp = (f' O DIFERENCIAL DE  TEMPERATURA: {diferenca:.1f}  ')
                print(f'{diferenca:.1f}')

            # temperatura ambiente de 0  a 18 

            if diferenca <= 7 and temperature <= 18 and rele == 0 and button.is_pressed:
                print(f'{diferenca:.1f}')
                pin_rele = 4
                GPIO.setmode(GPIO.BCM)
                GPIO.setup(pin_rele, GPIO.OUT)
                # Activate the relay (close the contact)
                GPIO.output(pin_rele, GPIO.LOW)
                program_state = "ARREFECEDOR ABERTO"
                time.sleep(2.5)
                GPIO.output(pin_rele, GPIO.HIGH)
                program_state = "ARREFECEDOR FECHADO"
                time.sleep(6)

            elif diferenca < 6 and temperature <= 18 and button.is_pressed:
                print(f'{diferenca:.1f}')
                pin_rele = 4
                GPIO.setmode(GPIO.BCM)
                GPIO.setup(pin_rele, GPIO.OUT)
                # Activate the relay (close the contact)
                GPIO.output(pin_rele, GPIO.LOW)
                program_state = "ARREFECEDOR ABERTO"
                time.sleep(2.5)
                GPIO.output(pin_rele, GPIO.HIGH)
                program_state = "ARREFECEDOR FECHADO"
                time.sleep(6)

            elif diferenca > 6 or temperature >= 30:

                # Deactivate the relay (open the contact)
                pin_rele = 4
                GPIO.setmode(GPIO.BCM)
                GPIO.setup(pin_rele, GPIO.OUT)
                GPIO.output(pin_rele, GPIO.HIGH)
                program_state = " EM ARREFECIMENTO  "
                time.sleep(0.5)
                rele = 0
            # temperatura ambiente de  18-20

            elif diferenca <= 6 and temperature <= 20 and rele == 0 and button.is_pressed:
                print(f'{diferenca:.1f}')
                pin_rele = 4
                GPIO.setmode(GPIO.BCM)
                GPIO.setup(pin_rele, GPIO.OUT)
                # Activate the relay (close the contact)
                GPIO.output(pin_rele, GPIO.LOW)
                program_state = "ARREFECEDOR ABERTO"
                time.sleep(2.5)
                GPIO.output(pin_rele, GPIO.HIGH)
                program_state = "ARREFECEDOR FECHADO"
                time.sleep(6)

            elif diferenca <= 6 and temperature <= 20 and button.is_pressed:
                print(f'{diferenca:.1f}')
                pin_rele = 4
                GPIO.setmode(GPIO.BCM)
                GPIO.setup(pin_rele, GPIO.OUT)
                # Activate the relay (close the contact)
                GPIO.output(pin_rele, GPIO.LOW)
                program_state = "ARREFECEDOR ABERTO"
                time.sleep(2.5)
                GPIO.output(pin_rele, GPIO.HIGH)
                program_state = "ARREFECEDOR FECHADO"
                time.sleep(6)

            elif diferenca > 6 and temperature <= 20:
                pin_rele = 4
                GPIO.setmode(GPIO.BCM)
                GPIO.setup(pin_rele, GPIO.OUT)
                GPIO.output(pin_rele, GPIO.HIGH)
                program_state = " EM ARREFECIMENTO.  "

                rele = 0



            # temperatura ambiente de  20-22

            elif diferenca < 5 and temperature <= 23 and rele == 0 and button.is_pressed:
                print(f'{diferenca:.1f}')
                pin_rele = 4
                GPIO.setmode(GPIO.BCM)
                GPIO.setup(pin_rele, GPIO.OUT)
                # Activate the relay (close the contact)
                GPIO.output(pin_rele, GPIO.LOW)
                program_state = "ARREFECEDOR ABERTO"
                time.sleep(2.5)
                GPIO.output(pin_rele, GPIO.HIGH)
                program_state = "ARREFECEDOR FECHADO"
                time.sleep(6)

            elif diferenca < 5 and temperature <= 23 and button.is_pressed:
                print(f'{diferenca:.1f}')
                pin_rele = 4
                GPIO.setmode(GPIO.BCM)
                GPIO.setup(pin_rele, GPIO.OUT)
                # Activate the relay (close the contact)
                GPIO.output(pin_rele, GPIO.LOW)
                program_state = "ARREFECEDOR ABERTO"
                time.sleep(2.5)
                GPIO.output(pin_rele, GPIO.HIGH)
                program_state = "ARREFECEDOR FECHADO"
                time.sleep(6)
                              
                rele = 1

            elif diferenca >= 5 and temperature <= 23:
                # Deactivate the relay (open the contact)
                pin_rele = 4
                GPIO.setmode(GPIO.BCM)
                GPIO.setup(pin_rele, GPIO.OUT)
                GPIO.output(pin_rele, GPIO.HIGH)
                program_state = " EM ARREFECIMENTO  "
                time.sleep(0.1)
                rele = 0

            # temperatura ambiente acima  22 -26 

            elif diferenca <= 3.5 and temperature <= 26 and rele == 0 and button.is_pressed:
                pin_rele = 4
                GPIO.setmode(GPIO.BCM)
                GPIO.setup(pin_rele, GPIO.OUT)
                # Activate the relay (close the contact)
                
                GPIO.output(pin_rele, GPIO.LOW)
                program_state = "ARREFECEDOR ABERTO"
                print("Relay activated. The device is on.")
                time.sleep(0.1)
                print(temperature)
                rele = 1
                

            elif diferenca <= 3.5 and temperature <= 26 and button.is_pressed:
                print(f'{diferenca:.1f}')
                pin_rele = 4
                GPIO.setmode(GPIO.BCM)
                GPIO.setup(pin_rele, GPIO.OUT)
                # Activate the relay (close the contact)

                GPIO.output(pin_rele, GPIO.LOW)
                program_state = "ARREFECEDOR ABERTO"
                time.sleep(2.5)
                GPIO.output(pin_rele, GPIO.HIGH)
                program_state = "ARREFECEDOR FECHADO"
                time.sleep(6)
                
                rele = 1

            elif diferenca > 3.5 and temperature <= 26:
                # Deactivate the relay (open the contact)
                pin_rele = 4
                GPIO.setmode(GPIO.BCM)
                GPIO.setup(pin_rele, GPIO.OUT)
                GPIO.output(pin_rele, GPIO.HIGH)
                program_state = " EM ARREFECIMENTO ... "
                time.sleep(0.1)
                rele = 0



        elif not button.is_pressed:

            time.sleep(0.1)
            now2 = datetime.datetime.now()
            program_state = " A ESPERA DE MATERIAL - ARREFECEDOR FECHADO "
            pin_rele = 4
            GPIO.setmode(GPIO.BCM)
            GPIO.setup(pin_rele, GPIO.OUT)
            GPIO.output(pin_rele, GPIO.HIGH)

        elif ((datetime.datetime.now() - now2)).total_seconds() < 5:
            program_state = "  ARREFECEDOR A ENCHER - ARREFECEDOR FECHADO "


# Create the Tkinter window
window = tk.Tk()
window.title("Program Control")
window.geometry("400x400")
window.attributes("-zoomed", True)

# Create buttons
start_button = tk.Button(window, text="Start Program", command=start_program, padx=10)
stop_button = tk.Button(window, text="Stop Program", command=stop_program, padx=10)

# Create a Label to display the state
state_label = tk.Label(window, text=program_state, pady=10, padx=10)
state_label2 = tk.Label(window, text=valor2, pady=10, padx=10)
state_label3 = tk.Label(window, text=temp_amb, pady=10, padx=10)
state_label4 = tk.Label(window, text=diferenca_temp, pady=10, padx=10)
# Pack buttons and labels
start_button.pack()
stop_button.pack()
state_label.pack()
state_label2.pack()
state_label3.pack()
state_label4.pack()
# Start updating the state labels
update_state_label()

# Start the Tkinter main loop
window.mainloop()
