#! python3
import os
import win32com.client
import fitz  # PyMuPDF
from PIL import Image
from escpos.printer import Usb
from time import sleep
from typing import Tuple, List
import winreg
import usb.core
import win32print
#import pyusb



"""
=====================================================
        PARA OBTENER LOS ID DE LAS IMPRESORAS USB 
=====================================================

"""
def get_usb_printer_ids():
    printer_ids = []
    wmi = win32com.client.GetObject("winmgmts:")
    
    for printer in wmi.InstancesOf("Win32_Printer"):
        if "USB" in printer.PortName:
            printer_ids.append(printer.PnPDeviceID)
    return printer_ids

def get_vendor_product_ids(pnp_device_id):
    vendor_id, product_id = None, None
    wmi = win32com.client.GetObject("winmgmts:")
    for item in wmi.ExecQuery("SELECT * FROM Win32_PnPEntity WHERE DeviceID='{}'".format(pnp_device_id)):
        hardware_id = item.HardwareID[0]
        if "VID_" in hardware_id and "PID_" in hardware_id:
            vendor_id = hardware_id[hardware_id.index("VID_")+4:hardware_id.index("PID_")].upper()
            product_id = hardware_id[hardware_id.index("PID_")+4:].upper()
            break
    return vendor_id, product_id




def obtener_ruta_descargas():
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
        downloads_directory = winreg.QueryValueEx(key, "{374DE290-123F-4565-9164-39C4925E467B}")[0]
    # Expandir la variable de entorno para obtener la ruta completa
    downloads_directory = os.path.expandvars(downloads_directory)
    return downloads_directory


def imprimir_pdf(pdf_path) -> None:
    printer_ids = get_usb_printer_ids()
    for pnp_device_id in printer_ids:
        vendor_id, product_id = get_vendor_product_ids(pnp_device_id)
        if vendor_id  and product_id:
            # Inicializa la impresora térmica
            #0x0416, 0x5011, 0,
            printer = Usb(vendor_id,product_id)
            # Abre el archivo PDF
            pdf_document = fitz.open(pdf_path)
            for page_number in range(len(pdf_document)):
                # Obtiene la página
                page = pdf_document.load_page(page_number)
                # Renderiza la página como una imagen (formato PNG)
                image = page.get_pixmap()
                # Abre la imagen usando PIL
                pil_image = Image.frombytes("RGB", [image.width, image.height], image.samples)
                # Escala la imagen para que se ajuste al ancho de la impresora térmica
                width, height = pil_image.size
                new_width = 384  
                new_height = int((new_width / width) * height)
                pil_image = pil_image.resize((new_width, new_height))
                # Convierte la imagen a escala de grises 
                pil_image = pil_image.convert("L")
                # Imprime la imagen en la impresora térmica
                printer.image(pil_image)
            # Corta el papel después de imprimir todas las imágenes
            printer.cut()
            # Cierra el documento PDF
            pdf_document.close()
            os.remove(path=pdf_path)            
        else:
           os.remove(path=pdf_path)
           raise Exception('**No se encontro ninguna impresora**')
    
def main():
    descargas = obtener_ruta_descargas()

    # Obtener una lista de archivos en el directorio de descargas
    while True:
        descargas = obtener_ruta_descargas()
        pdfs = list(archivo for archivo in  os.listdir(descargas) if archivo.endswith('.pdf'))
        for file in pdfs:
            # Verificación de tickets del colegio
            list_name_file = file.split('_')
            if 'jvtk' in list_name_file:
                path = rf'{descargas}\{file}'
                #print(path)
                try:
                    imprimir_pdf(path)
                except Exception as e:
                    print(e)
                break    
        # Actualización del estado cada 5 segundos
        sleep(5)


def get_usb_printer_names():
  """
  Obtiene los nombres de las impresoras USB conectadas.

  Retorna:
      Una lista con los nombres de las impresoras USB.
  """
  printer_names = []

  try:
   main()
  except Exception as e:
    print(f"Error al obtener la lista de impresoras: {e}")

  return printer_names


# Función para obtener las impresoras USB conectadas y sus Vendor ID y Product ID
def get_usb_printers():
    printers = []
    # Buscar todos los dispositivos USB
    devices = usb.core.find(find_all=True)
    # Filtrar los dispositivos que son impresoras
    for device in devices:
        # Verificar si el dispositivo es una impresora USB (clase 7)
        if device.bDeviceClass == 7:
            # Obtener Vendor ID y Product ID
            vendor_id = hex(device.idVendor)
            product_id = hex(device.idProduct)
            # Guardar la información de la impresora
            printers.append({"Vendor ID": vendor_id, "Product ID": product_id})
    return printers

# Función para extraer Vendor ID y Product ID del nombre del controlador
def extract_ids_from_driver_name(driver_name):
    # Dividir el nombre del controlador por espacios
    driver_name_parts = driver_name.split()
    vendor_id = None
    product_id = None
    # Buscar partes del nombre del controlador que contienen "VendorID" y "ProductID"
    for part in driver_name_parts:
        if "VendorID" in part:
            vendor_id = part.split("VendorID")[1]
        elif "ProductID" in part:
            product_id = part.split("ProductID")[1]
    return vendor_id, product_id

def get_usb_printers_():
    printers = []
    printers = [win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)[-1]]
    #return printers
    
    for printer_info in printers:
        # Verificar si la impresora es una impresora USB
        if  printer_info[2]:
            # Obtener el nombre de la impresora
            printer_name = printer_info[2]

            printer_handle = win32print.OpenPrinter(printer_name)
            # Obtener el Vendor ID y el Product ID de la impresora
            printer_info_2 = win32print.GetPrinter(printer_handle, 2)
            #printer_info_2 = win32print.GetPrinter(printer_name, 2)
            #return printer_info_2
            driver_name = printer_info_2['pDriverName']

            vendor_id, product_id = extract_ids_from_driver_name(driver_name)
            #printers.append({"Printer Name": printer_name, "Vendor ID": vendor_id, "Product ID": product_id})

            return {"Printer Name": printer_name, "Vendor ID": vendor_id, "Product ID": product_id}
            # El Vendor ID y Product ID están en la cadena pSecurityDescriptor en formato hexadecimal
            # El formato es "\\VendorIDxxxx&ProductIDyyyy#" (por ejemplo, "\\VendorID04A9&ProductID190D#")
            vendor_id_start = printer_info_2[2].find("VendorID") + len("VendorID")
            vendor_id_end = printer_info_2[2].find("&ProductID")
            product_id_start = vendor_id_end + len("&ProductID")
            product_id_end = printer_info_2[2].find("#")
            vendor_id = printer_info_2[2][vendor_id_start:vendor_id_end]
            product_id = printer_info_2[2][product_id_start:product_id_end]
            printers.append({"Printer Name": printer_name, "Vendor ID": vendor_id, "Product ID": product_id})
    return printers

if __name__ == "__main__":
    try:
        print(get_usb_printers_())
        #main()
        #bus = pyusb.usb.busses[0]
        #devices = bus.get_devices()
        #print(devices)
    except Exception as  e:
        print(f'**Error al ejecutar el script: {e}**')