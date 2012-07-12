#!/usr/bin/env python2
# -*- coding: utf-8 -*-
import os, dircache, ConfigParser
import xlrd, openpyxl
import datetime
import time
import email
from email.mime.text import MIMEText
import smtplib


# -------------------------------------------------------------- CONFIG READER
def read_config(file_path):
    """
    Lee la configuracion de "min_pines.cfg" y devuelve un
    diccionario con los valores leidos "Nombre":num_entero.
    """
    config = ConfigParser.RawConfigParser()
    config.read(file_path)
    parsed_config = {}

    for (name, value) in config.items("config"):
        parsed_config[name] = str(value)

    parsed_config["interval_min"] = int(parsed_config["interval_min"])

    hour =  int(parsed_config["daily_send_hour"].split(":")[0])
    min =  int(parsed_config["daily_send_hour"].split(":")[1])
    parsed_config["daily_send_hour"] = datetime.time(hour, min)

    hour =  int(parsed_config["weekly_send_hour"].split(":")[0])
    min =  int(parsed_config["weekly_send_hour"].split(":")[1])
    parsed_config["weekly_send_hour"] = datetime.time(hour, min)

    week_days = ["Mon", "Tue", "Wen", "Thu", "Fri", "Sat", "Sun"]
    parsed_config["weekly_send_day"] = week_days.index(
                                            parsed_config["weekly_send_day"])

    if not os.path.isdir(parsed_config["stock_dir"]):
        raise IOError("Stock dir erroneus")

    available_files = dircache.listdir(parsed_config["stock_dir"])
    available_files = available_files[:]
    dircache.annotate(parsed_config["stock_dir"], available_files)
    available_dirs = [ item[:-1] for item in available_files if '/' in item]
    parsed_config["min_pines"] = {}
    for (name, value) in config.items("num_min_pines"):
        parsed_config["min_pines"][name] = int(value)
        if name not in available_dirs:
            raise KeyError("Config name without dir")

    if len(available_dirs) != len(parsed_config["min_pines"]):
        raise KeyError("Dir without definition in config file")
    return parsed_config


# -------------------------------------------------------------- PROGRAM LOGIC
class Pintracker(object):
    def __init__(self, config):
        self.config = config
        self.last_email_time = None
        self.last_pines_below_min = {}

    def numero_pines(self, ruta_fichero_excel):
        """
        Lee el archivo excel indicado y por su numero de filas estima el
        numero de pines que tiene en ese instante. Devuelve un numero
        entero.
        """
        pinesfichero = 0
        if ruta_fichero_excel[-3:] == "xls":
            libro = xlrd.open_workbook(ruta_fichero_excel)
            hojas = libro.sheets()
            hoja = hojas[0]
            pinesfichero = hoja.nrows
        if ruta_fichero_excel[-4:] == "xlsx":
            libro = openpyxl.load_workbook(filename = ruta_fichero_excel)
            hoja = libro.get_active_sheet()
            pinesfichero = hoja.get_highest_row()
        return pinesfichero

    def check_dir_total_pines(self, dir_path):
        """
        Suma todo los pines de una carpeta
        """
        pines_carpeta = 0
        available_files = dircache.listdir(dir_path)
        for file in available_files:
            pines_carpeta = pines_carpeta + self.numero_pines(
                                os.path.join(dir_path, file))
        return pines_carpeta

    def estado_pines(self):
        """
        Busca en todas las carpetas en configuracion y devuelve un diccionario
        con el estado general de los pines.
        """
        base_path = self.config["stock_dir"]
        estado_dict = {}
        for dir in self.config["min_pines"].iterkeys():
            estado_dict[dir] = self.check_dir_total_pines(
                                                os.path.join(base_path,dir))
        return estado_dict

    def check_min_pines(self):
        """
        Busca en todas las carpetas en configuracion y si hay alguna que esta
        por debajo del minimo indicado la añade al diccionario que devuelve.
        """
        below_min = {}
        estado_dict = self.estado_pines()
        for dir in estado_dict.iterkeys():
            min_pines_dir = self.config["min_pines"][dir]
            actual_pines_dir = estado_dict[dir]
            if actual_pines_dir < min_pines_dir:
                below_min[dir] = actual_pines_dir - min_pines_dir
        return below_min

    def is_time_daily(self):
        now = datetime.datetime.now().time()
        send_time = self.config["daily_send_hour"]
        if (now < send_time and
            send_time.minute - now.minute < self.config["interval_min"]):
            return True
        else:
            return False

    def is_time_weekly(self):
        now = datetime.datetime.now()
        send_time = self.config["weekly_send_hour"]
        in_window = send_time.minute - now.minute < self.config["interval_min"]
        same_weekday = self.config["weekly_send_day"] == now.weekday()
        if now.time() < send_time and in_window and same_weekday:
            return True
        else:
            return False

    def run(self):
        if self.is_time_daily():
            below_min = self.check_min_pines()
            if below_min:
                PinEmail(self.config).send_daily(below_min)
                print "\n\nEMAIL SENT DAILY\n\n\n"
        if self.is_time_weekly():
            below_min = self.check_min_pines()
            estado_pines = self.estado_pines()
            PinEmail(self.config).send_weekly(estado_pines, below_min)
            print "\n\nEMAIL SENT WEEKLY\n\n\n"


class PinEmail(object):
    def __init__(self, config):
        self.config = config

    def send(self, subject, body):
        msg = MIMEText(body, 'html', 'utf-8')
        msg['Subject'] = subject
        msg['From'] = self.config["email"]
        msg['To'] = self.config["recipients"]

        s = smtplib.SMTP(self.config["smtp_server"])
        s.ehlo()
        s.esmtp_features['auth'] = 'LOGIN PLAIN'
        s.debuglevel = 5
        s.login(self.config["user"], self.config["password"])
        s.sendmail(self.config["email"],
                   self.config["recipients"],
                   msg.as_string())
        s.quit()

    def send_daily(self, pines_below_limit):
        subject = "[STOCK PINES] ALERTA diaria"
        body = """<p>Sumario de los pines que están por debajo
        de los limites establecidos:</p><br><ul>
        """
        for item in pines_below_limit.iterkeys():
            body = body + "<li>    %s: %i (limit %i)</li><br>" % (item,
                                                pines_below_limit[item],
                                                self.config["min_pines"][item])
        body = body + "</ul><br><br><p>Hora de revision: %i:%i<br>" % (datetime.datetime.now().hour,
                                                        datetime.datetime.now().minute)
        body = body + "Carpeta de STOCK: %s<br><br></p>" % self.config["stock_dir"]
        self.send(subject,body)

    def send_weekly(self, estado_pines, below_limit):
        subject = "[STOCK PINES] Informe semanal"
        body = """<p>Sumario del estado de los pines en stock (en rojo aquellos
        por debajo de los limites establecidos):</p><br><ul>
        """
        for item in estado_pines.iterkeys():
            if item in below_limit:
                body = body + '<li><b style="color:red;">    %s: %i (limite %i)</b></li><br>' % (
                                                    item,
                                                    estado_pines[item],
                                                    self.config["min_pines"][item])
            else:
                body = body + '<li>    %s: %i (limite %i)</li><br>' % (item,
                                                    estado_pines[item],
                                                    self.config["min_pines"][item])

        body = body + "</ul><br><br><p>Hora de revision: %i:%i<br>" % (datetime.datetime.now().hour,
                                                        datetime.datetime.now().minute)
        body = body + "Carpeta de STOCK: %s<br><br></p>" % self.config["stock_dir"]
        self.send(subject,body)


# ----------------------------------------------------------------------- MAIN
def main():
    # Deberiamos cambiar esta ruta a la estandar del SO
    config_file_path = os.path.join(os.path.abspath(os.path.dirname(__file__)),
                                                    "pintracker.cfg")
    config = read_config(config_file_path)
    pintrck = Pintracker(config)
    while True:
        pintrck.run()
        time.sleep(self.config["interval_min"]*60)

if __name__ == "__main__": main()
