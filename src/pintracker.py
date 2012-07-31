#!/usr/bin/env python2
# -*- coding: utf-8 -*-
import os, dircache, ConfigParser
import xlrd, openpyxl
import datetime
import time
import email
from email.mime.text import MIMEText
import smtplib
import glib
import gtk


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
        now = datetime.datetime.now()
        self.next_check_daily = datetime.datetime(
                now.year, now.month, now.day,
                self.config["daily_send_hour"].hour,
                self.config["daily_send_hour"].minute)
        # Calculate days to weekly check
        if self.config["weekly_send_day"] >= now.weekday():
            days_to_weekly = self.config["weekly_send_day"] - now.weekday()
        else:
            days_to_weekly = 7 - now.weekday() + self.config["weekly_send_day"]
        delta_days = datetime.timedelta(days=days_to_weekly)
        self.next_check_weekly = datetime.datetime(
                now.year, now.month, now.day,
                self.config["weekly_send_hour"].hour,
                self.config["weekly_send_hour"].minute) + delta_days

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
        now = datetime.datetime.now()
        seconds_from_deadline = abs((self.next_check_daily - now).total_seconds())
        if seconds_from_deadline < self.config["interval_min"]*60:
            delta_day = datetime.timedelta(days=1)
            self.next_check_daily = self.next_check_daily + delta_day
            return True
        else:
            return False

    def is_time_weekly(self):
        now = datetime.datetime.now()
        in_window = abs((self.next_check_weekly - now).total_seconds()) < self.config["interval_min"]*60
        same_weekday = self.config["weekly_send_day"] == now.weekday()
        if same_weekday and in_window:
            delta_week = datetime.timedelta(weeks=1)
            self.next_check_weekly = self.next_check_weekly + delta_week
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

    def send_estado_pines(self):
        below_min = self.check_min_pines()
        estado_pines = self.estado_pines()
        PinEmail(self.config).send_status(estado_pines, below_min)
        print "\n\nEMAIL SENT by petition\n\n\n"


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
                   self.config["recipients"].split(", "),
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


    def send_status(self, estado_pines, below_limit):
        subject = "[STOCK PINES] Informe por peticion"
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


# ------------------------------------------------------------------------ GUI
class PintrackerStatusIcon(object):

    def __init__(self, config):
        # hold a pintracker instance
        self.pintrck = Pintracker(config)
        self.pins_liststore = gtk.ListStore(str, int)
        self.update_pins_liststore()
        # add the time out callback
        self.interval = config["interval_min"]*60 #seconds
        glib.timeout_add_seconds(self.interval, self.timeout_repeat)

        # create a new Status Icon
        self.staticon = gtk.StatusIcon()
        self.staticon.set_from_stock(gtk.STOCK_EDIT)
        self.staticon.set_name("pintracker")
        self.staticon.set_title("Pintracker")
        #self.staticon.set_tooltip("Pintracker")
        #self.staticon.set_blinking(True)

        # create de popup menu
        self.menu = gtk.Menu()

        # Create the menu items
        status = gtk.MenuItem("Pin status")
        about = gtk.MenuItem("About")
        exit = gtk.MenuItem("Exit")


        #connect signals
        status.connect("activate", self.show_status_window)
        about.connect("activate", self.show_about_dialog)
        exit.connect("activate", gtk.main_quit)

        # Add them to the menu
        self.menu.append(status)
        self.menu.append(about)
        self.menu.append(exit)

        self.menu.show_all()

        #connect signals
        self.staticon.connect("popup-menu", self.popup_menu) #right click

        #show everything needed
        self.staticon.set_visible(True)

    def update_pins_liststore(self):
        dict_state = self.pintrck.estado_pines()
        if len(self.pins_liststore) == 0:
            # populate for the first time
            for path in dict_state.keys():
                self.pins_liststore.append([path, dict_state[path]])
        else:
            for row in self.pins_liststore:
                if row[0] in dict_state.keys():
                    row = [row[0], dict_state[row[0]]]

    def timeout_repeat(self):
        self.pintrck.run()
        glib.timeout_add_seconds(self.interval, self.timeout_repeat)
        return False


    def run(self):
        gtk.main()
        return False

    # popup menu callback
    def popup_menu(self, icon, button, time):
        self.menu.popup(None, None, gtk.status_icon_position_menu,
                    button, time, self.staticon)
        return False

    def delete_status_window(self, widget, event, data=None):
        self.window.hide()
        return False

    def show_status_window(self, widget):
        self.window = gtk.Window(gtk.WINDOW_TOPLEVEL)
        self.window.connect("delete_event", self.delete_status_window)
        self.window.set_border_width(10)
        self.window.set_title("Pin status")
        #self.window.set_size_request(200, 200)

        vbox = gtk.VBox(False, 10)
        hbox = gtk.HBox(False, 10)

        label = gtk.Label("Stock folder: %s" % self.pintrck.config["stock_dir"])

        hbox.pack_start(label, False, False, 0)
        label.show()

        vbox.pack_start(hbox, False, False, 0)

        hbox.show()

        # create the TreeView using liststore
        treeview = gtk.TreeView(self.pins_liststore)

        # create the TreeViewColumns to display the data
        tvcolumn_folders = gtk.TreeViewColumn('Folder')
        tvcolumn_pins = gtk.TreeViewColumn('Pines')

        # add columns to treeview
        treeview.append_column(tvcolumn_folders)
        treeview.append_column(tvcolumn_pins)

        # create a CellRenderers to render the data
        cell_folder = gtk.CellRendererText()
        cell_pins = gtk.CellRendererText()

        tvcolumn_folders.pack_start(cell_folder, True)
        tvcolumn_pins.pack_start(cell_pins, True)

        tvcolumn_folders.set_attributes(cell_folder, text=0)
        tvcolumn_pins.set_attributes(cell_pins, text=1)

        # make treeview searchable
        #treeview.set_search_column(0)

        # Allow sorting on the column
        tvcolumn_folders.set_sort_column_id(0)
        tvcolumn_pins.set_sort_column_id(1)

        # Allow drag and drop reordering of rows
        treeview.set_reorderable(True)

        # create a new scrolled window.
        scrolled_window = gtk.ScrolledWindow()
        scrolled_window.set_border_width(0)
        scrolled_window.set_size_request(400, 500)

        # the policy is one of POLICY AUTOMATIC, or POLICY_ALWAYS.
        # POLICY_AUTOMATIC will automatically decide whether you need
        # scrollbars, whereas POLICY_ALWAYS will always leave the scrollbars
        # there. The first one is the horizontal scrollbar, the second, the
        # vertical.
        scrolled_window.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_ALWAYS)
        scrolled_window.add_with_viewport(treeview)

        vbox.pack_start(scrolled_window, True, True, 0)
        #vbox.pack_start(treeview, False, False, 0)

        treeview.show()
        scrolled_window.show()

        separator = gtk.HSeparator()
        #separator.set_size_request(400, 5)

        separator.show()

        quitbox = gtk.HBox(False, 10)

        email_button = gtk.Button("Send email")
        qbutton = gtk.Button("Close")

        qbutton.connect("clicked", lambda w: self.window.hide())
        email_button.connect("clicked", lambda w: self.pintrck.send_estado_pines())
        quitbox.pack_end(qbutton, False, False, 0)
        quitbox.pack_end(email_button, False, False, 0)

        vbox.pack_end(quitbox, False, False, 0)
        vbox.pack_end(separator, False, True, 5)

        separator.show()

        self.window.add(vbox)

        email_button.show()
        qbutton.show()
        quitbox.show()

        vbox.show()
        self.window.show()
        return False


    def show_about_dialog(self, widget):
        about_dialog = gtk.AboutDialog()

        about_dialog.set_destroy_with_parent(True)
        about_dialog.set_name("Pintracker")
        about_dialog.set_version("0.1")
        about_dialog.set_authors(["Jesus de Mula Cano"])

        about_dialog.run()
        about_dialog.destroy()
        return False


# ----------------------------------------------------------------------- MAIN
def main():
    # Deberiamos cambiar esta ruta a la estandar del SO
    config_file_path = os.path.join(os.path.abspath(os.path.dirname(__file__)),
                                                    "pintracker.cfg")
    config = read_config(config_file_path)
    pintrck = PintrackerStatusIcon(config)
    pintrck.run()

if __name__ == "__main__": main()
