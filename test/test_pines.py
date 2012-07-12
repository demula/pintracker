#!/usr/bin/env python2
# -*- coding: utf-8 -*-
import os
import shutil
import time
import ConfigParser
import datetime
import unittest
from src.pintracker import read_config
from src.pintracker import Pintracker


class TestConfig(unittest.TestCase):

    def setUp(self):
        self.base_path = os.path.abspath(os.path.dirname(__file__))
        self.dirs = ["jazztel (activado)", "orange", "tarjetalia"]
        self.config_file_path = os.path.join(self.base_path, "sample.cfg")
        listdir = os.listdir(self.base_path)
        for d in self.dirs:
            if d not in listdir:
                os.mkdir(os.path.join(self.base_path, d))

        config = ConfigParser.RawConfigParser()
        config.read(self.config_file_path)
        config.set("config", "stock_dir", self.base_path)
        # Writing our configuration file to 'sample.cfg'
        with open(self.config_file_path, 'wb') as f:
            config.write(f)

    def tearDown(self):
        listdir = os.listdir(self.base_path)
        for d in self.dirs:
            if d in listdir:
                shutil.rmtree(os.path.join(self.base_path, d))

        if "one_too_many" in listdir:
            shutil.rmtree(os.path.join(self.base_path, "one_too_many"))

    def test_config(self):
        dict_sample_config = {
                "stock_dir": self.base_path,
                "smtp_server":"mail.example.com",
                "email":"user@example.com",
                "recipients":"user2@example.com, user3@example.com",
                "user":"login_user_name",
                "password":"1234",
                "interval_min":40,
                "daily_send_hour": datetime.time(18,00),
                "weekly_send_hour": datetime.time(8,00),
                "weekly_send_day": 4,
                "min_pines": {
                    "jazztel (activado)":100,
                    "orange":20,
                    "tarjetalia":10
                    }
                }
        dict_number_pins = read_config(self.config_file_path)
        for i in dict_sample_config.iterkeys():
            self.assertEqual(dict_sample_config[i], dict_number_pins[i])

    def test_num_pins_dir_short(self):
        config_file_path = os.path.join(self.base_path,"sample.cfg")
        # Remove one dir to check for failure
        os.rmdir(os.path.join(self.base_path, "orange"))
        time.sleep(0.5)
        self.assertRaises(KeyError, read_config, config_file_path)

    def test_num_pins_dir_long(self):
        config_file_path = os.path.join(self.base_path,"sample.cfg")
        # Add one dir to check for failure
        os.mkdir(os.path.join(self.base_path, "one_too_many"))
        time.sleep(0.5)
        self.assertRaises(KeyError, read_config, config_file_path)


class TestPinesReader(unittest.TestCase):

    def setUp(self):
        self.expected_results = {"sample_pins.xls":60, "sample_pins.xlsx":25}
        self.base_path = os.path.abspath(os.path.dirname(__file__))

    def test_num_pins(self):
        pintrck = Pintracker({})
        for file in self.expected_results.iterkeys():
            pines_file_path = os.path.join(self.base_path, file)
            pines_leidos = pintrck.numero_pines(pines_file_path)
            self.assertEqual(self.expected_results[file], pines_leidos)

class TestCheckPinesDir(unittest.TestCase):

    def setUp(self):
        self.base_path = os.path.abspath(os.path.dirname(__file__))
        self.dir = "sample_dir"
        self.files = ["sample_pins.xls", "sample_pins.xlsx"]
        self.expected_result = 85
        listdir = os.listdir(self.base_path)
        if self.dir not in listdir:
            os.mkdir(os.path.join(self.base_path, self.dir))
        for f in self.files:
            shutil.copyfile(
                    os.path.join(self.base_path, f),
                    os.path.join(self.base_path, self.dir, f))

    def tearDown(self):
        shutil.rmtree(os.path.join(self.base_path, self.dir))

    def test_check_dir_total_pines(self):
        pintrck = Pintracker({})
        num_pines = pintrck.check_dir_total_pines(
                                        os.path.join(self.base_path, self.dir))
        self.assertEqual(num_pines, self.expected_result)


class TestEstadoPines(unittest.TestCase):

    def setUp(self):
        self.base_path = os.path.abspath(os.path.dirname(__file__))
        self.dirs = ["jazztel (activado)", "orange", "tarjetalia"]
        self.files = ["sample_pins.xls", "sample_pins.xlsx"]
        listdir = os.listdir(self.base_path)
        for d in self.dirs:
            if d not in listdir:
                os.mkdir(os.path.join(self.base_path, d))
                for f in self.files:
                    shutil.copyfile(
                        os.path.join(self.base_path, f),
                        os.path.join(self.base_path, d, f))


    def tearDown(self):
        for d in self.dirs:
            shutil.rmtree(os.path.join(self.base_path, d))

    def test_estado_pines(self):
        dict_sample_config = {
                "stock_dir": self.base_path,
                "smtp_server":"mail.example.com",
                "email":"user@example.com",
                "password":"1234",
                "interval_min":40,
                "daily_send_hour": datetime.time(18,00),
                "weekly_send_hour": datetime.time(8,00),
                "weekly_send_day": 4,
                "min_pines": {
                    "jazztel (activado)":100,
                    "orange":20,
                    "tarjetalia":10
                    }
                }
        dict_expected_result = {
                "jazztel (activado)": 85,
                "orange":85,
                "tarjetalia":85
                }
        pintrck = Pintracker(dict_sample_config)
        actual_result = pintrck.estado_pines()
        self.assertEqual(actual_result, dict_expected_result)

class TestCheckPinesMin(unittest.TestCase):

    def setUp(self):
        self.base_path = os.path.abspath(os.path.dirname(__file__))
        self.dirs = ["jazztel (activado)", "orange", "tarjetalia"]
        self.files = ["sample_pins.xls", "sample_pins.xlsx"]
        listdir = os.listdir(self.base_path)
        for d in self.dirs:
            if d not in listdir:
                os.mkdir(os.path.join(self.base_path, d))
                for f in self.files:
                    shutil.copyfile(
                        os.path.join(self.base_path, f),
                        os.path.join(self.base_path, d, f))


    def tearDown(self):
        for d in self.dirs:
            shutil.rmtree(os.path.join(self.base_path, d))

    def test_min_pines(self):
        dict_sample_config = {
                "stock_dir": self.base_path,
                "smtp_server":"mail.example.com",
                "email":"user@example.com",
                "password":"1234",
                "interval_min":40,
                "daily_send_hour": datetime.time(18,00),
                "weekly_send_hour": datetime.time(8,00),
                "weekly_send_day": 4,
                "min_pines": {
                    "jazztel (activado)":100,
                    "orange":20,
                    "tarjetalia":10
                    }
                }
        dict_expected_result = {
                "jazztel (activado)": 85 - 100
                }
        pintrck = Pintracker(dict_sample_config)
        actual_result = pintrck.check_min_pines()
        self.assertEqual(actual_result, dict_expected_result)

class TestIsTime(unittest.TestCase):

    def test_is_time_daily(self):
        interval_min = 40
        now = datetime.datetime.now()

        send_time_date = now + datetime.timedelta(minutes=interval_min/2)
        send_time = send_time_date.time()
        dict_sample_config = {
                "interval_min": interval_min,
                "daily_send_hour": send_time,
                "weekly_send_hour": send_time,
                "weekly_send_day": 4,
                }
        expected_result = True
        pintrck = Pintracker(dict_sample_config)
        actual_result = pintrck.is_time_daily()
        self.assertEqual(actual_result, expected_result)

        send_time_date = now - datetime.timedelta(minutes=interval_min/2)
        send_time = send_time_date.time()
        dict_sample_config = {
                "interval_min": interval_min,
                "daily_send_hour": send_time,
                "weekly_send_hour": send_time,
                "weekly_send_day": 4,
                }
        expected_result = False
        pintrck = Pintracker(dict_sample_config)
        actual_result = pintrck.is_time_daily()
        self.assertEqual(actual_result, expected_result)

    def test_is_time_weekly(self):
        interval_min = 40
        now = datetime.datetime.now()

        send_time_date = now + datetime.timedelta(minutes=interval_min/2)
        send_time = send_time_date.time()
        dict_sample_config = {
                "interval_min": interval_min,
                "daily_send_hour": send_time,
                "weekly_send_hour": send_time,
                "weekly_send_day": now.weekday(),
                }
        expected_result = True
        pintrck = Pintracker(dict_sample_config)
        actual_result = pintrck.is_time_weekly()
        self.assertEqual(actual_result, expected_result)

        send_time_date = now - datetime.timedelta(minutes=interval_min/2)
        send_time = send_time_date.time()
        dict_sample_config = {
                "interval_min": interval_min,
                "daily_send_hour": send_time,
                "weekly_send_hour": send_time,
                "weekly_send_day": now.weekday(),
                }
        expected_result = False
        pintrck = Pintracker(dict_sample_config)
        actual_result = pintrck.is_time_weekly()
        self.assertEqual(actual_result, expected_result)

        send_time_date = now + datetime.timedelta(minutes=interval_min/2)
        send_time = send_time_date.time()
        dict_sample_config = {
                "interval_min": interval_min,
                "daily_send_hour": send_time,
                "weekly_send_hour": send_time,
                "weekly_send_day": (now.weekday() + 2) % 7,
                }
        expected_result = False
        pintrck = Pintracker(dict_sample_config)
        actual_result = pintrck.is_time_weekly()
        self.assertEqual(actual_result, expected_result)


