#!/usr/bin/env python2
# -*- coding: utf-8 -*-
import sys, os
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'test'))
sys.path.insert(0, os.path.dirname(__file__))

unittest.main(module='test_pines')
