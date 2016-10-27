#! /usr/bin/env python
# -*- coding: utf-8 -*-

import unittest
import subprocess

import mowgli


class DefaultTestCase(unittest.TestCase):
    def test_disable_au(self):
        mowgli.set_pref()
        o = subprocess.check_output(['defaults',
                                     'read',
                                     'com.microsoft.autoupdate2',
                                     'HowToCheck'])
        self.assertEquals(o.strip(), 'Manual')

    def test_enable_au(self):
        mowgli.set_pref(v='Automatic')
        o = subprocess.check_output(['defaults',
                                     'read',
                                     'com.microsoft.autoupdate2',
                                     'HowToCheck'])
        self.assertEquals(o.strip(), 'Automatic')


if __name__ == '__main__':
    unittest.main()
