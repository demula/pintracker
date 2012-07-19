pintracker
==========

Build on windows with pyinstaller

Install dependencies
--------------------

Python 2.7 - http://python.org/ftp/python/2.7.3/python-2.7.3.msi
Pywin32 extensions - http://sourceforge.net/projects/pywin32/files/pywin32/Build%20217/pywin32-217.win32-py2.7.exe/download
PyGTK - http://ftp.gnome.org/pub/GNOME/binaries/win32/pygtk/2.24/pygtk-all-in-one-2.24.2.win32-py2.7.msi
Setuptools (for python modules) - http://pypi.python.org/packages/2.7/s/setuptools/setuptools-0.6c11.win32-py2.7.exe#md5=57e1e64f6b7c7f1d2eddfc9746bbaf20
easy_install xlrd
easy_install openpyxl

Compile instructions
--------------------
cd C:\dir\to\pintracker\src
C:\Python27\python.exe C:\pyinstaller\Configure.py
C:\Python27\python.exe C:\pyinstaller\Makespec.py -F -w pintracker.py
C:\Python27\python.exe C:\pyinstaller\Build.py pintracker.spec

the one file executable is in the dist directory