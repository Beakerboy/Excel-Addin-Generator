# Excel-Addin-Generator
Tools to create an Excel Addin from VBA code

[![Coverage Status](https://coveralls.io/repos/github/Beakerboy/Excel-Addin-Generator/badge.svg?branch=master)](https://coveralls.io/github/Beakerboy/Excel-Addin-Generator?branch=master)[![Build Status](https://travis-ci.org/Beakerboy/Excel-Addin-Generator.svg?branch=master)](https://travis-ci.org/Beakerboy/Excel-Addin-Generator)


Preparation
------------

An Excel Addin (.xlam) is a zip archive of XML and binary files. Some of these XML files contain information specific to the computer that was used to create the xlam file. This script can be used to create a standard xlam file from the VBA code.

The script can be given a source xlam file, from which the binary file is extracted an repackaged, or the binary file can be provided alone.

**Usage**

`python excelAddinGenerator/main.py path/to/vbaProject.bin output/to/myAddin.xlam`

or

`python excelAddinGenerator/main.py path/to/source.xlam output/to/myAddin.xlam`
