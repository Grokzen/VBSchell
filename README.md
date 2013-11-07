========
VBSchell
========

VBSchell is a interactive interpreter for VBscript. (Similar to pythons interactive interpreter)



Supported features
==================

 - Input argument handling (Using docopt syntax standard)
 - VBScript style line continuations (Use _ at the end of each line)
 - Script import at startup
 - Import of script files during runtime.
 - Printing of variables (No need to write 'wscript.echo FooVar' just write 'FooVar')
 - Proper logging. (DEBUG/INFO/WARNING/ERROR/CRITICAL standard used)



Interactive mode commands
=========================
``` Shell
 ?           Prints this Help
 ?exit       Exit the shell
 ?import     Prompts to import a .vbs script
 ?reimport   Reimports the predefined scripts
 ?version    Prints the version
```


CLI usage
=========

``` Shell
Usage:
 vbsh.wsf [-v ... ] [--help]
 vbsh.wsf --version

Options:
 --help      Prints this help
 -v ...      Sets the logging level of this run. [Default: -vvvv --> INFO logging] [Supports level 1-5]
 --version   Prints the version of the application
```



Startup script import
=====================

Add the full path to all files you want to import to the array 'scripts'. At startup time you will be asked if you want to import them or not. All files will be imported and no cherry picking is available.

If you allways want the scripts to load at startup, set the variable 'allways_import_scripts' to True.



Supported runtime script imports
================================

 - .vbs files
 - NYI: .wsf files (Will be implemented in the future)



Upcoming features
=================

 - Python style line continuations
 - Automatic printing of function returns if possible
