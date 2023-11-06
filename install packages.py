#Author- Florian Wagner
#Description- Install python packages 

import adsk.core, adsk.fusion, adsk.cam, traceback
import subprocess, sys

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface

        PathToPyton = sys.path

        print("\n********************************************************\n")
        for ii in PathToPyton:
            print(ii)

        try:
            subprocess.check_call([PathToPyton[2] + '\\Python\\python.exe', "-m pip install --upgrade pip"])
        except:
            ui.messageBox("Failed to upgrade pip")
        try:
            subprocess.check_call([PathToPyton[2] + '\\Python\\python.exe', "-m", "pip", "install", "numpy"])
            subprocess.check_call([PathToPyton[2] + '\\Python\\python.exe', "-m", "pip", "install", "pandas"])
            ui.messageBox("Packages installed.")
        except:
            ui.messageBox("Failed to install packages.")

        try:
            import numpy as np
            import pandas as pd
            ui.messageBox("Package test import succeeded!")
        except:
            ui.messageBox("Failed when importing packages.")

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

