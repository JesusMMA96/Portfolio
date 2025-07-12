# -*- coding: utf-8 -*-
"""
@author: JesusMMA
"""
# Flag
ContinueProgram = True
# Load json
import json
def load_SAP_info(path="SAP_info.json"):
    with open(path, "r") as file:
        return json.load(file)

config = load_SAP_info()
