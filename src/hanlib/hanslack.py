# -*- coding: utf-8 -*-
"""
Created on Sat Nov 13 14:32:51 2021

@author: MarkTan
"""

import json
import requests
from dotenv import load_dotenv
import os
load_dotenv()

SLACK_API_KEY = os.getenv("SLACK_API_KEY")

def send_slack_message(payload, hook=SLACK_API_KEY):
    
    response = requests.post(hook,
                             data=json.dumps(payload))
    return response
