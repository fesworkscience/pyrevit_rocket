# -*- coding: utf-8 -*-
"""
CPSK Project Configuration
Global configuration variables for the project.
"""

import os

# Debug mode - disables authentication requirement
DEBUG = False

# Gemini API Token
GEMINI_TOKEN = os.environ.get("GEMINI_TOKEN", "")

# API Settings
API_BASE_URL = "https://api-cpsk-superapp.gip.su"
API_ROCKETREVIT_URL = API_BASE_URL + "/api/rocketrevit"
