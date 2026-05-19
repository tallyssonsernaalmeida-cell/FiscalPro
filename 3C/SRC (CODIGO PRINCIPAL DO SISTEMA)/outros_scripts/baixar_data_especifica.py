"""
FORÇAR DOWNLOAD DE UMA DATA ESPECÍFICA
Altere a variável DATA_ALVO abaixo
"""

import os
import re
import json
import time
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests
from bs4 import BeautifulSoup

# ================================================================
# CONFIGURAÇÃO ESPECÍFICA
# ================================================================
DATA_ALVO = "04/05/2026"  # <<< ALTERE A DATA AQUI

# ================================================================
# (copiar todo o resto do script, mas sem a lógica de data)
# ================================================================
# ... (todo o código do seu script, com a modificação abaixo)