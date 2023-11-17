# conig file 
# last revised ... August 18, 2023


import os

# Ottieni il percorso assoluto della directory del modulo corrente
current_dir = os.path.dirname(os.path.abspath(__file__))

# Il parent directory di automation24
basedir = 'C:/SaS/'

# La directory dell'agente
agentdir = basedir + os.path.sep + 'agent'

# Ricerca automatica del percorso per rst2odt.py all'interno dell'ambiente virtuale
#rst2odtpath = 'C:/SAS/agent/env/Scripts/rst2odt.py'
rst2odtpath = os.path.join(agentdir, 'env', 'Scripts', 'rst2odt.py')

# Parametri di configurazione GWS
gws_template_dir = basedir + os.path.sep + 'agent' + os.path.sep + 'gws' + os.path.sep + 'users.csv'
gws_receiver = 'angela.alaimo@iccarvico.edu.it'