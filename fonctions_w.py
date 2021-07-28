import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from io import StringIO
from html5lib import *
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches
from PIL import Image
import seaborn as sns
from unicodedata import normalize
import lxml
from bs4 import BeautifulSoup

#Récupération du fichier HTML transposition et formatage
s75_fp_fichier_agent = 'F:\DevPyt\PyStat\html_data\S75-ARM_03.html'
s75_fp_fichier_flux ='F:\DevPyt\PyStat\html_data\S75-F15_03.html'
s92_fp_fichier_agent = 'F:\DevPyt\PyStat\html_data\S92-ARM_03.html'
s92_fp_fichier_flux ='F:\DevPyt\PyStat\html_data\S92-F15_03.html'
s93_fp_fichier_agent = 'F:\DevPyt\PyStat\html_data\S93-ARM_03.html'
s93_fp_fichier_flux ='F:\DevPyt\PyStat\html_data\S93-F15_03.html'
s94_fp_fichier_agent = 'F:\DevPyt\PyStat\html_data\S94-ARM_03.html'
s94_fp_fichier_flux ='F:\DevPyt\PyStat\html_data\S94-F15_03.html'



def collecter_data_type_agent(file_path):
    data_type_agent = pd.read_html(file_path, decimal=',', thousands='.')
    tab_type_agent = data_type_agent[0]
    tab_type_agent["Statistique"] = tab_type_agent["Groupe"].map(str) + " - " + tab_type_agent["Statistique"]
    str_cat = tab_type_agent.loc[0, 'Objet'].split('_')
    type_agent = str_cat[2]
    samu_agent = str_cat[1]
    tab_type_agent = tab_type_agent.drop(columns=['Objet','Groupe'])
    tab_type_agent.insert(0, "type_data", type_agent, allow_duplicates=False)
    tab_type_agent.insert(0, "samu", samu_agent, allow_duplicates=False)
    return tab_type_agent





def collecter_data_type_flux(file_path):
    data_type_flux = pd.read_html(file_path, decimal=',', thousands='.')
    tab_type_flux = data_type_flux[0]
    tab_type_flux["Statistique"] = tab_type_flux["Groupe"].map(str) + " - " + tab_type_flux["Statistique"]
    str_cat = tab_type_flux.loc[0, 'Objet'].split('_')
    type_flux = str_cat[1]
    samu_flux = str_cat[0]
    tab_type_flux = tab_type_flux.drop(columns=['Objet','Groupe'])
    tab_type_flux.insert(0, "type_data", type_flux, allow_duplicates=False)
    tab_type_flux.insert(0, "samu", samu_flux, allow_duplicates=False)
    return tab_type_flux

#Fusion de tous les dataframes
pool = pd.concat([ collecter_data_type_agent(s75_fp_fichier_agent),collecter_data_type_flux(s75_fp_fichier_flux),collecter_data_type_agent(s92_fp_fichier_agent),collecter_data_type_flux(s92_fp_fichier_flux),collecter_data_type_agent(s93_fp_fichier_agent),collecter_data_type_flux(s93_fp_fichier_flux),collecter_data_type_agent(s94_fp_fichier_agent),collecter_data_type_flux(s94_fp_fichier_flux)])
pool[0:10]

sns.pairplot(pool, vars=pool.columns[:-1],hue="y")
