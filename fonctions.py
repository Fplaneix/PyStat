import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import locale
import seaborn as sns
import lxml
from io import StringIO
from html5lib import *
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches
from PIL import Image
from unicodedata import normalize
from bs4 import BeautifulSoup
from matplotlib import figure
sns.set()
#Récupération du fichier HTML transposition et formatage
def import_data():
    global s75_fp_fichier_agent
    s75_fp_fichier_agent = 'F:\DevPyt\PyStat\html_data\S75-ARM_03.html'
    global s75_fp_fichier_flux
    s75_fp_fichier_flux ='F:\DevPyt\PyStat\html_data\S75-F15_03.html'
    global s92_fp_fichier_agent
    s92_fp_fichier_agent = 'F:\DevPyt\PyStat\html_data\S92-ARM_03.html'
    global s92_fp_fichier_flux
    s92_fp_fichier_flux ='F:\DevPyt\PyStat\html_data\S92-F15_03.html'
    global s93_fp_fichier_agent
    s93_fp_fichier_agent = 'F:\DevPyt\PyStat\html_data\S93-ARM_03.html'
    global s93_fp_fichier_flux
    s93_fp_fichier_flux ='F:\DevPyt\PyStat\html_data\S93-F15_03.html'
    global s94_fp_fichier_agent
    s94_fp_fichier_agent = 'F:\DevPyt\PyStat\html_data\S94-ARM_03.html'
    global s94_fp_fichier_flux
    s94_fp_fichier_flux ='F:\DevPyt\PyStat\html_data\S94-F15_03.html'
    pass

def test_df(df):
    print("SHAPE => ")
    print(df.shape)
    print("NOM des colones => ")
    print(df.columns)
    print("Types de colones => ")
    print(df.dtypes)
    print(df[0:5])


#collecteur les données depuis un fichier HTML CCPLUSE contenant les agents
def collecter_data_type_agent(file_path):
    data_type_agent = pd.read_html(file_path, decimal=',', thousands='.')
    tab_type_agent = data_type_agent[0]
    tab_type_agent.reindex()
    tab_type_agent["Statistique"] = tab_type_agent["Groupe"].map(str) + " - " + tab_type_agent["Statistique"]
    str_cat = tab_type_agent.loc[0, 'Objet'].split('_')
    type_agent = str_cat[2]
    samu_agent = str_cat[1]
    tab_type_agent = tab_type_agent.drop(columns=['Objet','Groupe'])
    tab_type_agent = tab_type_agent.set_index('Statistique').transpose()
    #tab_type_agent.insert(0, "type_data", type_agent, allow_duplicates=False)
    tab_type_agent.insert(0, "samu", samu_agent, allow_duplicates=False)
    tab_type_agent.reset_index(level=0, inplace=True)
    tab_type_agent.rename(columns={'index': 'date'}, inplace=True)
    return tab_type_agent

#collecteur les données depuis un fichier HTML CCPLUSE contenant les flux
def collecter_data_type_flux(file_path):
    data_type_flux = pd.read_html(file_path, decimal=',', thousands='.')
    tab_type_flux = data_type_flux[0]
    tab_type_flux.reindex()
    tab_type_flux["Statistique"] = tab_type_flux["Groupe"].map(str) + " - " + tab_type_flux["Statistique"]
    str_cat = tab_type_flux.loc[0, 'Objet'].split('_')
    type_flux = str_cat[1]
    samu_flux = str_cat[0]
    tab_type_flux = tab_type_flux.drop(columns=['Objet','Groupe'])
    tab_type_flux = tab_type_flux.set_index('Statistique').transpose()
    #tab_type_flux.insert(0, "type_data", type_flux, allow_duplicates=False)
    tab_type_flux.insert(0, "samu", samu_flux, allow_duplicates=False)
    tab_type_flux.reset_index(level=0, inplace=True)
    tab_type_flux.rename(columns={'index': 'date'}, inplace=True)
    return tab_type_flux



def traitement_data():
    #Fusion de tous les dataframes
    pool_agents = pd.concat([ collecter_data_type_agent(s75_fp_fichier_agent),collecter_data_type_agent(s92_fp_fichier_agent),collecter_data_type_agent(s93_fp_fichier_agent),collecter_data_type_agent(s94_fp_fichier_agent)])
    pool_flux = pd.concat([ collecter_data_type_flux(s75_fp_fichier_flux),collecter_data_type_flux(s92_fp_fichier_flux),collecter_data_type_flux(s93_fp_fichier_flux),collecter_data_type_flux(s94_fp_fichier_flux)])
    #merge des dataframes entre eux
    pool_merge = pd.merge(pool_agents, pool_flux)
    global pool_selected
    #Sélection des colones
    pool_selected = pool_merge[['date', 'samu','Pourcentage en communication - % en com','Nb Agents Logués - Nb Moy Agents horaire','Nombre d\'appels - Entrés','QOS - Nouvelle QOS 60s','Abandons - Abandons 0-15s', 'Abandons - Total Abandons','Efficacite - Eff sans abandons 15s']]

    #Controle des types de chacunes des colones sélectionnées
    #pool_selected['date'] = pool_selected['date'].astype('timest')
    pool_selected['date'] =pd.to_datetime(pool_selected.date, format="%d/%m/%Y %H:%M:%S")
    pool_selected['Nb Agents Logués - Nb Moy Agents horaire'] = pool_selected['Nb Agents Logués - Nb Moy Agents horaire'].astype('float64')
    pool_selected['Pourcentage en communication - % en com'] = pool_selected['Pourcentage en communication - % en com'].astype('float64')
    pool_selected['Nombre d\'appels - Entrés'] = pool_selected['Nombre d\'appels - Entrés'].astype('float64')
    pool_selected['QOS - Nouvelle QOS 60s'] = pool_selected['QOS - Nouvelle QOS 60s'].astype('float64')
    pool_selected['Abandons - Abandons 0-15s'] = pool_selected['Abandons - Abandons 0-15s'].astype('float64')
    pool_selected['Abandons - Total Abandons'] = pool_selected['Abandons - Total Abandons'].astype('float64')
    pool_selected['Efficacite - Eff sans abandons 15s'] = pool_selected['Efficacite - Eff sans abandons 15s'].astype('float64')

    #Ajout des valeurs calculées
    #Injection de la colonne calculée appels à traiter
    pool_selected['Appels à traiter'] = pool_selected['Nombre d\'appels - Entrés'] - pool_selected['Abandons - Abandons 0-15s']
    #Injection de la colonne calculée  jour de la Semaine en FR
    locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
    pool_selected['jour'] = pd.DatetimeIndex(pool_selected['date']).day_name(locale = 'fr_FR.UTF-8')
    pool_selected['jour'] = pool_selected['jour'].astype('string')

    #Injection de la colonne calculée de la tranche horaire
    pool_selected['heures'] =pd.to_datetime(pool_selected['date']).dt.time
    pool_selected['heures'] = pool_selected['heures'].astype('string')
    #Injection de la colonne calculée de la tranche horaire
    pool_selected['Abandons - Abandons +15s'] = pool_selected['Abandons - Total Abandons'] - pool_selected['Abandons - Abandons 0-15s']
    global table_S75
    global table_S92
    global table_S93
    global table_S94
    #Table intermédiaires par SAMU
    table_S75 = pool_selected[pool_selected["samu"] =='S75']
    table_S92 = pool_selected[pool_selected["samu"] =='S92']
    table_S93 = pool_selected[pool_selected["samu"] =='S93']
    table_S94 = pool_selected[pool_selected["samu"] =='S94']
    pass
#création
def build_plot(df,io,samu):
    d_grph =  pd.pivot_table(df, index=['jour','heures'],aggfunc={'Appels à traiter':np.mean,'Abandons - Abandons +15s':np.mean,'QOS - Nouvelle QOS 60s':np.mean})
    d_grph = d_grph.transpose()
    d_grph.columns = [' '.join(col) if type(col) is tuple else col for col in d_grph.columns.values]
    d_grph = d_grph.transpose()
    d_grph = d_grph.reindex(io)
    d_grph.plot(figsize=(12, 12), subplots=True, sharey=False, rot=90,xlabel='Horaire/jour', ylabel='Valeurs', title=samu)
    d_grph.to_excel(r'F:\DevPyt\PyStat\xlsx\\'+samu+'.xlsx')
    pass

def build_toexcel(df,name):
    df = pd.pivot_table(df, index=['jour','heures'],aggfunc={'Appels à traiter':np.mean,'Abandons - Abandons +15s':np.mean,'QOS - Nouvelle QOS 60s':np.mean})
    df.to_excel(r'F:\DevPyt\PyStat\xlsx\\'+name+'.xlsx')
    pass

def b_toexcel(df,name):
    df.to_excel(r'F:\DevPyt\PyStat\xlsx\\'+name+'.xlsx')
    pass


import_data()
traitement_data()


#GRAPHIQUES
index_order = pd.array(['Lundi 00:00:00', 'Lundi 01:00:00', 'Lundi 02:00:00','Lundi 03:00:00', 'Lundi 04:00:00', 'Lundi 05:00:00','Lundi 06:00:00', 'Lundi 07:00:00', 'Lundi 08:00:00','Lundi 09:00:00', 'Lundi 10:00:00', 'Lundi 11:00:00','Lundi 12:00:00', 'Lundi 13:00:00', 'Lundi 14:00:00','Lundi 15:00:00', 'Lundi 16:00:00', 'Lundi 17:00:00','Lundi 18:00:00', 'Lundi 19:00:00', 'Lundi 20:00:00','Lundi 21:00:00', 'Lundi 22:00:00', 'Lundi 23:00:00','Mardi 00:00:00', 'Mardi 01:00:00', 'Mardi 02:00:00','Mardi 03:00:00', 'Mardi 04:00:00', 'Mardi 05:00:00','Mardi 06:00:00', 'Mardi 07:00:00', 'Mardi 08:00:00','Mardi 09:00:00', 'Mardi 10:00:00', 'Mardi 11:00:00','Mardi 12:00:00', 'Mardi 13:00:00', 'Mardi 14:00:00','Mardi 15:00:00', 'Mardi 16:00:00', 'Mardi 17:00:00','Mardi 18:00:00', 'Mardi 19:00:00', 'Mardi 20:00:00','Mardi 21:00:00', 'Mardi 22:00:00', 'Mardi 23:00:00','Mercredi 00:00:00', 'Mercredi 01:00:00', 'Mercredi 02:00:00','Mercredi 03:00:00', 'Mercredi 04:00:00', 'Mercredi 05:00:00','Mercredi 06:00:00', 'Mercredi 07:00:00', 'Mercredi 08:00:00','Mercredi 09:00:00', 'Mercredi 10:00:00', 'Mercredi 11:00:00','Mercredi 12:00:00', 'Mercredi 13:00:00', 'Mercredi 14:00:00','Mercredi 15:00:00', 'Mercredi 16:00:00', 'Mercredi 17:00:00','Mercredi 18:00:00', 'Mercredi 19:00:00', 'Mercredi 20:00:00','Mercredi 21:00:00', 'Mercredi 22:00:00', 'Mercredi 23:00:00','Jeudi 00:00:00', 'Jeudi 01:00:00', 'Jeudi 02:00:00','Jeudi 03:00:00', 'Jeudi 04:00:00', 'Jeudi 05:00:00','Jeudi 06:00:00', 'Jeudi 07:00:00', 'Jeudi 08:00:00','Jeudi 09:00:00', 'Jeudi 10:00:00', 'Jeudi 11:00:00','Jeudi 12:00:00', 'Jeudi 13:00:00', 'Jeudi 14:00:00','Jeudi 15:00:00', 'Jeudi 16:00:00', 'Jeudi 17:00:00','Jeudi 18:00:00', 'Jeudi 19:00:00', 'Jeudi 20:00:00','Jeudi 21:00:00', 'Jeudi 22:00:00', 'Jeudi 23:00:00','Vendredi 00:00:00', 'Vendredi 01:00:00', 'Vendredi 02:00:00','Vendredi 03:00:00', 'Vendredi 04:00:00', 'Vendredi 05:00:00','Vendredi 06:00:00', 'Vendredi 07:00:00', 'Vendredi 08:00:00','Vendredi 09:00:00', 'Vendredi 10:00:00', 'Vendredi 11:00:00','Vendredi 12:00:00', 'Vendredi 13:00:00', 'Vendredi 14:00:00','Vendredi 15:00:00', 'Vendredi 16:00:00', 'Vendredi 17:00:00','Vendredi 18:00:00', 'Vendredi 19:00:00', 'Vendredi 20:00:00','Vendredi 21:00:00', 'Vendredi 22:00:00', 'Vendredi 23:00:00','Samedi 00:00:00', 'Samedi 01:00:00', 'Samedi 02:00:00','Samedi 03:00:00', 'Samedi 04:00:00', 'Samedi 05:00:00','Samedi 06:00:00', 'Samedi 07:00:00', 'Samedi 08:00:00','Samedi 09:00:00', 'Samedi 10:00:00', 'Samedi 11:00:00','Samedi 12:00:00', 'Samedi 13:00:00', 'Samedi 14:00:00','Samedi 15:00:00', 'Samedi 16:00:00', 'Samedi 17:00:00','Samedi 18:00:00', 'Samedi 19:00:00', 'Samedi 20:00:00','Samedi 21:00:00', 'Samedi 22:00:00', 'Samedi 23:00:00','Dimanche 00:00:00', 'Dimanche 01:00:00', 'Dimanche 02:00:00','Dimanche 03:00:00', 'Dimanche 04:00:00', 'Dimanche 05:00:00','Dimanche 06:00:00', 'Dimanche 07:00:00', 'Dimanche 08:00:00','Dimanche 09:00:00', 'Dimanche 10:00:00', 'Dimanche 11:00:00','Dimanche 12:00:00', 'Dimanche 13:00:00', 'Dimanche 14:00:00','Dimanche 15:00:00', 'Dimanche 16:00:00', 'Dimanche 17:00:00','Dimanche 18:00:00', 'Dimanche 19:00:00', 'Dimanche 20:00:00','Dimanche 21:00:00', 'Dimanche 22:00:00', 'Dimanche 23:00:00'],dtype=object)
axe_x_order = pd.array([['Lundi 00:00:00','Mardi 00:00:00','Mercredi 00:00:00','Jeudi 00:00:00','Vendredi 00:00:00','Samedi 00:00:00','Dimanche 00:00:00']])





#build_toexcel(table_S75,'S75')
#build_plot(table_S75,index_order,'SAMU 75')
#build_plot(table_S92,index_order,'SAMU 92')
#build_plot(table_S93,index_order,'SAMU 93')
#build_plot(table_S94,index_order,'SAMU 94')





#
#pool_pp = sns.pairplot(pool_selected,palette="Set2", diag_kind="kde", hue="samu")
#pool_pp.savefig("F:\DevPyt\PyStat\img\Pool_paiplot.png")
#S75_pp = sns.pairplot(table_S75,palette="Set2", diag_kind="kde", hue="samu")
#S75_pp.savefig("F:\DevPyt\PyStat\img\S75_paiplot.png")
#S92_pp = sns.pairplot(table_S92,palette="Set2", diag_kind="kde", hue="samu")
#S92_pp.savefig("F:\DevPyt\PyStat\img\S92_paiplot.png")
#S93_pp = sns.pairplot(table_S93,palette="Set2", diag_kind="kde", hue="samu")
#S93_pp.savefig("F:\DevPyt\PyStat\img\S93_paiplot.png")
#S94_pp = sns.pairplot(table_S94,palette="Set2", diag_kind="kde", hue="samu")
#S94_pp.savefig("F:\DevPyt\PyStat\img\S94_paiplot.png")
