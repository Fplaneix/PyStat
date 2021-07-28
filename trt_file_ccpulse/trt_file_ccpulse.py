import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import locale
import seaborn as sns
import lxml
import os
import glob
from os import path
from io import StringIO
from html5lib import *
from unicodedata import normalize
from bs4 import BeautifulSoup
from matplotlib import figure


#Function de parcourir dossier et récupération de la liste des fichies présents ctrl hmlt file_path
#Fonction détermination si agent ou flux
#Fonction de déplacement de fichier
#Function collecter_data_type_flux collecter_data_type_agent
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

#Fonction traitement_data pool_merge + export xlsx

#Définition des fonctions
#fonction de récupération des fichiers


def typage_data(fichier):
    global pool_agents
    global pool_flux
    data_type = pd.read_html(fichier, decimal=',', thousands='.')
    type_data = data_type[0].Objet[0][0:2]
    type_data
    if data_type[0].Objet[0][0:2] == 'GA':
        data_a = collecter_data_type_agent(fichier)
        pool_agents = pd.concat([data_a])
        return pool_agents
    else :
        data_f = collecter_data_type_flux(fichier)
        pool_flux = pd.concat([data_f])
        return pool_flux

def compilation_data():
    global pool_merge
    path_dossier_fichiers_source =r'F:\DevPyt\PyStat\trt_file_ccpulse\in'
    path_dossier_fichiers_source = os.chdir(path_dossier_fichiers_source)
    list_fichiers_html = os.listdir(path_dossier_fichiers_source)
    for fichier in list_fichiers_html:
        typage_data(fichier)
    pool_merge = pd.merge(pool_agents, pool_flux)
    pool_merge['date'] =pd.to_datetime(pool_merge.date, format="%d/%m/%Y %H:%M:%S")
    pool_merge['Temps total en communication - Tps Entrant'] = pd.to_datetime(pool_merge['Temps total en communication - Tps Entrant']).dt.time
    pool_merge['Temps total en communication - Tps Sortant'] = pd.to_datetime(pool_merge['Temps total en communication - Tps Sortant']).dt.time
    pool_merge['Temps total en communication - Tps Interne'] = pd.to_datetime(pool_merge['Temps total en communication - Tps Interne']).dt.time
    pool_merge['Temps total en communication - Tps mise en attente'] = pd.to_datetime(pool_merge['Temps total en communication - Tps mise en attente']).dt.time
    pool_merge['Temps total en communication - Tps en double appels'] = pd.to_datetime(pool_merge['Temps total en communication - Tps en double appels']).dt.time
    pool_merge['Temps total par états - Logué'] = pd.to_datetime(pool_merge['Temps total par états - Logué']).dt.time
    pool_merge['Temps total par états - Disponible'] = pd.to_datetime(pool_merge['Temps total par états - Disponible']).dt.time
    pool_merge['Temps total par états - Pause Admin'] = pd.to_datetime(pool_merge['Temps total par états - Pause Admin']).dt.time
    pool_merge['Temps total par états - Indisponible'] = pd.to_datetime(pool_merge['Temps total par états - Indisponible']).dt.time
    pool_merge['Temps total par états - Retrait'] = pd.to_datetime(pool_merge['Temps total par états - Retrait']).dt.time
    pool_merge['Temps total par états - Fermeture position'] = pd.to_datetime(pool_merge['Temps total par états - Fermeture position']).dt.time
    pool_merge['Temps total par états - Autres retraits'] = pd.to_datetime(pool_merge['Temps total par états - Autres retraits']).dt.time
    pool_merge['Temps total par états - En communication'] = pd.to_datetime(pool_merge['Temps total par états - En communication']).dt.time
    pool_merge['Temps d\'attente - Total avant réponse'] = pd.to_datetime(pool_merge['Temps d\'attente - Total avant réponse']).dt.time
    pool_merge['Temps d\'attente - Max avant réponse'] = pd.to_datetime(pool_merge['Temps d\'attente - Max avant réponse']).dt.time
    pool_merge['Temps d\'attente - Moy avant réponse en s'] = pool_merge['Temps d\'attente - Moy avant réponse en s'].astype('float64')
    pool_merge['Temps total en communication - Total h/h'] = pool_merge['Temps total en communication - Total h/h'].astype('float64')
    pool_merge['Temps moyen en communication - Temps moyen en s'] =  pool_merge['Temps moyen en communication - Temps moyen en s'].astype('float64')
    pool_merge['Nb Agents Logués - Nb Moy Agents horaire'] = pool_merge['Nb Agents Logués - Nb Moy Agents horaire'].astype('float64')
    pool_merge['Pourcentage en communication - % en com'] = pool_merge['Pourcentage en communication - % en com'].astype('float64')
    pool_merge['Nombre d\'appels - Entrés'] = pool_merge['Nombre d\'appels - Entrés'].astype('float64')
    pool_merge['QOS - Nouvelle QOS 60s'] = pool_merge['QOS - Nouvelle QOS 60s'].astype('float64')
    pool_merge['Abandons - Abandons 0-15s'] = pool_merge['Abandons - Abandons 0-15s'].astype('float64')
    pool_merge['Abandons - Total Abandons'] = pool_merge['Abandons - Total Abandons'].astype('float64')
    pool_merge['Efficacite - Eff sans abandons 15s'] = pool_merge['Efficacite - Eff sans abandons 15s'].astype('float64')
    pool_merge['Nombre total d\'appels - Nb Entrant Ext'] = pool_merge['Nombre total d\'appels - Nb Entrant Ext'].astype('float64')
    pool_merge['Nombre total d\'appels - Nb Sortant Ext'] = pool_merge['Nombre total d\'appels - Nb Sortant Ext'].astype('float64')
    pool_merge['Nombre total d\'appels - Nb Entrant Int'] = pool_merge['Nombre total d\'appels - Nb Entrant Int'].astype('float64')
    pool_merge['Nombre total d\'appels - Nb Sortant Int'] = pool_merge['Nombre total d\'appels - Nb Sortant Int'].astype('float64')
    pool_merge['Nombre total d\'appels - Nb Double Appels Effectué'] = pool_merge['Nombre total d\'appels - Nb Double Appels Effectué'].astype('float64')
    pool_merge['Nombre total d\'appels - Nb Double Appels Arrivé'] = pool_merge['Nombre total d\'appels - Nb Double Appels Arrivé'].astype('float64')
    pool_merge['Nombre total d\'appels - Nb Transfert Effectué'] = pool_merge['Nombre total d\'appels - Nb Transfert Effectué'].astype('float64')
    pool_merge['Nombre total d\'appels - Nb Transfert Arrivé'] = pool_merge['Nombre total d\'appels - Nb Transfert Arrivé'].astype('float64')
    pool_merge['Nombre total d\'appels - Nb Conférence'] = pool_merge['Nombre total d\'appels - Nb Conférence'].astype('float64')
    pool_merge['Nombre total d\'appels - Nb Rona'] = pool_merge['Nombre total d\'appels - Nb Rona'].astype('float64')
    pool_merge['Nombre d\'appels - Distribués'] = pool_merge['Nombre d\'appels - Distribués'].astype('float64')
    pool_merge['Nombre d\'appels - Répondus'] = pool_merge['Nombre d\'appels - Répondus'].astype('float64')
    pool_merge['Nombre d\'appels distribué en X secondes - 20s'] = pool_merge['Nombre d\'appels distribué en X secondes - 20s'].astype('float64')
    pool_merge['Nombre d\'appels distribué en X secondes - 60s'] = pool_merge['Nombre d\'appels distribué en X secondes - 60s'].astype('float64')
    pool_merge['Nombre d\'appels distribué en X secondes - 61 sec et +'] = pool_merge['Nombre d\'appels distribué en X secondes - 61 sec et +'].astype('float64')
    pool_merge['QOS - Nouvelle QOS 20s'] = pool_merge['QOS - Nouvelle QOS 20s'].astype('float64')
    pool_merge['Abandons - Abandons 0-40s'] = pool_merge['Abandons - Abandons 0-40s'].astype('float64')
    pool_merge['Abandons - Abandons 41s et +'] = pool_merge['Abandons - Abandons 41s et +'].astype('float64')
    pool_merge['Efficacite - Eff sans abandons 40s'] = pool_merge['Efficacite - Eff sans abandons 40s'].astype('float64')
    pool_merge['Nb Agents Logués - Nb Moy Agents pas en retrait'] = pool_merge['Nb Agents Logués - Nb Moy Agents pas en retrait'].astype('float64')
    pool_merge['Temps total en communication - Total h/h'] = pool_merge['Temps total en communication - Total h/h'].astype('float64')
    pool_merge['Temps d\'attente - Moy avant réponse en s'] = pool_merge['Temps d\'attente - Moy avant réponse en s'].astype('float64')
    pool_merge['Pourcentage en communication - % en com'] = pool_merge['Pourcentage en communication - % en com'].astype('float64')
    pool_merge['Nb Agents Logués - Nb Moy Agents horaire'] = pool_merge['Nb Agents Logués - Nb Moy Agents horaire'].astype('float64')
    pool_merge['Nb Agents Logués - Nb Moy Agents pas en retrait'] = pool_merge['Nb Agents Logués - Nb Moy Agents pas en retrait'].astype('float64')
    pool_merge['Nombre d\'appels - Entrés'] = pool_merge['Nombre d\'appels - Entrés']
    pool_merge['Nombre d\'appels - Distribués'] = pool_merge['Nombre d\'appels - Distribués'].astype('float64')
    pool_merge['Nombre d\'appels - Répondus'] = pool_merge['Nombre d\'appels - Répondus'].astype('float64')
    pool_merge['Nombre d\'appels distribué en X secondes - 20s'] = pool_merge['Nombre d\'appels distribué en X secondes - 20s'].astype('float64')
    pool_merge['Nombre d\'appels distribué en X secondes - 60s'] = pool_merge['Nombre d\'appels distribué en X secondes - 60s'].astype('float64')
    pool_merge['Nombre d\'appels distribué en X secondes - 61 sec et +'] = pool_merge['Nombre d\'appels distribué en X secondes - 61 sec et +'].astype('float64')



def b_toexcel(df,name):
    df.to_excel(r'F:\DevPyt\PyStat\trt_file_ccpulse\save\\'+name+'.xlsx')
    pass


compilation_data()
pool_merge.shape
pool_merge

b_toexcel(pool_merge,'data_pool')
