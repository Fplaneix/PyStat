import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from io import StringIO
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches
from PIL import Image
import seaborn as sns

data_source = pd.read_excel ('F:\DevPyt\PyStat\data.xlsx')

data_source.shape
data_source.columns
data_source.head()

data_EA = data_source[['SAMU','Jour Semaine','Appel à traiter','Heures','DATE','Nombre d\'appels Entrés','Temps d\'attente Moy avant réponse en s','Abandons Abandons 15s et +','Nb Agents Logués Nb Moy Agents horaire','Pourcentage en communication % en com', 'Efficacite Eff sans abandons 15s','Temps moyen en communication Temps moyen en s','QOS Nouvelle QOS 60s']]

data_EA.shape
data_EA.columns
data_EA.dtypes
data_EA[0:2]
data_agg = data_EA.groupby(['SAMU','Jour Semaine','Heures']).mean()[['Appel à traiter','Abandons Abandons 15s et +','QOS Nouvelle QOS 60s']].reset_index()
data_agg[0:2]
data_agg.plot()

data_source[0:2]
data_source.dtypes
table = pd.pivot_table(data_source, index=['Jour Semaine','Heures'],columns=['SAMU'],aggfunc={'Appel à traiter':np.mean,'Abandons Abandons 15s et +':np.mean,'QOS Nouvelle QOS 60s':np.mean})
table.plot()

table_S75 = data_agg[data_agg["SAMU"] =='S75']
table_S75
S75_plot = pd.pivot_table(table_S75, index=['Jour Semaine','Heures'],columns=['SAMU'],aggfunc={'Appel à traiter':np.mean,'Abandons Abandons 15s et +':np.mean,'QOS Nouvelle QOS 60s':np.mean})
S75_plot.plot()

table_S92 = data_agg[data_agg["SAMU"] =='S92']
table_S92
S92_plot = pd.pivot_table(table_S92, index=['Jour Semaine','Heures'],columns=['SAMU'],aggfunc={'Appel à traiter':np.mean,'Abandons Abandons 15s et +':np.mean,'QOS Nouvelle QOS 60s':np.mean})
S92_plot.plot()

table_S93 = data_agg[data_agg["SAMU"] =='S93']
table_S93
S93_plot = pd.pivot_table(table_S93, index=['Jour Semaine','Heures'],columns=['SAMU'],aggfunc={'Appel à traiter':np.mean,'Abandons Abandons 15s et +':np.mean,'QOS Nouvelle QOS 60s':np.mean})
S93_plot.plot().legend(loc='center left',bbox_to_anchor=(1.0, 0.5))

table_S94 = data_agg[data_agg["SAMU"] =='S94']
table_S94
S94_plot = pd.pivot_table(table_S94, index=['Jour Semaine','Heures'],columns=['SAMU'],aggfunc={'Appel à traiter':np.mean,'Abandons Abandons 15s et +':np.mean,'QOS Nouvelle QOS 60s':np.mean})
S94_plot.plot()


table_S92_S94 = pd.concat([table_S94, table_S92])
table_S92_S94
S92_S94_plot = pd.pivot_table(table_S92_S94, index=['Jour Semaine','Heures'],columns=['SAMU'],aggfunc={'Appel à traiter':np.mean,'Abandons Abandons 15s et +':np.mean,'QOS Nouvelle QOS 60s':np.mean})

S92_S94_plot.plot().legend(loc='center left',bbox_to_anchor=(1.0, 0.5))

#Seaborn
sns.set_theme(style="darkgrid")
tips = sns.load_dataset("tips")
sns.relplot(x="total_bill", y="tip", data=tips);
table_S94.columns
sns.pairplot(data_EA,palette="Set2", diag_kind="kde", hue="SAMU")




prs.save('F:\DevPyt\PyStat\chart-01.pptx')
