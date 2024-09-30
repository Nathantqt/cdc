
Scoring : 
# -*- coding: utf-8 -*-
"""
Created on Wed Apr  5 10:11:37 2023

@author: amorlat
"""
#pilotage.py

import numpy as np
import pandas as pd
import xlwings as xw
import os
import seaborn as sns
#from openpyxl import load_workbook

import sys
import matplotlib.pyplot as plt
import sklearn
from sklearn.decomposition import PCA
import plotly.express as px

#fonction pour calculer le score en souhaitant maximiser ce critère
def high_score(critere,df,percent_high,percent_low,i):
    rep=0
    if df[critere][i]>np.percentile(np.array(df[critere].dropna()),percent_high):
            rep=1
    if df[critere][i]<np.percentile(np.array(df[critere].dropna()),percent_low):
            rep=-1
    if df[critere][i]>np.percentile(np.array(df[critere].dropna()),percent_low) and df[critere][i]<np.percentile(np.array(df[critere].dropna()),percent_high):
            rep=0
    return rep

#fonction pour calculer le score en souhaitant minimiser ce critère
def low_score(critere, df,percent_high,percent_low,i):
    rep=0
    if df[critere][i]>np.percentile(np.array(df[critere].dropna()),percent_high):
            rep=-1
    if df[critere][i]<np.percentile(np.array(df[critere].dropna()),percent_low):
            rep=1
    if df[critere][i]>np.percentile(np.array(df[critere].dropna()),percent_low) and df[critere][i]<np.percentile(np.array(df[critere].dropna()),percent_high):
            rep=0
    return rep

def normalize(vec):
    moy=np.mean(vec.dropna())
    stda=np.std(vec.dropna())
    rep=[]
    if pd.isna(stda):
        for i in vec:
            rep.append(0)
    elif stda ==0:
        for i in vec:
            rep.append(0)
    else :
        for i in vec:
            rep.append((i-moy)/stda)
    return rep

#fonction qui regarde si la clé en paramètre appartient à l'ensemble des clés du dictionnaire en paramètre
def dicMemberCheck(key, dicObj):
    if key in dicObj:
        return True
    else:
        return False

def calculate_score(df, percent_high, percent_low, selection_criteres=None, coeff=None, new_criteria=None, score_type=None):
    n = len(df["ISIN"])
    df["Score " + score_type + str(percent_high)] = 0.0
    
    for i in range(n):
        total_score = 0
        j = 0
        
        for critere in selection_criteres:
            if dicMemberCheck(critere, new_criteria):
                if new_criteria[critere] == 'L':
                    total_score += coeff[j] * low_score(critere, df, percent_high, percent_low, i)
                elif new_criteria[critere] == 'H':
                    total_score += coeff[j] * high_score(critere, df, percent_high, percent_low, i)
                else:
                    print("Erreur pour le critere", score_type, critere, "l'objectif n'est pas L ou H")
            j += 1
        
        rep = total_score / sum(coeff)
        df.loc[i, "Score " + score_type + str(percent_high)] = rep

def feuille_existe(classeur, nom_feuille):
    for feuille in classeur.sheets:
        if feuille.name == nom_feuille:
            return True
    return False

def CreaFeuille(wb,nom_feuille):
    if feuille_existe(wb, nom_feuille):
        resultats = wb.sheets[nom_feuille]
    else:
        resultats = wb.sheets.add(nom_feuille) #ajout d'une feuille
    return resultats

def SuppFeuille(wb,sheet_name):
    if sheet_name in [sheet.name for sheet in wb.sheets]:
        sheet_to_delete = wb.sheets[sheet_name]
        sheet_to_delete.delete()
        
def start(nomExcel,chemin_init,nomFichier,SeuilMin,SeuilMax,SeuilPas,PoidPerf,PoidRisk,PoidESG,SelectedCategory,ValAdjustFees,Perso = False,list_criteria_Perf = None, list_criteria_Risk = None, list_criteria_ESG = None, coeff_perf = None, coeff_risk = None, coeff_esg = None, dic_unknow_criteria_perf = None, dic_unknow_criteria_risk = None, dic_unknow_criteria_esg = None, graphique_option = False):
    print(chemin_init,nomFichier,SeuilMin,SeuilMax,SeuilPas,PoidPerf,PoidRisk,PoidESG,SelectedCategory,ValAdjustFees, coeff_perf, coeff_risk, coeff_esg, dic_unknow_criteria_perf, dic_unknow_criteria_risk, dic_unknow_criteria_esg, graphique_option)
    chemin_fichier = os.path.join(chemin_init+"\Extraction_Morningstar", nomFichier)
    print("Voici le chemin de nos fichiers",chemin_fichier)
    app2 = xw.App(visible=False) #ouvre la feuille sans l'afficher à l'utilisateur

    try:
        wb2 = xw.Book(chemin_fichier)
    except Exception as e:
        print("Une erreur s'est produite lors de l'ouverture du fichier Excel, fichier inconnu ")
        
    # Sélection de la feuille des fonds
    #if nomFichier == 'emd_global.xlsx':
    #    feuille = wb2.sheets['EMD Global (Test)  + ESG'] 
    #elif nomFichier == 'hy_europe.xlsx':
    #    feuille = wb2.sheets['HY Europe (test)'] 
    #else :
    feuille = wb2.sheets[0]
    
    #sélection des sous catégories des fonds
    plage_donnees = feuille.range('A1').expand()
    df_score = plage_donnees.options(pd.DataFrame, index=False, header=True).value
    if SelectedCategory == 'Toutes Catégories' or SelectedCategory == '':
        df_score = df_score
    else:
        try:
            df_score = df_score.loc[df_score['Morningstar Category'] == SelectedCategory]
            df_score = df_score.reset_index(drop=True)
        except KeyError:
            print("La colonne 'Morningstar Category' n'existe pas dans le DataFrame.")

    #fermeture du classeur et de l'application xlwings
    wb2.close()
    app2.quit()
    ValAdjustFees = ValAdjustFees/10

    if(Perso == False):
        main(nomExcel,chemin_init,df_score,SeuilMin,SeuilMax,SeuilPas,PoidPerf,PoidRisk,PoidESG,ValAdjustFees)
    else:
        main(nomExcel,chemin_init,df_score,SeuilMin,SeuilMax,SeuilPas,PoidPerf,PoidRisk,PoidESG,ValAdjustFees, True,list_criteria_Perf, list_criteria_Risk, list_criteria_ESG, coeff_perf, coeff_risk, coeff_esg, dic_unknow_criteria_perf, dic_unknow_criteria_risk, dic_unknow_criteria_esg, graphique_option)
    
def main(nomExcel,chemin_init,df_data,SeuilMin,SeuilMax,SeuilPas,PoidPerf,PoidRisk,PoidESG,ValAdjustFees, Perso = False, list_criteria_Perf = None, list_criteria_Risk = None, list_criteria_ESG = None, coeff_perf = None, coeff_risk = None, coeff_esg = None, dic_unknow_criteria_perf = None, dic_unknow_criteria_risk = None, dic_unknow_criteria_esg = None, graphique_option=False):
    cwd = chemin_init
    chemin_pilotage = nomExcel
    app = xw.App(visible=False)
    wb = xw.Book(chemin_pilotage)
    SuppFeuille(wb, "Données")
    SuppFeuille(wb, "Graphiques")
    SuppFeuille(wb, "Résultats")
    sheet = wb.sheets['Classements']
    #trouver la dernière ligne non vide dans la colonne A (ou une autre colonne de référence)
    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    #effacer le contenu de toutes les cellules sauf la première ligne
    sheet.range(f'A2:Z{last_row}').clear_contents()
    
    #feuille Données
    feuille_data = CreaFeuille(wb, 'Données')
    feuille_data.range('A1').value = df_data
    df_score = df_data.copy()


    #y=pd.DataFrame()
    nom_fichier = nomExcel
    feuille = 'Sélection des critères'
    wb = xw.Book(nom_fichier)
    feuille = wb.sheets[feuille]
    error_message_2 = ""
    error_0_criteria = False
    #Lever les erreurs lorsque le paramètre n'existe pas dans le fichier
    #Cas des erreurs dans perf
    criteria_errors = [] #contiendra l'ensemble des critères inexistants des trois catégories 
    new_criteria_perf = dict() #dictionnaire qui contient les critères de perf dont on ne connais pas la façon de calculer le score mais bien présent dans le dataset
    if(dic_unknow_criteria_perf is not None): #si il existe des critères non reconnus
        for key, value in dic_unknow_criteria_perf.items(): #on parcourt le dictionnaire et si une clé n'est pas présente dans notre dataset, je la supprime ainsi que son coefficient dans le tableau des coefficients
            if key not in df_score.columns or (key in df_score.columns and (df_score[key].nunique() < 2)) :
                criteria_errors.append(key)
                indice = list_criteria_Perf.index(key) 
                list_criteria_Perf.remove(key)
                del coeff_perf[indice]
            else: 
                print(key, "est valide") 
                new_criteria_perf[key] = value
                

    if(len(new_criteria_perf)==0):
        print("ERREUR, la list perf est vide, mettez au moins un critère valide dans perf")
        error_message_2 = "la list perf ne comporte aucun critère " 
        error_0_criteria = True
    #Cas des erreurs dans risques
    new_criteria_risk = dict() #dictionnaire qui contient les critères de risk dont on ne connais pas la façon de calculer le score mais bien présent dans le dataset
    if(dic_unknow_criteria_risk is not None): #si il existe des critères non reconnus
        for key, value in dic_unknow_criteria_risk.items():#on parcourt le dictionnaire et si une clé n'est pas présente dans notre dataset, je la supprime ainsi que son coefficient dans le tableau des coefficients
            if key not in df_score.columns or (key in df_score.columns and (df_score[key].nunique() < 2)) :
                criteria_errors.append(key)
                indice = list_criteria_Risk.index(key)
                list_criteria_Risk.remove(key)
                del coeff_risk[indice]
            else: 
                print(key, "est valide") 
                new_criteria_risk[key] = value
    
    if(len(new_criteria_risk)==0):
        print("ERREUR, la list risk est vide, mettez au moins un critère valide dans risk")
        error_message_2 = error_message_2 + ", la list risk ne comporte aucun critère "
        error_0_criteria = True

    #Cas des erreurs dans ESG
    new_criteria_esg = dict() #dictionnaire qui contient les critères d'esg dont on ne connais pas la façon de calculer le score mais bien présent dans le dataset
    if(dic_unknow_criteria_esg is not None): #si il existe des critères non reconnus
        for key, value in dic_unknow_criteria_esg.items(): #on parcourt le dictionnaire et si une clé n'est pas présente dans notre dataset, je la supprime ainsi que son coefficient dans le tableau des coefficients
            if key not in df_score.columns or (key in df_score.columns and (df_score[key].nunique() < 2)) :
                criteria_errors.append(key)
                indice = list_criteria_ESG.index(key)
                list_criteria_ESG.remove(key)
                del coeff_esg[indice]
            else: 
                print(key, "est valide") 
                new_criteria_esg[key] = value
    if(len(new_criteria_esg)==0):
        print("ERREUR, la list esg est vide, mettez au moins un critère valide dans esg")
        error_message_2 = error_message_2 + ", la list esg ne comporte aucun critère"
        error_0_criteria = True
    error_message = "Les critères suivants ne sont pas présents dans le fichier Morning star: " + ", ".join(criteria_errors)
    feuille['M6'].value = error_message #on affiche les erreurs dans la case M6
    feuille['M14'].value = error_message_2

    if(error_0_criteria== False):

        tab_column= []
        tab_column.extend(new_criteria_perf.keys())
        tab_column.extend(new_criteria_risk.keys())
        tab_column.extend(new_criteria_esg.keys())
        x=pd.DataFrame()
        for i in tab_column:
            x[i]=normalize(df_score[i])
        x.fillna(0,inplace=True)
    
        for i in np.arange(SeuilMin,SeuilMax,SeuilPas): #on parcours nos seuils
            if Perso :
                calculate_score(df_score, i, 100-i, list_criteria_Perf, coeff_perf, new_criteria_perf, "Perf")
                calculate_score(df_score, i, 100-i, list_criteria_Risk, coeff_risk, new_criteria_risk, "Risks")
                calculate_score(df_score, i, 100-i, list_criteria_ESG, coeff_esg, new_criteria_esg, "ESG")
        noms = df_score["Name"]
        df_score=df_score.set_index('Name')
        #rajoute des colonnes pour calculer les moyennes des scores de chaque dataset
        df_score["Score_Perf_Moy"] = 0.0
        df_score["Score_Risks_Moy"] = 0.0
        df_score["Score_ESG_Moy"] = 0.0
        for i in df_score.index:
            score_perf_columns = [col for col in df_score.columns if col.startswith("Score Perf")]
            df_score["Score_Perf_Moy"] = df_score[score_perf_columns].mean(axis=1)
            score_risks_columns = [col for col in df_score.columns if col.startswith("Score Risks")]
            df_score["Score_Risks_Moy"] = df_score[score_risks_columns].mean(axis=1)
            score_esg_columns = [col for col in df_score.columns if col.startswith("Score ESG")]
            df_score["Score_ESG_Moy"] = df_score[score_esg_columns].mean(axis=1)

        ######### Calcul of R^2 with Alpha
        def R2_Alpha_Calc(abscisse, ordonnee):
            Alpha_r_df=df_score[[ordonnee,abscisse,"Score Moyenne"]].dropna()
            
            Alpha_r_df["Score quali"] = ''
            Alpha_r_df.index = np.arange(0, len(Alpha_r_df), 1)
            
            for i in range(len(Alpha_r_df["Score Moyenne"])):
                score_moyenne = Alpha_r_df["Score Moyenne"][i]
                score_percentiles = np.array(Alpha_r_df["Score Moyenne"])
            
                if score_moyenne >= np.percentile(score_percentiles, 75):
                    Alpha_r_df.loc[i, "Score quali"] = "Très bon"
                elif score_moyenne < np.percentile(score_percentiles, 75) and score_moyenne >= np.percentile(score_percentiles, 50):
                    Alpha_r_df.loc[i, "Score quali"] = "Bon"
                elif score_moyenne < np.percentile(score_percentiles, 50) and score_moyenne >= np.percentile(score_percentiles, 25):
                    Alpha_r_df.loc[i, "Score quali"] = "Moyen"
                elif score_moyenne < np.percentile(score_percentiles, 25):
                    Alpha_r_df.loc[i, "Score quali"] = "Mauvais"
            return Alpha_r_df

        ######### Plot of R^2 with Alpha
        def R2_Alpha_Fig(Alpha_r_df, abscisse, ordonnee):
            name = abscisse+"_" +ordonnee+".png"
            print(abscisse)
            print(ordonnee)
            fig = plt.figure(figsize = (8,8))
            ax = fig.add_subplot(1,1,1) 
            ax.set_xlabel(abscisse, fontsize = 15)
            ax.set_ylabel(ordonnee, fontsize = 15)
            targets = ['Très bon','Bon', 'Moyen', 'Mauvais']
            colors = ['r', 'g', 'black','blue']
            for target, color in zip(targets,colors):
                indicesToKeep = Alpha_r_df["Score quali"] == target
                ax.scatter(Alpha_r_df.loc[indicesToKeep, abscisse]
                            , Alpha_r_df.loc[indicesToKeep, ordonnee]
                            , c = color
                            , s = 50)
                for i in range(Alpha_r_df.shape[0]):
                    plt.text(x=Alpha_r_df[abscisse][i],y=Alpha_r_df[ordonnee][i],s=str(i), fontdict=dict(color='black',size=10))
            ax.legend(targets)
            ax.grid()
            plt.savefig(os.path.join(cwd+"\Graphiques",name))
        
        ######### Calcul of PCA
        def PCA_Calc():
            pca = PCA(n_components=2)
            principalComponents = pca.fit_transform(x)
            principalDf = pd.DataFrame(data=principalComponents, columns=['principal component 1', 'principal component 2'])
            
            df_score2 = df_score.copy(deep=True)
            df_score2["Score quali"] = ''
            df_score2.index = np.arange(0, len(df_score2), 1)
            
            for i in range(0, len(df_score2["Score Moyenne"])):
                if df_score2.loc[i, "Score Moyenne"] >= np.percentile(np.array(df_score2["Score Moyenne"]), 75):
                    df_score2.loc[i, "Score quali"] = "Très bon"
                elif df_score2.loc[i, "Score Moyenne"] < np.percentile(np.array(df_score2["Score Moyenne"]), 75) and df_score2.loc[i, "Score Moyenne"] >= np.percentile(np.array(df_score2["Score Moyenne"]), 50):
                    df_score2.loc[i, "Score quali"] = "Bon"
                elif df_score2.loc[i, "Score Moyenne"] < np.percentile(np.array(df_score2["Score Moyenne"]), 50) and df_score2.loc[i, "Score Moyenne"] >= np.percentile(np.array(df_score2["Score Moyenne"]), 25):
                    df_score2.loc[i, "Score quali"] = "Moyen"
                else:
                    df_score2.loc[i, "Score quali"] = "Mauvais"
            finalDf=pd.concat([principalDf,df_score2[["Score quali"]]],axis=1)
            #print(pca.explained_variance_ratio_)
            return(finalDf)
        
        ######### Figure of PCA
        def PCA_Fig(finalDf):
            fig = plt.figure(figsize = (8,8))
            ax = fig.add_subplot(1,1,1) 
            ax.set_xlabel('Principal Component 1', fontsize = 15)
            ax.set_ylabel('Principal Component 2', fontsize = 15)
            ax.set_title('2 component PCA', fontsize = 20)
            targets = ['Très bon','Bon', 'Moyen', 'Mauvais']
            colors = ['r', 'g', 'black','blue']
            for target, color in zip(targets,colors):
                indicesToKeep = finalDf['Score quali'] == target
                ax.scatter(finalDf.loc[indicesToKeep, 'principal component 1']
                            , finalDf.loc[indicesToKeep, 'principal component 2']
                            , c = color
                            , s = 50)
            ax.legend(targets)
            ax.set_xlim(-5,5)
            ax.grid()
            plt.savefig(os.path.join(cwd+"\Graphiques",'PCA_2D.png'))
        
        ######### Calcul of Score_Perf_Moy with Score_Risks_Moy
        def Perf_Risk_Calc():
            Perf_risk_df=df_score[["ISIN","Score_Perf_Moy","Score_Risks_Moy","Score Moyenne"]].copy()
            Perf_risk_df.sort_values(by=["Score Moyenne"],inplace=True,ascending=False)
            #Version AM
            Perf_risk_df["Score quali"] = ''
            Perf_risk_df.index = np.arange(0, len(Perf_risk_df), 1)
            for i in range(0, len(Perf_risk_df["Score Moyenne"])):
                if Perf_risk_df.loc[i, "Score Moyenne"] >= np.percentile(np.array(Perf_risk_df["Score Moyenne"]), 75):
                    Perf_risk_df.loc[i, "Score quali"] = "Très bon"
                elif Perf_risk_df.loc[i, "Score Moyenne"] < np.percentile(np.array(Perf_risk_df["Score Moyenne"]), 75) and Perf_risk_df.loc[i, "Score Moyenne"] >= np.percentile(np.array(Perf_risk_df["Score Moyenne"]), 50):
                    Perf_risk_df.loc[i, "Score quali"] = "Bon"
                elif Perf_risk_df.loc[i, "Score Moyenne"] < np.percentile(np.array(Perf_risk_df["Score Moyenne"]), 50) and Perf_risk_df.loc[i, "Score Moyenne"] >= np.percentile(np.array(Perf_risk_df["Score Moyenne"]), 25):
                    Perf_risk_df.loc[i, "Score quali"] = "Moyen"
                else:
                    Perf_risk_df.loc[i, "Score quali"] = "Mauvais"
            return(Perf_risk_df)

            
        ######## Plot Score_Perf_Moy with Score_Risks_Moy
        def Perf_Risk_Fig(Perf_risk_df):
            fig = plt.figure(figsize = (8,8))
            ax = fig.add_subplot(1,1,1) 
            ax.set_xlabel('Score_Perf_Moy', fontsize = 15)
            ax.set_ylabel('Score_Risks_Moy', fontsize = 15)
            ax.set_title('Scatter Plot 2D', fontsize = 20)
            targets = ['Très bon','Bon', 'Moyen', 'Mauvais']
            colors = ['r', 'g', 'black','blue']
            for target, color in zip(targets,colors):
                indicesToKeep = Perf_risk_df["Score quali"] == target
                ax.scatter(Perf_risk_df.loc[indicesToKeep, 'Score_Perf_Moy']
                            , Perf_risk_df.loc[indicesToKeep, 'Score_Risks_Moy']
                            , c = color
                            , s = 50)
                for i in range(Perf_risk_df.shape[0]):
                    plt.text(x=Perf_risk_df["Score_Perf_Moy"][i],y=Perf_risk_df["Score_Risks_Moy"][i],s=str(i), fontdict=dict(color='black',size=10))
            ax.legend(targets)
            ax.grid()
            plt.savefig(os.path.join(cwd+"\Graphiques",'Score_Perf_Moy_Score_Risks_Moy.png'))
            
        ############### Adding Score_Perf_Moy with Score_ESG_Moy
        ######### Calcul of Score_Perf_Moy with Score_ESG_Moy
        def Perf_ESG_Calc():
            Perf_ESG_df=df_score[["ISIN","Score_Perf_Moy","Score_ESG_Moy","Score Moyenne"]].copy()
            Perf_ESG_df.sort_values(by=["Score Moyenne"],inplace=True,ascending=False)
            #Version AM
            Perf_ESG_df["Score quali"] = ''
            Perf_ESG_df.index = np.arange(0, len(Perf_ESG_df), 1)
            for i in range(0, len(Perf_ESG_df["Score Moyenne"])):
                if Perf_ESG_df.loc[i, "Score Moyenne"] >= np.percentile(np.array(Perf_ESG_df["Score Moyenne"]), 75):
                    Perf_ESG_df.loc[i, "Score quali"] = "Très bon"
                elif Perf_ESG_df.loc[i, "Score Moyenne"] < np.percentile(np.array(Perf_ESG_df["Score Moyenne"]), 75) and Perf_ESG_df.loc[i, "Score Moyenne"] >= np.percentile(np.array(Perf_ESG_df["Score Moyenne"]), 50):
                    Perf_ESG_df.loc[i, "Score quali"] = "Bon"
                elif Perf_ESG_df.loc[i, "Score Moyenne"] < np.percentile(np.array(Perf_ESG_df["Score Moyenne"]), 50) and Perf_ESG_df.loc[i, "Score Moyenne"] >= np.percentile(np.array(Perf_ESG_df["Score Moyenne"]), 25):
                    Perf_ESG_df.loc[i, "Score quali"] = "Moyen"
                else:
                    Perf_ESG_df.loc[i, "Score quali"] = "Mauvais"
            return(Perf_ESG_df)

            
        ######## Plot Score_Perf_Moy with Score_ESG_Moy
        def Perf_ESG_Fig(Perf_ESG_df):
            fig = plt.figure(figsize = (8,8))
            ax = fig.add_subplot(1,1,1) 
            ax.set_xlabel('Score_Perf_Moy', fontsize = 15)
            ax.set_ylabel('Score_ESG_Moy', fontsize = 15)
            ax.set_title('Scatter Plot 2D', fontsize = 20)
            targets = ['Très bon','Bon', 'Moyen', 'Mauvais']
            colors = ['r', 'g', 'black','blue']
            for target, color in zip(targets,colors):
                indicesToKeep = Perf_ESG_df["Score quali"] == target
                ax.scatter(Perf_ESG_df.loc[indicesToKeep, 'Score_Perf_Moy']
                            , Perf_ESG_df.loc[indicesToKeep, 'Score_ESG_Moy']
                            , c = color
                            , s = 50)
                for i in range(Perf_ESG_df.shape[0]):
                    plt.text(x=Perf_ESG_df["Score_Perf_Moy"][i],y=Perf_ESG_df["Score_ESG_Moy"][i],s=str(i), fontdict=dict(color='black',size=10))
            ax.legend(targets)
            ax.grid()
            plt.savefig(os.path.join(cwd+"\Graphiques",'Score_Perf_Moy_Score_ESG_Moy.png'))
        ############### End of Score_Perf_Moy with Score_ESG_Moy
        
        def Perf_fees_Calc():
            Perf_fees_df = df_score[["ISIN", "Score_Perf_Moy", "Score_Risks_Moy", "Score Moyenne", "Management Fee (ex Distribution fees) Actual"]]
            Perf_fees_df = Perf_fees_df.sort_values(by=["Score Moyenne"], ascending=False)
            Perf_fees_df.loc[:, "Fees Quali"] = ''
            Perf_fees_df["Fees Quali"] = ''
            for fund in Perf_fees_df.index:
                if Perf_fees_df.loc[fund,"Management Fee (ex Distribution fees) Actual"] is not None:
                    fee = float(Perf_fees_df.loc[fund,"Management Fee (ex Distribution fees) Actual"])
                else:
                    fee = np.nan
                fee_percentiles = np.array(Perf_fees_df["Management Fee (ex Distribution fees) Actual"].dropna(), dtype=float)
                if np.isnan(fee):
                    Perf_fees_df.loc[fund, "Fees Quali"] = 'Non précisées'
                elif fee >= np.percentile(fee_percentiles, 75):
                    Perf_fees_df.loc[fund, "Fees Quali"] = "Élevées +"
                elif fee >= np.percentile(fee_percentiles, 50):
                    Perf_fees_df.loc[fund, "Fees Quali"] = "Élevées"
                elif fee >= np.percentile(fee_percentiles, 25):
                    Perf_fees_df.loc[fund, "Fees Quali"] = "Moyennes"
                else:
                    Perf_fees_df.loc[fund, "Fees Quali"] = "Basses"
            return(Perf_fees_df)

        
        def Perf_fees_Fig(Perf_fees_df):
            fig = plt.figure(figsize = (8,8))
            ax = fig.add_subplot(1,1,1) 
            ax.set_xlabel('Score_Perf_Moy', fontsize = 15)
            ax.set_ylabel('Score_Risks_Moy', fontsize = 15)
            ax.set_title('Scatter Plot avec Fees', fontsize = 20)
            targets = ['Élevées +','Élevées', 'Moyennes', 'Basses',"Non précisées"]
            colors = ['red', 'green', 'black','blue',"grey"]
            for target, color in zip(targets,colors):
                indicesToKeep = Perf_fees_df["Fees Quali"] == target
                ax.scatter(Perf_fees_df.loc[indicesToKeep, 'Score_Perf_Moy']
                            , Perf_fees_df.loc[indicesToKeep, 'Score_Risks_Moy']
                            , c = color
                            , s = 50)
                for i in range(Perf_fees_df.shape[0]):
                    plt.text(x=Perf_fees_df["Score_Perf_Moy"][i],y=Perf_fees_df["Score_Risks_Moy"][i],s=str(i), fontdict=dict(color='black',size=10))
            ax.legend(targets)
            ax.grid()
            plt.savefig(os.path.join(cwd+"\Graphiques",'Score_Perf_fees.png'))
            
        w1 = PoidPerf
        w2 = PoidRisk
        w3 = PoidESG 
        print("poids de la performance ", w1)
        print("poids des risques ",w2)
        print("poids de l'ESG ",w3)
        df_score["Score Moyenne"]=w1*df_score["Score_Perf_Moy"]+w2*df_score["Score_Risks_Moy"]+w3*df_score["Score_ESG_Moy"]
        df_score.sort_values(by=["Score Moyenne"],inplace=True,ascending=False) #mettre dans l'ordre de classement
        coeff=20/(1+df_score["Score Moyenne"].iloc[0])
        tab_notes=[]
        for i in df_score["Score Moyenne"]:
            tab_notes.append(round(coeff*(1+i),2))
        df_score["Notes"]=tab_notes
        
        #feuille Résultats
        resultats = CreaFeuille(wb,'Résultats')
        #colonnes_ajoutees = df_score.iloc[:, -28:]
        #création d'une liste de préfixes des colonnes à sélectionner
        prefixes = ['ISIN','Notes','Score']
        #sélection des colonnes en fonction des préfixes
        selected_columns = [col for prefix in prefixes for col in df_score.filter(like=prefix).columns]
        #ajout de la colonne 'Max Drawdown' à la liste des colonnes sélectionnées

        #selected_columns.append('Max Drawdown 2023-03-01 to 2024-02-29 Base Currency')
        #extraction des colonnes sélectionnées à partir de df_score
        colonnes_ajoutees = df_score[selected_columns]
        #ecriture dans la Feuille de Resultats
        resultats.range('A1').value = colonnes_ajoutees

        Perf_fees_df = Perf_fees_Calc()
        ############################# Feuille Graphique
        if(graphique_option):
            abscisse = feuille['P5'].value
            ordonnee = feuille['P6'].value
            feuille['M12'].value = None
            feuille['M13'].value = None
            feuille.range(f'P5').color = (0, 255, 0)
            feuille.range(f'P6').color = (0, 255, 0)
            if abscisse not in df_score.columns:
                feuille['M11'].value = "Aucun graphique"
                graphique_option = False
                feuille['M12'].value = "Abscisse pas valide"
                feuille.range(f'P5').color = (255, 0, 0)
            if ordonnee not in df_score.columns :
                feuille['M11'].value = "Aucun graphique"
                graphique_option = False
                feuille['M13'].value = "Ordonnée pas valide"
                feuille.range(f'P6').color = (255, 0, 0)



        if(graphique_option):
            feuille_graphiques = CreaFeuille(wb,'Graphiques')
            FigLeftPos = 10
            FigTopPos = 10
            FigWidth = 350
            FigHeight = 350
            name = abscisse+"_" +ordonnee+".png"
            Alpha_r_df = R2_Alpha_Calc(abscisse, ordonnee)
            R2_Alpha_Fig(Alpha_r_df, abscisse, ordonnee)

            # Affichage du Graphique1 dans Excel
            graphique1 = feuille_graphiques.pictures.add(os.path.join(cwd+"\Graphiques",  name ), name='Graphique1', update=True)
            # Position sur la Feuille
            graphique1.left = FigLeftPos
            graphique1.top = FigTopPos
            FigTopPos += FigWidth
            # Redimensionnement 
            graphique1.width = FigWidth  # Largeur en points
            graphique1.height = FigHeight  # Hauteur en points
            
            finalDf = PCA_Calc()
            PCA_Fig(finalDf)
            
            graphique2 = feuille_graphiques.pictures.add(os.path.join(cwd+"\Graphiques", 'PCA_2D.png'), name='Graphique2', update=True)
            # Position sur la Feuille
            graphique2.left = FigLeftPos
            graphique2.top = FigTopPos
            FigTopPos += FigWidth
            # Redimensionnement 
            graphique2.width = FigWidth  # Largeur en points
            graphique2.height = FigHeight  # Hauteur en points
            ##wb.save(chemin_pilotage)

            Perf_risk_df = Perf_Risk_Calc()
            Perf_Risk_Fig(Perf_risk_df)
            graphique3 = feuille_graphiques.pictures.add(os.path.join(cwd+"\Graphiques", 'Score_Perf_Moy_Score_Risks_Moy.png'), name='Graphique3', update=True)
            # Position sur la Feuille
            graphique3.left = FigLeftPos
            graphique3.top = FigTopPos
            FigTopPos += FigWidth
            # Redimensionnement 
            graphique3.width = FigWidth  # Largeur en points
            graphique3.height = FigHeight  # Hauteur en points
            
            
            Perf_ESG_df = Perf_ESG_Calc()
            Perf_ESG_Fig(Perf_ESG_df)
            graphique4 = feuille_graphiques.pictures.add(os.path.join(cwd+"\Graphiques", 'Score_Perf_Moy_Score_ESG_Moy.png'), name='Graphique4', update=True)
            # Position sur la Feuille
            graphique4.left = FigLeftPos
            graphique4.top = FigTopPos
            FigTopPos += FigWidth
            # Redimensionnement 
            graphique4.width = FigWidth  # Largeur en points
            graphique4.height = FigHeight  # Hauteur en points
            
            
            
            Perf_fees_Fig(Perf_fees_df)
            graphique5 = feuille_graphiques.pictures.add(os.path.join(cwd+"\Graphiques", 'Score_Perf_fees.png'), name='Graphique5', update=True)
            # Position sur la Feuille
            graphique5.left = FigLeftPos
            graphique5.top = FigTopPos
            FigTopPos += FigWidth
            # Redimensionnement
            graphique5.width = FigWidth  # Largeur en points
            graphique5.height = FigHeight  # Hauteur en points
            def Find_color(data):
                col = (0,250,250)
                if(data=='Très bon'):
                    col = (255, 0, 0)    #Rouge
                if(data=='Bon'):
                    col = (0, 255, 0)     #Vert
                if(data=='Moyen'):
                    col = (0, 0, 0)   #Noir
                if(data=='Mauvais'):
                    col = (0, 0, 255)   #Bleu
                return col
            # Légende: Numéro des Fonds
            for i in range(0,len(noms)):
                couleur = Find_color(Perf_risk_df.iloc[i,-1])
                feuille_graphiques.range(f'I{i + 1}').color = couleur
                feuille_graphiques.range(f'J{i + 1}').value = i
                #feuille_graphiques.range(f'K{i + 1}').color = couleur
                feuille_graphiques.range(f'K{i + 1}').value = df_score.index[i]
                #feuille_graphiques.range(f'L{i + 1}').color = couleur
                feuille_graphiques.range(f'L{i + 1}').value = df_score.iloc[i,0]
            #Redimensionner les colonnes J et K en fonction de la taille du contenu
            feuille_graphiques.range('J:J').api.EntireColumn.AutoFit()
            feuille_graphiques.range('K:K').api.EntireColumn.AutoFit()
            feuille_graphiques.range('L:L').api.EntireColumn.AutoFit()

            ############################# Feuille Classement
        top = CreaFeuille(wb,'Classements')
        nbtop = 50
        merged_df = df_score.reset_index().merge(Perf_fees_df.reset_index(), on='Name').set_index('Name')
        merged_df.sort_values(by="Notes",ascending=False,inplace=True)
        colonnes_ajoutees = ['ISIN_x','Notes']
        top.range('A1').value = merged_df[colonnes_ajoutees]
        #.head(nbtop)
        def adjust_score(score, fees, max_fees, adjustment_factor):
            #adjusadjustment_factor = 0.25
            if fees is not None:
                adjustment = (fees / max_fees) * adjustment_factor
                adjusted_score = round(score * (1 - adjustment),2)
            else:
                adjusted_score=0
            return adjusted_score
        # Calculer la moyenne des valeurs non nulles
        mean_fee = merged_df[merged_df["Management Fee (ex Distribution fees) Actual_x"] > 0]["Management Fee (ex Distribution fees) Actual_x"].mean()
        median_fee = merged_df[merged_df["Management Fee (ex Distribution fees) Actual_x"] > 0]["Management Fee (ex Distribution fees) Actual_x"].median()
        # Remplacer les valeurs de 0 par la moyenne
        merged_df.loc[merged_df["Management Fee (ex Distribution fees) Actual_x"] == 0, "Management Fee (ex Distribution fees) Actual_x"] = mean_fee

        merged_df = merged_df[merged_df["Notes"] > 0]
        max_fees = 1.7
        merged_df["Notes Ajustées des frais"] = merged_df.apply(lambda row: adjust_score(row["Notes"], row["Management Fee (ex Distribution fees) Actual_x"],max_fees,ValAdjustFees), axis=1)
        merged_df.sort_values(by="Notes Ajustées des frais",ascending=False,inplace=True)
        
        coeff=20/(1+merged_df["Notes Ajustées des frais"].iloc[0])
        tab_notes=[]
        for i in merged_df["Notes Ajustées des frais"]:
            tab_notes.append(round(coeff*(1+i),2))
        merged_df["Notes avec Frais"]=tab_notes
            
        colonnes_ajoutees = ['ISIN_x','Notes avec Frais','Fees Quali']
        top.range('E1').value = merged_df[colonnes_ajoutees].head(nbtop)
        
        top.range('A:H').api.EntireColumn.AutoFit()
        
        ############### Statistiques Descriptives pour l'étude de la valeur de adjustment_factor
        # Analyse de sensibilité : On peut effectuer une analyse de sensibilité en faisant varier adjustment_factor
        # sur une plage de valeurs et en calculant des statistiques descriptives pour les notes ajustées 
        # (par exemple, moyenne, médiane, écart-type). 
        # Cela permet de voir comment les notes ajustées changent en fonction des différentes valeurs de adjustment_factor.
        if(graphique_option):
            list_adjustment_factors = [0,0.2,0.4,0.6,0.8,1]
            def sensitivity_analysis(adjustment_factors, df):
                results = []
            
                for factor in adjustment_factors:
                    df[f"Adjusted Notes {factor}"] = df.apply(
                        lambda row: adjust_score(
                            row["Notes"], row["Management Fee (ex Distribution fees) Actual_x"], max_fees, factor
                        ),
                        axis=1,
                    )
                    mean = df[f"Adjusted Notes {factor}"].mean()
                    median = df[f"Adjusted Notes {factor}"].median()
                    std_dev = df[f"Adjusted Notes {factor}"].std()
                    results.append({"Adjustment Factor": factor, "Mean": mean, "Median": median, "Standard Deviation": std_dev})
            
                return pd.DataFrame(results)

            sensitivity_results = sensitivity_analysis(list_adjustment_factors, merged_df)
            feuille_graphiques.range('N2').value = sensitivity_results
            # Analyse de corrélation : On peut calculer la corrélation entre les frais de gestion et les notes ajustées
            # pour différentes valeurs de adjustment_factor. Une corrélation négative indiquera que les frais de gestion ont un impact
            # négatif sur les notes ajustées, et plus la corrélation est négative, plus l'impact est important.
            def correlation_analysis(adjustment_factors, df):
                correlations = []
            
                for factor in adjustment_factors:
                    df[f"Adjusted Notes {factor}"] = df.apply(
                        lambda row: adjust_score(
                            row["Notes"], row["Management Fee (ex Distribution fees) Actual_x"], max_fees, factor
                        ),
                        axis=1,
                    )
                    if df["Management Fee (ex Distribution fees) Actual_x"].isnull().all():
                        corr = 0
                    else:
                        corr = df[[f"Adjusted Notes {factor}", "Management Fee (ex Distribution fees) Actual_x"]].corr().iloc[0, 1]
                    correlations.append({"Adjustment Factor": factor, "Correlation": corr})
            
                return pd.DataFrame(correlations)
            correlation_results = correlation_analysis(list_adjustment_factors, merged_df)
            # Visualisation de l'impact sur le classement : On peut créer un graphique qui montre comment le classement des
            # fonds change en fonction des différentes valeurs de adjustment_factor. Par exemple, un diagramme 
            # à barres empilées où chaque barre représente un fonds, et les segments de chaque barre représentent les notes 
            # ajustées pour différentes valeurs de adjustment_factor. Cela permet de voir comment les classements des 
            # fonds sont affectés par les différentes valeurs de adjustment_factor.
            def plot_ranking_impact(adjustment_factors, df):
                rankings = pd.DataFrame()
            
                for factor in adjustment_factors:
                    df[f"Adjusted Notes {factor}"] = df.apply(
                        lambda row: adjust_score(
                            row["Notes"], row["Management Fee (ex Distribution fees) Actual_x"], max_fees, factor
                        ),
                        axis=1,
                    )
                    df[f"Rank {factor}"] = df[f"Adjusted Notes {factor}"].rank(ascending=False)
                    rankings[f"Rank {factor}"] = df[f"Rank {factor}"]
            
                rankings.plot(kind='bar', figsize=(15, 6))
                plt.xlabel("Funds")
                plt.ylabel("Rank")
                plt.title("Impact of Adjustment Factors on Fund Ranking")
                plt.legend([f"Adjustment Factor: {factor}" for factor in adjustment_factors])
                plt.tight_layout()
                plt.savefig(os.path.join(cwd+"\Graphiques",'Impact_Adjustment_Factor.png'))
                
            plot_ranking_impact(list_adjustment_factors, merged_df)

            graphique10 = feuille_graphiques.pictures.add(os.path.join(cwd+"\Graphiques", 'Impact_Adjustment_Factor.png'), name='Graphique10', update=True)
            # Position sur la Feuille
            graphique10.left = 900
            graphique10.top = 200
        # Redimensionnement
        #§graphique6.width = 300  # Largeur en points
        #raphique6.height = 450  # Hauteur en points
        # Optimisation basée sur les objectifs : La métrique d'évaluation proposée mesure la corrélation
        # entre les notes ajustées des frais et les frais de gestion. 
        # L'idée de cette métrique est de maximiser la relation inverse entre les notes et les frais de gestion. 
        # En d'autres termes, on cherche à obtenir une situation où les fonds avec des frais de gestion plus élevés
        # ont des notes ajustées plus faibles et vice versa.
        def evaluation_metric(df):
            res = -1 * df["Notes ajustées des frais"].corr(df["Management Fee (ex Distribution fees) Actual_x"])
            return res
        
        def calculate_adjusted_notes(df, adjustment_factor, max_fees):
            df["Notes ajustées des frais"] = df.apply(lambda row: adjust_score(row["Notes"], row["Management Fee (ex Distribution fees) Actual_x"], max_fees, adjustment_factor), axis=1)
            return df
        
        def optimize_adjustment_factor(df, max_fees, possible_factors):
            best_factor = None
            best_metric = float('inf')
            
            for factor in possible_factors:
                adjusted_notes_df = calculate_adjusted_notes(df.copy(), factor, max_fees)
                metric_value = evaluation_metric(adjusted_notes_df)
                
                if metric_value < best_metric:
                    best_metric = metric_value
                    best_factor = factor
                    
            return best_factor
        
        possible_factors = np.linspace(0, 1, num=20)
        best_adjustment_factor = optimize_adjustment_factor(merged_df.copy(), max_fees, possible_factors)
        ###############
        df_fees = Perf_fees_Calc()
        print("L'exécution du code a été réalisée avec succès")

#fonction pour renvoyer le tableau avec la liste des coefficients dans l'ordre
def verify_coeff(nom_fichier, colonneCritere, colonneCoeff, feuille ="Sélection des critères"):
    wb = xw.Book(nom_fichier)
    feuille = wb.sheets[feuille]
    coeff = []
    derniere_ligne = feuille.range(colonneCritere + str(feuille.cells.last_cell.row)).end('up').row
    for i in range(2, derniere_ligne + 1):
        valeur = feuille.range(colonneCoeff + str(i)).value
        if valeur is None or (isinstance(valeur, int) and valeur < 0): #si pas de valeur alors le coefficient est égal à 1 par défaut 
            valeur = 1
        coeff.append(valeur)
    return coeff


#fonction pour vérifier si les mots saisis sont déjà connu par l'algorithme ou non
def verify_existence_word(liste_mot, nom_fichier,colonneCritere, colonneObject, feuille = 'Sélection des critères'):
    wb = xw.Book(nom_fichier)
    feuille = wb.sheets[feuille]
    listCriteria = []
    derniere_ligne = feuille.range(colonneCritere + str(feuille.cells.last_cell.row)).end('up').row
    error = False
    dic_unknow_criteria =  dict()
    for i in range(2, derniere_ligne + 1): #parcours les critères
        valeur = feuille.range(colonneCritere + str(i)).value
        if valeur not in liste_mot or valeur in listCriteria : #si le critère n'appartient ni à la liste connue ni est déjà  présente dans la colonne
            objectif = feuille.range(colonneObject + str(i)).value
            if objectif not in ['H', 'L']: #si c'est le cas alors je vérifie que la case objectif a été remplie correctement
                print(valeur, " est inconnu et vous n'avez pas renseigné la case objectif")
                feuille.range(colonneCritere + str(i)).color = (255, 0, 0)
                error = True
            else :
                dic_unknow_criteria[valeur] = objectif #si la case objectif a été remplie alors je l'ajoute au dictionnaire des critères inconnu et je l'associe à son objectif (L ou H)
                listCriteria.append(valeur)
                feuille.range(colonneCritere + str(i)).color = (200,32, 100)
        else :
            listCriteria.append(valeur)
            feuille.range(colonneCritere + str(i)).color = (0, 255, 0)
    print("la liste des critères est ",listCriteria)
    print("les éléments non connus par notre algorithme sont",  dic_unknow_criteria)
    return listCriteria, error, dic_unknow_criteria

#fonction afin de vérifier si les poids renseignés sont valides
def verify_weight(nom_fichier, feuille="Sélection des critères"):
    wb = xw.Book(nom_fichier)
    feuille = wb.sheets[feuille]
    PoidsPerf = round(feuille['J2'].value,2)
    PoidsRisk = round(feuille['K2'].value,2)
    PoidsESG = round(feuille['L2'].value,2)
    #erreur si un des poids n'a pas été renseigné
    if None in (PoidsPerf, PoidsRisk, PoidsESG):
        if PoidsPerf is None:
            feuille.range(f'J2').color = (255, 0, 0)
        if PoidsRisk is None:
            feuille.range(f'K2').color = (255, 0, 0)
        if PoidsESG is None:
            feuille.range(f'L2').color = (255, 0, 0)
        return PoidsPerf,PoidsRisk,PoidsESG,"Données manquantes dans poids"
    #erreur les poids ne sont pas compris entre 0 et 1
    if not all(0 <= poids <= 1 for poids in (PoidsPerf, PoidsRisk, PoidsESG)):
        feuille.range(f'J2').color = (255, 0, 0)
        feuille.range(f'K2').color = (255, 0, 0)  
        feuille.range(f'L2').color = (255, 0, 0)     
        return PoidsPerf,PoidsRisk,PoidsESG,"Poids doit être compris entre 0 et 100%"
    #erreur sur la somme des poids non égale à 1
    somme_poids = PoidsPerf + PoidsRisk + PoidsESG
    if somme_poids != 1:
        feuille.range(f'J2').color = (255, 0, 0)
        feuille.range(f'K2').color = (255, 0, 0)  
        feuille.range(f'L2').color = (255, 0, 0)     
        return PoidsPerf,PoidsRisk,PoidsESG,"Somme des poids != 100%"
    #aucune erreur pour les poids
    feuille.range(f'J2').color = (0, 255, 0)
    feuille.range(f'K2').color = (0, 255, 0)  
    feuille.range(f'L2').color = (0, 255, 0)
    return PoidsPerf,PoidsRisk,PoidsESG, None

#fonction afin de vérifier le nom du fichier pour importer la data
'''def verify_file_name(nom_fichier, feuille="Sélection des critères"):
    wb = xw.Book(nom_fichier)
    feuille = wb.sheets[feuille]
    error_name = ""
    data_file = feuille['M2'].value
    if(data_file != 'emd_global.xlsx' and data_file != 'hy_europe.xlsx'): #il n'y a que deux possibilités de nom de fichier soit emd_global.xlsx ou hy_europe.xlsx
        feuille.range(f'M2').color = (255, 0, 0)
        return "Nom de fichier invalide soit emd_global.xlsx ou hy_europe.xlsx", data_file
    feuille.range(f'M2').color = (0, 255, 0)
    return None, data_file'''

def graph_option(nom_fichier, feuille= "Sélection des critères"):
    wb = xw.Book(nom_fichier)
    feuille = wb.sheets[feuille]
    graphique_option = feuille['P4'].value
    if(graphique_option == 'y' or graphique_option == 'Y' or graphique_option == 'yes' or graphique_option == 'oui'):
        graphique_option = graphique_option.lower()
        feuille.range(f'P4').color = (0, 255, 0)
        return True
    else :
        feuille.range(f'P4').color = (255, 165, 0)
        return False

#fonction afin de vérifier la pertinence des seuils renseignés
def verify_seuil(nom_fichier, feuille="Sélection des critères"):
    wb = xw.Book(nom_fichier)
    feuille = wb.sheets[feuille]
    seuilmin = feuille['N2'].value
    seuilmax = feuille['O2'].value
    seuilpas = feuille['P2'].value
    if(seuilmin is None or seuilmax is None or seuilpas is None): #si un des seuils n'a pas de valeur
        feuille.range(f'N2').color = (255, 0, 0)
        feuille.range(f'O2').color = (255, 0, 0)
        feuille.range(f'P2').color = (255, 0, 0)
        return seuilmin, seuilmax, seuilpas, "Problème(s) sur les seuils, vous n'avez pas renseigné une valeur"
    if any(condition for condition in [seuilmin >= seuilmax,seuilmin < 0,seuilmax < 0,seuilmax > 1,seuilmin > 1,seuilpas > seuilmax - seuilmin]): #si ne respecte pas les conditions d'encadrement
        feuille.range(f'N2').color = (255, 0, 0)
        feuille.range(f'O2').color = (255, 0, 0)
        feuille.range(f'P2').color = (255, 0, 0)
        return seuilmin, seuilmax, seuilpas, "Problème(s) sur les seuils, vérifiez: 0 <seuilmin < seuilmax =<100%"
    #si aucun pb
    feuille.range(f'N2').color = (0, 255, 0)
    feuille.range(f'O2').color = (0, 255, 0)
    feuille.range(f'P2').color = (0, 255, 0)
    return seuilmin, seuilmax, seuilpas, None

#fonction pour vérifier que la sélection du fees est bien compris entre 0 et 10 inclus
def verify_fees(nom_fichier, feuille="Sélection des critères"):
    wb = xw.Book(nom_fichier)
    feuille = wb.sheets[feuille]
    ValAdjustFees = int(feuille['Q2'].value)
    if 0 <= ValAdjustFees <= 10:
        feuille.range(f'Q2').color = (0, 255, 0)
        return None, ValAdjustFees
    else: #aucune erreur
        feuille.range(f'Q2').color = (255, 0, 0)
        return "ValAdjFees incorrecte doit être entre 0 et 10", ValAdjustFees

#Vérifier la catégorie selectionnée parmi les 6 possibles
'''def verify_SelectedCategory(nom_fichier, feuille="Sélection des critères"):
    wb = xw.Book(nom_fichier)
    feuille = wb.sheets[feuille]
    cell = feuille['R2'].value

    liste_selectedcategory = ['EAA Fund Global Emerging Markets Bond - EUR Biased','EAA Fund Global Emerging Markets Bond', 'EAA Fund Global Emerging Markets Corporate Bond - EUR Biased','EAA Fund Global Emerging Markets Corporate Bond','EAA Fund EUR High Yield Bond',None]
    if cell not in liste_selectedcategory: #si la celulle contient un élément hors des possibilités
        feuille.range(f'R2').color = (255, 0, 0)
        return  "La catégorie n'est pas existante ou mettez un vide", cell
    if cell != None :
        feuille.range(f'R2').color = (0, 255, 0)
        return None, cell #si la cellule contient un élément qui fait partie de la liste
    else :
        feuille.range(f'R2').color = (0, 255, 0)
        return None, '' #si la cellule est vide, on prends toutes les catégories'''


def verify_data(nomExcel):
    #liste_mot est la variable qui contient l'ensemble des fonctions dont nous avons un moyen de calculer son score
    #liste_mot = ["Max Drawdown"]
    liste_mot = []
    nom_fichier = nomExcel
    feuille = 'Sélection des critères'
    wb = xw.Book(nom_fichier)
    feuille = wb.sheets[feuille]
    feuille['M14'].value = None
    #vérification de l'existence des critères
    list_criteria_Perf, error1, dic_unknow_criteria_perf = verify_existence_word(liste_mot, nom_fichier, 'A', 'C')
    list_criteria_Risk, error2, dic_unknow_criteria_risk = verify_existence_word(liste_mot, nom_fichier, 'D', 'F')
    list_criteria_ESG, error3, dic_unknow_criteria_esg= verify_existence_word(liste_mot, nom_fichier, 'G','I')
    #tableau qui contiendra l'ensemble des colonnes où il y a au moins une erreur
    columns_with_errors = []
    if error1:
        columns_with_errors.append("A")
    if error2:
        columns_with_errors.append("D")
    if error3:
        columns_with_errors.append("G")
    #vérification de la cohérence des poids
    PoidPerf, PoidRisk, PoidESG, error_weight = verify_weight(nom_fichier)
    #vérification nom fichier pour la data
    #error_name, nomFichier = verify_file_name(nom_fichier)
    nomFichier = feuille['M2'].value
    #vérification seuil 
    SeuilMin, SeuilMax, SeuilPas, error_seuil = verify_seuil(nom_fichier)
    #vérification fees
    error_fees, ValAdjustFees = verify_fees(nom_fichier)
    #vérification SelectedCategory
    #error_category, SelectedCategory = verify_SelectedCategory(nom_fichier)
    cell = feuille['R2'].value
    if cell != None :
        SelectedCategory = cell #si la cellule contient un élément qui fait partie de la liste
    else :
        SelectedCategory = '' #si la cellule est vide, on prends toutes les catégories

    #vérification graph_option
    graphique_option = graph_option(nom_fichier)
    #condition qui s'exécute lorsqu'il y a au moins une erreur, pour afficher les différents messages d'erreurs
    if len(columns_with_errors)!=0 or error_weight is not None  or error_seuil is not None or error_fees is not None:
        feuille['M5'].value = ""
        if(len(columns_with_errors)!=0):
            error_message = "Remplissez la colonne objectif pour l'élement en rouge colonne(s): " + ", ".join(columns_with_errors),
            feuille['M5'].value = error_message
        feuille['M6'].value = error_weight
        #feuille['M7'].value = error_name
        feuille['M8'].value = error_seuil
        feuille['M9'].value = error_fees
        #feuille['M10'].value = error_category 
        raise ValueError
    #réinitialiser les valeurs des cases qui affiche les erreurs si il n'y a pas d'erreur
    feuille['M5'].value = "Aucune erreur de remplissage"
    feuille['M6'].value = None
    feuille['M7'].value = None
    feuille['M8'].value = None
    feuille['M9'].value = None
    feuille['M10'].value = None
    if graphique_option :
        feuille['M11'].value = "Le graphique sera tracé"
    else:
        feuille['M11'].value = "Aucun graphique"
    
    #récupère les coefficients des critères
    coeff_perf = verify_coeff(nom_fichier, 'A', 'B')
    coeff_risk = verify_coeff(nom_fichier, 'D', 'E')
    coeff_esg = verify_coeff(nom_fichier, 'G', 'H')
    chemin_init = os.path.dirname(os.path.abspath(__file__))
    #chemin_init = "U:\\GDA\\PFC\\02_Ressources\\Scoring"
    #traiter le cas du EAA Fund EUR High Yield Bond qui n'est disponible que dans le fichier hy_europe.xlsx
    #if(nomFichier == 'emd_global.xlsx' and SelectedCategory == 'EAA Fund EUR High Yield Bond'):
     #   nomFichier = 'hy_europe.xlsx
    return(nomExcel,chemin_init,nomFichier, int(round(SeuilMin*100,0)),int(round(SeuilMax*100,0)), int(round(SeuilPas*100,0)), int(round(PoidPerf*100,0)), int(round(PoidRisk*100,0)), int(round(PoidESG*100,0)), SelectedCategory, ValAdjustFees,list_criteria_Perf, list_criteria_Risk, list_criteria_ESG, coeff_perf, coeff_risk, coeff_esg, dic_unknow_criteria_perf, dic_unknow_criteria_risk, dic_unknow_criteria_esg, graphique_option)

def startperso(nomExcel):
    error = False
    try:
        nomExcel, chemin_init, nomFichier, SeuilMin, SeuilMax, SeuilPas, PoidPerf, PoidRisk, PoidESG, SelectedCategory, ValAdjustFees, list_criteria_Perf, list_criteria_Risk, list_criteria_ESG, coeff_perf, coeff_risk, coeff_esg, dic_unknow_criteria_perf, dic_unknow_criteria_risk, dic_unknow_criteria_esg, graphique_option = verify_data(nomExcel)
    except ValueError:
        print("Oops!  That was not a valid input.  Try again...")
        error = True
    if error == False:
        start(nomExcel,chemin_init, nomFichier, SeuilMin, SeuilMax, SeuilPas, PoidPerf, PoidRisk, PoidESG, SelectedCategory, ValAdjustFees, True, list_criteria_Perf, list_criteria_Risk, list_criteria_ESG, coeff_perf, coeff_risk, coeff_esg, dic_unknow_criteria_perf, dic_unknow_criteria_risk, dic_unknow_criteria_esg, graphique_option)

#pour faire démarrer directement en exécutant depuis VS
#start(r'U:\GDA\PFC\02_Ressources\Scoring','emd_global.xlsx',80,90,5,100,0,0,'Toutes Catégories',0)

Sur vba :
Sub Classement_perso()
    Dim dynamicPath As String
    dynamicPath = ThisWorkbook.Path
    dynamicPath = Replace(dynamicPath, "\", "\\")
    RunPython "import sys; sys.path.insert(0, '" & dynamicPath & "'); import Pilotage; Pilotage.startperso('" & ThisWorkbook.Name & "')"
    ' Exécute Pilotage.py en utilisant xlwings et en passant nomFichier en argument
    ' RunPython "import sys; sys.path.insert(0, 'U:\\GDA\\PFC\\02_Ressources\\scoring-tmp'); import Pilotage; Pilotage.startperso()"
End Sub


Scoring MX3
import os
import datetime
import shutil
import pandas as pd

def select_files(path_input, path_output, date_list):
    selected_files = {}
    original_data_path = os.path.join(path_output, "originally_data_files")
    if not os.path.exists(original_data_path):
        os.makedirs(original_data_path)
    
    for date_input in date_list:
        date_obj = datetime.datetime.strptime(date_input, '%Y%m%d')
        
        max_timestamp = 0
        selected_file = None
        
        for filename in os.listdir(path_input):
            if filename.startswith('Stock_Action_DOPE_'):
                file_split = filename.split('_')[3]
                if file_split.endswith('.csv'):
                    file_split = file_split.replace('.csv', '')

                file_date = datetime.datetime.strptime(file_split, '%Y%m%d')
                
                if file_date.date() == date_obj.date():
                    timestamp = int(filename.split('_')[-1][:-4])
                    
                    if timestamp > max_timestamp:
                        max_timestamp = timestamp
                        selected_file = filename
        
        selected_files[date_input] = selected_file
        
        if selected_file is not None:
            source_file = os.path.join(path_input, selected_file)
            destination_file = os.path.join(original_data_path, selected_file)
            shutil.copy(source_file, destination_file)
    
    return selected_files

def process_selected_files(selected_files, path_output, identifier):

    for date, filename in selected_files.items():
        if filename is not None:

            source_file = os.path.join(path_output, "originally_data_files", filename)
            
            identifier_path = os.path.join(path_output, "portfolio", identifier)
            if not os.path.exists(identifier_path):
                os.makedirs(identifier_path)
            

            extract_data_path = os.path.join(identifier_path, "extract_data")
            if not os.path.exists(extract_data_path):
                os.makedirs(extract_data_path)
            
            date_path = os.path.join(extract_data_path,date)
            if not os.path.exists(date_path):
                os.makedirs(date_path)

            df = pd.read_csv(source_file, sep=';')
            filtered_df = df[df['Portefeuille'] == identifier]
            if(len(filtered_df)==0):
                raise Exception(f"Le portefeuille '{identifier}' n'existe pas à la date '{date}'.")
            output_file = os.path.join(date_path, f"{identifier}_{date}.xlsx")
            filtered_df.to_excel(output_file, index=False)



def construction_morning_star_files(path_input, numero_ptf, date_list, portfolio_name, portfolio_id, path_output):
    selected_files = select_files(path_input, path_output, date_list)
    process_selected_files(selected_files, path_output, numero_ptf)
    
    for date in date_list:
       
        extract_folder = os.path.join(path_output, "portfolio", numero_ptf, "extract_data",date)
        file_path_ms = os.path.join(path_output, "portfolio", numero_ptf, "files_morningstar")
        os.makedirs(file_path_ms, exist_ok=True)

        for filename in os.listdir(extract_folder):
            if filename.endswith('.xlsx'):
                extract_data_file = os.path.join(extract_folder, filename)
                df = pd.read_excel(extract_data_file)
                
                new_filename = f"{filename.split('.')[0]}_morningstar.xlsx"
                new_file_path = os.path.join(file_path_ms, new_filename)
                
                df = df.rename(columns={
                    'Libelle Titre': 'Nom ss jacents',
                    'Code Isin': 'ISIN',
                    'Code sedol': 'sedol',
                    'Nb Titre': 'qt'
                })
                date =datetime.datetime.strptime(date, "%Y%m%d")
                date = date.strftime("%d/%m/%Y")

                df['date'] = date
                df['Nom portefeuille'] = portfolio_name
                df['ID portefeuille'] = portfolio_id
                
                df = df[['date','Nom portefeuille', 'ID portefeuille', 'Nom ss jacents', 'ISIN', 'sedol', 'qt']]
                df.to_excel(new_file_path, index=False, engine='openpyxl')



def construction_bloomberg_files(path_input, numero_ptf, date_list, portfolio_name, portfolio_id, path_output):
    selected_files = select_files(path_input, path_output, date_list)
    process_selected_files(selected_files, path_output, numero_ptf)
    
    for date in date_list:
        extract_folder = os.path.join(path_output, "portfolio", numero_ptf, "extract_data",date)
        file_path_bloomberg = os.path.join(path_output, "portfolio", numero_ptf, "files_bloomberg")
        os.makedirs(file_path_bloomberg, exist_ok=True)

        for filename in os.listdir(extract_folder):
            if filename.endswith('.xlsx'):
                extract_data_file = os.path.join(extract_folder, filename)
                df = pd.read_excel(extract_data_file)
                
                new_filename = f"{filename.split('.')[0]}_bloomberg.xlsx"
                new_file_path = os.path.join(file_path_bloomberg, new_filename)
                #'code sedol': 'SEDOL',
                df = df.rename(columns={
                    'Libelle Titre': 'SECURITY NAME',
                    'Code sedol': 'SEDOL',
                    'Code Isin': 'ISIN',
                    'Nb Titre': 'QUANTITY',
                    'Prix revient Moyen': 'Cost Price'
                })
                date_obj = datetime.datetime.strptime(date, "%Y%m%d")
                date_str = date_obj.strftime("%m/%d/%Y")

                df['PORTFOLIO NAME'] = portfolio_name
                df['As of Date'] = date_str
                
                df['QUANTITY'] = df['QUANTITY'].str.replace(',', '.').astype(float)
                df['Cost Price'] = df['Cost Price'].str.replace(',', '.').astype(float)
                
                df = df[['PORTFOLIO NAME', 'SEDOL','ISIN', 'SECURITY NAME', 'QUANTITY', 'Cost Price', 'As of Date']]
                df.to_excel(new_file_path, index=False, engine='openpyxl')



path_input = r"U:\TRANSFERTS_FICHIERS\PROD\SIGMA\Com_StockAction"

path_output = r"U:\GDA\PFC\02_Ressources\Scoring-ESG & Extraction MX3"
date_list = ["20240822"]
numero_ptf = 'GF150'
portfolio_name = "CXA"

portfolio_id = "CXA"
construction_morning_star_files(path_input,numero_ptf, date_list,portfolio_name, portfolio_id, path_output)
construction_bloomberg_files(path_input, numero_ptf, date_list, portfolio_name, portfolio_id, path_output)

Tableau de bord – Forward
Sub BoutonUnique_Cliquer()
    Call ImportShiller2("EquityEU")
    Call Download_HY_US2("HYCorpoUS")
    Call Download_HY_EU2("HYCorpoEU")
    Call Bouton_Cliquer("IGCorpoUS")
    Call Bouton_Cliquer("IGCorpoEU")
    Call Bouton_Cliquer("HYCorpoUS")
    Call Bouton_Cliquer("HYCorpoEU")
    ' Sheets("EquityEU").Calculate
    Sheets("Performancev2").Activate


End Sub

Sub Download_HY_US2(sheetName As String)
    Dim FichierCSV As Workbook
    Dim FeuilleCible As Worksheet
    Dim FeuilleSource As Worksheet
    Dim NombreLignes As Long
    Dim DestRange As Range
    Dim CheminFichier As String
    Dim Fichiers() As String
    Dim i As Long
    Dim DateFormat As String
    Dim DateMaJ As Long
    Dim FichierMaJ As String
    Dim FichiersCount As Long

    ' Change le chemin qui correspond au dossier où il y a les fichiers de data
    CheminFichier = "U:\GDA\PFC\02_Ressources\Tableau de bord - Forward\"
    
    ' Récupérer la liste des fichiers dans le dossier
    Dim Fichier As String
    Fichier = Dir(CheminFichier & "GDAPFC*", vbNormal)
    i = 0
    While Fichier <> ""
        ReDim Preserve Fichiers(i)
        Fichiers(i) = Fichier
        Fichier = Dir
        i = i + 1
    Wend
    FichiersCount = i
    DateMaJ = 19000101
    FichierMaJ = ""
    For i = 0 To UBound(Fichiers)
        DateFormat = Mid(Fichiers(i), InStrRev(Fichiers(i), "-20") + 1, 8)
        If DateFormat > DateMaJ Then
            DateMaJ = DateFormat
            FichierMaJ = Fichiers(i)
        End If
    Next i
    MsgBox "Le fichier sélectionné est : " & FichierMaJ
    Set FichierCSV = Workbooks.Open(CheminFichier & FichierMaJ, , True)
    Set FeuilleCible = FichierCSV.Worksheets(1)
    Set FeuilleSource = ThisWorkbook.Worksheets(sheetName)
    
    FeuilleSource.Range("Z1:AF14").Borders.LineStyle = xlNone
    FeuilleSource.Range("Z1:AF14").ClearContents
    NombreLignes = FeuilleCible.UsedRange.Rows.count
    FeuilleCible.Range("A1:Z12").Copy Destination:=FeuilleSource.Range("Z1")
    FichierCSV.Close SaveChanges:=False
    
    ' Renommer colonne A
    FeuilleSource.Range("A2") = "3M"
    FeuilleSource.Range("A3") = "6M"
    For i = 4 To 13
        FeuilleSource.Range("A" & i) = i - 3
    Next i
    
    ' Mettre la date de la data en A1
    FeuilleSource.Range("A1") = FeuilleSource.Range("AB4").Value
    For i = 2 To 8
        FeuilleSource.Range("B" & i) = FeuilleSource.Range("AF" & i)
    Next i
    For i = 10 To 13
        FeuilleSource.Range("B" & i) = FeuilleSource.Range("AF" & i - 1)
    Next i
    
    FeuilleSource.Range("B1") = "YTM"
    FeuilleSource.Range("B9").FormulaLocal = "=MOYENNE(B8;B10)"
    FeuilleSource.Range("A1:B13").Borders.LineStyle = xlContinuous
    FeuilleSource.Range("A1:B13").WrapText = True
    FeuilleSource.Range("A1:B13").HorizontalAlignment = xlCenter
    FeuilleSource.Range("A1:B13").VerticalAlignment = xlCenter
End Sub

Sub Download_HY_EU2(sheetName As String)
    Dim FichierCSV As Workbook
    Dim FeuilleCible As Worksheet
    Dim FeuilleSource As Worksheet
    Dim NombreLignes As Long
    Dim DestRange As Range
    Dim CheminFichier As String
    Dim Fichiers() As String
    Dim i As Long
    Dim DateFormat As String
    Dim DateMaJ As Long
    Dim FichierMaJ As String
    Dim FichiersCount As Long

    ' Change le chemin qui correspond au dossier où il y a les fichiers de data
    CheminFichier = "U:\GDA\PFC\02_Ressources\Tableau de bord - Forward\"
    
    ' Récupérer la liste des fichiers dans le dossier
    Dim Fichier As String
    Fichier = Dir(CheminFichier & "GDAPFC*", vbNormal)
    i = 0
    While Fichier <> ""
        ReDim Preserve Fichiers(i)
        Fichiers(i) = Fichier
        Fichier = Dir
        i = i + 1
    Wend
    FichiersCount = i
    DateMaJ = 19000101
    FichierMaJ = ""
    For i = 0 To UBound(Fichiers)
        DateFormat = Mid(Fichiers(i), InStrRev(Fichiers(i), "-20") + 1, 8)
        If DateFormat > DateMaJ Then
            DateMaJ = DateFormat
            FichierMaJ = Fichiers(i)
        End If
    Next i
    MsgBox "Le fichier sélectionné est : " & FichierMaJ
    Set FichierCSV = Workbooks.Open(CheminFichier & FichierMaJ, , True)
    Set FeuilleCible = FichierCSV.Worksheets(1)
    Set FeuilleSource = ThisWorkbook.Worksheets(sheetName)
    
    FeuilleSource.Range("Z1:AF14").Borders.LineStyle = xlNone
    FeuilleSource.Range("Z1:AF14").ClearContents
    NombreLignes = FeuilleCible.UsedRange.Rows.count
    FeuilleCible.Range("A13:Z23").Copy Destination:=FeuilleSource.Range("Z1")
    FichierCSV.Close SaveChanges:=False
    
    ' Renommer colonne A
    FeuilleSource.Range("A2") = "3M"
    FeuilleSource.Range("A3") = "6M"
    For i = 4 To 13
        FeuilleSource.Range("A" & i) = i - 3
    Next i
    
    ' Mettre la date de la data en A1
    FeuilleSource.Range("A1") = FeuilleSource.Range("AB4").Value
    For i = 2 To 8
        FeuilleSource.Range("B" & i) = FeuilleSource.Range("AF" & i - 1)
    Next i
    For i = 10 To 13
        FeuilleSource.Range("B" & i) = FeuilleSource.Range("AF" & i - 2)
    Next i
    
    FeuilleSource.Range("B1") = "YTM"
    FeuilleSource.Range("B9").FormulaLocal = "=MOYENNE(B8;B10)"
    FeuilleSource.Range("A1:B13").Borders.LineStyle = xlContinuous
    FeuilleSource.Range("A1:B13").WrapText = True
    FeuilleSource.Range("A1:B13").HorizontalAlignment = xlCenter
    FeuilleSource.Range("A1:B13").VerticalAlignment = xlCenter
End Sub



Sub Bouton_Cliquer(sheetName As String)

    Worksheets(sheetName).Activate
    SolverOk SetCell:="$O$2", MaxMinVal:=3, ValueOf:="0", ByChange:="$C$5"
    SolverSolve True
    
    SolverOk SetCell:="$P$2", MaxMinVal:=3, ValueOf:="0", ByChange:="$C$6"
    SolverSolve True
    
    SolverOk SetCell:="$Q$2", MaxMinVal:=3, ValueOf:="0", ByChange:="$C$7"
    SolverSolve True
    
    SolverOk SetCell:="$R$2", MaxMinVal:=3, ValueOf:="0", ByChange:="$C$8"
    SolverSolve True
    
    SolverOk SetCell:="$S$2", MaxMinVal:=3, ValueOf:="0", ByChange:="$C$9"
    SolverSolve True
    
    SolverOk SetCell:="$T$2", MaxMinVal:=3, ValueOf:="0", ByChange:="$C$10"
    SolverSolve True
    
    SolverOk SetCell:="$U$2", MaxMinVal:=3, ValueOf:="0", ByChange:="$C$11"
    SolverSolve True
    
    SolverOk SetCell:="$V$2", MaxMinVal:=3, ValueOf:="0", ByChange:="$C$12"
    SolverSolve True
    
    SolverOk SetCell:="$W$2", MaxMinVal:=3, ValueOf:="0", ByChange:="$C$13"
    SolverSolve True
    
    Dim seuil_sup As Double
    seuil_sup = Range("B35").Value
    Dim seuil_inf As Double
    seuil_inf = Range("B33").Value
    Dim rngA As Range, cell As Range
    Dim ligne As Long
    Set rngA = Range("A4:A13")
    For Each cell In rngA
        If cell.Value = seuil_inf Then
            ligne = cell.Row
            Range("C33").Value = Range("C" & ligne).Value
        End If
    Next cell
    For Each cell In rngA
        If cell.Value = seuil_sup Then
            ligne = cell.Row
            Range("C35").Value = Range("C" & ligne).Value
        End If
    Next cell
End Sub

Sub ImportShiller2(sheetName As String)
    Dim ws As Worksheet
    Dim csvFilePath As String
    Dim wbCSV As Workbook
    Dim rawDate As Long
    Dim endDate As Date
    Dim lastDateInCSV As Date
    Dim i As Long
    Worksheets(sheetName).Activate
    Set ws = ThisWorkbook.Sheets("EquityEU")
    csvFilePath = "U:\GDA\PFC\02_Ressources\Tableau de bord - Forward\Historic-cape-ratios.csv"

    rawDate = ws.Range("P14").Value

    endDate = DateSerial(Left(rawDate, 4), Mid(rawDate, 5, 2), Mid(rawDate, 7, 2))

    endDate = DateSerial(Year(endDate), Month(endDate), 0)
    ws.Columns("B:C").Clear
    ws.Columns("D").Clear

    Set wbCSV = Workbooks.Open(Filename:=csvFilePath, Format:=2) ' Format:=2 spécifie un délimiteur de virgule

    With wbCSV.Sheets(1)
        
        For i = 1 To .Cells(.Rows.count, 1).End(xlUp).Row
            If IsDate(.Cells(i, 1).Value) Then
                lastDateInCSV = .Cells(i, 1).Value
                If lastDateInCSV > endDate Then Exit For
            End If
        Next i
        
        .Range("A1:B" & i - 1).Copy ws.Range("B1")
    End With


    wbCSV.Close SaveChanges:=False

    ws.Columns("D").Clear
    ws.Activate
End Sub

Modèle arbitrage :
import numpy as np
import pandas as pd
from tia.bbg import LocalTerminal
from collections import OrderedDict
import pickle
from sklearn.linear_model import LinearRegression
import matplotlib.pyplot as plt
from sklearn.metrics import r2_score 
from datetime import datetime 
from dateutil.relativedelta import relativedelta
from tia.bbg import LocalTerminal
import datetime
import os
import luigi
import shutil

from functions_arbitrage import Results,get_periods_per_year,trades_visu, trades_visu,strategy_weights,tradoffline_visu,trading_rule
from global_parameter import ClassGlobal, SelectData
from gestion_hdf5 import h5load, h5store
from growth_estimation import get_GrowthRate_LongTerm, get_GrowthRate_ShortTerm
from edr.get_edr import get_EDR
from vola_download.vola_data_download import Collecte_data_Model_Vola

class Arbitrage(luigi.Task):
    Doc_name = luigi.Parameter()
    class_path_output = luigi.Parameter()
    class_path_output=os.path.join(ClassGlobal().path_output,"Simulation\{}".format(ClassGlobal().simulation))
    date1 = SelectData().date1
        
    def output(self):
        return luigi.LocalTarget(os.path.join(self.class_path_output,"Results\Returns"+ClassGlobal().Doc_name+".xlsx")),luigi.LocalTarget(os.path.join(self.class_path_output,"Results\Volatility"+ClassGlobal().Doc_name+".xlsx"))
    
    def run(self):
        path_data = os.path.join(ClassGlobal().path_output,"Simulation\{}\Assets\Assets_{}".format(ClassGlobal().simulation,datetime.datetime.now().strftime("%Y_%m_%d")))
        file_Prices = "Prices_"+self.Doc_name+".h5"
        file_Vol = "Vol_"+self.Doc_name+".h5"
        file_Yields = "Yields_"+self.Doc_name+".h5"
        prices = pd.read_hdf(path_data+"\\"+file_Prices)
        prices,metadata = h5load(path_data+"\\"+file_Prices)
        volatilities = pd.read_hdf(path_data+"\\"+file_Vol)
        yields = pd.read_hdf(path_data+"\\"+file_Yields)

        #Slope
        esp_vol = volatilities.dropna()
        exp_ret = yields.dropna()
        X = esp_vol
        Y = exp_ret
        intercept_values = exp_ret.filter(like="Cash")[exp_ret.filter(like="Cash").columns[0]]

        first_date_vol = esp_vol.index[0]
        slopes = [] 
        for date in exp_ret.index:
            if(date >= first_date_vol):
                X_date = esp_vol.loc[date].values.reshape(-1, 1)
                y_date = exp_ret.loc[date].values
                intercept_value = intercept_values.loc[date]
                y_adjusted = y_date - intercept_value
                model = LinearRegression(fit_intercept=False)
                model.fit(X_date, y_adjusted)
                slope = model.coef_[0]
                slopes.append(slope)
            else:
                slopes.append(0)

        slopes_df = pd.DataFrame(data={'slope': slopes}, index=exp_ret.index)
        #Affichage Graphique
        #         plt.figure(figsize=(12, 6))
        #         plt.plot(slopes_df.index, slopes_df['slope'])
        #         plt.xlabel('Date')
        #         plt.ylabel('Slope')
        #         plt.title('Slope vs Date')
        #         plt.show()
        slopeReg = slopes_df['slope']
        slope=slopeReg.to_frame("slope")
        theo_ret = esp_vol.multiply(slope["slope"], axis=0).add(intercept_values, axis=0)
        theo_ret.dropna(inplace=True)
        diff_ret = exp_ret - theo_ret
        plot = diff_ret.plot()
        fig = plot.get_figure()
        #################
        fig.savefig(os.path.join(self.class_path_output,'Ecart_Droite_'+ClassGlobal().Doc_name+'.png'))
        diff_ret.to_excel(os.path.join(self.class_path_output,'Results\Ecart_Droite_'+ClassGlobal().Doc_name+'.xlsx'))
        h5store(os.path.join(self.class_path_output,"Results\Ecart"+ClassGlobal().Doc_name+".h5"), esp_vol, metadata)
        #################
        date1 = self.date1
        tradoffline_visu(date1,slope,exp_ret,esp_vol,SelectData().zone,ClassGlobal().simulation,ClassGlobal().Doc_name,self.class_path_output)
        exp_ret.reindex(index=exp_ret.index[::-1]).to_excel(os.path.join(self.class_path_output,"Results\Returns"+ClassGlobal().Doc_name+".xlsx"))
        # exp_ret.to_hdf(os.path.join(self.class_path_output,"Results\Returns"+ClassGlobal().Doc_name+".h5"))
        h5store(os.path.join(self.class_path_output,"Results\Returns"+ClassGlobal().Doc_name+".h5"), exp_ret, metadata)
        esp_vol.reindex(index=esp_vol.index[::-1]).to_excel(os.path.join(self.class_path_output,"Results\Volatility"+ClassGlobal().Doc_name+".xlsx"))
        # esp_vol.to_hdf(os.path.join(self.class_path_output,"Results\Volatility"+ClassGlobal().Doc_name+".h5"))
        h5store(os.path.join(self.class_path_output,"Results\Volatility"+ClassGlobal().Doc_name+".h5"), esp_vol, metadata)
        #################
        trades = diff_ret.apply(trading_rule, axis=1) # Sur les colonnes
        trades.dropna(inplace=True)
        trades_visu(trades,SelectData().period,SelectData().zone,SelectData().Earnings_Provider,self.class_path_output)
        trades.to_excel(os.path.join(self.class_path_output,'Results\Trades_LongShort_'+ClassGlobal().Doc_name+'.xlsx'))
        #################
        weights = strategy_weights(trades,SelectData().strategy,prices,esp_vol,get_periods_per_year(SelectData().period),exp_ret)
        res = Results()
        res.compute(trades,weights,prices,exp_ret,get_periods_per_year(SelectData().period))
        plt.clf()
        df = res.cumulative_performance.to_frame("Stratégie "+SelectData().strategy) 
        df = df[df.index>"2000"]
        df_perf = prices.join(df).dropna()/prices.join(df).dropna().iloc[0,:]
        # plt.plot(res.cumulative_performance,label="Performance cumulative de la stratégie "+ClassGlobal().strategy+" en zone "+ClassGlobal().zone)
        df_perf.plot()
        plt.savefig(os.path.join(self.class_path_output,'Cumulative_Perf'+ClassGlobal().Doc_name+'.png'))

if __name__ == "__main__":
    luigi.build([ClassGlobal(), SelectData()],local_scheduler = True)
    Collecte_data_Model_Vola()
    luigi.build([Arbitrage(Doc_name=ClassGlobal().Doc_name)],local_scheduler = True,no_lock=True)
    
    # filename1=r"U:\GDA\PFC\03_Gerants\03_08_AM\pyfinanceadvise\Simulation\Simulation_007\DataSet\Model_EDR_Data_EU_mArbitrageFact_EU007.h5"
    # filename2=r"U:\GDA\PFC\03_Gerants\03_08_AM\pyfinanceadvise\Simulation\Simulation_03_07_US\DataSet\Model_EDR_Data_US_mArbitrageShiller_US.h5"
    # df,meta = h5load(filename1)
    # print(df)
    # print(meta)


    Preparation dataset:
import luigi
import pandas as pd
import os
import datetime
import matplotlib.pyplot as plt
from gestion_hdf5 import h5store, h5load
from global_parameter import ClassGlobal, SelectData
from CPI_download.CPI_data_download import Collecte_CPI
from earnings_download.earnings_data_download import Collecte_Earnings
from rates_download.rates_data_download import Collecte_Rate
from prices_download.prices_data_download import Collecte_Prix
from real_rates_download.real_rates_data_download import Collecte_Infla
from vola_download.vola_data_download import Collecte_data_Model_Vola
import time

if __name__ == "__main__":
    luigi.build([ClassGlobal(), SelectData()],local_scheduler = True)
    Collecte_Infla()
    Collecte_CPI()
    Collecte_Earnings()
    Collecte_Prix()
    Collecte_Rate()
Gestion hdf55 

import pandas as pd

def h5store(filename, df, dic):
    store = pd.HDFStore(filename)
    store.put('mydata', df)
    store.get_storer('mydata').attrs.metadata = dic
    store.close()

def h5load(filename):
    with pd.HDFStore(filename) as store:
        data = store['mydata']
        metadata = store.get_storer('mydata').attrs.metadata
        data.attrs = metadata
    return data, metadata     

luigi.cfg 

[ClassGlobal]
path_output=U:\GDA\PFC\02_Ressources\Modele_Arbitrage
simulation=20240906-Simulation-US-NB
Doc_name=20240906-Simulation-US-NB

[SelectData]
zone=US
period =MONTHLY
start=19931231
#start=19931231
#chine mettre start = 20130430
end=20240930
#Earnings_Provider=Bloomberg
#Earnings_Provider=Factset
Earnings_Provider=Shiller
Model_Vol=Historical
strategy=Long_Only
date1=2024-09-30
#date1 doit être dernier jour d'un mois

[StrategyPypfopt]
known_future_cov_matrix = False
known_future_expected_returns = False
rolling = 25
# (mois)

Rates_download

import luigi
import datetime
import os
from collections import OrderedDict
from global_parameter import ClassGlobal
from functions_data import _get_generic_series, end_month_data
from gestion_hdf5 import h5store
import calendar
import pandas as pd

class Recuperation_Rate_Bloomberg(luigi.Task): #classe parent
    class_path_output = luigi.Parameter()
    zone = luigi.Parameter()
    period= luigi.Parameter()
    indice = luigi.Parameter(default = "GECU10YR Index") #indice devient un paramètre
    def _get_px_last(self,Rate): #recup prix
        assets = OrderedDict()
        assets["A1"]={}
        assets["A1"]["ticker"]= Rate
        assets["A1"]["des"]="Rate"
        Prices = _get_generic_series(assets, ["PX_LAST"], self.period)
        print(Prices)
        return Prices
    def output(self): #nom du fichier de sortie
        suffix = ""
        if self.period == "DAILY":
            suffix ="_d"
        if self.period == "MONTHLY" :
            suffix="_m" 
        if self.period == "QUARTERLY" :
            suffix="_q" 
        if self.period == "YEARLY":
            suffix = "_y"
        return luigi.LocalTarget(f'{self.class_path_output}\Rate_'+ self.zone +suffix+'.h5')
    def fetch_data(self):
        return self._get_px_last(self.indice), self.indice
    def run(self):
        data,indice1 = self.fetch_data()
        if self.period != "DAILY" :
            data = end_month_data(data, force_last_date=True)

        if self.period == "QUARTERLY":
            last_day_of_month = calendar.monthrange(data.index[-1].year, data.index[-1].month)[1]
            last_index = data.index[-1].replace(day=last_day_of_month)
            data = data.set_index(data.index[:-1].append(pd.Index([last_index])))

        metadata = dict(data="Rate",provider="Bloomberg",indice=indice1) 
        h5store(self.output().path,data,metadata)  
        excel_file = os.path.splitext(self.output().path)[0] + '.xlsx'
        data.to_excel(excel_file)

def Collecte_Rate():#fonction pour collecter les taux US et EUR
    path_init = os.path.join(ClassGlobal().path_output,"Data\{}\Rates".format(datetime.datetime.now().strftime("%Y_%m_%d")))
    class_path_output = path_init
    if not os.path.exists(class_path_output):
        os.makedirs(class_path_output)
    
    task_eu_d = Recuperation_Rate_Bloomberg(class_path_output, zone="EU", period="DAILY",indice="GECU10YR Index")
    task_eu_m = Recuperation_Rate_Bloomberg(class_path_output, zone="EU", period="MONTHLY",indice="GECU10YR Index")
    task_eu_q = Recuperation_Rate_Bloomberg(class_path_output, zone="EU", period="QUARTERLY",indice="GECU10YR Index")
    task_eu_y = Recuperation_Rate_Bloomberg(class_path_output, zone="EU", period="YEARLY",indice="GECU10YR Index")
    task_us_d = Recuperation_Rate_Bloomberg(class_path_output, zone="US", period="DAILY",indice="USGG10YR Index")
    task_us_m = Recuperation_Rate_Bloomberg(class_path_output, zone="US", period="MONTHLY",indice="USGG10YR Index")
    task_us_q = Recuperation_Rate_Bloomberg(class_path_output, zone="US", period="QUARTERLY",indice="USGG10YR Index")
    task_us_y = Recuperation_Rate_Bloomberg(class_path_output, zone="US", period="YEARLY",indice="USGG10YR Index")
    task_uk_d =Recuperation_Rate_Bloomberg(class_path_output, zone="UK", period="DAILY",indice="GUKG10 Index")
    task_uk_m =Recuperation_Rate_Bloomberg(class_path_output, zone="UK", period="MONTHLY",indice="GUKG10 Index")
    task_uk_q =Recuperation_Rate_Bloomberg(class_path_output, zone="UK", period="QUARTERLY",indice="GUKG10 Index")
    task_uk_y =Recuperation_Rate_Bloomberg(class_path_output, zone="UK", period="YEARLY",indice="GUKG10 Index")
    task_cn_d =Recuperation_Rate_Bloomberg(class_path_output, zone="CN", period="DAILY",indice="GCNY10YR Index")
    task_cn_m =Recuperation_Rate_Bloomberg(class_path_output, zone="CN", period="MONTHLY",indice="GCNY10YR Index")
    task_cn_q =Recuperation_Rate_Bloomberg(class_path_output, zone="CN", period="QUARTERLY",indice="GCNY10YR Index")
    task_cn_y =Recuperation_Rate_Bloomberg(class_path_output, zone="CN", period="YEARLY",indice="GCNY10YR Index")


    luigi.build([task_cn_d,task_cn_m, task_cn_q, task_cn_y, task_eu_d,task_eu_m,task_eu_q,task_eu_y,task_us_d,task_us_m,task_us_q,task_us_y,task_uk_d,task_uk_m,task_uk_y,task_uk_q],local_scheduler = True)


prices download :
import os
import datetime
from global_parameter import ClassGlobal
import luigi
from collections import OrderedDict
from functions_data import _get_generic_series, end_month_data
from gestion_hdf5 import h5store
import calendar
import pandas as pd


class Recuperation_Prices_Bloomberg(luigi.Task): #classe parent qui ressemble à Recuperation_Earnings_Bloomberg afin de récuperer les prix des assets           
    class_path_output = luigi.Parameter()
    zone = luigi.Parameter()
    period= luigi.Parameter()
    indice = luigi.Parameter(default="SXXP Index") 
    indiceTR = luigi.Parameter(default="SXXR Index")

    def _get_px_last(self, indice,indiceTR): #fonction qui retourne les prix des deux indices mis en paramètres
        assets = OrderedDict() 
        assets["A1"]={}
        assets["A1"]["ticker"]= indice
        assets["A1"]["des"]="Prices Bloomberg"
        assets["A2"]={}
        assets["A2"]["ticker"]= indiceTR
        assets["A2"]["des"]="Prices Bloomberg TR"
        Prices = _get_generic_series(assets, ["PX_LAST"], self.period)
        return Prices #comme précédemment

    def output(self): #nom du fichier de sortie qui varie selon si monthly ou quarterly
        suffix = ""
        if self.period == "DAILY" :
            suffix="_d" 
        if self.period == "MONTHLY" :
            suffix="_m" 
        if self.period == "QUARTERLY" :
            suffix="_q" 
        if self.period == "YEARLY":
            suffix="_y" 
        return luigi.LocalTarget(f'{self.class_path_output}\Bloomberg_'+ self.zone +suffix+'.h5')
    def fetch_data(self):
        return self._get_px_last(self.indice,self.indiceTR),self.indice,self.indiceTR
    def run(self):
        data,indice_ExDiv,indice_TR = self.fetch_data()
        if self.period != "DAILY" :
            data = end_month_data(data, force_last_date=True) #vérifie l'exactitude des mois
        # data.to_hdf(self.output().path,key='df',mode="w")
        if self.period == "QUARTERLY":
            last_day_of_month = calendar.monthrange(data.index[-1].year, data.index[-1].month)[1]
            last_index = data.index[-1].replace(day=last_day_of_month)
            data = data.set_index(data.index[:-1].append(pd.Index([last_index])))

        metadata = dict(data="Prices",provider="Bloomberg",indice=indice_ExDiv,indiceTR=indice_TR) 
        h5store(self.output().path,data,metadata)   #on enregistre les data final dans un hdf5
        excel_file = os.path.splitext(self.output().path)[0] + '.xlsx'
        data.to_excel(excel_file)
        

def Collecte_Prix():
    path_init = os.path.join(ClassGlobal().path_output,"Data\{}\Prices".format(datetime.datetime.now().strftime("%Y_%m_%d")))
    class_path_output = path_init
    if not os.path.exists(class_path_output):
        os.makedirs(class_path_output)
    task_prices_eu_d = Recuperation_Prices_Bloomberg(class_path_output, zone = "EU", period="DAILY", indice ="SXXP Index", indiceTR = "SXXR Index")
    task_prices_us_d  = Recuperation_Prices_Bloomberg(class_path_output,zone = "US",period="DAILY", indice = "SPX Index", indiceTR = "SPXT Index")
    task_prices_uk_d  = Recuperation_Prices_Bloomberg(class_path_output,zone = "UK",period="DAILY", indice = "UKX Index", indiceTR = "UKXDUK Index")
    task_prices_cn_d  = Recuperation_Prices_Bloomberg(class_path_output,zone = "CN",period="DAILY", indice = "M3CN Index", indiceTR = "NDEUCHF Index")

    task_prices_eu_m = Recuperation_Prices_Bloomberg(class_path_output, zone = "EU", period="MONTHLY", indice ="SXXP Index", indiceTR = "SXXR Index")
    task_prices_us_m  = Recuperation_Prices_Bloomberg(class_path_output,zone = "US",period="MONTHLY", indice = "SPX Index", indiceTR = "SPXT Index")
    task_prices_uk_m  = Recuperation_Prices_Bloomberg(class_path_output,zone = "UK",period="MONTHLY", indice = "UKX Index", indiceTR = "UKXDUK Index")
    task_prices_cn_m  = Recuperation_Prices_Bloomberg(class_path_output,zone = "CN",period="MONTHLY", indice = "M3CN Index", indiceTR = "NDEUCHF Index")

    task_prices_eu_q = Recuperation_Prices_Bloomberg(class_path_output, zone = "EU", period="QUARTERLY", indice ="SXXP Index", indiceTR = "SXXR Index")
    task_prices_us_q  = Recuperation_Prices_Bloomberg(class_path_output,zone = "US",period="QUARTERLY", indice = "SPX Index", indiceTR = "SPXT Index")
    task_prices_uk_q  = Recuperation_Prices_Bloomberg(class_path_output,zone = "UK",period="QUARTERLY", indice = "UKX Index", indiceTR = "UKXDUK Index")
    task_prices_cn_q  = Recuperation_Prices_Bloomberg(class_path_output,zone = "CN",period="QUARTERLY", indice = "M3CN Index", indiceTR = "NDEUCHF Index")
    
    task_prices_eu_y = Recuperation_Prices_Bloomberg(class_path_output, zone = "EU", period="YEARLY", indice ="SXXP Index", indiceTR = "SXXR Index")
    task_prices_us_y  = Recuperation_Prices_Bloomberg(class_path_output,zone = "US",period="YEARLY", indice = "SPX Index", indiceTR = "SPXT Index")
    task_prices_uk_y  = Recuperation_Prices_Bloomberg(class_path_output,zone = "UK",period="YEARLY", indice = "UKX Index", indiceTR = "UKXDUK Index")
    task_prices_cn_y  = Recuperation_Prices_Bloomberg(class_path_output,zone = "CN",period="YEARLY", indice = "M3CN Index", indiceTR = "NDEUCHF Index")

    luigi.build([task_prices_eu_d,task_prices_us_d,task_prices_uk_d,task_prices_cn_d,task_prices_eu_m,task_prices_us_m,task_prices_uk_m,task_prices_cn_m,task_prices_eu_q,task_prices_us_q,task_prices_uk_q,task_prices_cn_q,task_prices_eu_y,task_prices_uk_y,task_prices_us_y,task_prices_cn_y],local_scheduler = True)


real rates download : 
import os
import datetime
import luigi 
from global_parameter import ClassGlobal
from collections import OrderedDict
from functions_data import _get_generic_series, end_month_data
from gestion_hdf5 import h5store
import calendar
import pandas as pd

class Recuperation_Infla_Bloomberg(luigi.Task): #classe parent pour récupérer l'inflation des pays
    class_path_output = luigi.Parameter()
    zone = luigi.Parameter()
    period= luigi.Parameter()
    indice = luigi.Parameter(default ="GFRN10 Index")
    indiceTIPS = luigi.Parameter(default = "GFRGEN10 Index")
    def _get_px_last(self,TS,TIPS):#récup les prix à la fois des oat et oat sur inflation
        assets = OrderedDict()
        assets["A1"]={}
        assets["A1"]["ticker"]= TS
        assets["A1"]["des"]="TS"
        assets["A2"]={}
        assets["A2"]["ticker"]= TIPS
        assets["A2"]["des"]="TIPS"
        Prices = _get_generic_series(assets, ["PX_LAST"], self.period)
        Prices["Expected Infla"]= Prices["TS"]-Prices["TIPS"] #expected inflation = TS-TIPS
        return Prices

    def output(self): #fonction pour fichier de sortie
        suffix = ""
        if self.period == "DAILY":
            suffix ="_d"
        if self.period == "MONTHLY" :
            suffix="_m" 
        if self.period == "QUARTERLY" :
            suffix="_q" 
        if self.period == "YEARLY":
            suffix = "_y"
        return luigi.LocalTarget(f'{self.class_path_output}\ExpInf_'+ self.zone +suffix+'.h5')
    
    def fetch_data(self):
        return self._get_px_last(self.indice,self.indiceTIPS),self.indice,self.indiceTIPS  # Cette méthode doit être implémentée par chaque sous-classe
    
    def run(self):
        data,indice1,indiceTIPS = self.fetch_data()
        if self.period != "DAILY" :
            data = end_month_data(data, force_last_date=True) #vérifier la fin de mois
        
        if self.period == "QUARTERLY":
            last_day_of_month = calendar.monthrange(data.index[-1].year, data.index[-1].month)[1]
            last_index = data.index[-1].replace(day=last_day_of_month)
            data = data.set_index(data.index[:-1].append(pd.Index([last_index])))

        metadata = dict(data="Inflation",provider="Bloomberg",indice=indice1,indiceTIPS=indiceTIPS) 
        h5store(self.output().path,data,metadata)  #stocker en HDF5
        excel_file = os.path.splitext(self.output().path)[0] + '.xlsx'
        data.to_excel(excel_file)



class Recuperation_Infla_Bloomberg_CN(luigi.Task): #classe parent pour récupérer l'inflation des pays
    class_path_output = luigi.Parameter()
    zone = luigi.Parameter()
    period= luigi.Parameter()
    indice = luigi.Parameter(default ="GCNY10YR Index")
    indiceCPI = luigi.Parameter(default = "ECOPCNN Index")



    def _get_px_last(self,bond,cpi):#récup les prix à la fois du cpi et du bond
        assets = OrderedDict()
        assets["A1"]={}
        assets["A1"]["ticker"]= bond
        assets["A1"]["des"]="BOND"
        assets["A2"]={}
        assets["A2"]["ticker"]= cpi
        assets["A2"]["des"]="CPI"

        Prices = _get_generic_series(assets, ["PX_LAST"], self.period)
        Prices= Prices.ffill()
        Prices = Prices.dropna()
        Prices = end_month_data(Prices)
        Prices = Prices[~Prices.index.duplicated(keep='first')]

        index_prices = Prices.index
        if(self.period == 'MONTHLY'):
            y = 12
        if(self.period =='QUARTERLY'):
            y = 4
        if(self.period == 'YEARLY'):
            y = 1

        Prices = Prices.reset_index(drop=False)
        for i in range(y,len(Prices.index)) : 
            exp_infla = ((Prices.loc[i, 'CPI'] /Prices.loc[i-y, 'CPI'])-1)*100
            Prices.loc[i,'Expected Infla'] = Prices.loc[i,'BOND'] - (exp_infla) 
        Prices = Prices.set_index('date')
        return Prices

    def output(self): #fonction pour fichier de sortie
        suffix = ""
        if self.period == "DAILY":
            suffix ="_d"
        if self.period == "MONTHLY" :
            suffix="_m" 
        if self.period == "QUARTERLY" :
            suffix="_q" 
        if self.period == "YEARLY":
            suffix = "_y"
        return luigi.LocalTarget(f'{self.class_path_output}\ExpInf_'+ self.zone +suffix+'.h5')
    
    def fetch_data(self):
        return self._get_px_last(self.indice,self.indiceCPI),self.indice,self.indiceCPI  # Cette méthode doit être implémentée par chaque sous-classe
    
    def run(self):
        data,indice1,indiceTIPS = self.fetch_data()
        if self.period != "DAILY" :
            data = end_month_data(data, force_last_date=True) #vérifier la fin de mois
        
        if self.period == "QUARTERLY":
            last_day_of_month = calendar.monthrange(data.index[-1].year, data.index[-1].month)[1]
            last_index = data.index[-1].replace(day=last_day_of_month)
            data = data.set_index(data.index[:-1].append(pd.Index([last_index])))

        metadata = dict(data="Inflation",provider="Bloomberg",indice=indice1,indiceTIPS=indiceTIPS) 
        h5store(self.output().path,data,metadata)  #stocker en HDF5
        excel_file = os.path.splitext(self.output().path)[0] + '.xlsx'
        data.to_excel(excel_file)






class Recuperation_InflaSwap_Bloomberg(luigi.Task): #classe pour recup inflaswap
    path_output = luigi.Parameter()
    zone = luigi.Parameter()
    period= luigi.Parameter()
    InflaSwap = luigi.Parameter(default = "EUSWI10 BGN Curncy")
    def _get_px_last(self,InflaSwap): #recup les prix
        assets = OrderedDict()
        assets["A1"]={}
        assets["A1"]["ticker"]= InflaSwap
        assets["A1"]["des"]="InflaSwap"
        
        InflaSwap = _get_generic_series(assets, ["PX_LAST"], self.period)
        return InflaSwap
    def output(self): #définition du nom du fichier
        suffix = ""
        if self.period == "DAILY":
            suffix ="_d"
        if self.period == "MONTHLY" :
            suffix="_m" 
        if self.period == "QUARTERLY" :
            suffix="_q"
        if self.period == "YEARLY":
            suffix = "_y"

        return luigi.LocalTarget(f'{self.path_output}\InflaSwap_'+ self.zone +suffix+'.h5')
    def fetch_data(self):
        return self._get_px_last(self.InflaSwap),self.InflaSwap
    def run(self):
        data,indice1 = self.fetch_data()
        if self.period != "DAILY" : 
            #data = end_month_data(data)
            data = end_month_data(data, force_last_date=True)

        if self.period == "QUARTERLY":
            last_day_of_month = calendar.monthrange(data.index[-1].year, data.index[-1].month)[1]
            last_index = data.index[-1].replace(day=last_day_of_month)
            data = data.set_index(data.index[:-1].append(pd.Index([last_index])))
        # data.to_hdf(self.output().path,key='df',mode="w")
        metadata = dict(data="Inflation",provider="Bloomberg",indice=indice1) 
        h5store(self.output().path,data,metadata) #stocker dans un hdf5      
        excel_file = os.path.splitext(self.output().path)[0] + '.xlsx'
        data.to_excel(excel_file)

def Collecte_Infla(): #fonction qui calculer les tasks inflation
    path_init = os.path.join(ClassGlobal().path_output,"Data\{}\Inflation".format(datetime.datetime.now().strftime("%Y_%m_%d")))
    # dirctory_date = "Infla_{}".format(datetime.datetime.now().strftime("%Y_%m_%d"))
    # class_path_output = os.path.join(path_init,dirctory_date)
    class_path_output = path_init
    if not os.path.exists(class_path_output):
        os.makedirs(class_path_output)
    task_infla_eu_d = Recuperation_Infla_Bloomberg(class_path_output, zone="EU", period="DAILY", indice = "GFRN10 Index", indiceTIPS = "GFRGEN10 Index")
    task_infla_eu_m = Recuperation_Infla_Bloomberg(class_path_output, zone="EU", period="MONTHLY", indice = "GFRN10 Index", indiceTIPS = "GFRGEN10 Index")
    task_infla_eu_q = Recuperation_Infla_Bloomberg(class_path_output, zone="EU", period="QUARTERLY", indice = "GFRN10 Index", indiceTIPS = "GFRGEN10 Index")
    task_infla_eu_y = Recuperation_Infla_Bloomberg(class_path_output, zone="EU", period="YEARLY", indice = "GFRN10 Index", indiceTIPS = "GFRGEN10 Index")

    
    task_infla_us_d = Recuperation_Infla_Bloomberg(class_path_output, zone="US", period="DAILY", indice = "USGG10YR Index", indiceTIPS = "USGGT10Y Index")
    task_infla_us_m = Recuperation_Infla_Bloomberg(class_path_output, zone="US", period="MONTHLY", indice = "USGG10YR Index", indiceTIPS = "USGGT10Y Index")
    task_infla_us_q = Recuperation_Infla_Bloomberg(class_path_output, zone="US", period="QUARTERLY", indice = "USGG10YR Index", indiceTIPS = "USGGT10Y Index")
    task_infla_us_y = Recuperation_Infla_Bloomberg(class_path_output, zone="US", period="YEARLY", indice = "USGG10YR Index", indiceTIPS = "USGGT10Y Index")

    #task_infla_cn_d = Recuperation_Infla_Bloomberg_CN(class_path_output, zone="CN", period="DAILY", indice = "GCNY10YR Index", indiceCPI = "ECOPCNN Index")
    task_infla_cn_m = Recuperation_Infla_Bloomberg_CN(class_path_output, zone="CN", period="MONTHLY", indice = "GCNY10YR Index", indiceCPI = "ECOPCNN Index")
    task_infla_cn_q = Recuperation_Infla_Bloomberg_CN(class_path_output, zone="CN", period="QUARTERLY", indice = "GCNY10YR Index", indiceCPI = "ECOPCNN Index")
    task_infla_cn_y = Recuperation_Infla_Bloomberg_CN(class_path_output, zone="CN", period="YEARLY", indice = "GCNY10YR Index", indiceCPI = "ECOPCNN Index")


    task_inflaswap_eu_d = Recuperation_InflaSwap_Bloomberg(class_path_output,zone="EU",period="DAILY",  InflaSwap="EUSWI10 BGN Curncy")
    task_inflaswap_eu_m = Recuperation_InflaSwap_Bloomberg(class_path_output,zone="EU",period="MONTHLY",  InflaSwap="EUSWI10 BGN Curncy")
    task_inflaswap_eu_q = Recuperation_InflaSwap_Bloomberg(class_path_output,zone="EU",period="QUARTERLY",  InflaSwap="EUSWI10 BGN Curncy")
    task_inflaswap_eu_y = Recuperation_InflaSwap_Bloomberg(class_path_output,zone="EU",period="YEARLY",  InflaSwap="EUSWI10 BGN Curncy")
    
    luigi.build([task_infla_cn_m,task_infla_cn_q,task_infla_cn_y,task_infla_eu_d, task_infla_eu_m,task_infla_eu_q,task_infla_eu_y,task_infla_us_d,task_infla_us_m,task_infla_us_q,task_infla_us_y, task_inflaswap_eu_d,task_inflaswap_eu_m,task_inflaswap_eu_q,task_inflaswap_eu_y],local_scheduler = True)

functions_assets : 
import pandas as pd
from functions_data import _get_generic_series2
from tia.bbg import LocalTerminal
from collections import OrderedDict
import numpy as np


def get_assets(zone):
    assets = OrderedDict()
    if (zone=="US"):
        assets['LUATTRUU'] = {}
        assets['LUATTRUU']['ticker'] = "LUATTRUU Index"
        assets['LUATTRUU']['des'] = "Treasuries US"
        assets['LUACTRUU'] = {}
        assets['LUACTRUU']['ticker'] = "LUACTRUU Index"
        assets['LUACTRUU']['des'] = "HG Corpo US"
        assets['LF98TRUU'] = {}
        assets['LF98TRUU']['ticker'] = "LF98TRUU Index"
        assets['LF98TRUU']['des'] = "HY Corpo US"
        assets['SPXT'] = {}
        assets['SPXT']['ticker'] = "SPXT Index"
        assets['SPXT']['des'] = "Equity US"
        assets['US0001M'] = {}
        # assets['USC0TR03']['ticker'] = "USC0TR03 Index"
        assets['US0001M']['ticker'] = "US0001M Index"
        assets['US0001M']['des'] = "Cash US"
    if (zone=="EU"):
        assets['LEATTREU'] = {}
        assets['LEATTREU']['ticker'] = "LEATTREU Index"
        assets['LEATTREU']['des'] = "Treasuries EU"
        assets['LECPTREU'] = {}
        assets['LECPTREU']['ticker'] = "LECPTREU Index"
        assets['LECPTREU']['des'] = "HG Corpo EU"
        assets['I02501EU'] = {}
        assets['I02501EU']['ticker'] = "I02501EU Index"
        assets['I02501EU']['des'] = "HY Corpo EU"
        assets['SXXP'] = {}
        assets['SXXP']['ticker'] = "SXXR Index"
        assets['SXXP']['des'] = "Equity EU"
        
        # assets['Bloom'] = {}
        # assets['Bloom']['ticker'] = "SX5T Index"
        # assets['Bloom']['des'] = "Equity EU Bloom"
        # assets['SCXP'] = {}
        # assets['SCXP']['ticker'] = "SCXP Index"
        # assets['SCXP']['des'] = "Small Cap EU"
        # assets['EUROT'] = {}
        # assets['EUROT']['ticker'] = "EUROT Index"
        # assets['EUROT']['des'] = "Large Cap EU"

        assets['ECC0TR03'] = {}
        # assets['ECC0TR03']['ticker'] = "ECC0TR03 Index"
        assets['ECC0TR03']['ticker'] = "ECC0TR03 Index"
        assets['ECC0TR03']['des'] = "Cash EU"
    if (zone=="FR"):
        assets['LTFRTREU'] = {}
        assets['LTFRTREU']['ticker'] = "LTFRTREU Index"
        assets['LTFRTREU']['des'] = "Treasuries FR"
        assets['LECPTREU'] = {}
        assets['LECPTREU']['ticker'] = "LECPTREU Index"
        assets['LECPTREU']['des'] = "HG Corpo FR"
        assets['I02501EU'] = {}
        assets['I02501EU']['ticker'] = "I02501EU Index"
        assets['I02501EU']['des'] = "HY Corpo FR"
        assets['NCAC'] = {}
        assets['NCAC']['ticker'] = "NCAC Index"
        assets['NCAC']['des'] = "Equity FR"
        assets['ECC0TR03'] = {}
        assets['ECC0TR03']['ticker'] = "ECC0TR03 Index"
        assets['ECC0TR03']['des'] = "Cash FR"
    if (zone=="DE"):
        assets['LETGTREU'] = {}
        assets['LETGTREU']['ticker'] = "LETGTREU Index"
        assets['LETGTREU']['des'] = "Treasuries DE"
        assets['LECPTREU'] = {}
        assets['LECPTREU']['ticker'] = "LECPTREU Index"
        assets['LECPTREU']['des'] = "HG Corpo DE"
        assets['I02501EU'] = {}
        assets['I02501EU']['ticker'] = "I02501EU Index"
        assets['I02501EU']['des'] = "HY Corpo DE"
        assets['DAX'] = {}
        assets['DAX']['ticker'] = "DAX Index"
        assets['DAX']['des'] = "Equity DE"
        assets['ECC0TR03'] = {}
        assets['ECC0TR03']['ticker'] = "ECC0TR03 Index"
        assets['ECC0TR03']['des'] = "Cash DE"
    if (zone=="UK"):
        assets['H09027CH'] = {}
        assets['H09027CH']['ticker'] = "H09027CH Index"
        # assets['FTFIRDY7']['ticker'] = "FTFIRDY7 Index"
        assets['H09027CH']['des'] = "Treasuries UK"
        assets['I17389GB'] = {}
        assets['I17389GB']['ticker'] = "I17389GB Index"
        assets['I17389GB']['des'] = "HG Corpo UK"
        assets['CBPDHYI'] = {}
        assets['CBPDHYI']['ticker'] = "I05892GB Index"
        assets['CBPDHYI']['des'] = "HY Corpo UK"
        assets['TUKXG'] = {}
        assets['TUKXG']['ticker'] = "TUKXG Index"    
        assets['TUKXG']['des'] = "Equity UK"
        assets['DBMMSONI'] = {}
        assets['DBMMSONI']['ticker'] = "DBMMSONI Index"
        assets['DBMMSONI']['des'] = "Cash UK"

    if (zone=="CN"):
        assets['I32561US'] = {}
        assets['I32561US']['ticker'] = "I32561US Index"
        assets['I32561US']['des'] = "Treasuries CN"
        assets['JBMXCNTR'] = {}
        assets['JBMXCNTR']['ticker'] = "JBMXCNTR Index"
        assets['JBMXCNTR']['des'] = "HG Corpo CN"
        assets['JBMQCNTR'] = {}
        assets['JBMQCNTR']['ticker'] = "JBMQCNTR Index"
        assets['JBMQCNTR']['des'] = "HY Corpo CN"
        assets['M3CN'] = {}
        assets['M3CN']['ticker'] = "M3CN Index"    
        assets['M3CN']['des'] = "Equity CN"
        assets['CHBM7D'] = {}
        assets['CHBM7D']['ticker'] = "CHBM7D Index"
        assets['CHBM7D']['des'] = "Cash CN"
    return assets

def get_prices(start, end, period,zone): 
    assets = get_assets(zone)
    print(period)
    prices =  _get_generic_series2(assets, ['PX_LAST'], start, end, period)
    prices.attrs = assets
    print('okkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkk')
    print(prices)
    return prices

def get_short_term_yields_us(start, end, period, **kwargs):
    indexes = ['US0001M Index', 'SOFRRATE Index']
    resp = LocalTerminal.get_historical(indexes, 'PX_LAST',
        start=start, end=end, period=period, **kwargs)
    data = resp.as_frame()
    data.columns = data.columns.droplevel(1)
    str_series = pd.Series(index=data.index, name='ST index', dtype='float64')
    ind = np.isnan(data[indexes[1]])
    str_series.loc[~ind] = data.loc[~ind, indexes[1]]
    str_series.loc[ind] = data.loc[ind, indexes[0]]
    return str_series

def get_short_term_yields_eu(start, end, period, **kwargs):
    indexes = ['EONIA Index', 'ESTRON Index']
    resp = LocalTerminal.get_historical(indexes, 'PX_LAST',
        start=start, end=end, period=period, **kwargs)
    data = resp.as_frame()
    data.columns = data.columns.droplevel(1)
    str_series = pd.Series(index=data.index, name='ST index', dtype='float64')
    ind = np.isnan(data[indexes[1]])
    str_series.loc[~ind] = data.loc[~ind, indexes[1]]
    str_series.loc[ind] = data.loc[ind, indexes[0]]
    return str_series

def get_short_term_yields_uk(start, end, period, **kwargs):
    indexes = OrderedDict()
    indexes['SONIO'] = {}
    indexes['SONIO']['ticker'] = "SONIO/N Index"
    indexes['SONIO']['des'] = "Cash UK"
    str_series = _get_generic_series2(indexes, ['PX_LAST'], start, end, period)
    return str_series

def get_short_term_yields_cn(start, end, period, **kwargs):
    indexes = OrderedDict()
    indexes['CHBM7D'] = {}
    indexes['CHBM7D']['ticker'] = "CHBM7D Index"
    indexes['CHBM7D']['des'] = "Cash CN"
    str_series = _get_generic_series2(indexes, ['PX_LAST'], start, end, period)
    return str_series

def get_short_term_yields(zone,start, end, period, **kwargs):
    if(zone=="US"):
        yields = get_short_term_yields_us(start, end, period)
    if(zone=="EU"):
        yields = get_short_term_yields_eu(start, end, period)
    if(zone=="UK"):
        yields = get_short_term_yields_uk(start, end, period)
    if(zone=="CN"):
        yields = get_short_term_yields_cn(start, end, period)
    return yields

def get_assets_yields(assets, start, end, period):
    return _get_generic_series2(assets, ['BX207'], start, end, period)

def get_yields(start, end, period,zone,BDD):
    assets = get_assets(zone)
    yields = get_assets_yields(assets, start, end, period)
    
    yields["Cash "+zone] = get_short_term_yields(zone, start, end, period)
    # For HG Corporate Bonds : Yield less 50bp to take into account the average historical loss due to downgrades.
    if(zone=='CN'):
        assets ={}
        assets['JBMXCNSW'] = {}
        assets['JBMXCNSW']['ticker'] = "JBMXCNSW Index"
        assets['JBMXCNSW']['des'] = "IG Spread"
        assets['JBMQCNSW'] = {}
        assets['JBMQCNSW']['ticker'] = "JBMQCNSW Index"
        assets['JBMQCNSW']['des'] = "HY Spread"                      
        assets['GCNY10YR'] = {}
        assets['GCNY10YR']['ticker'] = "GCNY10YR Index"
        assets['GCNY10YR']['des'] = "Bond 10Y"
        hy_ig = _get_generic_series2(assets, ['PX_LAST'], start, end, period)
        yields["HG Corpo "+zone] =(hy_ig['IG Spread']/100)+ hy_ig['Bond 10Y']
        yields["HY Corpo "+zone] =(hy_ig['HY Spread']/100)+ hy_ig['Bond 10Y']


    yields["HG Corpo "+zone]=yields["HG Corpo "+zone]-0.5
    # For HY Corporate Bonds : Yield less 250bp to take into account the average historical loss due to defaults.
    yields["HY Corpo "+zone]=yields["HY Corpo "+zone]-2.5
    #     yields["Equity "+zone] = get_EDR("SPX Index",zone,"MONTHLY",start,end)
    yields.attrs = assets
    return yields

function_data :

import numpy as np
import pandas as pd
from tia.bbg import LocalTerminal
from collections import OrderedDict
import pickle
from sklearn.linear_model import LinearRegression
import matplotlib.pyplot as plt
from sklearn.metrics import r2_score 
from datetime import datetime 
from tia.bbg import LocalTerminal
import datetime 
from datetime import date
import os
import luigi
import matplotlib.dates as mdates


def _get_generic_series2(assets, fields, start, end, period, **kwargs):
    resp = LocalTerminal.get_historical([assets[i]["ticker"] for i in assets.keys()],#utilise tia de bloom pour récupérer données historiques pour l'ensemble des assets détenus dans l'indice
        fields, start=start, end=end, period=period, **kwargs)  #à partir d'une date de début et de fin
    data = resp.as_frame() #stocke dans data frame
    data.columns = data.columns.droplevel(1) #on supprime la col 1
    data.ffill(inplace=True) #rempli les valeurs Nan avec la méthode forward fill
    data.rename(columns=_build_easy_map(assets), inplace=True) #rename les colonnes avec la méthode _build_easy_map
    return data #retourne le dataframe

  
def _get_generic_series(assets, fields, period, **kwargs):
    #if(fields[0] =='T12_EPS_AGGTE' and assets['']["ticker"] == 'MXCN Index' ): #case where we need to convert in CNY currency
     #   resp = LocalTerminal.get_historical([assets[i]["ticker"] for i in assets.keys()],fields, start='1/1/1900', end = '1/1/2030', period=period, currency ='CNY')
   #else:
    resp = LocalTerminal.get_historical([assets[i]["ticker"] for i in assets.keys()],fields, start='1/1/1900', end = '1/1/2030', period=period, **kwargs)  #à partir d'une date de début et de fin
    data = resp.as_frame() #stocke dans data frame
    data.columns = data.columns.droplevel(1) #on supprime la col 1
    print(data)
    data.ffill(inplace=True) #rempli les valeurs Nan avec la méthode forward fill
    data.rename(columns=_build_easy_map(assets), inplace=True) #rename les colonnes avec la méthode _build_easy_map
    return data #retourne le dataframe

def _build_easy_map(assets):
    assets_easy_map = {}
    for k in assets.keys(): #récupère les tickers des assets
        assets_easy_map[assets[k]['ticker']] = assets[k]['des']
    return assets_easy_map

def end_month_data(df, force_last_date=False): #vérifie si les mois sont bien sont succesifs
    df = df.reset_index("date")
    last_date = df["date"].iloc[-1]
    print(last_date)
    if force_last_date:
        df["date"] = df["date"].apply(lambda x: x+pd.offsets.MonthEnd(1) if (x.month == (x+pd.offsets.MonthEnd(1)).month) else x)
    else:
        df["date"] = df["date"].apply(lambda x: x+pd.offsets.MonthEnd(1) if (x.month == (x+pd.offsets.MonthEnd(1)).month and x!=last_date) else x)
    df = df.set_index("date")
    return df


'''assets = OrderedDict()
assets[""]={}
assets[""]["ticker"]= "EACPI Index"
assets[""]["des"]="CPI Bloomberg"
Prices = _get_generic_series(assets, ["PX_LAST"], "MONTHLY")
print(Prices)'''

convert to pdf : 
 import os
from selenium import webdriver
from selenium.webdriver.edge.service import Service
import json
import time
from selenium.common.exceptions import TimeoutException
import shutil
from datetime import datetime, timedelta



#initialise les paths correspondants à nos chemins
edge_service = Service(r"C:\Users\nbonneau\Downloads\edgedriver_win64\msedgedriver.exe")
path_files_to_convert = r"C:\Local\xhtml2pdf"
savefile_directory = r"C:\Local\xhtml2pdf\dl"
errors = savefile_directory + r"\errors_convert_to_pdf.txt"


#state pour choisir l'imprimante save as pdf
app_state = {
    "recentDestinations": [
        {
            "id": "Save as PDF",
            "origin": "local",
            "account": ""
        }
    ],
    "selectedDestinationId": "Save as PDF",
    "version": 2
}



i = 0
# itérer nos files présents
directory = path_files_to_convert
for filename in os.listdir(directory):
    if filename.endswith(".html") or filename.endswith(".xhtml"):

        #new
        prefs = {
    'printing.print_preview_sticky_settings.appState': json.dumps(app_state),
    'savefile.default_directory': savefile_directory,
    'savefile.default_filename': filename
}
        options = webdriver.EdgeOptions()
        options.add_experimental_option('prefs', prefs)
        options.add_argument('--kiosk-printing')
        

        #end
        i= i+1
        print("helloooo i ==",i)
        if(i>0):
            file_path = os.path.join(directory, filename)


            print(f"Processing file: {filename}")
            #poids du file
            file_size_bytes = os.path.getsize(file_path)
            file_size_mb = file_size_bytes / (1024 ** 2)
            print(file_size_mb)

            driver = webdriver.Edge(service=edge_service, options=options)
            driver.get(file_path)
            #driver.set_script_timeout(10)
            #driver.execute_async_script('window.print()')

            if file_size_mb < 5 : 
                driver.execute_script('window.print()')

            else :
                try:
                    driver.set_script_timeout(35+file_size_mb)
                    driver.execute_async_script('window.print()')
                    
                except TimeoutException:
                    print(f"Timed out while printing {filename}")
            driver.quit()

            old_filename = max([savefile_directory + "\\" + f for f in os.listdir(savefile_directory)],key=os.path.getctime)
            current_time = datetime.now()
            file_creation_time = datetime.fromtimestamp(os.path.getctime(old_filename))
            time_difference = current_time - file_creation_time
            if time_difference.total_seconds() <= 32+file_size_mb:
                new_filename = filename + '.pdf'
                shutil.move(old_filename, os.path.join(savefile_directory, new_filename))
                print(f"Le fichier '{os.path.basename(old_filename)}' a été renommé en '{new_filename}'.")
            else:
                print("Aucun fichier trouvé avec une heure de création égale à l'heure actuelle ou à la dernière minute.")






            full_name = os.path.join(savefile_directory, filename) +".pdf"
            if os.path.exists(full_name):
                pass
            else:
                with open(errors , "a") as error_log:
                    error_log.write(f"{filename} \n")



        print(f"Finished processing file: {filename}")

Jp download : 
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
import time
import config as cf
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import os
import datetime
import shutil
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

def market_jp(path_for_saving, path_default_download, date):
    options = webdriver.EdgeOptions()
    options.add_argument(f"user-data-dir={cf.local['userDataDir']}")
    service = Service(cf.local["executablePath"])
    driver = webdriver.Edge(service=service, options=options)
    date_init = datetime.datetime.now()
        

    #ouvrir le site

    driver.get("https://markets.jpmorgan.com/research/CFP?page=EMBI")
    wait = WebDriverWait(driver, 10)
    #cliquer sur le calendrier
    td_element = driver.find_element(By.ID, "historicalDate")
    td_element.click()
    #identifier la date mise en paramètre
    year = int(date[:4])
    month = int(date[4:6].lstrip('0'))
    day = int(date[6:8].lstrip('0'))
    wait = WebDriverWait(driver, 10)
    try : 
        #choisir le bon mois
        select_month = wait.until(
            EC.visibility_of_element_located((By.CLASS_NAME, "ui-datepicker-month"))
        )
        month_option = select_month.find_element(By.XPATH, f".//option[@value='{month-1}']")
        month_option.click()

        #choisir la bonne année
        select_year = wait.until(
            EC.visibility_of_element_located((By.CLASS_NAME, "ui-datepicker-year"))
        )
        year_option = select_year.find_element(By.XPATH, f".//option[@value='{year}']")
        year_option.click()
        #choisir le bon jour
        day_td = wait.until(EC.element_to_be_clickable((By.XPATH, f"//td[@data-handler='selectDay' and .//a[text()='{day}']]")))
        day_td.click()


    except Exception as e:
        #affiche la date si elle n'est pas disponible sur le site (que 3 mois d'historical data disponible )
        print(f"Une erreur est survenue : {e}")
        print("La date d/m/y suivante n'est pas disponible", day,"/",month, "/", year)
    
    time.sleep(5)
    
    date_init = datetime.datetime.now()
    #EMBI GLOBAL DIV

    embi_global_row = wait.until(
        EC.presence_of_element_located((By.XPATH, "//td[text()='EMBI Global Div.']/.."))
    )

    #embi global div composition 
    composition_cell = embi_global_row.find_element(By.CLASS_NAME, "EMBI_index_historical_data_COMPOSITION")
    download_link = composition_cell.find_element(By.TAG_NAME, "a")
    download_link.click()
    time.sleep(2)
    #embi global div sub-index 
    composition_cell = embi_global_row.find_element(By.CLASS_NAME, "EMBI_index_historical_data_SUB_INDEX")
    download_link = composition_cell.find_element(By.TAG_NAME, "a")
    download_link.click()
    time.sleep(2)
    #embi global div monthly country weights
    composition_cell = embi_global_row.find_element(By.CLASS_NAME, "EMBI_index_historical_data_MONTHLY_COUNTRY_WEIGHTS")
    download_link = composition_cell.find_element(By.TAG_NAME, "a")
    download_link.click() 
    time.sleep(2)
    #embi global div monthly country preview 
    composition_cell = embi_global_row.find_element(By.CLASS_NAME, "EMBI_index_historical_data_PREVIEW")
    download_link = composition_cell.find_element(By.TAG_NAME, "a")
    download_link.click() 
    time.sleep(2)

    #EMBI GLOBAL 
    embi_global_row = wait.until(
        EC.presence_of_element_located((By.XPATH, "//td[text()='EMBI Global']/.."))
    )

    #embi global composition 
    composition_cell = embi_global_row.find_element(By.CLASS_NAME, "EMBI_index_historical_data_COMPOSITION")
    download_link = composition_cell.find_element(By.TAG_NAME, "a")
    download_link.click()
    time.sleep(2)
    #embi global sub-index 
    composition_cell = embi_global_row.find_element(By.CLASS_NAME, "EMBI_index_historical_data_SUB_INDEX")
    download_link = composition_cell.find_element(By.TAG_NAME, "a")
    download_link.click()
    time.sleep(2)
    #embi global monthly country weights
    composition_cell = embi_global_row.find_element(By.CLASS_NAME, "EMBI_index_historical_data_MONTHLY_COUNTRY_WEIGHTS")
    download_link = composition_cell.find_element(By.TAG_NAME, "a")
    download_link.click() 
    time.sleep(2)
    #embi global monthly country preview 
    composition_cell = embi_global_row.find_element(By.CLASS_NAME, "EMBI_index_historical_data_PREVIEW")
    download_link = composition_cell.find_element(By.TAG_NAME, "a")
    download_link.click() 
    time.sleep(2)


    #EMBI+
    embi_global_row = wait.until(
        EC.presence_of_element_located((By.XPATH, "//td[text()='EMBI Global']/.."))
    )

    #embi + composition 
    composition_cell = embi_global_row.find_element(By.CLASS_NAME, "EMBI_index_historical_data_COMPOSITION")
    download_link = composition_cell.find_element(By.TAG_NAME, "a")
    download_link.click()
    time.sleep(2)
    #embi + sub-index 
    composition_cell = embi_global_row.find_element(By.CLASS_NAME, "EMBI_index_historical_data_SUB_INDEX")
    download_link = composition_cell.find_element(By.TAG_NAME, "a")
    download_link.click()
    time.sleep(2)
    #embi + monthly country preview 
    composition_cell = embi_global_row.find_element(By.CLASS_NAME, "EMBI_index_historical_data_PREVIEW")
    download_link = composition_cell.find_element(By.TAG_NAME, "a")
    download_link.click() 
    time.sleep(2)


    #EMBIG Core
    embi_global_row = wait.until(
        EC.presence_of_element_located((By.XPATH, "//td[text()='EMBIG Core']/.."))
    )

    # EMBIG Core composition 
    composition_cell = embi_global_row.find_element(By.CLASS_NAME, "EMBI_index_historical_data_COMPOSITION")
    download_link = composition_cell.find_element(By.TAG_NAME, "a")
    download_link.click()
    time.sleep(2)
    #EMBIG Core monthly country preview 
    composition_cell = embi_global_row.find_element(By.CLASS_NAME, "EMBI_index_historical_data_PREVIEW")
    download_link = composition_cell.find_element(By.TAG_NAME, "a")
    download_link.click() 
    time.sleep(2)
    time.sleep(10)

    #créer des tuples avec le nom des fichiers et la date de création du fichier
    file_names_with_timestamps = []
    download_folder = path_default_download
    now = datetime.datetime.now()
    num_files_downloaded = 0
    for filename in os.listdir(download_folder):
        if os.path.getmtime(os.path.join(download_folder, filename)) > date_init.timestamp() and os.path.getmtime(os.path.join(download_folder, filename)) < now.timestamp():
            file_names_with_timestamps.append((filename, os.path.getmtime(os.path.join(download_folder, filename))))
    #faire une liste avec le nom des fichiers rangés par date d'anciennete dans file_name (du plus vieux au plus jeune)
    file_names_with_timestamps.sort(key=lambda x: x[1], reverse=True)
    file_names = [x[0] for x in file_names_with_timestamps]
    file_names = file_names[::-1]
    print(f"Nombre de fichiers téléchargés: {len(file_names)}")
    #13 fichiers attendus
    if(len(file_names)!=13):
        print("erreur dans le chargement des fichiers, ces fichiers ont été downloads correctement :")
        print(file_names)
        raise Exception("Impossible de télécharger tous les fichiers nécessaires.")

    for file_name in file_names:
        print(file_name)
    

    #identifier les noms des dossiers
    embi_global_div_dir = os.path.join(path_for_saving, "EMBI Global Div.")
    embi_global_dir = os.path.join(path_for_saving, "EMBI Global")
    embi_plus_dir = os.path.join(path_for_saving, "EMBI+")
    embig_core_dir = os.path.join(path_for_saving, "EMBIG Core")

    embi_global_div_compo = os.path.join(embi_global_div_dir, "Composition")
    embi_global_div_sub = os.path.join(embi_global_div_dir, "Sub-Index")
    embi_global_div_monthly_c_w = os.path.join(embi_global_div_dir, "Monthly-Country-Weights")
    embi_global_div_preview = os.path.join(embi_global_div_dir, "Preview")

    embi_global_compo = os.path.join(embi_global_dir, "Composition")
    embi_global_sub = os.path.join(embi_global_dir, "Sub-Index")
    embi_global_monthly_c_w = os.path.join(embi_global_dir, "Monthly-Country-Weights")
    embi_global_preview = os.path.join(embi_global_dir, "Preview")

    embi_plus_compo = os.path.join(embi_plus_dir, "Composition")
    embi_plus_sub = os.path.join(embi_plus_dir, "Sub-Index")
    embi_plus_preview = os.path.join(embi_plus_dir, "Preview")

    embig_core_compo = os.path.join(embig_core_dir, "Composition")
    embig_core_preview = os.path.join(embig_core_dir, "Preview")


    #création des dossiers à condition qu'il n'existe déja pas
    if not os.path.exists(embi_global_div_dir):
        os.makedirs(embi_global_div_dir)
    if not os.path.exists(embi_global_dir):
        os.makedirs(embi_global_dir)
    if not os.path.exists(embi_global_div_monthly_c_w):
        os.makedirs(embi_global_div_monthly_c_w)
    if not os.path.exists(embig_core_dir):
        os.makedirs(embig_core_dir)
    
    if not os.path.exists(embi_global_div_compo):
        os.makedirs(embi_global_div_compo)
    if not os.path.exists(embi_global_div_sub):
        os.makedirs(embi_global_div_sub)
    if not os.path.exists(embi_plus_dir):
        os.makedirs(embi_plus_dir)
    if not os.path.exists(embi_global_div_preview):
        os.makedirs(embi_global_div_preview)
    
    if not os.path.exists(embi_global_compo):
        os.makedirs(embi_global_compo)
    if not os.path.exists(embi_global_sub):
        os.makedirs(embi_global_sub)
    if not os.path.exists(embi_global_monthly_c_w):
        os.makedirs(embi_global_monthly_c_w)
    if not os.path.exists(embi_global_preview):
        os.makedirs(embi_global_preview)

    if not os.path.exists(embi_plus_compo):
        os.makedirs(embi_plus_compo)
    if not os.path.exists(embi_plus_sub):
        os.makedirs(embi_plus_sub)
    if not os.path.exists(embi_plus_preview):
        os.makedirs(embi_plus_preview)

    if not os.path.exists(embig_core_compo):
        os.makedirs(embig_core_compo)
    if not os.path.exists(embig_core_preview):
        os.makedirs(embig_core_preview)

    #renommer les fichiers correctement
    new_file_names = []
    for i, file_name in enumerate(file_names):
        if i == 0:
            new_file_name = f"{date}-jpmm-EMBI Global Div.-Composition.xlsx"
            new_file_names.append(new_file_name)
            source_path = os.path.join(download_folder, file_name)
            excel_file = os.path.splitext(source_path)[0] + ".xlsx"
            df = pd.read_csv(source_path, delimiter=',')
            df.to_excel(excel_file, index=False)
            source_path = excel_file
            target_path = os.path.join(embi_global_div_compo, new_file_name)
            shutil.move(source_path, target_path)
        elif i == 1:
            new_file_name = f"{date}-jpmm-EMBI Global Div.-Sub-Index.xlsx"
            new_file_names.append(new_file_name)
            source_path = os.path.join(download_folder, file_name)
            excel_file = os.path.splitext(source_path)[0] + ".xlsx"
            df = pd.read_csv(source_path, delimiter=',')
            df.to_excel(excel_file, index=False)
            source_path = excel_file
            target_path = os.path.join(embi_global_div_sub, new_file_name)
            shutil.move(source_path, target_path)
        elif i == 2:
            new_file_name = f"{date}-jpmm-EMBI Global Div.-Monthly-Country-Weights.xlsx"
            new_file_names.append(new_file_name)
            source_path = os.path.join(download_folder, file_name)
            excel_file = os.path.splitext(source_path)[0] + ".xlsx"
            df = pd.read_csv(source_path, delimiter=',')
            df.to_excel(excel_file, index=False)
            source_path = excel_file
            target_path = os.path.join(embi_global_div_monthly_c_w, new_file_name)
            shutil.move(source_path, target_path)
        elif i == 3:
            new_file_name = f"{date}-jpmm-EMBI Global Div.-Preview.xlsx"
            new_file_names.append(new_file_name)
            source_path = os.path.join(download_folder, file_name)
            excel_file = os.path.splitext(source_path)[0] + ".xlsx"
            df = pd.read_csv(source_path, delimiter=',')
            df.to_excel(excel_file, index=False)
            source_path = excel_file
            target_path = os.path.join(embi_global_div_preview, new_file_name)
            shutil.move(source_path, target_path)
        elif i == 4:
            new_file_name = f"{date}-jpmm-EMBI Global-Composition.xlsx"
            new_file_names.append(new_file_name)
            source_path = os.path.join(download_folder, file_name)
            excel_file = os.path.splitext(source_path)[0] + ".xlsx"
            df = pd.read_csv(source_path, delimiter=',')
            df.to_excel(excel_file, index=False)
            source_path = excel_file
            target_path = os.path.join(embi_global_compo, new_file_name)
            shutil.move(source_path, target_path)
        elif i == 5:
            new_file_name = f"{date}-jpmm-EMBI Global-Sub-Index.xlsx"
            new_file_names.append(new_file_name)
            source_path = os.path.join(download_folder, file_name)
            excel_file = os.path.splitext(source_path)[0] + ".xlsx"
            df = pd.read_csv(source_path, delimiter=',')
            df.to_excel(excel_file, index=False)
            source_path = excel_file
            target_path = os.path.join(embi_global_sub, new_file_name)
            shutil.move(source_path, target_path)
        elif i == 6:
            new_file_name = f"{date}-jpmm-EMBI Global-Monthly-Country-Weights.xlsx"
            new_file_names.append(new_file_name)
            source_path = os.path.join(download_folder, file_name)
            excel_file = os.path.splitext(source_path)[0] + ".xlsx"
            df = pd.read_csv(source_path, delimiter=',')
            df.to_excel(excel_file, index=False)
            source_path = excel_file
            target_path = os.path.join(embi_global_monthly_c_w, new_file_name)
            shutil.move(source_path, target_path)
        elif i == 7:
            new_file_name = f"{date}-jpmm-EMBI Global-Preview.xlsx"
            new_file_names.append(new_file_name)
            source_path = os.path.join(download_folder, file_name)
            excel_file = os.path.splitext(source_path)[0] + ".xlsx"
            df = pd.read_csv(source_path, delimiter=',')
            df.to_excel(excel_file, index=False)
            source_path = excel_file
            target_path = os.path.join(embi_global_preview, new_file_name)
            shutil.move(source_path, target_path)
        elif i == 8:
            new_file_name = f"{date}-jpmm-EMBI+-Composition.xlsx"
            new_file_names.append(new_file_name)
            source_path = os.path.join(download_folder, file_name)
            excel_file = os.path.splitext(source_path)[0] + ".xlsx"
            df = pd.read_csv(source_path, delimiter=',')
            df.to_excel(excel_file, index=False)
            source_path = excel_file
            target_path = os.path.join(embi_plus_compo, new_file_name)
            shutil.move(source_path, target_path)
        elif i == 9:
            new_file_name = f"{date}-jpmm-EMBI+-Sub-Index.xlsx"
            new_file_names.append(new_file_name)
            source_path = os.path.join(download_folder, file_name)
            excel_file = os.path.splitext(source_path)[0] + ".xlsx"
            df = pd.read_csv(source_path, delimiter=',')
            df.to_excel(excel_file, index=False)
            source_path = excel_file
            target_path = os.path.join(embi_plus_sub, new_file_name)
            shutil.move(source_path, target_path)
        elif i == 10:
            new_file_name = f"{date}-jpmm-EMBI+-Preview.xlsx"
            new_file_names.append(new_file_name)
            source_path = os.path.join(download_folder, file_name)
            excel_file = os.path.splitext(source_path)[0] + ".xlsx"
            df = pd.read_csv(source_path, delimiter=',')
            df.to_excel(excel_file, index=False)
            source_path = excel_file
            target_path = os.path.join(embi_plus_preview, new_file_name)
            shutil.move(source_path, target_path)
        elif i == 11:
            new_file_name = f"{date}-jpmm-EMBIG Core-Composition.xlsx"
            new_file_names.append(new_file_name)
            source_path = os.path.join(download_folder, file_name)
            excel_file = os.path.splitext(source_path)[0] + ".xlsx"
            df = pd.read_csv(source_path, delimiter=',')
            df.to_excel(excel_file, index=False)
            source_path = excel_file
            target_path = os.path.join(embig_core_compo, new_file_name)
            shutil.move(source_path, target_path)
        elif i == 12:
            new_file_name = f"{date}-jpmm-EMBIG Core-Preview.xlsx"
            new_file_names.append(new_file_name)
            source_path = os.path.join(download_folder, file_name)
            excel_file = os.path.splitext(source_path)[0] + ".xlsx"
            df = pd.read_csv(source_path, delimiter=',')
            df.to_excel(excel_file, index=False)
            source_path = excel_file
            target_path = os.path.join(embig_core_preview, new_file_name)
            shutil.move(source_path, target_path)
        else:
            break

    driver.quit()


download_folder = os.path.expandvars(r"%userprofile%\Downloads")
market_jp(r"U:\GDA\PFC\03_Gerants\03_12_NB\JP", download_folder, '20240830')




MS download :
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
import time
import config as cf
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import os
import datetime
import shutil
import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re


def matrix_ms(path_for_saving, path_default_download):
    options = webdriver.EdgeOptions()
    options.add_argument(f"user-data-dir={cf.local['userDataDir']}")

    download_dir = path_default_download
    options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,  # Prevents the prompt for download
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

    service = Service(cf.local["executablePath"])
    driver = webdriver.Edge(service=service, options=options)
    date_init = datetime.datetime.now()

    #supprimer l'historique de téléchargement
    driver.get("edge://downloads/all")

    try :
        wait = WebDriverWait(driver, 10)
        first_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[title="Effacer tout"]')))
        first_button.click()
        second_button = wait.until(EC.element_to_be_clickable((By.ID, 'confirmModalPrimaryButton')))
        second_button.click()
    except Exception as e:
        print("historique déjà vide:")

    #commencement du téléchargement
    driver.get("https://ny.matrix.ms.com/eqr/research/docs/content/equitystrategy/ui/index.html")

    #trouver la date
    time.sleep(5)
    view_info_div = driver.find_element(By.CLASS_NAME, 'view-info')
    full_text = view_info_div.text
    match = re.search(r'of\s(.*?)\s\(', full_text)
    if match:
        extracted_text = match.group(1)
        print(extracted_text)
        date_obj = datetime.datetime.strptime(extracted_text, '%d %B %Y')
        yesterday = date_obj.strftime('%Y%m%d')
    
    #initialisation des téléchargements

    i = 1
    while i <8:
        print(i)
        date_before = datetime.datetime.now()
        dropdownbox = driver.find_elements(by=By.TAG_NAME, value="Option")
        dropdownbox[i].click()
        #vérifie que le fichier est présent dans téléchargement 
        time.sleep(1)
        fname = []
        while(len(fname)!=1):
            now = datetime.datetime.now()
            for filename in os.listdir(path_default_download):
                if os.path.getmtime(os.path.join(path_default_download, filename)) > date_before.timestamp() and os.path.getmtime(os.path.join(path_default_download, filename)) < now.timestamp():
                    fname.append((filename, os.path.getmtime(os.path.join(path_default_download, filename))))
            time.sleep(1)

        #je vérifie que le fichier à son telechargement fini
        name = dropdownbox[i].text
        driver.get("edge://downloads/all")
        downloads_list = driver.find_elements(By.XPATH, "//div[@role='listitem']")
        for download_item in downloads_list :
            while download_item.text[-7:]!='dossier' :
                time.sleep(1)
        driver.get("https://ny.matrix.ms.com/eqr/research/docs/content/equitystrategy/ui/index.html")
        if(name=='QuantIndia'):
            break
        i = i+1


    time.sleep(5)
    #clic et download du bouton MOST (3 month Horizon)
    date_before = datetime.datetime.now()
    buttons = driver.find_elements(By.TAG_NAME, 'button')
    for button in buttons:
        if 'downloadXLSIcon' in button.get_attribute('class') and 'enabled' in button.get_attribute('class'):
            button.click()
    fname = []
    while(len(fname)!=1):
        now = datetime.datetime.now()
        for filename in os.listdir(path_default_download):
            if os.path.getmtime(os.path.join(path_default_download, filename)) > date_before.timestamp() and os.path.getmtime(os.path.join(path_default_download, filename)) < now.timestamp():
                fname.append((filename, os.path.getmtime(os.path.join(path_default_download, filename))))
        time.sleep(1)
    #je vérifie que le fichier à son telechargement fini
    driver.get("edge://downloads/all")
    downloads_list = driver.find_elements(By.XPATH, "//div[@role='listitem']")
    for download_item in downloads_list :
        while download_item.text[-7:]!='dossier' :
            time.sleep(1)

    driver.get("https://ny.matrix.ms.com/eqr/research/docs/content/equitystrategy/ui/index.html")
    time.sleep(5)
    date_before = datetime.datetime.now()
    buttons = driver.find_elements(By.TAG_NAME, 'button')
    div_element = driver.find_element(By.XPATH, "//div[label[text()='BEST']]")
    div_element.click()
    for button in buttons:
        if 'downloadXLSIcon' in button.get_attribute('class') and 'enabled' in button.get_attribute('class'):
            button.click()
    fname = []
    while(len(fname)!=1):
        now = datetime.datetime.now()
        for filename in os.listdir(path_default_download):
            if os.path.getmtime(os.path.join(path_default_download, filename)) > date_before.timestamp() and os.path.getmtime(os.path.join(path_default_download, filename)) < now.timestamp():
                fname.append((filename, os.path.getmtime(os.path.join(path_default_download, filename))))
        time.sleep(1)

    driver.get("edge://downloads/all")
    downloads_list = driver.find_elements(By.XPATH, "//div[@role='listitem']")
    for download_item in downloads_list :
        while download_item.text[-7:]!='dossier' :
            time.sleep(1)


    file_names_with_timestamps = []
    file_names = []
    now = datetime.datetime.now()
    num_files_downloaded = 0
    for filename in os.listdir(download_folder):
        if os.path.getmtime(os.path.join(download_folder, filename)) > date_init.timestamp() and os.path.getmtime(os.path.join(download_folder, filename)) < now.timestamp():
            file_names_with_timestamps.append((filename, os.path.getmtime(os.path.join(download_folder, filename))))
    #faire une liste avec le nom des fichiers rangés par date d'anciennete dans file_name (du plus vieux au plus jeune)
    file_names_with_timestamps.sort(key=lambda x: x[1], reverse=True)
    file_names = [x[0] for x in file_names_with_timestamps]
    file_names = file_names[::-1]

    print(f"Nombre de fichiers téléchargés: {len(file_names)}")

    if(len(file_names)!=9):
        print("erreur dans le chargement des fichiers, ces fichiers ont été downloads correctement :")
        print(file_names)
        raise Exception("Impossible de télécharger tous les fichiers nécessaires.")

    for file_name in file_names:
        print(file_name)
    #renommer les fichiers correctement
    new_file_names = []
    for i, file_name in enumerate(file_names):
        if i == 0:
            new_file_name = f"{yesterday}-Weekly Global Model Rankings-Global-DM.xlsx"
            new_file_names.append(new_file_name)
        elif i == 1:
            new_file_name = f"{yesterday}-Weekly Global Model Rankings-Emerging.xlsx"
            new_file_names.append(new_file_name)
        elif i == 2:
            new_file_name = f"{yesterday}-Weekly Global Model Rankings-China A.xlsx"
            new_file_names.append(new_file_name)
        elif i == 3:
            new_file_name = f"{yesterday}-QuanTopix SAS - HC.xlsx"
            new_file_names.append(new_file_name)
        elif i == 4:
            new_file_name = f"{yesterday}-QuantETF_model_output.xlsx"
            new_file_names.append(new_file_name)
        elif i == 5:
            new_file_name = f"{yesterday}-QuantChina_model.xlsx"
            new_file_names.append(new_file_name)
        elif i == 6:
            new_file_name = f"{yesterday}-QuantIndia_model.xlsx"
            new_file_names.append(new_file_name)
        elif i == 7:
            new_file_name = f"{yesterday}-EquityStrategyCompanies-Most3.xlsx"
            new_file_names.append(new_file_name)
        elif i == 8:
            new_file_name = f"{yesterday}-EquityStrategyCompanies-Best24.xlsx"
            new_file_names.append(new_file_name)
        else:
            break
        old_path = os.path.join(download_folder, file_name)
        new_path = os.path.join(download_folder, new_file_name)
        shutil.move(old_path, new_path)
        print(f"Fichier {file_name} renommé en {new_file_name}")

#créer le dossier où l'on va déplacer nos fichiers
    target_directory = os.path.join(path_for_saving, f"{yesterday} - MS matrix Data")
    if not os.path.exists(target_directory):
        os.makedirs(target_directory)

    #déplacer les fichiers renommés dans le nouveau dossier
    for new_file_name in new_file_names:
        source_path = os.path.join(download_folder, new_file_name)
        target_path = os.path.join(target_directory, new_file_name)
        shutil.move(source_path, target_path)
        print(f"Fichier {new_file_name} a été déplacé")
    driver.quit()


download_folder = os.path.expandvars(r"%userprofile%\Downloads")

#download_folder = r"U:\GDA\PFC\03_Gerants\03_12_NB\ok"
matrix_ms(r"U:\GDA\PFC\03_Gerants\03_12_NB\MS",  download_folder)





Config :

local = {
    "executablePath" : r"C:\Users\nbonneau\Downloads\edgedriver_win64\msedgedriver.exe",
    "userDataDir" : r"C:\Users\nbonneau\Downloads\edgedriver_win64\profile",
}


Nordic 
from tia.bbg import LocalTerminal
import pandas as pd
import datetime
from xbbg import blp
from blp import blp as bld
import pybbg as pybbg
import matplotlib.pyplot as plt
import os
from dateutil.relativedelta import relativedelta
import numpy as np
from datetime import datetime, date
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor, RandomForestClassifier
from sklearn.metrics import accuracy_score, confusion_matrix
import plotly.graph_objects as go
import pypfopt
from collections import OrderedDict
from pypfopt.efficient_frontier import EfficientFrontier
from pandas import Timestamp
import json

from pypfopt import plotting
import cvxpy as cp
from pypfopt.objective_functions import ex_ante_tracking_error

def end_month_data(df, force_last_date=False): #vérifie si les mois sont bien sont succesifs
    df = df.reset_index("date")
    last_date = df["date"].iloc[-1]
    print(last_date)
    if force_last_date:
        df["date"] = df["date"].apply(lambda x: x+pd.offsets.MonthEnd(1) if (x.month == (x+pd.offsets.MonthEnd(1)).month) else x)
    else:
        df["date"] = df["date"].apply(lambda x: x+pd.offsets.MonthEnd(1) if (x.month == (x+pd.offsets.MonthEnd(1)).month and x!=last_date) else x)
    df = df.set_index("date")
    return df

def expected_returns(start_date, end_date):
    bbg = pybbg.Pybbg()
    bquery = bld.BlpQuery().start() 
    #retrouver le taux attendu en allemagne
    crp_msci_europe = blp.bdp("VOW3 GY Equity", "COUNTRY_RISK_MARKET_RETURN", COUNTRY_RISK_EFFECTIVE_DATE=end_date).iloc[0,0]/100
    #retrouver le taux attendu sur les mids cap
    beta_mid_europe = blp.bdp("MCXR Index", "BETA_ADJ_OVERRIDABLE", BETA_OVERRIDE_REL_INDEX="MXEUSXNE Index",BETA_OVERRIDE_END_DT = end_date,BETA_OVERRIDE_START_DT = start_date).iloc[0,0]
    crp_mid_europe = beta_mid_europe*crp_msci_europe
        
    #retrouver le taux attendu sur les small cap
    beta_small_europe = blp.bdp("SCXR Index", "BETA_ADJ_OVERRIDABLE", BETA_OVERRIDE_REL_INDEX="MXEUSXNE Index",BETA_OVERRIDE_END_DT = end_date,BETA_OVERRIDE_START_DT = start_date).iloc[0,0]
    crp_small_europe = beta_small_europe*crp_msci_europe

    #retrouver le taux attendus sur letf msci nordic tr
    #sweden crp
    crp_sweden = blp.bdp("VOLVB SS Equity", "COUNTRY_RISK_MARKET_RETURN", COUNTRY_RISK_EFFECTIVE_DATE=end_date).iloc[0,0]/100
    #denmark crp
    crp_denmark = blp.bdp("NOVOB DC Equity", "COUNTRY_RISK_MARKET_RETURN", COUNTRY_RISK_EFFECTIVE_DATE=end_date).iloc[0,0]/100
    #norway crp
    crp_norway = blp.bdp("DNB NO Equity", "COUNTRY_RISK_MARKET_RETURN", COUNTRY_RISK_EFFECTIVE_DATE=end_date).iloc[0,0]/100
    #finland crp
    crp_finland = blp.bdp("NDA SS Equity", "COUNTRY_RISK_MARKET_RETURN", COUNTRY_RISK_EFFECTIVE_DATE=end_date).iloc[0,0]/100
    #recup les poids de chaque pays dans le MSCI Nordic
    datefin_format_tirets = f"{end_date[:4]}-{end_date[4:6]}-{end_date[6:]}"
    poids = bquery.bql(f"for(holdings('XDN0 GY Equity',dates='{datefin_format_tirets}')) get(sum(group(id().weights,COUNTRY_FULL_NAME)))")
    denmark_weight = poids.loc[poids['secondary_value']=='DENMARK']['value'].iloc[0]/100
    sweden_weight = poids.loc[poids['secondary_value']=='SWEDEN']['value'].iloc[0]/100
    finland_weight = poids.loc[poids['secondary_value']=='FINLAND']['value'].iloc[0]/100
    norway_weight = poids.loc[poids['secondary_value']=='NORWAY']['value'].iloc[0]/100
    crp_msci_nordic = denmark_weight*crp_denmark + sweden_weight*crp_sweden +crp_norway* norway_weight+finland_weight*crp_finland
    return crp_msci_europe, crp_mid_europe, crp_msci_nordic, crp_small_europe 


def h5store(filename, df, dic):
    store = pd.HDFStore(filename)
    store.put('mydata', df)
    store.get_storer('mydata').attrs.metadata = dic
    store.close()

def h5load(filename):
    with pd.HDFStore(filename) as store:
        data = store['mydata']
        metadata = store.get_storer('mydata').attrs.metadata
        data.attrs = metadata
    return data, metadata     


def _build_easy_map(assets):
    assets_easy_map = {}
    for k in assets.keys(): #récupère les tickers des assets
        assets_easy_map[assets[k]['ticker']] = assets[k]['des']
    return assets_easy_map

def download_data(start_date, end_date):
    assets = OrderedDict()
    assets['MXEUSXNE'] = {}
    assets['MXEUSXNE']['ticker'] = "MXEUSXNE Index"
    assets['MXEUSXNE']['des'] = "MSCI Europe"
    
    assets['MCXR'] = {}
    assets['MCXR']['ticker'] = "MCXR Index"
    assets['MCXR']['des'] = "MCXR Index"

    assets['XDN0 GY'] = {}
    assets['XDN0 GY']['ticker'] = "MSDENCN Index"
    assets['XDN0 GY']['des'] = "MSCI Nordic ETF"

    assets['SCXR'] = {}
    assets['SCXR']['ticker'] = "SCXR Index"
    assets['SCXR']['des'] = "SCXR Index"

    fields = ['PX_LAST']
    #Telechargement des prix et des returns en daily
    period = 'DAILY'
    resp = LocalTerminal.get_historical([assets[i]["ticker"] for i in assets.keys()],
        fields, start=start_date, end=end_date, period=period)
    daily_price = resp.as_frame()
    daily_price.columns = daily_price.columns.droplevel(1)
    daily_price.ffill(inplace=True)
    daily_price
    daily_price.rename(columns=_build_easy_map(assets), inplace=True)
    daily_return = daily_price.pct_change()
    daily_return = daily_return[start_date:]
    index_daily = daily_return.index
    
    #prochaine date de fin de mois
    today = datetime.date.today()
    prec_month = datetime.date(today.year, today.month, 1) - datetime.timedelta(days=1)
    prec_month = prec_month.strftime("%Y-%m-%d")
    next_month = datetime.date(today.year, today.month, 1) + relativedelta(months=+1) - datetime.timedelta(days=1)
    next_month = next_month.strftime("%Y-%m-%d")
    #Telechargement des prix et des returns en monthly
    period1 = 'MONTHLY'
    resp = LocalTerminal.get_historical([assets[i]["ticker"] for i in assets.keys()],
        fields, start=start_date, end=end_date, period=period1)
    monthly_price = resp.as_frame()
    monthly_price.columns = monthly_price.columns.droplevel(1)
    monthly_price.ffill(inplace=True)
    monthly_price
    monthly_price.rename(columns=_build_easy_map(assets), inplace=True)
    monthly_return = monthly_price.pct_change()
    monthly_return = monthly_return[start_date:]
   
    index_monthly = monthly_return.index
    expected_returns_df = pd.DataFrame(columns = ['date','MSCI Europe', 'MCXR Index', 'MSCI Nordic ETF','SCXR Index'])
#on enregistre la covariance du mois dans le dictionnaire covariance_dict avec des données daily, la clé correspond au mois où la covariance est observée
    for date in monthly_return.index:
        crp_msci_europe, crp_mid_europe, crp_msci_nordic, crp_small_europe = expected_returns(start_date, date.strftime('%Y%m%d'))
        expected_returns_df = expected_returns_df._append({'date': date, 'MSCI Europe' :crp_msci_europe , 'MCXR Index':crp_mid_europe , 'MSCI Nordic ETF':crp_msci_nordic,'SCXR Index':crp_small_europe}, ignore_index= True)

    #forcer la dernière valeur sur la fin du prochain mois
    expected_returns_next_month = expected_returns(start_date, today.strftime('%Y%m%d'))
    last_index = expected_returns_df.index[-1]+1
    expected_returns_df.loc[last_index,'date'] = next_month
    expected_returns_df.loc[last_index,'MSCI Europe'] = expected_returns_next_month[0]
    expected_returns_df.loc[last_index,'MCXR Index'] = expected_returns_next_month[1]
    expected_returns_df.loc[last_index,'MSCI Nordic ETF'] = expected_returns_next_month[2]
    expected_returns_df.loc[last_index,'SCXR Index'] = expected_returns_next_month[3]
    monthly_return.loc[next_month,'MSCI Europe'] = (daily_price.iloc[-1]['MSCI Europe']-daily_price.loc[prec_month,'MSCI Europe'])/daily_price.loc[prec_month,'MSCI Europe']
    monthly_return.loc[next_month,'MCXR Index'] = (daily_price.iloc[-1]['MCXR Index']-daily_price.loc[prec_month,'MCXR Index'])/daily_price.loc[prec_month,'MCXR Index']
    monthly_return.loc[next_month,'SCXR Index'] = (daily_price.iloc[-1]['SCXR Index']-daily_price.loc[prec_month,'SCXR Index'])/daily_price.loc[prec_month,'SCXR Index']
    monthly_return.loc[next_month,'MSCI Nordic ETF'] = (daily_price.iloc[-1]['MSCI Nordic ETF']-daily_price.loc[prec_month,'MSCI Nordic ETF'])/daily_price.loc[prec_month,'MSCI Nordic ETF']
    
    expected_returns_df = expected_returns_df.set_index('date')
    expected_returns_df= expected_returns_df.ffill()
    print(expected_returns_df)
    
    dossier_initial = r'U:\GDA\PFC\03_Gerants\03_12_NB\nordic'
    monthly_return = monthly_return.reset_index(drop=False)
    file_name_h5 = f"df_monthly_return.h5"
    df_member_path_h5 = os.path.join(dossier_initial, file_name_h5)
    metadata = dict(data="data monthly return",provider="Bloomberg",indice="MSCIEU, MXCR, SXCR, MSCI Nordic") 
    h5store(df_member_path_h5 ,monthly_return,metadata)
    file_name_xlsx = f"df_monthly_return.xlsx"
    df_member_path = os.path.join(dossier_initial, file_name_xlsx)
    monthly_return.to_excel(df_member_path, index=False)

    expected_returns_df = expected_returns_df.reset_index(drop=False)
    file_name_h5 = f"df_expected_returns.h5"
    df_member_path_h5 = os.path.join(dossier_initial, file_name_h5)
    metadata = dict(data="expected returns",provider="Bloomberg",indice="MSCIEU, MXCR, SXCR, MSCI Nordic") 
    h5store(df_member_path_h5 ,expected_returns_df,metadata)
    file_name_xlsx = f"df_expected_returns.xlsx"
    df_member_path = os.path.join(dossier_initial, file_name_xlsx)
    expected_returns_df.to_excel(df_member_path, index=False)

    daily_return = daily_return.reset_index(drop=False)
    file_name_h5 = f"df_daily_return.h5"
    df_member_path_h5 = os.path.join(dossier_initial, file_name_h5)
    metadata = dict(data="daily return",provider="Bloomberg",indice="MSCIEU, MXCR, SXCR, MSCI Nordic") 
    h5store(df_member_path_h5 ,daily_return,metadata)
    file_name_xlsx = f"df_daily_return.xlsx"
    df_member_path = os.path.join(dossier_initial, file_name_xlsx)
    daily_return.to_excel(df_member_path, index=False)
    
    

#download_data("20150902","20240630") #date final correspond à la date du jour






def compute_optimal_ptf(start_date, end_date, an, rolling = False):
    dossier_initial = r'U:\GDA\PFC\03_Gerants\03_12_NB\nordic'

    file_name_h5 = f"df_monthly_return.h5"
    monthly_return, metadata1 = h5load(os.path.join(dossier_initial, file_name_h5))
    monthly_return = monthly_return.set_index('date')


    file_name_h5 = f"df_expected_returns.h5"
    expected_returns_df, metadata2 = h5load(os.path.join(dossier_initial, file_name_h5))
    expected_returns_df = expected_returns_df.set_index('date')

    file_name_h5 = f"df_daily_return.h5"
    daily_return, metadata3 = h5load(os.path.join(dossier_initial, file_name_h5))
    daily_return = daily_return.set_index('date')

    
    monthly_return = monthly_return[start_date:end_date]
    daily_return = daily_return[start_date:monthly_return.index[-1]]
    index_monthly = monthly_return.index
    expected_returns_df = expected_returns_df[start_date:end_date]

    #mettre les expected returns au meme niveau
    #expected_returns_df['MCXR Index'] = expected_returns_df['MSCI Europe']
    #expected_returns_df['SCXR Index'] = expected_returns_df['MSCI Europe']
    #expected_returns_df['MSCI Nordic ETF'] = expected_returns_df['MSCI Europe']



    index_daily = daily_return.index
    covariance_dict = {}
    cor_matrix = {}
    for date in monthly_return.index:
        copy_daily_return = daily_return.loc[(index_daily.month == date.month) & (index_daily.year == date.year)]
        copy_daily_return = copy_daily_return.dropna()
        cov_1_month = copy_daily_return.cov(min_periods=20)
        covariance_dict[date] = cov_1_month*252
        cor_matrix[date] = copy_daily_return.corr(min_periods = 20)
    print(covariance_dict)

    #download taux sans risque europe
    indexes = ['EONIA Index', 'ESTRON Index']
    resp = LocalTerminal.get_historical(indexes, 'PX_LAST',start=start_date, end=end_date, period='MONTHLY')
    data = resp.as_frame()
    data.columns = data.columns.droplevel(1)
    str_series = pd.Series(index=pd.to_datetime(data.index), name='ST index', dtype='float64')
    ind = np.isnan(data[indexes[1]])
    str_series.loc[~ind] = data.loc[~ind, indexes[1]]
    str_series.loc[ind] = data.loc[ind, indexes[0]]
    print(str_series)

    monthly_return = monthly_return.iloc[1:]
    weights = []
    exp_ptf_perf = []
    date_tab = []
    datedeb = monthly_return.index[an+1]
    datedeb_perf = monthly_return.index[an]
    for i, date in enumerate(monthly_return.index[:-1]):
        if i>0:  
            if(date.strftime("%Y-%m-%d")<monthly_return.index[an].strftime("%Y-%m-%d") ):
                continue
            #monthly_returns_know contient toutes les returns monthly avant la date d'aujourd'hui donc le mois en cours et les mois d'avants
            monthly_returns_know= monthly_return.loc[:date]
            if rolling :
                if(len(monthly_returns_know) >an): #conservation des an dernières années pr l'estimation de la cov matrix
                    monthly_returns_know = monthly_returns_know[-an:]
                cov_matrix  = monthly_returns_know.cov()*12
            else :
                cov_matrix = covariance_dict[Timestamp(monthly_return.index[i])]

            #la cov est calculée uniquement avec les données daily du mois en cours
            #cov_matrix = covariance_dict[Timestamp(date)]
            

            #constraint tracking error
            
            benchmark_weights =  np.array([0.75,0.125,0,0.125])
            #dict_init = {'initvals': benchmark_weights}
            #dict_init = {"feastol": 0.0001}
            dict_init = {}
            ef = EfficientFrontier(expected_returns_df.loc[date], cov_matrix, verbose=True, solver_options = dict_init)
            ex = lambda w: ex_ante_tracking_error(w, ef.cov_matrix, benchmark_weights) <= 0.015
            #ef.add_constraint(ex)            
            #ef.add_constraint(ex_ante_tracking_error, cov_matrix=ef.cov_matrix, benchmark_weights=benchmark_weights)
            #ligne lorsque l'on ne connais pas les vrais expected YTM du prochain mois
            #ef = pypfopt.efficient_frontier.EfficientFrontier(expected_returns_aligned.loc[date], cov_matrix, weight_bounds=(0, 1), solver=None, verbose=False, solver_options=None)
            
            new_keys = ['MSCI Europe', 'MCXR Index', 'MSCI Nordic ETF', 'SCXR Index']
            try :
                riskfree = str_series.loc[date]
                poid = ef.max_sharpe(risk_free_rate=riskfree/100)
                poids = OrderedDict(zip(new_keys, poid.values()))
                weights.append({'date': date, 'weights': poids})
                print(date)
                exp_ptf_perf.append({'date':date, 'perf':ef.portfolio_performance(risk_free_rate=riskfree/100)})
            except Exception as e :
                print("error for date", date)
                #poids = 0
                poids = weights[-1]['weights']  # Get the previous weights
                weights.append({'date': date, 'weights': poids})
                #perf attendue sur le portefeuille pareil que celle d'avant car les poids ne bouge pas car l'optimiseur n'a pas réussi à résourdre le pb d'opti
                ptf_perf  = exp_ptf_perf[-1]['perf']
                exp_ptf_perf.append({'date':date, 'perf':ptf_perf})
    
    data_ptf = pd.DataFrame()
    for element in exp_ptf_perf:
        date = element['date']  # extraction de la date
        ptf_perf = element['perf'][0]
        ptf_sd = element['perf'][1]
        ratio_sortino = element['perf'][2]  # extraction des perf
        data_ptf.loc[date,'perf'] = ptf_perf
        data_ptf.loc[date,'ptf_sd'] = ptf_sd
        data_ptf.loc[date,'ratio_sortino'] = ratio_sortino
    data_ptf.index = monthly_return[datedeb_perf:monthly_return.index[-2]].index
    print(data_ptf)
    
    data = {}
    for element in weights:
        date = element['date']  # extraction de la date
        poids = element['weights']  # extraction des poids
        data[date] = poids

    df = pd.DataFrame(data)
    df = df.transpose()
    df_weights = df.copy()
    df_weights.index = monthly_return[datedeb:].index
    weighted_returns = monthly_return[datedeb:]*df_weights 
    strategy_returns = weighted_returns.sum(axis=1) #somme les poids * les returns
    strategy_returns_df = pd.DataFrame(data={'strategy_returns': strategy_returns})
    print(strategy_returns_df)
    cumulative_performance = (1 + strategy_returns_df['strategy_returns']).cumprod()
    print(cumulative_performance)
    years = (strategy_returns_df.index[-1] - strategy_returns_df.index[0]).days / 365
    annualized_return = (cumulative_performance[-1])**(1/years) - 1 
    print(annualized_return)
    print(df_weights)

    #cas LONG ONLY MSCI Europe
    copy = df
    copy['MSCI Europe'] =1
    copy['MCXR Index'] = 0
    copy['MSCI Nordic ETF'] = 0
    copy['SCXR Index'] = 0
    print(copy)
    copy.index = monthly_return[datedeb:].index
    weighted_returns_mscieu = monthly_return[datedeb:]*copy
    strategy_returns_mscieu = weighted_returns_mscieu.sum(axis=1)
    strategy_returns_df_mscieu = pd.DataFrame(data={'strategy_returns': strategy_returns_mscieu})
    print(strategy_returns_df_mscieu)
    cumulative_performance_mscieu = (1 + strategy_returns_df_mscieu['strategy_returns']).cumprod()

    #cas LONG ONLY MCXR Index
    copy_mid = df
    copy_mid['MSCI Europe'] =0
    copy_mid['MCXR Index'] = 1
    copy_mid['MSCI Nordic ETF'] = 0
    copy_mid['SCXR Index'] = 0
    print(copy_mid)
    copy_mid.index = monthly_return[datedeb:].index
    weighted_returns_mid = monthly_return[datedeb:]*copy_mid
    strategy_returns_mid = weighted_returns_mid.sum(axis=1)
    strategy_returns_df_mid = pd.DataFrame(data={'strategy_returns': strategy_returns_mid})
    print(strategy_returns_df_mid)
    cumulative_performance_mid = (1 + strategy_returns_df_mid['strategy_returns']).cumprod()
    
    #cas LONG ONLY MSCI Nordic ETF
    copy_nordic= df
    copy_nordic['MSCI Europe'] =0
    copy_nordic['MCXR Index'] = 0
    copy_nordic['MSCI Nordic ETF'] = 1
    copy_nordic['SCXR Index'] = 0
    print(copy_nordic)
    copy_nordic.index = monthly_return[datedeb:].index
    weighted_returns_nordic = monthly_return[datedeb:]*copy_nordic
    strategy_returns_nordic = weighted_returns_nordic.sum(axis=1)
    strategy_returns_df_nordic = pd.DataFrame(data={'strategy_returns': strategy_returns_nordic})
    print(strategy_returns_df_nordic)
    cumulative_performance_nordic = (1 + strategy_returns_df_nordic['strategy_returns']).cumprod()

    #cas LONG ONLY small
    copy_small= df
    copy_small['MSCI Europe'] =0
    copy_small['MCXR Index'] = 0
    copy_small['MSCI Nordic ETF'] = 0
    copy_small['SCXR Index'] = 1
    print(copy_small)
    copy_small.index = monthly_return[datedeb:].index
    weighted_returns_small = monthly_return[datedeb:]*copy_small
    strategy_returns_small = weighted_returns_small.sum(axis=1)
    strategy_returns_df_small = pd.DataFrame(data={'strategy_returns': strategy_returns_small})
    print(strategy_returns_df_small)
    cumulative_performance_small = (1 + strategy_returns_df_small['strategy_returns']).cumprod()


     #cas LONG portfeuille europe cdc
    copy_ptf= df
    copy_ptf['MSCI Europe'] =0.75
    copy_ptf['MCXR Index'] = 0.125
    copy_ptf['MSCI Nordic ETF'] = 0
    copy_ptf['SCXR Index'] = 0.125
    print(copy_ptf)
    copy_ptf.index = monthly_return[datedeb:].index
    weighted_returns_ptf = monthly_return[datedeb:]*copy_ptf
    strategy_returns_ptf = weighted_returns_small.sum(axis=1)
    strategy_returns_df_ptf = pd.DataFrame(data={'strategy_returns': strategy_returns_ptf})
    print(strategy_returns_df_ptf)
    cumulative_performance_ptf = (1 + strategy_returns_df_ptf['strategy_returns']).cumprod()
    #création du graphe de performance cumulative
    plt.figure(figsize=(10, 6))  
    #plt.figure(figsize=(10, 6))
    plt.plot(cumulative_performance.index, cumulative_performance.values, label='Performance cumulative', color='brown')
    plt.plot(cumulative_performance_small.index, cumulative_performance_small.values, label='Performance cumulative small', color='green')
    plt.plot(cumulative_performance_nordic.index, cumulative_performance_nordic.values, label='Performance cumulative nordic', color='lightskyblue')
    plt.plot(cumulative_performance_mid.index, cumulative_performance_mid.values, label='Performance cumulative mid', color='red')
    plt.plot(cumulative_performance_mscieu.index, cumulative_performance_mscieu.values, label='Performance cumulative MSCI EUROPE', color='black')
    plt.plot(cumulative_performance_ptf.index, cumulative_performance_ptf.values, label='Performance cumulative Portefeuille action EUROPE', color='violet')
    plt.title('Performance cumulative')
    plt.xlabel('Date')
    plt.ylabel('Valeur cumulative')
    plt.legend()
    plt.grid(True)
    destination_file = r'U:\GDA\PFC\03_Gerants\03_12_NB\nordic'
    name =  f"performance_cumulative_{start_date}_{end_date}.png"
    plt.savefig(os.path.join(destination_file, name))
    plt.show()




    fig, ax = plt.subplots(figsize=(10, 6))
    colors = ['black','red','lightskyblue','green']
    i = 0
    for col in df_weights.columns:
        ax.plot(df_weights.index, df_weights[col], label=col,color=colors[i] )
        i +=1
    ax.legend()
    ax.set_title('Évolution des poids au fil du temps with ')
    ax.set_xlabel('Date')
    ax.set_ylabel('Poids')
    #plt.savefig(os.path.join(destination_folder, "weight_matrix.png"))
    #print("La figure a été enregistrée sous :", destination_file)
    name =  f"weight_matrix_{start_date}_{end_date}.png"
    plt.savefig(os.path.join(destination_file, name))
    plt.show()

    #plot les expected returns et les expected standard deviation du ptf
    dates = data_ptf.index.tolist()
    perf = data_ptf['perf'].tolist()
    ptf_sd = data_ptf['ptf_sd'].tolist()
    plt.plot(dates, perf, label='Expected Portfolio Return')
    plt.plot(dates, ptf_sd, label='Expected Portfolio Semivariance')
    plt.xlabel('Date')
    plt.ylabel('Valeur')
    plt.title('Expected Portfolio Performance et Semivariance en fonction du temps')
    plt.legend()
    name =  f"expected_ptf_perf_{start_date}_{end_date}.png"
    plt.savefig(os.path.join(destination_file, name))
    plt.show()


    #ratio de sortino 
    dates = data_ptf.index.tolist()
    sortino_ratio = data_ptf['ratio_sortino'].tolist()
    plt.plot(dates, sortino_ratio, label='Expected Portfolio Ratio de Sortino')
    plt.xlabel('Date')
    plt.ylabel('Valeur')
    plt.title('Expected Portfolio Ratio de Sortino en fonction du temps')
    plt.legend()
    name =  f"sortino_ratio_{start_date}_{end_date}.png"
    plt.savefig(os.path.join(destination_file, name))
    plt.show()


    #pour tracer la frontière efficiente

    S = cov_matrix
    mu = expected_returns_df.loc[end_date]
    n_samples = 1000
    w = np.random.dirichlet(np.ones(len(mu)), n_samples)
    rets = w.dot(mu)
    stds = np.sqrt((w.T * (S @ w.T)).sum(axis=0))
    sharpes = rets / stds
    ef =  pypfopt.efficient_frontier.EfficientFrontier(mu, S)
    fig, ax = plt.subplots()
    plotting.plot_efficient_frontier(ef, ax=ax, show_assets=False)
    ''# Find and plot the tangency portfolio
    ef2 =  pypfopt.efficient_frontier.EfficientFrontier(mu, S)
    ef2.max_sharpe()
    ret_tangent, std_tangent, _ = ef2.portfolio_performance()
    # Plot random portfolios
    ax.scatter(stds, rets, marker=".", c=sharpes, cmap="gray")
    print(mu)
    print(S.iloc[0,0])
    ax.set_title("Efficient Frontier with random portfolios with ")
    #max_sharpe =monthly_returns[end_date]*df.loc[end_date]
    #ordonnee = max_sharp.sum(axis=1)
    ax.scatter(std_tangent, ret_tangent, marker="o", color='r', label='Portfolio Optimal')
    ax.legend()
    plt.tight_layout()
    plt.savefig(os.path.join(destination_file, "Efficient Frontier.png"))
    plt.close(fig)

    print("Les derniers poids optimaux proposés sont :")
    print(df_weights.iloc[-1])

compute_optimal_ptf("20151030","20240630",60,False)








Skew :
from tia.bbg import LocalTerminal
import pandas as pd
import datetime
from xbbg import blp
import pybbg as pybbg
import matplotlib.pyplot as plt
import os
from dateutil.relativedelta import relativedelta
import numpy as np

def h5store(filename, df, dic):
    store = pd.HDFStore(filename)
    store.put('mydata', df)
    store.get_storer('mydata').attrs.metadata = dic
    store.close()

def h5load(filename):
    with pd.HDFStore(filename) as store:
        data = store['mydata']
        metadata = store.get_storer('mydata').attrs.metadata
        data.attrs = metadata
    return data, metadata     


def download_vol_data(date_hist): 


    #charger l'ensemble des tickers existants
    bbg = pybbg.Pybbg()
    #tickers_dispo =  blp.bds('SPX Index','OPT_CHAIN')
    date_str = date_hist.strftime('%Y%m%d')
    tickers_dispo =  blp.bds('SPX Index','OPT_CHAIN',SINGLE_DATE_OVERRIDE=date_str)
    tickers_dispo = tickers_dispo['security_description']
    print(tickers_dispo)
    #récupérer le vrai nom du ticker
    tickers_names = blp.bdp(tickers_dispo, "SECURITY_DES")
    tickers_names = tickers_names['security_des'].tolist()
    print(tickers_names)

    dates_availables = []



    for ticker in tickers_names:
        parts = ticker.split()
        for part in parts:
            if "/" in part and len(part) == 8:
                dates_availables.append(part)
                
#for i in range(0,len(tickers_names)) : 
 #       ticker = tickers_names[i]
    #    dates_availables.append(ticker[-14:-6])
    dates_availables_unique = list(set(dates_availables)) #recupérer que les éléments uniques
    print(dates_availables_unique)
    
    #cherche le ticker le plus proche dans deux mois
    #today = datetime.date.today()
    today = date_hist
    next_date = today + datetime.timedelta(days=60)
    next_date_format = next_date.strftime('%m/%d/%y')
    

    tomorrow = today + datetime.timedelta(days=1)
    tomorrow_format_compact  = tomorrow.strftime('%Y%m%d')

    yesterday = today - datetime.timedelta(days=1)
    yesterday_format_compact  = yesterday.strftime('%Y%m%d')



    trouve = False
    sens_chrono = True
    jour = 1
    while(trouve==False):
        if next_date_format in dates_availables_unique:
            trouve = True
            print("l'option la plus proche de la maturité dans 2 mois est à la date ", next_date_format)
        else : 
            if(sens_chrono):
                next_date = next_date + datetime.timedelta(days=jour)
                sens_chrono = False
            else : 
                next_date = next_date - datetime.timedelta(days=jour)
                sens_chrono = True
            jour +=1
            next_date_format = next_date.strftime('%m/%d/%y')

    #recup les noms des tickers qui correspond a next_date_format
    tickers_date = [tickers_names[i] for i, date in enumerate(dates_availables) if date == next_date_format]
    print(tickers_date)


    today_format_compact  = today.strftime('%Y%m%d')
    spot1 = LocalTerminal.get_historical('SPX Index', 'PX_LAST',f'{today_format_compact}',f'{today_format_compact}').as_frame()
    spot = spot1.iloc[0,0]
    print("Le prix du SPX pour le jour ", today_format_compact, " est égal : ",  spot)


    #cas du call
    df_call = pd.DataFrame()
    df_call['ticker'] = [tickers_date[i] for i, date in enumerate(tickers_date) if date[-5] == 'C']
    df_call['strike'] = df_call['ticker'].apply(lambda x: x[-4:])
    df_call = df_call.sort_values('strike', ascending = True)
    #cas du put
    df_put = pd.DataFrame()
    df_put['ticker'] = [tickers_date[i] for i, date in enumerate(tickers_date) if date[-5] == 'P']
    df_put['strike'] = df_put['ticker'].apply(lambda x: x[-4:])
    df_put = df_put.sort_values('strike', ascending = True)
    #calcul des bornes inf et sup en fonction du spot et ne garder que ca dans le dataframe

    inf = int(round(spot*0.8,-1))
    sup = int(round(spot*1.2,-1)) #remettre 1.2
    #cas du call
    df_call['strike'] = df_call['strike'].astype(int)
    df_call = df_call[(df_call['strike'] >= inf) & (df_call['strike'] <= sup)]
    #cas du put
    df_put['strike'] = df_put['strike'].astype(int)
    df_put = df_put[(df_put['strike'] >= inf) & (df_put['strike'] <= sup)]


    #mettre le nombre de contrats ouverts 

    df_call['ticker'] = df_call['ticker'] + ' Index'
    df_put['ticker'] = df_put['ticker'] + ' Index'

    #cas du call


    resp1 = pd.DataFrame()
    tickers1 = df_call['ticker'].tolist()
    #resp1 = LocalTerminal.get_reference_data(tickers1, 'OPEN_INT', today_format_compact, 'days=a').as_frame()
    #resp1 = resp1.rename(columns={'OPEN_INT': 'agreement'})
    #resp1['ticker'] = resp1.index
    #df_call = pd.merge(df_call, resp1, on='ticker')

    #source de warnings
    resp1 = pd.DataFrame(index=tickers1, columns=['OPEN_INT'])

    for ticker in tickers1:
        try:
            df = bbg.bdh([ticker], 'OPEN_INT', today_format_compact, tomorrow_format_compact)
            resp1.loc[ticker, 'OPEN_INT'] = df.iloc[0,0]
        except:
            try:
                df = bbg.bdh([ticker], 'OPEN_INT', yesterday_format_compact, today_format_compact)
                resp1.loc[ticker, 'OPEN_INT'] = df.iloc[0,0]
            except:
                resp1.loc[ticker, 'OPEN_INT'] = 0

    resp1 = resp1.rename(columns={'OPEN_INT': 'agreement'})
    resp1['ticker'] = resp1.index
    df_call = pd.merge(df_call, resp1, on='ticker')


    #cas du put
    resp2 = pd.DataFrame()
    tickers2 = df_put['ticker'].tolist()
    #resp2 = LocalTerminal.get_reference_data(tickers2, 'OPEN_INT', today_format_compact, 'days=a').as_frame()
    #resp2 = resp2.rename(columns = {'OPEN_INT': 'agreement'})
    #resp2['ticker'] = resp2.index
    #df_put = pd.merge(df_put, resp2, on='ticker')

    resp2 = pd.DataFrame(index=tickers2, columns=['OPEN_INT'])

    for ticker in tickers2:
        try:
            df2 = bbg.bdh([ticker], 'OPEN_INT', today_format_compact, tomorrow_format_compact)
            resp2.loc[ticker, 'OPEN_INT'] = df2.iloc[0,0]
        except:
            try:
                df2 = bbg.bdh([ticker], 'OPEN_INT', yesterday_format_compact, today_format_compact)
                resp2.loc[ticker, 'OPEN_INT'] = df2.iloc[0,0]
            except:
                resp2.loc[ticker, 'OPEN_INT'] = 0

    resp2 = resp2.rename(columns={'OPEN_INT': 'agreement'})
    resp2['ticker'] = resp2.index
    df_put = pd.merge(df_put, resp2, on='ticker')


    #rajouter une colonne avec la volatilité implicite
    #cas du call
    resp4 = pd.DataFrame()
    tickers4 = df_call['ticker'].tolist()
    if(datetime.date.today()==date_hist):
        resp4 = LocalTerminal.get_historical(tickers4, 'IVOL_MID',f'{today_format_compact}').as_frame()
    else: 
        resp4 = LocalTerminal.get_historical(tickers4, 'IVOL_MID',f'{today_format_compact}',f'{today_format_compact}').as_frame()
    resp4 = resp4.iloc[0]
    new_index = [index[0] for index in resp4.index]
    resp4 = resp4.copy()
    resp4.index = new_index
    resp4 = resp4.to_frame() #car il s'agit d'une series
    resp4.columns = [ 'implied_vol']
    resp4['ticker'] = resp4.index
    df_call= pd.merge(df_call, resp4, on='ticker')


    resp3 = pd.DataFrame()
    tickers3 = df_put['ticker'].tolist()
    if(datetime.date.today()==date_hist):
        resp3 = LocalTerminal.get_historical(tickers3, 'IVOL_MID',f'{today_format_compact}').as_frame()
    else :
        resp3 = LocalTerminal.get_historical(tickers3, 'IVOL_MID',f'{today_format_compact}',f'{today_format_compact}').as_frame()

    resp3 = resp3.iloc[0]
    new_index = [index[0] for index in resp3.index]
    resp3 = resp3.copy()
    resp3.index = new_index
    resp3 = resp3.to_frame() #car il s'agit d'une series
    resp3.columns = [ 'implied_vol']
    resp3['ticker'] = resp3.index
    df_put= pd.merge(df_put, resp3, on='ticker')

#faire des interpolations lorsque le nombre de contrats est égale à 0 pour l'index en question
    df_call = clean_dataframe_interpolation(df_call)
    df_put = clean_dataframe_interpolation(df_put)
    #faire un merge sur les strikes en commun 

    df_put2 = df_put.rename(columns={'ticker':'ticker put','strike': 'strike', 'agreement':'agreement put', 'implied_vol' :'implied_vol_put'})
    df_call2 = df_call.rename(columns={'ticker':'ticker call','strike': 'strike', 'agreement':'agreement call', 'implied_vol' :'implied_vol_call'})
    df_call_put = pd.merge(df_call2, df_put2, on = 'strike',how='inner') #avec correspondance uniquement
    df_call_put = df_call_put.dropna()

    #tracer le graphe skew call et put
    dossier_initial = r'U:\GDA\PFC\03_Gerants\03_12_NB\skew-download'
    dossier_path = os.path.join(dossier_initial, f"{today_format_compact}_Options_Expiring_{next_date_format.replace('/', '-')}")
    os.makedirs(dossier_path, exist_ok=True)
    plt.plot(df_call_put['strike'], df_call_put['implied_vol_call'], 'b', label='Implied Volatility du Call')
    plt.plot(df_call_put['strike'], df_call_put['implied_vol_put'], 'r', label='Implied Volatility du Put')
    
    plt.axvline(x=spot, color='g', linestyle='--', label='Spot')
    plt.text(spot, 0.5, 'ATM', color='g', va='center', ha='right')
    
    
    plt.xlabel('Strike')
    plt.ylabel('Volatility')
    title = f'Implied Volatility : Call vs Put ({next_date_format})'
    plt.title(title)
    plt.legend()
    graphique_path = os.path.join(dossier_path, 'call-put-skew.png')
    plt.savefig(graphique_path)
    plt.close()

    df_call_path = os.path.join(dossier_path, 'df_call.xlsx')
    df_call.to_excel(df_call_path, index=False)

    df_call_path_h5 = os.path.join(dossier_path, 'df_call.h5')
    metadata = dict(data="Ivol call",provider="Bloomberg",indice="SPX") 
    h5store(df_call_path_h5,df_call,metadata)


    df_put_path = os.path.join(dossier_path, 'df_put.xlsx')
    df_put.to_excel(df_put_path, index=False)
    df_put_path_h5 = os.path.join(dossier_path, 'df_put.h5')
    metadata = dict(data="Ivol put",provider="Bloomberg",indice="SPX") 
    h5store(df_put_path_h5,df_put,metadata)


    df_call_put_path = os.path.join(dossier_path, 'df_call_put.xlsx')
    df_call_put.to_excel(df_call_put_path, index=False)
    df_call_put_path_h5 = os.path.join(dossier_path, 'df_call_put.h5')
    metadata = dict(data="Ivol callput",provider="Bloomberg",indice="SPX") 
    h5store(df_call_put_path_h5 ,df_call_put,metadata)

    spot_path_h5 = os.path.join(dossier_path, 'spot.h5')
    metadata = dict(data="Spot SPX",provider="Bloomberg",indice="SPX") 
    h5store(spot_path_h5,spot1,metadata)

    plt.show()
    return dossier_path

def AMB_mean(df_call_put, spot):
    inf_bas_strike = int(round(spot * 0.8, -1))
    sup_bas_strike = int(round(spot * 0.95, -1))
    inf_haut_strike = int(round(spot * 1.05, -1))
    sup_haut_strike = int(round(spot * 1.2, -1))
    
    partie_gauche = df_call_put[
        (df_call_put['strike'] >= inf_haut_strike) & (df_call_put['strike'] <= sup_haut_strike)
    ]['implied_vol_put'].mean() + df_call_put[
        (df_call_put['strike'] >= inf_haut_strike) & (df_call_put['strike'] <= sup_haut_strike)
    ]['implied_vol_call'].mean()
    
    partie_droite = df_call_put[
        (df_call_put['strike'] >= inf_bas_strike) & (df_call_put['strike'] <= sup_bas_strike)
    ]['implied_vol_put'].mean() + df_call_put[
        (df_call_put['strike'] >= inf_bas_strike) & (df_call_put['strike'] <= sup_bas_strike)
    ]['implied_vol_call'].mean()
    
    AMB = (partie_gauche - partie_droite) / 2
    return AMB

def COMA_mean(df_call_put, spot):
    
    inf_bas_strike = int(round(spot*0.8,-1))
    sup_bas_strike = int(round(spot*0.95,-1))
    inf_atm = int(round(spot*0.95,-1))
    sup_atm = int(round(spot*1.05,-1))
    inf_haut_strike = int(round(spot*1.05,-1))
    sup_haut_strike = int(round(spot*1.2,-1))
    coma = df_call_put[(df_call_put['strike']>= inf_haut_strike)&(df_call_put['strike']<=sup_haut_strike)]['implied_vol_call'].mean() - df_call_put[(df_call_put['strike']>= inf_atm)&(df_call_put['strike']<=sup_atm)]['implied_vol_call'].mean()
    return coma

def POMA_mean(df_call_put, spot):
    inf_bas_strike = int(round(spot*0.8,-1))
    sup_bas_strike = int(round(spot*0.95,-1))
    inf_atm = int(round(spot*0.95,-1))
    sup_atm = int(round(spot*1.05,-1))
    poma = df_call_put[(df_call_put['strike']>= inf_bas_strike)&(df_call_put['strike']<=sup_bas_strike)]['implied_vol_put'].mean() - df_call_put[(df_call_put['strike']>= inf_atm)&(df_call_put['strike']<=sup_atm)]['implied_vol_put'].mean()
    return poma

def CW_mean(df_call_put, spot):
    inf_atm = int(round(spot*0.95,-1))
    sup_atm = int(round(spot*1.05,-1))
    cw = df_call_put[(df_call_put['strike']>= inf_atm)&(df_call_put['strike']<=sup_atm)]['implied_vol_call'].mean() - df_call_put[(df_call_put['strike']>= inf_atm)&(df_call_put['strike']<=sup_atm)]['implied_vol_put'].mean()
    return cw

def ZZX_mean(df_call_put, spot):
    inf_bas_strike = int(round(spot*0.8,-1))
    sup_bas_strike = int(round(spot*0.95,-1))
    inf_atm = int(round(spot*0.95,-1))
    sup_atm = int(round(spot*1.05,-1))
    zzx = df_call_put[(df_call_put['strike']>= inf_bas_strike)&(df_call_put['strike']<=sup_bas_strike)]['implied_vol_put'].mean() - df_call_put[(df_call_put['strike']>= inf_atm)&(df_call_put['strike']<=sup_atm)]['implied_vol_call'].mean()
    return zzx



def COMA_nearest(df_call_put, spot):
    inf_haut_strike = int(round(spot * 1.05, -1))
    sup_haut_strike = int(round(spot * 1.2, -1))
    inf_atm = int(round(spot*0.95,-1))
    sup_atm = int(round(spot*1.05,-1))
    df_call_atm = df_call_put[(df_call_put['strike']>= inf_atm) & (df_call_put['strike']<=sup_atm)].sort_values(by ='strike',ascending = True)
    spot_strike = df_call_atm['strike'].sub(spot).abs()
    idx_nearest_strike = spot_strike.idxmin()
    call_atm = df_call_atm.loc[idx_nearest_strike,'implied_vol_call']
    df_call_otm = df_call_put[(df_call_put['strike'] >= inf_haut_strike) & (df_call_put['strike'] <= sup_haut_strike)].sort_values(by ='strike',ascending=True)
    call_otm = df_call_otm.loc[df_call_otm.index[0], 'implied_vol_call']
    COMA = call_otm - call_atm
    return COMA

def POMA_nearest(df_call_put, spot):
    inf_bas_strike = int(round(spot*0.8,-1))
    sup_bas_strike = int(round(spot*0.95,-1))
    inf_atm = int(round(spot*0.95,-1))
    sup_atm = int(round(spot*1.05,-1))
    df_put_atm = df_call_put[(df_call_put['strike']>= inf_atm) & (df_call_put['strike']<=sup_atm)].sort_values(by ='strike',ascending = True)
    spot_strike = df_put_atm['strike'].sub(spot).abs()
    idx_nearest_strike = spot_strike.idxmin()
    put_atm = df_put_atm.loc[idx_nearest_strike,'implied_vol_put']
    df_put_otm = df_call_put[(df_call_put['strike']>= inf_bas_strike)&(df_call_put['strike']<=sup_bas_strike)].sort_values(by ='strike',ascending=False)
    put_otm = df_put_otm.loc[df_put_otm.index[0], 'implied_vol_put']
    POMA = put_otm - put_atm
    return POMA


def CW_nearest(df_call_put, spot):
    inf_atm = int(round(spot*0.95,-1))
    sup_atm = int(round(spot*1.05,-1))
    df_atm = df_call_put[(df_call_put['strike']>= inf_atm) & (df_call_put['strike']<=sup_atm)].sort_values(by ='strike',ascending = True)
    spot_strike = df_atm['strike'].sub(spot).abs()
    idx_nearest_strike = spot_strike.idxmin()
    put_atm = df_atm.loc[idx_nearest_strike,'implied_vol_put']
    call_atm = df_atm.loc[idx_nearest_strike,'implied_vol_call']
    CW = call_atm - put_atm
    return CW

def ABM_nearest_strike(df_call_put, spot):
    inf_haut_strike = int(round(spot * 1.05, -1))
    inf_bas_strike = int(round(spot * 0.8, -1))
    sup_haut_strike = int(round(spot * 1.2, -1))
    sup_bas_strike = int(round(spot * 0.95, -1))

    df_put_itm = df_call_put[(df_call_put['strike'] >= inf_haut_strike) & (df_call_put['strike'] <= sup_haut_strike)].sort_values(by ='strike',ascending=True)
    put_itm = df_put_itm.loc[df_put_itm.index[0], 'implied_vol_put']

    df_call_otm = df_call_put[(df_call_put['strike'] >= inf_haut_strike) & (df_call_put['strike'] <= sup_haut_strike)].sort_values(by ='strike',ascending=True)
    call_otm = df_call_otm.loc[df_call_otm.index[0], 'implied_vol_call']

    df_put_otm = df_call_put[(df_call_put['strike'] <= sup_bas_strike) & (df_call_put['strike'] >= inf_bas_strike)].sort_values(by ='strike',ascending=False)
    put_otm = df_put_otm.loc[df_put_otm.index[0], 'implied_vol_put']

    df_call_itm = df_call_put[(df_call_put['strike'] <= sup_bas_strike) & (df_call_put['strike'] >= inf_bas_strike)].sort_values(by ='strike',ascending=False)
    call_itm = df_call_itm.loc[df_call_itm.index[0], 'implied_vol_call']
    AMB = ((put_itm + call_otm) - (call_itm + put_otm)) / 2
    return AMB

def ZZX_nearest(df_call_put, spot):
    inf_bas_strike = int(round(spot*0.8,-1))
    sup_bas_strike = int(round(spot*0.95,-1))
    inf_atm = int(round(spot*0.95,-1))
    sup_atm = int(round(spot*1.05,-1))
    df_call_atm = df_call_put[(df_call_put['strike']>= inf_atm) & (df_call_put['strike']<=sup_atm)].sort_values(by ='strike',ascending = True)
    spot_strike = df_call_atm['strike'].sub(spot).abs()
    idx_nearest_strike = spot_strike.idxmin()
    call_atm = df_call_atm.loc[idx_nearest_strike,'implied_vol_call']
    df_put_otm = df_call_put[(df_call_put['strike']>= inf_bas_strike)&(df_call_put['strike']<=sup_bas_strike)].sort_values(by ='strike',ascending=False)
    put_otm = df_put_otm.loc[df_put_otm.index[0], 'implied_vol_put']
    ZZX = put_otm - call_atm
    return ZZX


def compute_metrics(dossier_path, date_hist):
    df_call_put, metadata1 = h5load(dossier_path+'\\df_call_put.h5')
    df_call, metadata2 = h5load(dossier_path+ '\\df_call.h5')
    df_put, metadata3 = h5load(dossier_path+ '\\df_put.h5')
    df_spot, metadata4 = h5load(dossier_path+ '\\spot.h5')
    spot = df_spot.iloc[0,0]
    date_str = date_hist.strftime('%Y%m%d')
    df_metrics = pd.DataFrame()
    df_metrics.loc[0,'Date'] = date_str
    
    #calcul de l'AMB
    df_metrics.loc[0,'Mean AMB'] = AMB_mean(df_call_put, spot)
    df_metrics.loc[0,'AMB'] = ABM_nearest_strike(df_call_put, spot)
    #calcul du COMA
    df_metrics.loc[0,'Mean COMA'] = COMA_mean(df_call_put, spot)
    df_metrics.loc[0,'COMA'] = COMA_nearest(df_call_put, spot)
    #calcul du POMA
    df_metrics.loc[0,'Mean POMA'] = POMA_mean(df_call_put, spot)
    df_metrics.loc[0,'POMA'] = POMA_nearest(df_call_put, spot)
    #calcul du CW
    df_metrics.loc[0, 'Mean CW'] = CW_mean(df_call_put, spot)
    df_metrics.loc[0,'CW'] = CW_nearest(df_call_put, spot)
    #calcul du ZZX
    df_metrics.loc[0, 'Mean ZZX'] = ZZX_mean(df_call_put, spot)
    df_metrics.loc[0, 'ZZX'] = ZZX_nearest(df_call_put, spot)

    #calcul des returns du mois suivants
    #l'objectif de ces 5 lignes c'est d'avoir la date de fin de mois ouvré du mois suivant pour pouvoir calculer la perf du mois suivant
    date_plus_2month = date_hist+ relativedelta(months=2)
    first_day_next_month = datetime.date(date_plus_2month.year, date_plus_2month.month, 1)
    date_suiv = first_day_next_month - datetime.timedelta(days=10)
    date_plus_2month = date_plus_2month.strftime('%Y%m%d')
    liste = fin_de_mois_jours_ouvres( date_suiv, date_plus_2month)


    spot_in_1month = LocalTerminal.get_historical('SPX Index', 'PX_LAST',f'{liste[0]}',f'{liste[0]}').as_frame()
    spot_in_1month = spot_in_1month.iloc[0,0]
    return_1m = (spot_in_1month-spot)/spot
    df_metrics.loc[0,'Future return'] = return_1m 

    return df_metrics



def clean_dataframe_interpolation(df):
    #si des strikes sont en doublons, je garde la ligne qui contient le nombre d'agreements le plus important
    keep_lines = {}
    for index, row in df.iterrows():
        strike = row["strike"]
        if strike in keep_lines:
            if row["agreement"] > keep_lines[strike]["agreement"]:
                keep_lines[strike] = row
        else:
            keep_lines[strike] = row


    df = pd.DataFrame.from_dict(keep_lines, orient="index").reset_index(drop=True)
    i = 0
    while(df.loc[i,'agreement'] == 0) :
        df = df.drop(i,axis=0)
        print("supression de la première ligne contenant un agreement = 0 , interpolation impossible, pas de data avant")
        i +=1


    i = len(df.index)-1
    while(df.loc[i,'agreement'] == 0) :
        df = df.drop(i,axis=0)
        print("supression de la dernière ligne contenant un agreement = 0 , interpolation impossible, pas de data après")
        i -=1

    for index_ticker in df.index: 
        if df.loc[index_ticker,'agreement'] == 0 :
            prec_vol = df.loc[index_ticker- 1, 'implied_vol']
            prec_strike  = df.loc[index_ticker - 1, 'strike']
            i = index_ticker +1
            #on doit selectionner la volatilité suivante dont le nombre d'agreement est =! 0
            while(df.loc[i,'agreement'] == 0):
                i+=1
            suiv_vol = df.loc[i, 'implied_vol']
            suiv_strike  = df.loc[i, 'strike']
            interpo = prec_vol+(suiv_vol-prec_vol)*((df.loc[index_ticker, 'strike']-prec_strike)/(suiv_strike-prec_strike))
            df.loc[index_ticker, 'implied_vol'] = interpo                 
    return df
    


def fin_de_mois_jours_ouvres(datedeb, datefin):
    bbg = pybbg.Pybbg()
    df= pd.DataFrame()
    df = bbg.bdh("SPX Index", 'PX_LAST', datedeb, datefin)
    df['Date'] = df.index
    df = df.reset_index(drop=True)
    liste = []
    for i in range(0,len(df.index)-2) :
        if(df.loc[i,'Date'].month != df.loc[i+1,'Date'].month):
            liste.append(df.loc[i,'Date'])
    return liste


def download_data(datedeb, datefin):
    liste_fin_mois = fin_de_mois_jours_ouvres(datedeb, datefin)
    for i in liste_fin_mois :
        download_vol_data(i.date())
    return


#download_data(20240101,20240605)




df_metrics_decembre =compute_metrics(r'U:\GDA\PFC\03_Gerants\03_12_NB\skew-download\20231229_Options_Expiring_02-16-24', datetime.date(2023,12,29))
df_metrics_janvier =compute_metrics(r'U:\GDA\PFC\03_Gerants\03_12_NB\skew-download\20240131_Options_Expiring_03-15-24', datetime.date(2024,1,31))
df_metrics_fevrier =compute_metrics(r'U:\GDA\PFC\03_Gerants\03_12_NB\skew-download\20240229_Options_Expiring_04-19-24', datetime.date(2024,2,29))
df_metrics_mars = compute_metrics(r'U:\GDA\PFC\03_Gerants\03_12_NB\skew-download\20240328_Options_Expiring_05-17-24', datetime.date(2024,3,28))
df_metrics_avril =  compute_metrics(r'U:\GDA\PFC\03_Gerants\03_12_NB\skew-download\20240430_Options_Expiring_06-21-24',datetime.date(2024, 4, 30))



dataframes = [df_metrics_decembre, df_metrics_janvier, df_metrics_fevrier, df_metrics_mars, df_metrics_avril]
df_merged = pd.concat(dataframes, ignore_index=True)
print(df_merged)

dossier_path = r'U:\GDA\PFC\03_Gerants\03_12_NB\skew-download'
df_merged_path= os.path.join(dossier_path, 'resultats_metrics_IVOL.xlsx')
df_merged.to_excel(df_merged_path, index=False)



df = df_merged
df['Date'] = pd.to_datetime(df['Date'], format='%Y%m%d')
plt.figure(figsize=(12, 6))
x = np.arange(len(df))  
width = 0.15
plt.bar(x - 2*width, df['AMB'], width=width, label='AMB')
plt.bar(x - width, df['ZZX'], width=width, label='ZZX')
plt.bar(x, df['COMA'], width=width, label='COMA')
plt.bar(x + width, df['POMA'], width=width, label='POMA')
plt.bar(x + 2*width, df['CW'], width=width, label='CW')

for i, future_return in enumerate(df['Future return']):
    plt.text(i, df['AMB'][i] + df['ZZX'][i] + df['COMA'][i] + df['POMA'][i] + df['CW'][i] + 0.2, f"{future_return*100:.2f}", ha='center')
plt.plot(x, 100*df['Future return'], color='red', label='Future return en %')
plt.xticks(x, df['Date'], rotation=45)
plt.legend()
plt.title('relation entre les métriques sur la vol imp  et le rendement mensuel du mois suivant')
plt.xlabel('Date')
plt.ylabel('Valeur')
plt.show()


Small vs large : 
#importation des librairies nécessaires
from tia.bbg import LocalTerminal
import pandas as pd
import datetime
from xbbg import blp
from blp import blp as bld
import pybbg as pybbg
import matplotlib.pyplot as plt
import os
from sklearn.metrics import mean_absolute_error
from dateutil.relativedelta import relativedelta
import numpy as np
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score, confusion_matrix
import shap 
from sklearn.metrics import mean_squared_error
from sklearn.metrics import accuracy_score
from sklearn.linear_model import LogisticRegressionCV
from datetime import datetime, timedelta, date
from tqdm import tqdm
import shutil
import re

working_folder = r"U:\GDA\PFC\03_Gerants\03_12_NB\Small VS Large"
class PrevisionError(Exception):
    pass


def create_or_replace_directory(directory_path):
    if os.path.exists(directory_path):
        shutil.rmtree(directory_path)
    os.makedirs(directory_path)

def h5store(filename, df, dic):
    store = pd.HDFStore(filename)
    store.put('mydata', df)
    store.get_storer('mydata').attrs.metadata = dic
    store.close()

def h5load(filename):
    with pd.HDFStore(filename) as store:
        data = store['mydata']
        metadata = store.get_storer('mydata').attrs.metadata
        data.attrs = metadata
    return data, metadata     

def find_latest_file(working_folder, prefix): #trouver le dernier (le + récent) fichier .hdf5 qui comportent les données des index LCXR et SCXR
    data_index_folder = os.path.join(working_folder, 'data_index')
    files = os.listdir(data_index_folder)
    filtered_files = [f for f in files if f.startswith(f'df_{prefix}') and f.endswith('.h5')]

    def extract_date(filename):
        match = re.search(r'(\d{8})\.h5$', filename)
        if match:
            return match.group(1)
        return None
    latest_file = max(filtered_files, key=lambda f: extract_date(f))
    latest_file_path = os.path.join(data_index_folder, latest_file)

    return latest_file_path


#PARTIE CHARGEMENT DE LA DATA
#à partir de deux dates, on trouve la date de début de mois où il y a eu une cotation
def find_prochaine_date_mois(datedeb, datefin):
    bbg = pybbg.Pybbg()
    df= pd.DataFrame()
    df = bbg.bdh("LCXR Index", 'PX_LAST', datedeb, datefin)
    df['Date'] = df.index
    df = df.reset_index(drop=True)
    for i in range(1,len(df.index)) :
        if(df.loc[i,'Date'].month != df.loc[i-1,'Date'].month):
             return df.loc[i,'Date']
    raise ValueError("Aucun changement de mois n'a été trouvé entre les dates fournies (rendement inconnus à cette date) ",datefin)



    
def download_data_per(index,datefin, working_folder = working_folder) :
    #récup les poids  dans l'index à une certaine date
    datefin_format_tirets = f"{datefin[:4]}-{datefin[4:6]}-{datefin[6:]}"
    bbg = pybbg.Pybbg()
    bquery = bld.BlpQuery().start()
    df_member = bquery.bql(f"for(MEMBERS('{index}',DATES='{datefin_format_tirets}')) get(ID().WEIGHTS)")
    df_member = df_member.rename(columns={'security': 'index_member'})
    df_member = df_member.drop(['field', 'secondary_name', 'secondary_value'], axis=1)
    df_member = df_member.rename(columns={'value': 'weight'})
    print(df_member)
    
    #recup les secteurs
    sector = blp.bdp(df_member['index_member'], "GICS_SECTOR_NAME")
    sector['index_member'] = sector.index
    df_member = pd.merge(df_member, sector, on = 'index_member', how='left')

    #recup la dette short et long terme
    lt_debt = bquery.bql(f"for(members('{index}',dates='{datefin_format_tirets}')) get(bs_lt_borrow(dates='{datefin_format_tirets}', AE=A, fill=prev, currency=EUR).value)")
    st_debt = bquery.bql(f"for(members('{index}',dates='{datefin_format_tirets}')) get(bs_st_borrow(dates='{datefin_format_tirets}', AE=A, fill=prev, currency=EUR).value)")
    
    lt_debt= lt_debt.groupby('security').agg(list)
    st_debt= st_debt.groupby('security').agg(list)

    for i in range(max(st_debt['value'].str.len())):
        st_debt[f'value_{i+1}'] = st_debt['value'].apply(lambda x: x[i] if len(x) > i else '')
    for i in range(max(lt_debt['value'].str.len())):
        lt_debt[f'value_{i+1}'] = lt_debt['value'].apply(lambda x: x[i] if len(x) > i else '')

    lt_debt = lt_debt.drop('value', axis=1)
    lt_debt = lt_debt.rename(columns={
        'value_1': 'lt_debt',
    })
    st_debt = st_debt.drop('value', axis=1)
    st_debt = st_debt.rename(columns={
        'value_1': 'st_debt',
    })

    lt_debt= lt_debt.drop(['field', 'secondary_name', 'secondary_value'], axis=1)
    st_debt= st_debt.drop(['field', 'secondary_name', 'secondary_value'], axis=1)
    df_member = pd.merge(df_member, lt_debt,left_on='index_member', right_on='security', how='left')
    df_member = pd.merge(df_member, st_debt,left_on='index_member', right_on='security', how='left')

    #calcul la somme lt debt et st term debt
    df_member['st_lt_debt'] = df_member['lt_debt'] + df_member['st_debt']

    #recup le wacc
    wacc = blp.bdh(df_member['index_member'],'WACC_COST_DEBT',datefin, datefin,  Days = 'NON_TRADING_WEEKDAYS' )
    wacc['index_member'] = wacc.index
    wacc = wacc.iloc[0]
    new_index = [index[0] for index in wacc.index]
    new_index  = new_index.copy()
    wacc.index = new_index
    wacc['index_member'] = wacc.index
    wacc = wacc[:-1]
    wacc = wacc.rename('wacc')
    df_member = pd.merge(df_member, wacc,left_on='index_member', right_on=wacc.index, how='left')


    #recup le currency adjusted enterprise value
    crncy = bquery.bql(f"for(members('{index}',dates='{datefin_format_tirets}')) get(curr_entp_val(dates='{datefin_format_tirets}', fill=prev, currency = EUR).value)")
    crncy= crncy.groupby('security').agg(list)
    for i in range(max(crncy['value'].str.len())):
        crncy[f'value_{i+1}'] = crncy['value'].apply(lambda x: x[i] if len(x) > i else '')
    crncy = crncy.drop('value', axis=1)
    crncy= crncy.rename(columns={'value_1': 'crncy_adj_ev' })
    crncy= crncy.drop(['field', 'secondary_name', 'secondary_value'], axis=1)
    df_member = pd.merge(df_member, crncy,left_on='index_member', right_on='security', how='left')

    #recup le long terme price earnings
    lt_pe = blp.bdh(df_member['index_member'],'LONG_TERM_PRICE_EARNINGS_RATIO',datefin, datefin,  Days = 'NON_TRADING_WEEKDAYS' )
    lt_pe['index_member'] = lt_pe.index
    if(len(lt_pe)!= 0):
        lt_pe= lt_pe.iloc[0]
        new_index = [index[0] for index in lt_pe.index]
        new_index  = new_index.copy()
        lt_pe.index = new_index
        lt_pe['index_member'] = lt_pe.index
        lt_pe= lt_pe[:-1]
        lt_pe = lt_pe.rename('lt_pe')
        df_member= pd.merge(df_member, lt_pe,left_on='index_member', right_on=lt_pe.index, how='left')

    #recup l'earnings, pe ratio et best pe ratio 
    bbg = pybbg.Pybbg()
    bquery = bld.BlpQuery().start()
    #df_pe = bquery.bql(f"for(members('{index}',dates='{datefin_format_tirets}')) get(cur_mkt_cap(dates='{datefin_format_tirets}', fill=prev, currency=EUR).value,(pe_ratio(FPT=LTM,FPO=0, AE=E,  dates='{datefin_format_tirets}', fill=prev).value),(pe_ratio(dates = '{datefin_format_tirets}', fill=prev).value))")
    df_pe = bquery.bql(f"for(members('{index}',dates='{datefin_format_tirets}')) get(cur_mkt_cap(dates='{datefin_format_tirets}', fill=prev, currency=EUR).value,(pe_ratio(FPT=LTM,FPO=1, AE=E,  as_of_date='{datefin_format_tirets}', fill=prev).value),(pe_ratio(as_of_date = '{datefin_format_tirets}', fill=prev).value))")

    print(df_member)
    df_pe = df_pe.groupby('security').agg(list)
    for i in range(max(df_pe['value'].str.len())):
        df_pe[f'value_{i+1}'] = df_pe['value'].apply(lambda x: x[i] if len(x) > i else '')

    df_pe = df_pe.drop('value', axis=1)
    df_pe = df_pe.rename(columns={
        'value_1': 'cur_mrk_cap',
        'value_2': 'best_pe_ratio',
        'value_3': 'pe_ratio'
    })

    df_pe = df_pe.drop(['field', 'secondary_name', 'secondary_value'], axis=1)
    df_member = pd.merge(df_member, df_pe,left_on='index_member', right_on='security', how='left')
    print(df_member)

    df_member.replace('', np.nan, inplace=True)
    #df_member['best_pe_ratio'] = df_member['best_pe_ratio'].fillna(df_member['pe_ratio'])
    for i, value in enumerate(df_member['best_pe_ratio']):

        if pd.isna(value) or value == '' or value == 'NaN'  or value =='nan':
            df_member.at[i, 'best_pe_ratio'] = df_member.at[i, 'pe_ratio']

    for i, value in enumerate(df_member['pe_ratio']):

        if pd.isna(value) or value == '' or value == 'NaN'  or value =='nan':
            df_member.at[i, 'pe_ratio'] = df_member.at[i, 'best_pe_ratio']

    dossier_deb = working_folder 
    dossier_initial = os.path.join(dossier_deb, f"members_{datefin}")
    os.makedirs(dossier_initial, exist_ok=True)

    dossier_path = os.path.join(dossier_initial, f"{index}_members_")
    os.makedirs(dossier_path, exist_ok=True)
    file_name_xlsx = f"df_member_{index}_{datefin_format_tirets}.xlsx"
    df_member_path = os.path.join(dossier_path, file_name_xlsx)
    df_member.to_excel(df_member_path, index=False)

    file_name_h5 = f"df_member_{index}_{datefin_format_tirets}.h5"
    df_member_path_h5 = os.path.join(dossier_path, file_name_h5)
    metadata = dict(data="data index member",provider="Bloomberg",indice="SXXR") 
    h5store(df_member_path_h5 ,df_member,metadata)

#fonction pour calculer une nouvelle version du pe
def compute_pe_u(index,datefin, compute_futur_return = True, working_folder = working_folder):
    datefin_format_tirets = f"{datefin[:4]}-{datefin[4:6]}-{datefin[6:]}"
    bbg = pybbg.Pybbg()
    bquery = bld.BlpQuery().start()
    file_name_h5 = f"df_member_{index}_{datefin_format_tirets}.h5"
    dossier_deb = working_folder 
    dossier_initial = os.path.join(dossier_deb, f"members_{datefin}")
    os.makedirs(dossier_initial, exist_ok=True)

    dossier_path = os.path.join(dossier_initial, f"{index}_members_")
    df_member, metadata1 = h5load(os.path.join(dossier_path, file_name_h5))
    print(df_member)

    if(compute_futur_return==True):
        date_obj = datetime.strptime(datefin, "%Y%m%d")
        date_next_month = date_obj + relativedelta(days=35)
        date1mois = date_next_month.strftime("%Y%m%d")
        next_date = find_prochaine_date_mois(datefin, date1mois)
        next_date = next_date.strftime("%Y%m%d")
        next_month_return = LocalTerminal.get_reference_data(df_member['index_member'],"CUST_TRR_RETURN_HOLDING_PER", CUST_TRR_END_DT=next_date, CUST_TRR_START_DT=datefin, CUST_TRR_CRNCY="EUR").as_frame()
        df_member = pd.merge(df_member, next_month_return, left_on='index_member', right_on=next_month_return.index, how='left')
        df_member = df_member.rename(columns={'CUST_TRR_RETURN_HOLDING_PER': 'next_month_return'})

    #convertir les lignes en numeric
    cols_to_convert = ['cur_mrk_cap', 'st_lt_debt', 'crncy_adj_ev', 'wacc', 'best_pe_ratio', 'pe_ratio']
    df_member[cols_to_convert] = df_member[cols_to_convert].apply(pd.to_numeric, errors='coerce')

    df_member.replace('', np.nan, inplace=True)
    #df_member['best_pe_ratio'] = df_member['best_pe_ratio'].fillna(df_member['pe_ratio'])
    
    #calcul du pe_u à partir de la formule donnée
    df_member['pe_u'] = df_member.apply(lambda row: 
                                    row['cur_mrk_cap'] / 
                                    ((row['st_lt_debt'] / row['crncy_adj_ev']) * 
                                      (((row['wacc']/100) * row['cur_mrk_cap']) - (row['cur_mrk_cap'] / row['best_pe_ratio'])) + 
                                     (row['cur_mrk_cap'] / row['best_pe_ratio'])) if not np.isnan(row['cur_mrk_cap']) and not np.isnan(row['st_lt_debt']) and not np.isnan(row['crncy_adj_ev']) and not np.isnan(row['wacc']) and not np.isnan(row['best_pe_ratio']) else np.nan, 
                                    axis=1)
    #df_member['pe_u'] = df_member['cur_mrk_cap']/((df_member['st_lt_debt']/df_member['crncy_adj_ev'])*(df_member['wacc']*df_member['cur_mrk_cap']-(df_member['cur_mrk_cap']/df_member['best_pe_ratio']))+(df_member['cur_mrk_cap']/df_member['best_pe_ratio'])) 

    file_name_xlsx = f"df_pe_u_{index}_{datefin_format_tirets}.xlsx"
    df_member_path = os.path.join(dossier_path, file_name_xlsx)
    df_member.to_excel(df_member_path, index=False)
    
    file_name_h5 = f"df_pe_u_{index}_{datefin_format_tirets}.h5"
    df_member_path_h5 = os.path.join(dossier_path, file_name_h5)
    metadata = dict(data="data index member pe_u",provider="Bloomberg",indice="SXXR") 
    h5store(df_member_path_h5 ,df_member,metadata)
    return 0


#compare_pe_u qui permet de créer un fichier .h5 et excel pour chaque date de début de mois commun aux larges, mids et smalls caps.
#avec notamment le pe median,le pe_u median, les returns du mois suivant par catégorie (large, small et mids) et par secteur
def compare_pe_u(datefin, compute_futur_return=True, working_folder = working_folder):
    datefin_format_tirets = f"{datefin[:4]}-{datefin[4:6]}-{datefin[6:]}"
    bbg = pybbg.Pybbg()
    bquery = bld.BlpQuery().start()

    dossier_deb = working_folder 
    dossier_initial = os.path.join(dossier_deb, f"members_{datefin}")
    os.makedirs(dossier_initial, exist_ok=True)


    dossier_path_scxr = os.path.join(dossier_initial, f"SCXR Index_members_")
    file_name_h5_scxr = f"df_pe_u_SCXR Index_{datefin_format_tirets}.h5"
    file_scxr = os.path.join(dossier_path_scxr, file_name_h5_scxr)
    if os.path.exists(file_scxr):
        print("Le fichier SCXR existe.")
    else:
        print("Le fichier SCXR n'existe pas à la date demandée.")
        return 0

    dossier_path_mcxr = os.path.join(dossier_initial, f"MCXR Index_members_")
    file_name_h5_mcxr = f"df_pe_u_MCXR Index_{datefin_format_tirets}.h5"
    file_mcxr = os.path.join(dossier_path_mcxr, file_name_h5_mcxr)
    if os.path.exists(file_mcxr):
        print("Le fichier MCXR existe.")
    else:
        print("Le fichier MCXR n'existe pas à la date demandée.")
        return 0

    dossier_path_lcxr = os.path.join(dossier_initial, f"LCXR Index_members_")
    file_name_h5_lcxr = f"df_pe_u_LCXR Index_{datefin_format_tirets}.h5"
    file_lcxr = os.path.join(dossier_path_lcxr, file_name_h5_lcxr)
    if os.path.exists(file_lcxr):
        print("Le fichier LCXR existe.")
    else:
        print("Le fichier LCXR n'existe pas à la date demandée.")
        return 0 

    df_scxr, metadata1 = h5load(os.path.join(dossier_path_scxr, file_name_h5_scxr))
    df_mcxr, metadata2 = h5load(os.path.join(dossier_path_mcxr, file_name_h5_mcxr))
    df_lcxr, metadata3 = h5load(os.path.join(dossier_path_lcxr, file_name_h5_lcxr))


    df_scxr[df_scxr == 'NaN'] = np.nan
    df_mcxr[df_mcxr == 'NaN'] = np.nan
    df_lcxr[df_lcxr == 'NaN'] = np.nan
    df_scxr.replace('', np.nan, inplace=True)
    df_mcxr.replace('', np.nan, inplace=True)
    df_lcxr.replace('', np.nan, inplace=True)

    df = pd.DataFrame()
    categories = df_scxr['gics_sector_name'].unique()
    index = 0
    for categorie in categories : 

        if categorie != 'Financials':
            #sélection des dataframes correspondants
            print(categorie)
            category_scxr = df_scxr[df_scxr['gics_sector_name']==categorie] 
            category_scxr = category_scxr[category_scxr['pe_u'].notna()]
            category_scxr = category_scxr[category_scxr['pe_u'] != '']
            category_mcxr = df_mcxr[df_mcxr['gics_sector_name']==categorie] 
            category_mcxr = category_mcxr[category_mcxr['pe_u'].notna()]
            category_mcxr = category_mcxr[category_mcxr['pe_u'] != '']
            category_lcxr = df_lcxr[df_lcxr['gics_sector_name']==categorie] 
            category_lcxr = category_lcxr[category_lcxr['pe_u'].notna()]
            category_lcxr = category_lcxr[category_lcxr['pe_u'] != '']
            df.loc[index, 'secteur'] = categorie


            
            #calcul du pe median par catégorie
            df.loc[index, 'Median SCXR pe'] = category_scxr['pe_ratio'].median()
            df.loc[index, 'Median MCXR pe'] = category_mcxr['pe_ratio'].median()
            df.loc[index, 'Median LCXR pe'] = category_lcxr['pe_ratio'].median()
            
            #calcul du pe_u médian par catégorie
            df.loc[index, 'Median SCXR pe_u'] = category_scxr['pe_u'].median()
            df.loc[index, 'Median MCXR pe_u'] = category_mcxr['pe_u'].median()
            df.loc[index, 'Median LCXR pe_u'] = category_lcxr['pe_u'].median()
            
            #calcul de pe en faisant une pondération par market cap par catégorie
            df.loc[index, 'SCXR pe mrk cap weighted'] =  ((category_scxr['cur_mrk_cap'] / category_scxr['cur_mrk_cap'].sum()) * category_scxr['pe_ratio']).sum()
            df.loc[index, 'MCXR pe mrk cap weighted'] =  ((category_mcxr['cur_mrk_cap'] / category_mcxr['cur_mrk_cap'].sum()) * category_mcxr['pe_ratio']).sum()
            df.loc[index, 'LCXR pe mrk cap weighted'] =  ((category_lcxr['cur_mrk_cap'] / category_lcxr['cur_mrk_cap'].sum()) * category_lcxr['pe_ratio']).sum()
            
            #calcul de pe en faisant une pondération par market cap par catégorie
            df.loc[index, 'SCXR pe_u mrk cap weighted'] = ((category_scxr['cur_mrk_cap'] / category_scxr['cur_mrk_cap'].sum()) * category_scxr['pe_u']).sum()
            df.loc[index, 'MCXR pe_u mrk cap weighted'] = ((category_mcxr['cur_mrk_cap'] / category_mcxr['cur_mrk_cap'].sum()) * category_mcxr['pe_u']).sum()
            df.loc[index, 'LCXR pe_u mrk cap weighted'] = ((category_lcxr['cur_mrk_cap'] / category_lcxr['cur_mrk_cap'].sum()) * category_lcxr['pe_u']).sum()
            
            #calcul returns du mois suivant de la cat
            if(compute_futur_return == True):
                category_scxr = category_scxr[category_scxr['next_month_return'].notna()]
                category_scxr = category_scxr[category_scxr['next_month_return'] != '']
                category_mcxr = category_mcxr[category_mcxr['next_month_return'].notna()]
                category_mcxr = category_mcxr[category_mcxr['next_month_return'] != '']
                category_lcxr = category_lcxr[category_lcxr['next_month_return'].notna()]
                category_lcxr = category_lcxr[category_lcxr['next_month_return'] != '']

                df.loc[index, 'SCXR next month return'] = ((category_scxr['cur_mrk_cap'] / category_scxr['cur_mrk_cap'].sum()) * category_scxr['next_month_return']).sum()
                df.loc[index, 'MCXR next month return'] = ((category_mcxr['cur_mrk_cap'] / category_mcxr['cur_mrk_cap'].sum()) * category_mcxr['next_month_return']).sum()
                df.loc[index, 'LCXR next month return'] = ((category_lcxr['cur_mrk_cap'] / category_lcxr['cur_mrk_cap'].sum()) * category_lcxr['next_month_return']).sum()

            index +=1
    

    df_scxr = df_scxr[df_scxr['pe_u'].notna() & (df_scxr['pe_u'] != '') & (df_scxr['pe_u'] != 'NaN') & (df_scxr['pe_u'] != 'Na') & (~df_scxr['pe_u'].isin(['nan']))]
    df_scxr = df_scxr[df_scxr['pe_ratio'].notna() & (df_scxr['pe_ratio'] != '') & (df_scxr['pe_ratio'] != 'NaN') & (df_scxr['pe_ratio'] != 'Na') & (~df_scxr['pe_ratio'].isin(['nan']))]

    df_mcxr = df_mcxr[df_mcxr['pe_u'].notna() & (df_mcxr['pe_u'] != '') & (df_mcxr['pe_u'] != 'NaN') & (df_mcxr['pe_u'] != 'Na') & (~df_mcxr['pe_u'].isin(['nan']))]
    df_mcxr = df_mcxr[df_mcxr['pe_ratio'].notna() & (df_mcxr['pe_ratio'] != '') & (df_mcxr['pe_ratio'] != 'NaN') & (df_mcxr['pe_ratio'] != 'Na') & (~df_mcxr['pe_ratio'].isin(['nan']))]

    df_lcxr = df_lcxr[df_lcxr['pe_u'].notna() & (df_lcxr['pe_u'] != '') & (df_lcxr['pe_u'] != 'NaN') & (df_lcxr['pe_u'] != 'Na') & (~df_lcxr['pe_u'].isin(['nan']))]
    df_lcxr = df_lcxr[df_lcxr['pe_ratio'].notna() & (df_lcxr['pe_ratio'] != '') & (df_lcxr['pe_ratio'] != 'NaN') & (df_lcxr['pe_ratio'] != 'Na') & (~df_lcxr['pe_ratio'].isin(['nan']))]

    df.loc[index, 'secteur'] = 'Full Index'
    df.loc[index, 'Median SCXR pe'] = df_scxr['pe_ratio'].median()
    df.loc[index, 'Median MCXR pe'] = df_mcxr['pe_ratio'].median()
    df.loc[index, 'Median LCXR pe'] = df_lcxr['pe_ratio'].median()
    df.loc[index, 'Median SCXR pe_u'] = df_scxr['pe_u'].median()
    df.loc[index, 'Median MCXR pe_u'] = df_mcxr['pe_u'].median()
    df.loc[index, 'Median LCXR pe_u'] = df_lcxr['pe_u'].median()
    df.loc[index, 'SCXR pe mrk cap weighted'] =  ((df_scxr['cur_mrk_cap'] / df_scxr['cur_mrk_cap'].sum()) * df_scxr['pe_ratio']).sum()
    df.loc[index, 'MCXR pe mrk cap weighted'] =  ((df_mcxr['cur_mrk_cap'] / df_mcxr['cur_mrk_cap'].sum()) * df_mcxr['pe_ratio']).sum()
    df.loc[index, 'LCXR pe mrk cap weighted'] =  ((df_lcxr['cur_mrk_cap'] / df_lcxr['cur_mrk_cap'].sum()) * df_lcxr['pe_ratio']).sum()
    df.loc[index, 'SCXR pe_u mrk cap weighted'] = ((df_scxr['cur_mrk_cap'] / df_scxr['cur_mrk_cap'].sum()) * df_scxr['pe_u']).sum()
    df.loc[index, 'MCXR pe_u mrk cap weighted'] = ((df_mcxr['cur_mrk_cap'] / df_mcxr['cur_mrk_cap'].sum()) * df_mcxr['pe_u']).sum()
    df.loc[index, 'LCXR pe_u mrk cap weighted'] = ((df_lcxr['cur_mrk_cap'] / df_lcxr['cur_mrk_cap'].sum()) * df_lcxr['pe_u']).sum()
   
    if(compute_futur_return == True):
        df_scxr = df_scxr[df_scxr['next_month_return'].notna() & (df_scxr['next_month_return'] != '')]
        df_mcxr = df_mcxr[df_mcxr['next_month_return'].notna() & (df_mcxr['next_month_return'] != '')]
        df_lcxr = df_lcxr[df_lcxr['next_month_return'].notna() & (df_lcxr['next_month_return'] != '')]

        df.loc[index, 'SCXR next month return'] = ((df_scxr['cur_mrk_cap'] / df_scxr['cur_mrk_cap'].sum()) * df_scxr['next_month_return']).sum()
        df.loc[index, 'MCXR next month return'] = ((df_mcxr['cur_mrk_cap'] / df_mcxr['cur_mrk_cap'].sum()) * df_mcxr['next_month_return']).sum()
        df.loc[index, 'LCXR next month return'] = ((df_lcxr['cur_mrk_cap'] / df_lcxr['cur_mrk_cap'].sum()) * df_lcxr['next_month_return']).sum()

    file_name_xlsx = f"df_score_pe_{datefin_format_tirets}.xlsx"
    df_score_path = os.path.join(dossier_initial, file_name_xlsx)
    df.to_excel(df_score_path, index=False)

    file_name_h5 = f"df_score_pe_{datefin_format_tirets}.h5"
    df_score_path_h5 = os.path.join(dossier_initial, file_name_h5)
    metadata = dict(data="results pe_u",provider="Bloomberg",indice="SXXR") 
    h5store(df_score_path_h5 ,df,metadata)
    print(df)
    return 0

#fonction à appeler pour charger la data entre deux données et mettre la data dans un folder
def deb_de_mois_jours_ouvres(datedeb, datefin):
    bbg = pybbg.Pybbg()
    df= pd.DataFrame()
    df = bbg.bdh("LCXR Index", 'PX_LAST', datedeb, datefin)
    df['Date'] = df.index
    df = df.reset_index(drop=True)
    liste = []
    for i in range(1,len(df.index)) :
        if(df.loc[i,'Date'].month != df.loc[i-1,'Date'].month):
            liste.append(df.loc[i,'Date'])
    liste = [date.strftime('%Y%m%d') for date in liste]
    for i in liste :
        download_data_per('SCXR Index', i)
        download_data_per('LCXR Index', i)
        download_data_per('MCXR Index', i)
        now = datetime.now()
        date_obj = datetime.strptime(i, "%Y%m%d")
        if date_obj.month == now.month and date_obj.year == now.year:
            compute_pe_u('SCXR Index', i, False)
            compute_pe_u('MCXR Index', i, False)
            compute_pe_u('LCXR Index', i, False)
            compare_pe_u(i, False)
        else : 
            compute_pe_u('SCXR Index', i)
            compute_pe_u('MCXR Index', i)
            compute_pe_u('LCXR Index', i)
            compare_pe_u(i)
        
#fonction à appeler pour télécharger l'ensemble des données entre la date1 et la date2 en pas mensuel début de mois de jour ouvrés
#deb_de_mois_jours_ouvres(20240625,20240728)

#PARTIE CALCUL POUR LA REGRESSION LOGISTIQUE ET LA PREDICTION 

def sharpe_ratio(df_pred_test,df_result_test,diff_real_large_small, name, directory_path):
    #calculer le ratio de sharpe de la difference entre large reconstitué - small reconstitué
    returns_test_reconstituted_large_small = df_pred_test['next_month_return_large'] - df_pred_test['next_month_return_small']
    sharpe_ratio_test_reconstituted_large_small = np.mean(returns_test_reconstituted_large_small ) / np.std(returns_test_reconstituted_large_small)
    #calculer le ratio de sharpe de la différence entre small reconstitué - large reconstitué
    returns_test_reconstituted_small_large = df_pred_test['next_month_return_small'] - df_pred_test['next_month_return_large']
    sharpe_ratio_test_reconstituted_small_large = np.mean(returns_test_reconstituted_small_large ) / np.std(returns_test_reconstituted_small_large)
    #calculer le ratio de sharpe de la stratégie appliqué sur le large-small reconstitué (si prevision = 1) et sur le small-large reconstitué (si prévision = 0)
    returns_test_reconstituted_strat = df_result_test['result']
    sharpe_ratio_test_reconstituted_strat = np.mean(returns_test_reconstituted_strat) / np.std(returns_test_reconstituted_strat)
    #calculer le ratio de sharpe de la stratégie appliqué sur le vrai index LCXR-SCXR (si prévision =1)et sur le SCXR-LCXR (si prévision =0)
    returns_test_real_diff_large_small = diff_real_large_small
    sharpe_ratio_test_real_diff_large_small = np.mean(returns_test_real_diff_large_small) / np.std(returns_test_real_diff_large_small)

    sharpe_ratios = [
        sharpe_ratio_test_reconstituted_large_small,
        sharpe_ratio_test_reconstituted_small_large,
        sharpe_ratio_test_reconstituted_strat,
        sharpe_ratio_test_real_diff_large_small
    ]

    labels = [
        'Reconstituted Large - Small',
        'Reconstituted Small - Large',
        'Reconstituted Strategy',
        'Strategy on Real Large - Small']
    #ploter avec la bonne légende les ratios de sharpe sur un même graphe dans le bon dossier
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(labels, sharpe_ratios, 'o')
    ax.set_title('Ratios de Sharpe')
    ax.set_xlabel('')
    ax.set_ylabel('Ratio de Sharpe')
    ax.grid(True)
    full_directory_sharpe = os.path.join(directory_path,'Sharpe Ratio -'+ name +'.png')
    plt.savefig(full_directory_sharpe)
    plt.close()

#ratio information moyenne des écart de rendemments avec le bench(excess returns) sur l'écart type 
#attention ratio d'information mensuel
def IF_ratio(portfolio_returns,benchmark_returns):
    excess_returns = portfolio_returns - benchmark_returns
    information_ratio = excess_returns.mean() / excess_returns.std()
    return information_ratio
#plot l'information ratio
def information_ratio(df_result_test,df_pred_test,diff_real_large_small, name,directory_path):
    #ratio d'information de la stratégie appliquée aux index reconstitués (long large short small ou inversement) avec comme benchmark le large-small reconstitué
    test_reconstituted_strat_returns = df_result_test['result']
    reconstitued_large_small_benchmark_returns = df_pred_test['next_month_return_large'] - df_pred_test['next_month_return_small']
    information_ratio_test_reconstituted_strat_vs_reconstitued_large_small = IF_ratio(test_reconstituted_strat_returns,reconstitued_large_small_benchmark_returns)
    #ratio d'information de la stratégie appliquée aux index reconstitués (long large short small ou inversement) avec comme benchmark le small-large reconstitué
    reconstitued_small_large_benchmark_returns = df_pred_test['next_month_return_small'] - df_pred_test['next_month_return_large']
    information_ratio_test_reconstituted_strat_vs_reconstitued_small_large = IF_ratio(test_reconstituted_strat_returns,reconstitued_small_large_benchmark_returns)
    #ratio d'information de la stratégié appliqué aux vrais index (long LCXR short SCXR ou inversement) avec comme benchmark le large-small reconstitué
    test_strat_returns_real_large_small =diff_real_large_small
    information_ratio_test_strat_real_vs_reconstitued_large_small = IF_ratio(test_strat_returns_real_large_small,reconstitued_large_small_benchmark_returns)
    #ratio d'information de la stratégié appliqué aux vrais index (long LCXR short SCXR ou inversement) avec comme benchmark le small-large reconstitué
    information_ratio_test_strat_real_vs_reconstitued_small_large = IF_ratio(test_strat_returns_real_large_small,reconstitued_small_large_benchmark_returns)
    x = [1, 2, 3, 4]
    y = [information_ratio_test_reconstituted_strat_vs_reconstitued_large_small,
        information_ratio_test_reconstituted_strat_vs_reconstitued_small_large,
        information_ratio_test_strat_real_vs_reconstitued_large_small,
        information_ratio_test_strat_real_vs_reconstitued_small_large]
    #plot la figure
    plt.figure(figsize=(10, 6))
    plt.scatter(x, y)
    plt.xticks(x, ['Strat Reconst I & b=Reconst L-S',
                'Strat Reconst I & b=Reconst S-L',
                'Strat Real I & b=Reconst L-S',
                'Strat Real I & b=Reconst S-L'])
    plt.xlabel('Comparison')
    plt.ylabel('Information Ratio')
    plt.title('Information Ratios')
    plt.grid()
    full_directory_IF = os.path.join(directory_path,'information_ratios'+ name+'.png')
    plt.savefig(full_directory_IF)
    plt.close()
    

def lasso_data(datedeb, datefin, working_folder = working_folder):
    bbg = pybbg.Pybbg()
    df= pd.DataFrame()
    df = bbg.bdh("LCXR Index", 'PX_LAST', datedeb, datefin)
    df['Date'] = df.index
    df = df.reset_index(drop=True)
    liste = []
    for i in range(1,len(df.index)) :
        if(df.loc[i,'Date'].month != df.loc[i-1,'Date'].month):
            liste.append(df.loc[i,'Date'])
    liste = [date.strftime('%Y%m%d') for date in liste]
    score_dict = {}
    #vérification de l'existence des dossiers
    for i in liste : 
        i_format_tirets = f"{i[:4]}-{i[4:6]}-{i[6:]}"
        dossier_deb = working_folder
        dossier_initial = os.path.join(dossier_deb, f"members_{i}")
        file_name_h5 = f"df_score_pe_{i_format_tirets}.h5"
        file = os.path.join(dossier_initial, file_name_h5)
        if os.path.exists(file):
            pass
            print("Le fichier score existe à la date : ", i)
        else:
            print("Le fichier score à la date : ", i," n'existe pas.")
            return 0
        df_score, metadata1 = h5load(file)
        score_dict[i] = df_score
    print("L'ensemble des fichiers score vont pouvoir être download")
    dates =  []
    df_scores = []
    for date_str, df_score in score_dict.items():
        dates.append(date_str)
        df_scores.append(df_score)
    dates= dates[1:]
    df_scores = df_scores[:-1]
    print(dates)
    print(df_scores)
    #objectif c'est davoir rdt qui contient les rendements du mois précédent (dans nos dataframes initiaux, il contenait les rdt du mois suivant)
    rdt = {}
    for i, date in enumerate(dates):
        num_df = df_scores[i]
        df_rdt = pd.DataFrame()
        df_rdt['secteur'] = num_df['secteur']
        df_rdt['SCXR return'] = num_df['SCXR next month return']
        df_rdt['MCXR return'] = num_df['MCXR next month return']
        df_rdt['LCXR return'] = num_df['LCXR next month return']
        rdt[date] = df_rdt
    print(rdt)
    return dates, df_scores, rdt, score_dict


#fonction qui charge les données du GDP europe et du GDP world en daily et qui renvoie les deux dataframes correspondants
def retrieve_gdp():
    resp_eu = LocalTerminal.get_historical('EHGDEUY Index', 'PX_LAST',start="19001231", period='DAILY')
    data_eu = resp_eu.as_frame()
    resp_wrld = LocalTerminal.get_historical('GDPGAWLD Index', 'PX_LAST',start="19001231", period='DAILY')
    data_wrld = resp_wrld.as_frame()
    return data_eu, data_wrld

#fonction qui sert à charger les fichiers df_LCXR (qui contient l'ensemble des prix en daily) et df_SCXR disponible en .h5 sur data_index folder
#cela retourne les rendemments à chaque debut de mois du LCXR et SCXR
def retrieve_index(file,date):
    merged_df1 = pd.DataFrame()
    merged_df1, metadata3 = h5load(file)
    merged_df1.index = pd.to_datetime(merged_df1.index).strftime('%Y%m%d')
    merged_df1 = merged_df1.loc[merged_df1.index.isin(date)]
    merged_df1 = merged_df1.pct_change()
    return merged_df1


def lasso(date1, date2,mois, trinaire, decalage_interpolation_gdp=0, working_folder =working_folder):
    #on modifie la date de fin (date2) qd on souhaite un décalage de l'interpolation 
    #car sinon nous allons au delà de décembre 2023 et après décembre 2023 nous ne connaisons
    #pas le gdp 2024 pour faire l'interpolation
    date = datetime.strptime(str(date2), "%Y%m%d")
    new_date = date - timedelta(days=decalage_interpolation_gdp*30)
    new_date_2 = new_date.strftime("%Y%m%d")
    #prepare les datasets avec la fonction lasso_date
    dates, df_scores, rdt, score_dict = lasso_data(date1, new_date_2)
    #recupere les données du gdp pour procéder aux interpolations
    data_eu, data_wrld = retrieve_gdp()
    #on calcule l'ensemble des rendements  pour estimer la distribution des rendemments large-small pour trouver le quantile 50%
    perf_largevssmall = []
    for date, df in rdt.items():
        largevssmall = (1+(df.loc[df['secteur'] == 'Full Index', 'LCXR return'].values[0]/100) ) - (1+(df.loc[df['secteur'] == 'Full Index', 'SCXR return'].values[0]/100))
        perf_largevssmall.append(largevssmall)
    print(perf_largevssmall)

    #on définit un seuil à la moitié de la distribution au dessus on attribuera 1 (large qui perf par rapport aux smalls) en dessous 0 (small qui perf par rapport aux larges)
    perf_largevssmall.sort()
    n = len(perf_largevssmall)
    seuil = perf_largevssmall[int((n)/2)]

    print(f"Seuil : {seuil}")
    rep_gdp = []

    dataframe_tree = pd.DataFrame()
    index = 0
    for date, score_df in tqdm(score_dict.items()):

        #calcul de la diff gdp eur vs gdp world
        date_obj = datetime.strptime(date, '%Y%m%d')

        #décalage de decalage_interpolation_gdp mois
        #ca décale l'interpolation du gdp
        date_obj = date_obj + timedelta(days=decalage_interpolation_gdp*30)
        #on cherche la prochaine année à retrouver dans le dataframe
        month = date_obj.month
        year = date_obj.year
        previous_year= year-1
        previous_date = f"{previous_year}-12-31"
        next_date = f"{year}-12-31"
        #on mets en pourcentage le gdp et on trouve le point du prochain gdp en n+1
        previous_value_gdp_eu = data_eu.loc[previous_date].iloc[0]/100
        previous_value_gdp_wrld = data_wrld.loc[previous_date].iloc[0]/100
        next_value_gdp_eu = data_eu.loc[next_date].iloc[0]/100
        next_value_gdp_wrld = data_wrld.loc[next_date].iloc[0]/100
        #réalisation de l'interpolation
        interpolation_gdp_eu = previous_value_gdp_eu+(month-1)*((next_value_gdp_eu-previous_value_gdp_eu)/11)
        interpolation_gdp_wrld = previous_value_gdp_wrld+(month-1)*((next_value_gdp_wrld-previous_value_gdp_wrld)/11)
        #on exprime la différence de gdp entre europe et world
        diff_gdp = interpolation_gdp_eu-interpolation_gdp_wrld
        rep_gdp.append(diff_gdp) #gros dataframe avec toutes les diff de gdp

        dataframe_tree.loc[index,'Date'] = date
        rmse = 0
        diff_l= 0
        diff_s = 0
        #on parcours l'ensemble des valeurs de pe_u pour chaque secteur
        sectors = ['Utilities','Consumer Discretionary', 'Industrials', 'Energy', 'Information Technology', 'Materials', 'Real Estate', 'Consumer Staples', 'Communication Services', 'Health Care']
        for sector in sectors:
            value_peu_s = score_df.loc[score_df['secteur'] == sector, 'Median SCXR pe_u'].values[0]
            value_peu_l = score_df.loc[score_df['secteur'] == sector, 'Median LCXR pe_u'].values[0]
            value_pe_s = score_df.loc[score_df['secteur'] == sector, 'Median SCXR pe'].values[0]
            value_pe_l = score_df.loc[score_df['secteur'] == sector, 'Median LCXR pe'].values[0]
            if (sector != 'Full Index'):
                rmse += (value_peu_l - value_peu_s)**2 #on calcule le rmse comme étant la somme des écarts au carré des pe_u large et small pour chaque secteur
                #diff_l += (value_peu_l-value_pe_l)**2 #on calcule la diff_l comme étant la somme des écarts au carré des pe_u et pe des larges pour chaque secteur
                #diff_s += (value_peu_s-value_pe_s)**2 #on calcule la diff_l comme étant la somme des écarts au carré des pe_u et pe des smalls pour chaque secteur
        rmse = (rmse)**(0.5)
        #rmse_l = (diff_l)**(0.5)
        #rmse_s = (diff_s)**(0.5)


        #on rajoute en variable le pe_u médian des larges de l'ensemble des secteurs (donc full index)
        #on rajoute en variable le pe_u médian des smalls de l'ensemble des secteurs (donc full index)
        dataframe_tree.loc[index, 'Median SCXR pe_u Full Index'] = score_df.loc[score_df['secteur'] == 'Full Index', 'Median SCXR pe_u'].values[0]
        dataframe_tree.loc[index, 'Median LCXR pe_u Full Index'] = score_df.loc[score_df['secteur'] == 'Full Index', 'Median LCXR pe_u'].values[0]

        dataframe_tree.loc[index, 'rmse'] = rmse #rmse en variable
        #dataframe_tree.loc[index, 'diff LXCR'] = rmse_l #si on ne le mets pas en commentaire alors il sera considérer comme variable dans la régression logistique
        #dataframe_tree.loc[index, 'diff SXCR'] = rmse_s #si on ne le mets pas en commentaire alors il sera considérer comme variable dans la régression logistique
        dataframe_tree.loc[index, 'diff GDP'] = diff_gdp #diff du gdp en variable
        #rajouter le rendement passé (du dernier mois) en variable pr la reg logisitique
        if (index!=0):
            dff = rdt[dates[index-1]] 
            dataframe_tree.loc[index, 'rdt 0'] = (1+(dff.loc[dff['secteur'] == 'Full Index', 'LCXR return'].values[0]/100) ) - (1+(dff.loc[dff['secteur'] == 'Full Index', 'SCXR return'].values[0]/100))
        #recuperer le rendement le mois suivant
        if((index+mois)> len(dates)-1):
            break
        else:
            df = rdt[dates[index+mois]]
        #on stocke les returns du mois suivant pour ensuite pouvoir faire les calculs de perf plus facilement (ils seront enlevés lors de la régression logistiques)
        dff1 = rdt[dates[index]] 
        #rdt du mois suivant uniquement large -small reconstitué
        dataframe_tree.loc[index, 'month_return'] = (1+(dff1.loc[dff1['secteur'] == 'Full Index', 'LCXR return'].values[0]/100) ) - (1+(dff1.loc[dff1['secteur'] == 'Full Index', 'SCXR return'].values[0]/100))
        #rdt du mois suivant uniquement large reconstitué
        dataframe_tree.loc[index, 'next_month_return_large'] = dff1.loc[dff1['secteur'] == 'Full Index', 'LCXR return'].values[0]/100
        #rdt du mois suivant uniquement small reconstitué
        dataframe_tree.loc[index, 'next_month_return_small'] = dff1.loc[dff1['secteur'] == 'Full Index', 'SCXR return'].values[0]/100      
        #on transforme le rdt en score 1 ou 0 (en fonction de sa position dans la distribution historique ) et l'objectif est que le modèle prédise ses valeurs
        largevssmall_score = (1+(df.loc[df['secteur'] == 'Full Index', 'LCXR return'].values[0]/100) ) - (1+(df.loc[df['secteur'] == 'Full Index', 'SCXR return'].values[0]/100))
        dataframe_tree.loc[index,'lvs'] = largevssmall_score
        if largevssmall_score > seuil: #si supérieur au seuil on est à droite de la distribution donc bon signe => 1
            dataframe_tree.loc[index,'score large_vs_small'] = 1
        else:
            dataframe_tree.loc[index,'score large_vs_small'] = 0

        #cas pour rdt0 ajouté à notre dataframe (variables)
        if(index==0):
            dataframe_tree = dataframe_tree.iloc[1:, :]
        index +=1
    dataframe_tree = dataframe_tree.iloc[:-1]
    
    print(dataframe_tree)
    
    #rendre discrete la variable gdp, si ternaire vaut true je converti le gdp en variable ternaire
    #à l'aide des quantiles 0.33 et 0.66
    if trinaire : 
        seuil1 = np.quantile(rep_gdp, 0.33)
        seuil2 = np.quantile(rep_gdp, 0.66)
        for index, row in dataframe_tree.iterrows():
            if row['diff GDP'] < seuil1:
                dataframe_tree.loc[index, 'diff GDP'] = -1
            if row['diff GDP'] > seuil2:
                dataframe_tree.loc[index, 'diff GDP'] = 1
            else:
                dataframe_tree.loc[index, 'diff GDP'] = 0

    #change le dataset de test et de train : le train est la période la plus récente et on essaie de prédire les scores large vs small des années plus vieille (train)
    num_rows = dataframe_tree.shape[0]
    #split_index = int(num_rows * 0.35)
    #df_part1 = dataframe_tree.iloc[split_index+mois:]
    #df_part2 = dataframe_tree.iloc[:split_index+mois]

    #par défaut les années anciennes (2006-2017) prédisent les rendements des années récentes (2017-2023)
    split_index = int(num_rows * 0.65)
    df_part1 = dataframe_tree.iloc[:split_index+mois]
    df_part2 = dataframe_tree.iloc[split_index+mois:]

    df_pred_train = pd.DataFrame()
    df_pred_train = df_part1.copy()
    df_part1_equi= df_part1.drop(['lvs', 'Date','month_return','next_month_return_large','next_month_return_small'], axis=1)
    df_part1 = df_part1.drop(['lvs', 'Date','month_return','next_month_return_large','next_month_return_small'], axis=1)
    
    #si on souhaite diviser le dataset de train en supprimant une ligne sur 3 par exemple
    #df_part1 = df_part1.loc[df_part1.index %3 != 0, :]
    #df_part1= df_part1.sample(frac=0.8, random_state=31)
    
    #L'objectif du code ci dessous est d'avoir un dataset de train équipondéré
    #avoir le même nb de 0 et de 1 dans df_part1
    df_large_vs_small_0 = df_part1_equi[df_part1_equi['score large_vs_small'] == 0]
    df_large_vs_small_1 = df_part1_equi[df_part1_equi['score large_vs_small'] == 1]
    num_rows_0 = len(df_large_vs_small_0)
    num_rows_1 = len(df_large_vs_small_1)
    if num_rows_0 > num_rows_1:
        df_large_vs_small_0 = df_large_vs_small_0.iloc[:num_rows_1]
    else:
        df_large_vs_small_1 = df_large_vs_small_1.iloc[:num_rows_0]
    df_part1_equi = pd.concat([df_large_vs_small_0, df_large_vs_small_1], ignore_index=True)

    #différentes copie de dataset et on garde comme variables pour la régression logistique uniquement 
    #rmse : écart entre le pe_u large et pe_u small de chaque secteur au carré
    #Median SCXR pe_u Full Index : le pe_u médian sur l'ensemble des secteurs des smalls
    #Median LCXR pe_u Full Index : le pe_u médian sur l'ensemble des secteurs des larges
    #diff GDP : l'interpolation du gdp du mois correspondant (avec ou sans décalage en fonction des paramètres)
    #rdt0 : rendement du large-small reconstitué du mois passé

    df_pred_test = pd.DataFrame()
    df_pred_test = df_part2.copy() #sert de copie pour faire les calculs de perf
    df_part2= df_part2.drop(['lvs', 'Date','month_return','next_month_return_large','next_month_return_small'], axis=1)

    #dataset de train équipondéré (qui sert à fit le modèle)
    y = df_part1_equi['score large_vs_small']
    X = df_part1_equi.drop('score large_vs_small', axis=1)
    features_list1 = list(X.columns)
    X = np.array(X)

    #dataset de train non équipondéré (en entier mais seulement une partie a entrainer le modèle)
    y_full_train = df_part1['score large_vs_small']
    X_full_train = df_part1.drop('score large_vs_small', axis=1)
    features_list1_full = list(X_full_train.columns)
    X_full_train = np.array(X_full_train)

    #dataset de test 
    labels2 = df_part2['score large_vs_small']
    features2 = df_part2.drop('score large_vs_small', axis=1)
    features_list2 = list(features2.columns)
    features2 = np.array(features2)

    #régression logistique avec crossvalidation sur le dataset de train équipondéré
    #la cross validation permet de divierser l'ensemble de donnée en k plis de taille égale
    #ensuite le modèle est entrainé k fois en utilisant k-1 plis comme ensemble d'apprentissage et le pli restant comme test
    #après on a k estimations de la perf du modèle et on peut faire une estimation finale du modèle
    class_weights = {0: 1, 1: 1}
    reg = LogisticRegressionCV(cv=20, random_state=7, penalty='l2', solver='lbfgs', multi_class='multinomial', class_weight = class_weights,max_iter=1000)
    reg.fit(X, y)
    y_pred = reg.predict(features2)

    #voir le degré d'importance des variables avec shap
    explainer = shap.KernelExplainer(reg.predict_proba, X)
    shap_values = explainer.shap_values(features2)
    shap.summary_plot(shap_values, features2, feature_names=features_list1, plot_type="bar", show=False)
    fig = plt.gcf()
    fig.savefig('summary_shap_plot.png', bbox_inches='tight')
    plt.close()
    
    #précision de la prédiction avec nos données de test uniquement
    accuracy = accuracy_score(labels2, y_pred)
    print(y_pred)
    print(f"Taux de réussite sur l'échantillon de test uniquement : {accuracy:.2f}")


    #perf cumulative échantillon de test uniquement
    df_pred_test['pred'] = y_pred
    df_pred_test['pred'] = df_pred_test['pred'].replace(0, -1)
    df_result_test = pd.DataFrame(index=df_pred_test.index)
    df_result_test['result'] = df_pred_test['month_return'] * df_pred_test['pred']

    #objectif est d'avoir datum_test (une liste de date) qui permet de mesurer la performance de notre stratégie
    #nos prévisions s'effectuent en T pour la période T+1
    #nous rajoutons donc un mois sur les index après notre dernière prédiction pour calculer les performances
    datum_test = df_pred_test['Date'].to_list()
    date_obj1 = datetime.strptime(datum_test[-1], "%Y%m%d")
    date_next_month = date_obj1 + relativedelta(days=35)
    date1mois = date_next_month.strftime("%Y%m%d")
    next_date = find_prochaine_date_mois(datum_test[-1], date1mois)
    next_date = next_date.strftime("%Y%m%d")
    datum_test.append(next_date)

    #charger data LCXR et SCXR, qui correspond aux index réels
    lcxr_latest_path = find_latest_file(working_folder, 'LCXR')
    scxr_latest_path = find_latest_file(working_folder, 'SCXR')
    LCXR_index = retrieve_index(lcxr_latest_path, datum_test)
    SCXR_index = retrieve_index(scxr_latest_path, datum_test)
    LCXR_index = LCXR_index['LCXR Index'].squeeze()
    SCXR_index = SCXR_index['SCXR Index'].squeeze()
    #supprime la première ligne qui contient un Nan car on effectue le calcul du rdt avec .pct_change()
    LCXR_index = LCXR_index.iloc[1:]
    SCXR_index = SCXR_index.iloc[1:]
    #on mets les mêmes index que df_result_test pour simplifier les calculs 
    diff_real_large_small = (df_pred_test['pred']).reset_index(drop=True) *(LCXR_index - SCXR_index).reset_index(drop=True)
    diff_real_large_small.index = df_result_test.index

    #sharpe ratio for test only 
    sharpe_ratio(df_pred_test,df_result_test,diff_real_large_small, 'test',  working_folder )

    #information ratio for test only
    information_ratio(df_result_test,df_pred_test,diff_real_large_small,'test', working_folder )
    
    #cumulatives performances :
    #cumulative_performance_test_reconstituted_strat : la performance cumulative de notre stratégie appliqué sur le large-small reconstitué et small-large reconstitué sur notre dataframe de test uniquement
    #cumulative_performance_diff_reconstituted_large_small : performance cumulative de l'index reconstitué large - small (si on fait tjrs long large et short small pendant toute la durée de notre dataframe de test)
    #cumulative_performance_diff_reconstituted_small_large : performance cumulative de l'index reconstitué small - large (si on fait tjrs long small et short large pendant toute la durée de notre dataframe de test)
    #cumulative_performance_test_real_diff_large_small :  la performance cumulative de notre stratégie appliqué sur le large-small réel (LCXR et SCXR sur bloom) et small-large réel sur notre dataframe de test uniquement
    cumulative_performance_test = pd.DataFrame({
    'cumulative_performance_test_reconstituted_strat': (1 + df_result_test['result']).cumprod(),
    'cumulative_performance_diff_reconstituted_large_small': (1 + (df_pred_test['next_month_return_large'] - df_pred_test['next_month_return_small'])).cumprod(),
    'cumulative_performance_diff_reconstituted_small_large': (1 + (df_pred_test['next_month_return_small'] - df_pred_test['next_month_return_large'])).cumprod(),
    'cumulative_performance_test_real_diff_large_small' : (1+diff_real_large_small).cumprod()})
    #on change les index car les prévision sont en T pour la période T+1 donc on doit décaler d'un mois le debut et la fin des dates pour voir les performances de la stratégie
    cumulative_performance_test = cumulative_performance_test.reset_index(drop=True)
    cumulative_performance_test.index = datum_test[1:]

    #LONG RECONSTITUED LARGE (if pred =1) or LONG reconstitued small (if pred = -1)
    df_pred_test.loc[df_pred_test['pred'] == 1, 'long_large_or_small'] = df_pred_test['next_month_return_large']
    df_pred_test.loc[df_pred_test['pred'] == -1, 'long_large_or_small'] = df_pred_test['next_month_return_small']
    #LONG le rendement le plus élévé à chaque date entre large reconstitué et small reconstitué 
    df_pred_test['max_large_or_small'] = df_pred_test[['next_month_return_large', 'next_month_return_small']].max(axis=1)
    #50% long le large reconstitué et 50% long le small reconstitué pendant l'ensemble du dataframe
    df_pred_test['50_50_large_small'] = df_pred_test['next_month_return_large']*0.5 + 0.5*df_pred_test['next_month_return_small']

    #cumulative_performance_test_reconstituted_large : la performance cumulative de l'index du large reconstitué (long large only) sur le dataframe de test
    #cumulative_performance_test_reconstituted_small : perfomance cumulative de l'index du small reconstitué (long small only) sur le dataframe de test 
    #cumulative_performance_long_large_or_small_reconstituted : on applique la stratégie de cette facon : si la prévision c'est 1 : long large reconstitué sinon (prévision = 0) long small reconstitué
    #cumulative_performance_max_large_or_small_reconstituted : on long chaque mois le rendement le plus élévé entre le large reconstitué et le small reconstitué
    #cumulative_performance_50_50_large_small_reconstituted : on fait un 50% long le large reconstitué et 50% long le small reconstitué
    cumulative_performance_test_long = pd.DataFrame({
    'cumulative_performance_test_reconstituted_large': (1 + df_pred_test['next_month_return_large']).cumprod(),
    'cumulative_performance_test_reconstituted_small': (1 + df_pred_test['next_month_return_small']).cumprod(),
    'cumulative_performance_long_large_or_small_reconstituted': (1 +df_pred_test['long_large_or_small']).cumprod(),
    'cumulative_performance_max_large_or_small_reconstituted': (1 + df_pred_test['max_large_or_small']).cumprod(),
    'cumulative_performance_50_50_large_small_reconstituted': (1 + df_pred_test['50_50_large_small']).cumprod(),

    })
    #décale l'index de 1 car on prédit à T pour T+1
    cumulative_performance_test_long = cumulative_performance_test_long.reset_index(drop=True)
    cumulative_performance_test_long.index = datum_test[1:]
    
    #affichage du premier dataframe cumulative_performance_test
    plt.figure(figsize=(12, 6))
    cumulative_performance_test.plot()
    plt.title('Performance Cumulative Small VERSUS Large - Dataset Test')
    plt.xlabel('Date')
    plt.ylabel('Cumulative Return')
    plt.legend(['Stratégie sur Index reconstitué', 'Long Index reconstitué Large-Small', 'Long Index reconstitué Small-Large', 'Stratégie sur Index réel'])
    file_path = os.path.join(working_folder, 'Performance Cumulative Stratégie diff SVL -Test.png')
    plt.savefig(file_path)
    plt.close()
    #affichage du second dataframe cumulative_performance_test_long
    plt.figure(figsize=(12, 6))
    cumulative_performance_test_long.plot()
    plt.title('Performance Cumulative LONG Small VERSUS Large- Dataset Test')
    plt.xlabel('Date')
    plt.ylabel('Cumulative Return')
    plt.legend(['Index reconstitué Large', 'Index reconstitué Small', 'Long Large or Long Small with predictions', 'Long max between Large et Long return','50% large et 50% small reconstitué'])
    file_path = os.path.join(working_folder, 'Performance Cumulative LONG SVL - Test.png')
    plt.savefig(file_path)
    plt.close()


    #calcul precision échantillon train (donc équipondéré)
    y_pred_train_equi = reg.predict(X)
    accuracy_train_equi = accuracy_score(y, y_pred_train_equi)
    print(f"Taux de réussite sur l'échantillon de train équipondéré uniquement : {accuracy_train_equi:.2f}")

    #calcul precision échantillon train 
    #perf cumulative échantillon train (l'ensemble du dataset pas seulement le dataset equi) + test
    y_pred_train_full = reg.predict(X_full_train)
    accuracy_train_full = accuracy_score(y_full_train, y_pred_train_full)
    print(f"Taux de réussite sur l'échantillon de train full (modèle entrainé sur l'échantillon train équipondéré) uniquement : {accuracy_train_full:.2f}")

    
    #on reconstruit les prédictions sur l'ensemble du dataframe (train + test)
    df_pred_train['pred'] = y_pred_train_full
    df_pred_train['pred'] = df_pred_train['pred'].replace(0, -1)
    df_pred = pd.concat([df_pred_train, df_pred_test], ignore_index=True)
    df_result = pd.DataFrame(index=df_pred.index)
    #on calcule la performance de la stratégie sur l'ensemble du dataframe
    df_result['result'] = df_pred['month_return'] * df_pred['pred']

   #objectif est d'avoir datum_full (une liste de date) qui permet de mesurer la performance de notre stratégie
    #nos prévisions s'effectuent en T pour la période T+1
    #nous rajoutons donc un mois sur les index après notre dernière prédiction pour calculer les performances

    datum_full = df_pred['Date'].to_list()
    date_obj1 = datetime.strptime(datum_full[-1], "%Y%m%d")
    date_next_month = date_obj1 + relativedelta(days=35)
    date1mois = date_next_month.strftime("%Y%m%d")
    next_date = find_prochaine_date_mois(datum_full[-1], date1mois)
    next_date = next_date.strftime("%Y%m%d")
    datum_full.append(next_date)

   #charger data LCXR et SCXR, qui correspond aux index réels
    LCXR_index = retrieve_index(lcxr_latest_path, datum_full)
    SCXR_index = retrieve_index(scxr_latest_path, datum_full)
    LCXR_index = LCXR_index['LCXR Index'].squeeze()
    SCXR_index = SCXR_index['SCXR Index'].squeeze()
    LCXR_index = LCXR_index.iloc[1:]
    SCXR_index = SCXR_index.iloc[1:]
    #on mets les mêmes index que df_result pour simplifier les calculs 
    diff_real_large_small = (df_pred['pred']).reset_index(drop=True) *(LCXR_index - SCXR_index).reset_index(drop=True)
    diff_real_large_small.index = df_pred['pred'].index

    #sharpe ratio et information ratio  for test+ train only 
    sharpe_ratio(df_pred,df_result,diff_real_large_small, 'Test + Train',working_folder)
    information_ratio(df_result,df_pred,diff_real_large_small,'Test + Train',working_folder)

    #cumulatives performances :
    #cumulative_performance_full_reconstituted_strat : la performance cumulative de notre stratégie appliqué sur le large-small reconstitué et small-large reconstitué sur notre dataframe de train et test
    #cumulative_performance_full_reconstituted_diff_large_small : performance cumulative de l'index reconstitué large - small (si on fait tjrs long large et short small pendant toute la durée de notre dataframe )
    #cumulative_performance_full_diff_reconstituted_small_large : performance cumulative de l'index reconstitué small - large (si on fait tjrs long small et short large pendant toute la durée de notre dataframe)
    #cumulative_performance_full_real_diff_large_small :  la performance cumulative de notre stratégie appliqué sur le large-small réel (LCXR et SCXR sur bloom) et small-large réel sur notre dataframe 
    cumulative_performance_full = pd.DataFrame({
    'cumulative_performance_full_reconstituted_strat': (1 + df_result['result']).cumprod(),
    'cumulative_performance_full_reconstituted_diff_large_small': (1 + (df_pred['next_month_return_large'] - df_pred['next_month_return_small'])).cumprod(),
    'cumulative_performance_full_diff_reconstituted_small_large': (1 + (df_pred['next_month_return_small'] - df_pred['next_month_return_large'])).cumprod(),
    'cumulative_performance_full_real_diff_large_small' : (1+diff_real_large_small).cumprod()
    })

#on change les index car les prévision sont en T pour la période T+1 donc on doit décaler d'un mois le debut et la fin des dates pour voir les performances de la stratégie
    cumulative_performance_full = cumulative_performance_full.reset_index(drop=True)
    cumulative_performance_full.index = datum_full[1:]

#LONG RECONSTITUED LARGE (if pred =1) or LONG reconstitued small (if pred = -1)
    df_pred.loc[df_pred['pred'] == 1, 'long_large_or_small'] = df_pred['next_month_return_large']
    df_pred.loc[df_pred['pred'] == -1, 'long_large_or_small'] = df_pred['next_month_return_small']
    #LONG le rendement le plus élévé à chaque date entre large reconstitué et small reconstitué 
    df_pred['max_large_or_small'] = df_pred[['next_month_return_large', 'next_month_return_small']].max(axis=1)
    #50% long le large reconstitué et 50% long le small reconstitué pendant l'ensemble du dataframe
    df_pred['50_50_large_small'] = df_pred['next_month_return_large']*0.5 + 0.5*df_pred['next_month_return_small']

    #cumulative_performance_full_reconstituted_large : la performance cumulative de l'index du large reconstitué (long large only) sur le dataframe de test et train
    #cumulative_performance_full_reconstituted_small : perfomance cumulative de l'index du small reconstitué (long small only) sur le dataframe de test et train
    #cumulative_performance_long_large_or_small_reconstituted : on applique la stratégie de cette facon : si la prévision c'est 1 : long large reconstitué sinon (prévision = 0) long small reconstitué
    #cumulative_performance_max_large_or_small_reconstituted : on long chaque mois le rendement le plus élévé entre le large reconstitué et le small reconstitué
    #cumulative_performance_50_50_large_small_reconstituted : on fait un 50% long le large reconstitué et 50% long le small reconstitué
    cumulative_performance_full_long = pd.DataFrame({
    'cumulative_performance_full_reconstituted_large': (1 + df_pred['next_month_return_large']).cumprod(),
    'cumulative_performance_full_reconstituted_small': (1 + df_pred['next_month_return_small']).cumprod(),
    'cumulative_performance_long_large_or_small_reconstituted': (1 + df_pred['long_large_or_small']).cumprod(),
    'cumulative_performance_max_large_or_small_reconstituted': (1 + df_pred['max_large_or_small']).cumprod(),
    'cumulative_performance_50_50_large_small_reconstituted': (1 + df_pred['50_50_large_small']).cumprod(),

    })   

    #décale l'index de 1 car on prédit à T pour T+1
    cumulative_performance_full_long = cumulative_performance_full_long.reset_index(drop=True)
    cumulative_performance_full_long.index = datum_full[1:]

    #plot le dataframe cumulative_performance_full
    plt.figure(figsize=(12, 6))
    cumulative_performance_full.plot()
    plt.title('Performance Cumulative Small VERSUS Large- Dataset Train et Test')
    plt.xlabel('Date')
    plt.ylabel('Cumulative Return')
    plt.legend(['Stratégie sur Index reconstitué', 'Long Index reconstitué Large-Small', 'Long Index reconstitué Small-Large', 'Stratégie sur Index réel'])
    file_path = os.path.join(working_folder,'Performance Cumulative Stratégie diff SVL - Train et Test.png')
    plt.savefig(file_path)
    plt.close()
    #plot le dataframe cumulative_performance_full_long
    plt.figure(figsize=(12, 6))
    cumulative_performance_full_long.plot()
    plt.title('Performance Cumulative LONG Small VERSUS Large- Dataset Train et Test')
    plt.xlabel('Date')
    plt.ylabel('Cumulative Return')
    plt.legend(['Index reconstitué Large', 'Index reconstitué Small', 'Long Large or Long Small with predictions', 'Long max between Large et Long return','50% large et 50% small reconstitué'])
    file_path = os.path.join(working_folder,'Performance Cumulative LONG SVL - Train et Test.png')
    plt.savefig(file_path)
    plt.close()    
    return reg, seuil


##fonction afin de réaliser des simulaires sur des données sans connaisance du réel gdp
def simulation_utilisation_modele(date1, date2, mois,reg,seuil1, trinaire,decalage_interpolation_gdp=0, rep_gdp=0, data_eu=0, data_wrld=0, working_folder = working_folder):

    #prepare les datasets avec la fonction lasso_date
    dates, df_scores, rdt, score_dict = lasso_data(date1, date2)
    dataframe_tree = pd.DataFrame()
    index = 0
    if trinaire == False :
        rep_gdp = []
    for date, score_df in score_dict.items():
        if trinaire == True:
            if len(rep_gdp)!=len(dates)+1:
                raise PrevisionError(f"Erreur: il faut spécifier {len(dates)+1} prévisions dans rep.")
        if trinaire == False : 
            date_obj = datetime.strptime(date, '%Y%m%d')
            #décalage de decalage_interpolation_gdp mois
            #ca décale l'interpolation du gdp
            date_obj = date_obj + timedelta(days=decalage_interpolation_gdp*30)
            #on cherche la prochaine année à retrouver dans le dataframe
            month = date_obj.month
            year = date_obj.year
            previous_year= year-1
            previous_date = f"{previous_year}-12-31"
            next_date = f"{year}-12-31"
            previous_value_gdp_eu = data_eu.loc[previous_date].iloc[0]/100
            previous_value_gdp_wrld = data_wrld.loc[previous_date].iloc[0]/100
            next_value_gdp_eu = data_eu.loc[next_date].iloc[0]/100
            next_value_gdp_wrld = data_wrld.loc[next_date].iloc[0]/100
            interpolation_gdp_eu = previous_value_gdp_eu+(month-1)*((next_value_gdp_eu-previous_value_gdp_eu)/11)
            interpolation_gdp_wrld = previous_value_gdp_wrld+(month-1)*((next_value_gdp_wrld-previous_value_gdp_wrld)/11)
            #on exprime la différence de gdp entre europe et world
            diff_gdp = interpolation_gdp_eu-interpolation_gdp_wrld
            rep_gdp.append(diff_gdp) #gros dataframe avec toutes les diff de gdp
        dataframe_tree.loc[index,'Date'] = date
        rmse = 0
        diff_l= 0
        diff_s = 0
        #on parcours l'ensemble des valeurs de pe_u pour chaque secteur
        sectors = ['Utilities','Consumer Discretionary', 'Industrials', 'Energy', 'Information Technology', 'Materials', 'Real Estate', 'Consumer Staples', 'Communication Services', 'Health Care']
        for sector in sectors:
            value_peu_s = score_df.loc[score_df['secteur'] == sector, 'Median SCXR pe_u'].values[0]
            value_peu_l = score_df.loc[score_df['secteur'] == sector, 'Median LCXR pe_u'].values[0]
            value_pe_s = score_df.loc[score_df['secteur'] == sector, 'Median SCXR pe'].values[0]
            value_pe_l = score_df.loc[score_df['secteur'] == sector, 'Median LCXR pe'].values[0]
            if (sector != 'Full Index'):
                rmse += (value_peu_l - value_peu_s)**2 #on calcule le rmse comme étant la somme des écarts au carré des pe_u large et small pour chaque secteur
                #diff_l += (value_peu_l-value_pe_l)**2 #on calcule la diff_l comme étant la somme des écarts au carré des pe_u et pe des larges pour chaque secteur
                #diff_s += (value_peu_s-value_pe_s)**2 #on calcule la diff_l comme étant la somme des écarts au carré des pe_u et pe des smalls pour chaque secteur
        rmse = (rmse)**(0.5)
        #rmse_l = (diff_l)**(0.5)
        #rmse_s = (diff_s)**(0.5)


        #on rajoute en variable le pe_u médian des larges de l'ensemble des secteurs (donc full index)
        #on rajoute en variable le pe_u médian des smalls de l'ensemble des secteurs (donc full index)
        dataframe_tree.loc[index, 'Median SCXR pe_u Full Index'] = score_df.loc[score_df['secteur'] == 'Full Index', 'Median SCXR pe_u'].values[0]
        dataframe_tree.loc[index, 'Median LCXR pe_u Full Index'] = score_df.loc[score_df['secteur'] == 'Full Index', 'Median LCXR pe_u'].values[0]

        dataframe_tree.loc[index, 'rmse'] = rmse #rmse en variable
        #dataframe_tree.loc[index, 'diff LXCR'] = rmse_l #si on ne le mets pas en commentaire alors il sera considérer comme variable dans la régression logistique
        #dataframe_tree.loc[index, 'diff SXCR'] = rmse_s #si on ne le mets pas en commentaire alors il sera considérer comme variable dans la régression logistique
        
        dataframe_tree.loc[index, 'diff GDP'] = rep_gdp[index] #diff du gdp en variable
        
        #rajouter le rendement passé (du dernier mois) en variable pr la reg logisitique
        if (index!=0):
            dff = rdt[dates[index-1]] 
            dataframe_tree.loc[index, 'rdt 0'] = (1+(dff.loc[dff['secteur'] == 'Full Index', 'LCXR return'].values[0]/100) ) - (1+(dff.loc[dff['secteur'] == 'Full Index', 'SCXR return'].values[0]/100))
        #recuperer le rendement le mois suivant
        
        if((index+mois)<= len(dates)-1):
            shape = rdt[dates[index+mois]].shape
            df = rdt[dates[index+mois]]
        #on stocke les returns du mois suivant pour ensuite pouvoir faire les calculs de perf plus facilement (ils seront enlevés lors de la régression logistiques)
            dff1 = rdt[dates[index]] 
            #rdt du mois suivant uniquement large -small reconstitué
            dataframe_tree.loc[index, 'month_return'] = (1+(dff1.loc[dff1['secteur'] == 'Full Index', 'LCXR return'].values[0]/100) ) - (1+(dff1.loc[dff1['secteur'] == 'Full Index', 'SCXR return'].values[0]/100))
            #rdt du mois suivant uniquement large reconstitué
            dataframe_tree.loc[index, 'next_month_return_large'] = dff1.loc[dff1['secteur'] == 'Full Index', 'LCXR return'].values[0]/100
            #rdt du mois suivant uniquement small reconstitué
            dataframe_tree.loc[index, 'next_month_return_small'] = dff1.loc[dff1['secteur'] == 'Full Index', 'SCXR return'].values[0]/100      
            #on transforme le rdt en score 1 ou 0 (en fonction de sa position dans la distribution historique ) et l'objectif est que le modèle prédise ses valeurs
            largevssmall_score = (1+(df.loc[df['secteur'] == 'Full Index', 'LCXR return'].values[0]/100) ) - (1+(df.loc[df['secteur'] == 'Full Index', 'SCXR return'].values[0]/100))
            dataframe_tree.loc[index,'lvs'] = largevssmall_score
            if largevssmall_score > seuil1: #si supérieur au seuil on est à droite de la distribution donc bon signe => 1
                dataframe_tree.loc[index,'score large_vs_small'] = 1
            else:
                dataframe_tree.loc[index,'score large_vs_small'] = 0
        else : 
            df = np.nan
            dff1 = pd.DataFrame(np.zeros(shape))
        #cas pour rdt0 ajouté à notre dataframe (variables)
        if(index==0):
            dataframe_tree = dataframe_tree.iloc[1:, :]
        index +=1
 
    print(dataframe_tree)
    if trinaire : 
        name_folder = "Test-sans-connaissance-GDP_Trinaire_"+ str(date1)+ "_" + str(date2)
    else : 
        name_folder = "Test-sans-connaissance-GDP_Interpolation_"+ str(date1)+ "_" + str(date2)

    directory_path = working_folder
    full_directory_path = os.path.join(directory_path, name_folder)
    create_or_replace_directory(full_directory_path)
    #change le dataset de test et de train : le train est la période la plus récente et on essaie de prédire les scores large vs small des années plus vieille (train)
    num_rows = dataframe_tree.shape[0]
    #split_index = int(num_rows * 0.35)
    #df_part1 = dataframe_tree.iloc[split_index+mois:]
    df_part2 = dataframe_tree.copy()

    #différentes copie de dataset et on garde comme variables pour la régression logistique uniquement 
    #rmse : écart entre le pe_u large et pe_u small de chaque secteur au carré
    #Median SCXR pe_u Full Index : le pe_u médian sur l'ensemble des secteurs des smalls
    #Median LCXR pe_u Full Index : le pe_u médian sur l'ensemble des secteurs des larges
    #diff GDP : l'interpolation du gdp du mois correspondant (avec ou sans décalage en fonction des paramètres)
    #rdt0 : rendement du large-small reconstitué du mois passé    
    df_pred_test = pd.DataFrame() 
    df_pred_test = dataframe_tree.copy() #sert de copie pour faire les calculs de perf
    df_part2= df_part2.drop(['lvs', 'Date','month_return','next_month_return_large','next_month_return_small'], axis=1) 
    #dataset de test 
    labels2 = df_part2['score large_vs_small']
    features2 = df_part2.drop('score large_vs_small', axis=1)
    features_list2 = list(features2.columns)
    features2 = np.array(features2)
                                       
    #régression logistique avec crossvalidation sur le dataset de train équipondéré
    #la cross validation permet de divierser l'ensemble de donnée en k plis de taille égale
    #ensuite le modèle est entrainé k fois en utilisant k-1 plis comme ensemble d'apprentissage et le pli restant comme test
    #après on a k estimations de la perf du modèle et on peut faire une estimation finale du modèle
    y_pred = reg.predict(features2)
    #précision de la prédiction avec nos données de test uniquement
    if np.isnan(labels2[len(labels2)]) :
        labels2 = labels2.iloc[:-1]
        last_prediction = y_pred[-1]
        y_pred = y_pred[:-1]
        print("La dernière valeur prédite pour la période ", dates[-1], " est ", last_prediction)
        print("Rappel 1 = Large superforme Small et 0 = Small superforme Large")
        df_pred_test = df_pred_test.iloc[:-1]
    accuracy = accuracy_score(labels2, y_pred)
    print(y_pred)
    print(f"Taux de réussite sur l'échantillon de test uniquement : {accuracy:.2f}")



    #perf cumulative échantillon de test uniquement

    df_pred_test['pred'] = y_pred
    df_pred_test['pred'] = df_pred_test['pred'].replace(0, -1)
    df_result_test = pd.DataFrame(index=df_pred_test.index)
    df_result_test['result'] = df_pred_test['month_return'] * df_pred_test['pred']

    #objectif est d'avoir datum_test (une liste de date) qui permet de mesurer la performance de notre stratégie
    #nos prévisions s'effectuent en T pour la période T+1
    #nous rajoutons donc un mois sur les index après notre dernière prédiction pour calculer les performances
    datum_test = df_pred_test['Date'].to_list()
    date_obj1 = datetime.strptime(datum_test[-1], "%Y%m%d")
    date_next_month = date_obj1 + relativedelta(days=35)
    date1mois = date_next_month.strftime("%Y%m%d")
    next_date = find_prochaine_date_mois(datum_test[-1], date1mois)
    next_date = next_date.strftime("%Y%m%d")
    datum_test.append(next_date)

    #charger data LCXR et SCXR, qui correspond aux index réels
    lcxr_latest_path = find_latest_file(working_folder, 'LCXR')
    scxr_latest_path = find_latest_file(working_folder, 'SCXR')
    LCXR_index = retrieve_index(lcxr_latest_path, datum_test)
    SCXR_index = retrieve_index(scxr_latest_path, datum_test)
    LCXR_index = LCXR_index['LCXR Index'].squeeze()
    SCXR_index = SCXR_index['SCXR Index'].squeeze()
    #supprime la première ligne qui contient un Nan car on effectue le calcul du rdt avec .pct_change()
    LCXR_index = LCXR_index.iloc[1:]
    SCXR_index = SCXR_index.iloc[1:]
    #on mets les mêmes index que df_result_test pour simplifier les calculs 
    diff_real_large_small = (df_pred_test['pred']).reset_index(drop=True) *(LCXR_index - SCXR_index).reset_index(drop=True)
    diff_real_large_small.index = df_result_test.index

    #sharpe ratio for test only 
    sharpe_ratio(df_pred_test,df_result_test,diff_real_large_small, 'test', full_directory_path)

    #information ratio for test only
    information_ratio(df_result_test,df_pred_test,diff_real_large_small,'test',full_directory_path )
    
    #cumulatives performances :
    #cumulative_performance_test_reconstituted_strat : la performance cumulative de notre stratégie appliqué sur le large-small reconstitué et small-large reconstitué sur notre dataframe de test uniquement
    #cumulative_performance_diff_reconstituted_large_small : performance cumulative de l'index reconstitué large - small (si on fait tjrs long large et short small pendant toute la durée de notre dataframe de test)
    #cumulative_performance_diff_reconstituted_small_large : performance cumulative de l'index reconstitué small - large (si on fait tjrs long small et short large pendant toute la durée de notre dataframe de test)
    #cumulative_performance_test_real_diff_large_small :  la performance cumulative de notre stratégie appliqué sur le large-small réel (LCXR et SCXR sur bloom) et small-large réel sur notre dataframe de test uniquement
    cumulative_performance_test = pd.DataFrame({
    'cumulative_performance_test_reconstituted_strat': (1 + df_result_test['result']).cumprod(),
    'cumulative_performance_diff_reconstituted_large_small': (1 + (df_pred_test['next_month_return_large'] - df_pred_test['next_month_return_small'])).cumprod(),
    'cumulative_performance_diff_reconstituted_small_large': (1 + (df_pred_test['next_month_return_small'] - df_pred_test['next_month_return_large'])).cumprod(),
    'cumulative_performance_test_real_diff_large_small' : (1+diff_real_large_small).cumprod()})
    #on change les index car les prévision sont en T pour la période T+1 donc on doit décaler d'un mois le debut et la fin des dates pour voir les performances de la stratégie
    cumulative_performance_test = cumulative_performance_test.reset_index(drop=True)
    cumulative_performance_test.index = datum_test[1:]

    #LONG RECONSTITUED LARGE (if pred =1) or LONG reconstitued small (if pred = -1)
    df_pred_test.loc[df_pred_test['pred'] == 1, 'long_large_or_small'] = df_pred_test['next_month_return_large']
    df_pred_test.loc[df_pred_test['pred'] == -1, 'long_large_or_small'] = df_pred_test['next_month_return_small']
    #LONG le rendement le plus élévé à chaque date entre large reconstitué et small reconstitué 
    df_pred_test['max_large_or_small'] = df_pred_test[['next_month_return_large', 'next_month_return_small']].max(axis=1)
    #50% long le large reconstitué et 50% long le small reconstitué pendant l'ensemble du dataframe
    df_pred_test['50_50_large_small'] = df_pred_test['next_month_return_large']*0.5 + 0.5*df_pred_test['next_month_return_small']

    #cumulative_performance_test_reconstituted_large : la performance cumulative de l'index du large reconstitué (long large only) sur le dataframe de test
    #cumulative_performance_test_reconstituted_small : perfomance cumulative de l'index du small reconstitué (long small only) sur le dataframe de test 
    #cumulative_performance_long_large_or_small_reconstituted : on applique la stratégie de cette facon : si la prévision c'est 1 : long large reconstitué sinon (prévision = 0) long small reconstitué
    #cumulative_performance_max_large_or_small_reconstituted : on long chaque mois le rendement le plus élévé entre le large reconstitué et le small reconstitué
    #cumulative_performance_50_50_large_small_reconstituted : on fait un 50% long le large reconstitué et 50% long le small reconstitué
    cumulative_performance_test_long = pd.DataFrame({
    'cumulative_performance_test_reconstituted_large': (1 + df_pred_test['next_month_return_large']).cumprod(),
    'cumulative_performance_test_reconstituted_small': (1 + df_pred_test['next_month_return_small']).cumprod(),
    'cumulative_performance_long_large_or_small_reconstituted': (1 +df_pred_test['long_large_or_small']).cumprod(),
    'cumulative_performance_max_large_or_small_reconstituted': (1 + df_pred_test['max_large_or_small']).cumprod(),
    'cumulative_performance_50_50_large_small_reconstituted': (1 + df_pred_test['50_50_large_small']).cumprod(),

    })
    #décale l'index de 1 car on prédit à T pour T+1
    cumulative_performance_test_long = cumulative_performance_test_long.reset_index(drop=True)
    cumulative_performance_test_long.index = datum_test[1:]
    
    #affichage du premier dataframe cumulative_performance_test
    plt.figure(figsize=(12, 6))
    cumulative_performance_test.plot()
    plt.title('Performance Cumulative Small VERSUS Large')
    plt.xlabel('Date')
    plt.ylabel('Cumulative Return')
    plt.legend(['Stratégie sur Index reconstitué', 'Long Index reconstitué Large-Small', 'Long Index reconstitué Small-Large', 'Stratégie sur Index réel'])
    
    full_directory_path_cum = os.path.join(full_directory_path,"Performance Cumulative Stratégie diff SVL -TEST sur periode sans connaisance GDP.png")
    plt.savefig(full_directory_path_cum)
    plt.close()
    #affichage du second dataframe cumulative_performance_test_long
    plt.figure(figsize=(12, 6))
    cumulative_performance_test_long.plot()
    plt.title('Performance Cumulative LONG Small VERSUS Large')
    plt.xlabel('Date')
    plt.ylabel('Cumulative Return')
    plt.legend(['Index reconstitué Large', 'Index reconstitué Small', 'Long Large or Long Small with predictions', 'Long max between Large et Long return','50% large et 50% small reconstitué'])
    full_directory_path_cum2 = os.path.join(full_directory_path,"Performance Cumulative LONG SVL - TEST sur periode sans connaisance GDP.png")
    plt.savefig(full_directory_path_cum2)
    plt.close()




#lasso(date1, date2,mois, trinaire, decalage_interpolation_gdp)
#date1 correspond à la date de début pour la réalisation du dataframe 
#date2 correspond à la date de fin pour la réalisation du dataframe
#mois : on réalise le calcul des métriques(rmse (sur les pe_u),le dernier rdt des 20 derniers jours, une interpolation du gdp...) à la date t et nous souhaitons prédire les rendements à la date t+mois
# /!\ par défaut mois = 0 ca veut dire on souhaite prédire les rendement en t+1mois
#trinaire : si trinaire = false alors on utilise l'interpolation du gdp (connaisance du futur) si trinaire = true alors on transforme l'information de la différence de gdp en 1, 0 ou -1 en fonction des quantiles 0.33 et 0.66
#decalage_interpolation_gdp : est un nombre de mois par defaut vaut 0 ca correspond au décalage de l'interpolation à effectuer sur le gdp
# par exemple si on met decalage_interpolation_gdp = 2 alors la valeur du gdp au mois de mai 2023 correspondra à l'interpolation entre le gdp 2022 et gdp 2023 mais du mois de juillet



#A modifier 

#en trinaire
reg, seuil= lasso(20060401, 20231231, 0, True, 0)
rep = [1,1,1,1,1,1,1,1,-1,1,1,0,1,1,1]
simulation_utilisation_modele(20230505,20240810,0,reg,seuil,True,0,rep, 0,0)

#def simulation_utilisation_modele(date1, date2, mois,reg,seuil, trinaire,decalage_interpolation_gdp=0, rep_gdp=0, data_eu=0, data_wrld=0):
#def lasso(date1, date2,mois, trinaire, decalage_interpolation_gdp=0):

#en non trinaire

reg, seuil= lasso(20060401, 20231231, 0, False, 0)

data_eu, data_wrld = retrieve_gdp()
print(data_eu)
print(data_wrld)
data_eu.loc[pd.to_datetime('2024-12-31')] = float(0.7)
data_wrld.loc[pd.to_datetime('2024-12-31')] = float(3.10)
simulation_utilisation_modele(20240101,20240810,0,reg,seuil,False,0,0, data_eu, data_wrld)

download data scxr et lcxr 
from tia.bbg import LocalTerminal
import pandas as pd
import datetime
from xbbg import blp
from blp import blp as bld
import pybbg as pybbg
import matplotlib.pyplot as plt
import os



def h5store(filename, df, dic):
    store = pd.HDFStore(filename)
    store.put('mydata', df)
    store.get_storer('mydata').attrs.metadata = dic
    store.close()

working_folder = r'U:\GDA\PFC\03_Gerants\03_12_NB\Small VS Large'

def retrieve_index(index1, index2,date_1, date_2, dossier_initial =working_folder):
    data_index_folder = os.path.join(working_folder, 'data_index')
    if not os.path.exists(data_index_folder):
        os.makedirs(data_index_folder)
    resp_index1 = LocalTerminal.get_historical(index1, 'PX_LAST',start=date_1, end= date_2)
    merged_df1 = resp_index1.as_frame()
            
    resp_index2 = LocalTerminal.get_historical(index2, 'PX_LAST',start=date_1, end= date_2)
    merged_df2 = resp_index2.as_frame()

    #SAVE first index
    file_name_h5 = f"df_{index1}_{date_1}_{date_2}.h5"
    df_index1_path_h5 = os.path.join(data_index_folder, file_name_h5)
    metadata = dict(data="data monthly return",provider="Bloomberg",indice=index1) 
    h5store(df_index1_path_h5 ,merged_df1,metadata)
    file_name_index1_xlsx = f"df_{index1}_{date_1}_{date_2}.xlsx"
    df_index1_path = os.path.join(data_index_folder, file_name_index1_xlsx)
    merged_df1.to_excel(df_index1_path)

    #SAVE the second index
    file_name_h5 = f"df_{index2}_{date_1}_{date_2}.h5"
    df_index2_path_h5 = os.path.join(data_index_folder, file_name_h5)
    metadata = dict(data="data monthly return",provider="Bloomberg",indice=index2) 
    h5store(df_index2_path_h5 ,merged_df2,metadata)
    file_name_index2_xlsx = f"df_{index2}_{date_1}_{date_2}.xlsx"
    df_index2_path = os.path.join(data_index_folder, file_name_index2_xlsx)
    merged_df2.to_excel(df_index2_path)

    return merged_df1, merged_df2


def h5load(filename):
    with pd.HDFStore(filename) as store:
        data = store['mydata']
        metadata = store.get_storer('mydata').attrs.metadata
        data.attrs = metadata
    return data, metadata     

retrieve_index('LCXR Index', 'SCXR Index', '20051201', '20240720')

pypfopt modele arbitrage
import pypfopt
import pypfopt.objective_functions

from gestion_hdf5 import h5load
from tia.bbg import LocalTerminal
import numpy as np
import pandas as pd
from tia.bbg import LocalTerminal
from collections import OrderedDict
import pickle
from sklearn.linear_model import LinearRegression
import matplotlib.pyplot as plt
from sklearn.metrics import r2_score 
from datetime import datetime 
from dateutil.relativedelta import relativedelta
from tia.bbg import LocalTerminal
import datetime
from pandas import Timestamp
import os
import luigi
from global_parameter import ClassGlobal, SelectData, StrategyPypfopt
from functions_assets import get_assets
from functions_data import _build_easy_map, end_month_data


class Strategy_Pypfopt(luigi.Task):
    #zone = luigi.Parameter()
    #period = luigi.Parameter()
    #Doc_name = luigi.Parameter()

    def run(self):
        known_future_cov_matrix= StrategyPypfopt().known_future_cov_matrix == "True"
        known_future_expected_returns= StrategyPypfopt().known_future_expected_returns == "True"
        rolling = int(StrategyPypfopt().rolling)
        start_date = '2001-05-31'   #start_date minimum possible pour l'EU : 2001-01-31
        end_date = '2024-06-30'

        #pour la chine
        #start_date = '2019-07-31'
        

        #mettre une date de fin de mois
        #(pour know_future_cov_known) : end_date doit être il y a plus de deux mois par rapport à la date du jour car la cov du mois suivant doit être connue

        #chargement de la data

        #Expected returns estimés (EDR + YTM historique moins la prime de risque) 
        path_exp_returns = os.path.join(ClassGlobal().path_output,"Simulation\{}\Assets\Assets_{}".format(ClassGlobal().simulation,datetime.datetime.now().strftime("%Y_%m_%d")))
        file_exp_returns = "Yields_"+ClassGlobal().Doc_name+".h5" #Expected YTM (donc vrai YTM - un certain pourcentage %)
        expected_returns, metadata1 = h5load(path_exp_returns + "\\" + file_exp_returns)

        #transforme YTM en vrai YTM enregistrés par bloom avec la colonne equity qui correspond à l'EDR (on restructuera l'ensemble du dataframe par la suite)
        YTM, metadata = h5load(path_exp_returns + "\\" + file_exp_returns)  
        YTM['HG Corpo '+SelectData().zone] =  YTM['HG Corpo '+SelectData().zone] +0.005
        YTM['HY Corpo '+SelectData().zone] =  YTM['HY Corpo '+SelectData().zone] +0.025

        #Prix des index enregistrés par bloom 
        path_prices = os.path.join(ClassGlobal().path_output,"Simulation\{}\Assets\Assets_{}".format(ClassGlobal().simulation,datetime.datetime.now().strftime("%Y_%m_%d")))
        file_prices = "Prices_"+ClassGlobal().Doc_name+".h5" 
        prices, metabase2 = h5load(path_prices + "\\" + file_prices)

        path_vol = os.path.join(ClassGlobal().path_output,"Simulation\{}\Assets\Assets_{}".format(ClassGlobal().simulation,datetime.datetime.now().strftime("%Y_%m_%d")))
        file_vol = "Vol_"+ClassGlobal().Doc_name+".h5"
        vol, metabase3 = h5load(path_vol+ "\\" + file_vol)

        #Affichage de la data
        print("Expected Returns:") 
        print(expected_returns)
        print("YTM:")
        print(YTM)
        print("Prices:")
        print(prices)
        print("Vol :")
        print(vol)

        #nom pour les fichiers finaux
        nom = ''
        if(known_future_cov_matrix):
            nom =  nom +'_known_cov_+1m' #je connais déjà la matrice de variance covariance du mois suivant
        if(known_future_cov_matrix==False):
            nom = nom+'_cov_estimated' #j'estime la matrice de variance covariance du mois suivant par la matrice variance covariance historique
        if(known_future_expected_returns):
            nom = nom + '_known_expReturns_+1m' #je connais déjà les rendemments du mois suivant
        if(known_future_expected_returns==False):
            nom = nom+'_expReturns_estimated' #j'estime les rendemments du mois suivant (EDR + data historique moins prime de risques)
        if(rolling and known_future_cov_matrix==False):
            nom = nom+'_roll_'+str(rolling)+'_M'
        if(rolling== 1 and known_future_cov_matrix==False):
            nom = nom+'_cov_1monthInDaily' #on estime la covariance en prenant des données daily sur le mois et la cov dans le modele correspond à la cov du mois actuel seulement
        
        #calcul des rendemments des index en mensuel
        monthly_returns = prices.pct_change() 
        print("Monthly Returns : ")
        print(monthly_returns)
        covariance_dict = {} #dictionnaire de covariance qui contiendra la covariance mensuelle estimée
        #restructure le fichier YTM en gardant uniquement les index : on fait une annualisation du return mensuel constasté pour l'ensemble des classes d'actifs
        #UTILE UNIQUEMENT lorsque je connais en avance les vrais returns du m+1 (car sinon j'utilise le dataframe expected_returns)
        for date in YTM.index:
            YTM.loc[date,'Treasuries '+ SelectData().zone] = ((1+ monthly_returns.loc[date,'Treasuries '+SelectData().zone])**(12))-1
            YTM.loc[date, 'Equity '+SelectData().zone] =((1+ monthly_returns.loc[date,'Equity '+SelectData().zone])**(12))-1
            YTM.loc[date, 'HG Corpo '+SelectData().zone] =((1+ monthly_returns.loc[date,'HG Corpo '+SelectData().zone])**(12))-1
            YTM.loc[date, 'HY Corpo '+SelectData().zone] =((1+ monthly_returns.loc[date,'HY Corpo '+SelectData().zone])**(12))-1

        #chargement des données pour les prices en daily

        assets = OrderedDict()
        assets = get_assets(SelectData().zone)
        assets.popitem()  #nous allons recalculer nous-même cette colonne car erreurs pour le cash
        fields = ['PX_LAST']
        period = 'DAILY'
        resp = LocalTerminal.get_historical([assets[i]["ticker"] for i in assets.keys()],
            fields, start=start_date, end=end_date, period=period)
        daily_price = resp.as_frame()
        daily_price.columns = daily_price.columns.droplevel(1)
        daily_price.ffill(inplace=True)
        daily_price
        daily_price.rename(columns=_build_easy_map(assets), inplace=True)
        daily_return = daily_price.pct_change()
        daily_return = daily_return[start_date:]
        index = daily_return.index

        #on enregistre la covariance du mois dans le dictionnaire covariance_dict avec des données daily, la clé correspond au mois où la covariance est observée
        for date in monthly_returns.index:
            copy_daily_return = daily_return.loc[(index.month == date.month) & (index.year == date.year)]
            copy_daily_return = copy_daily_return.dropna()
            cov_1_month = copy_daily_return.cov(min_periods=20)
            covariance_dict[date] = cov_1_month*252

            
        #chargement des données pr le cash EU, utile aussi pour le paramètre du taux sans risque dans le .max_sharpe
        if SelectData().zone == 'EU':
            indexes = ['EONIA Index', 'ESTRON Index']
        if SelectData().zone == 'US':
            indexes = ['US0001M Index', 'SOFRRATE Index']
        if SelectData().zone == 'CN':
            indexes = ['CHBM7D Index']
        resp = LocalTerminal.get_historical(indexes, 'PX_LAST',start=start_date, end=end_date, period='DAILY')
        data = resp.as_frame()
        if SelectData().zone=='CN':
            data = data.rename(columns={'CHBM7D Index': 'ST index'})
        data.columns = data.columns.droplevel(1)
        str_series = pd.Series(index=pd.to_datetime(data.index), name='ST index', dtype='float64')
        if SelectData().zone !='CN':
            ind = np.isnan(data[indexes[1]])
            str_series.loc[~ind] = data.loc[~ind, indexes[1]]
            str_series.loc[ind] = data.loc[ind, indexes[0]]
        else : 
            str_series = data['ST index']

        print(str_series)
        dataframe_cash = str_series.to_frame()
        print(dataframe_cash["ST index"]/100) # on annualise les perf journalières
        dataframe_cash["perf_annualise"] = ((1 + dataframe_cash["ST index"] / 100) ** (1 / 365)) - 1
        dataframe_cash["cum_perf"] = dataframe_cash["perf_annualise"]
        dataframe_cash["cum_perf"] = (1 + dataframe_cash["perf_annualise"]).cumprod() #on calcule le produit cumulative
        print(dataframe_cash)
        dataframe_cash["daily_return"] = dataframe_cash["cum_perf"].pct_change() 

        #on choisi des sous datasets pour avoir la meme date de depart et de fin 
        #end_date = min(monthly_returns.index[-1], expected_returns.index[-1]) #
        monthly_returns_aligned = monthly_returns.loc[start_date:end_date]
        expected_returns_aligned = expected_returns.loc[start_date:end_date]
        print(monthly_returns_aligned.head())
        print(expected_returns_aligned.head())

        #drop colonnes cash ~dans expected_returns_aligned car (initile dans le calcul du portfolio opti)
        monthly_returns_aligned.drop('Cash '+ SelectData().zone, axis=1, inplace=True)
        expected_returns_aligned.drop('Cash '+ SelectData().zone, axis=1, inplace=True)
        YTM.drop('Cash '+ SelectData().zone, axis=1, inplace=True)


        #on calcule la perf cumulative du cash car mauvaise initialement (utile pr la partie où  l'on va long cash only)
        #cum_perf est comme un indice en daily, nous allons maintenant récupérer les valeurs de cet indice fictif mensuellement.
        for date in monthly_returns_aligned.index[1:] :
            month = date.month
            year = date.year
            data_for_month = dataframe_cash[(dataframe_cash.index.month == month) & (dataframe_cash.index.year == year)]
            cash_end_month = data_for_month.iloc[-1]["cum_perf"]
            date_precedent = date - relativedelta(months=1)
            data_for_month_prev = dataframe_cash[(dataframe_cash.index.month == date_precedent.month ) & (dataframe_cash.index.year == date_precedent.year)]
            cash_prev_month = data_for_month_prev.iloc[-1]["cum_perf"]
            #monthly_returns_aligned.at[date, 'Cash EU']  = (cash_eu_end_month - cash_eu_prev_month) /(cash_eu_prev_month)
            monthly_returns.at[date, 'Cash '+SelectData().zone]  = (cash_end_month - cash_prev_month) /(cash_prev_month)

        '''with pd.ExcelWriter('alignés.xlsx') as writer:
            monthly_returns_aligned.to_excel(writer, sheet_name='Monthly Returns Aligned')
            expected_returns_aligned.to_excel(writer, sheet_name='Expected Returns Aligned')'''

        #pour obtenir le risk free rate utile dans le parametre du max_sharpe
        resp2 = LocalTerminal.get_historical(indexes, 'PX_LAST',start=start_date, end=end_date, period='MONTHLY')
        data2 = resp2.as_frame()
        data2.columns = data2.columns.droplevel(1)

        if SelectData().zone=='CN':
            data2 = data2.rename(columns={'CHBM7D Index': 'ST index'})

        str_series2 = pd.Series(index=pd.to_datetime(data2.index), name='ST index', dtype='float64')

        if SelectData().zone !='CN':
            ind2 = np.isnan(data2[indexes[1]])
            str_series2.loc[~ind2] = data2.loc[~ind2, indexes[1]]
            str_series2.loc[ind2] = data2.loc[ind2, indexes[0]]
        else : 
            str_series2 = data2['ST index']


        str_series2 = end_month_data(str_series2, True)

        weights = []
        date_tab = []
        estimated_vol = pd.DataFrame(index=monthly_returns_aligned.index)
        datedeb = monthly_returns_aligned.index[rolling+1]
        for i, date in enumerate(monthly_returns_aligned.index[:-1]):
            if i>0:  
                if(date.strftime("%Y-%m-%d")<monthly_returns_aligned.index[rolling].strftime("%Y-%m-%d")):
                    continue
                if(date.strftime("%Y-%m-%d") =='2022-09-30'):
                    print("ok")
                #monthly_returns_know contient toutes les returns monthly avant la date d'aujourd'hui donc le mois en cours et les mois d'avants
                monthly_returns_know= monthly_returns_aligned.loc[:date]
                #cov_matrix du mois d'après que l'on suppose déjà connu
                if(known_future_cov_matrix):
                    cov_matrix = covariance_dict[Timestamp(monthly_returns_aligned.index[i+1])]
                else :
                    if rolling >1 :
                        if(len(monthly_returns_know) >rolling): #conservation des 5 dernières années pr l'estimation de la cov matrix
                            monthly_returns_know = monthly_returns_know[-rolling:]
                        cov_matrix  = monthly_returns_know.cov()*12
                    else :
                        cov_matrix = covariance_dict[Timestamp(monthly_returns_aligned.index[i])]

                #trace the volatilities
                estimated_vol.loc[date, 'Treasuries '+SelectData().zone] = np.sqrt(cov_matrix.loc['Treasuries '+SelectData().zone, 'Treasuries '+SelectData().zone]*252)
                estimated_vol.loc[date, 'HG Corpo '+SelectData().zone] = np.sqrt(cov_matrix.loc['HG Corpo '+SelectData().zone, 'HG Corpo '+SelectData().zone]*252)
                estimated_vol.loc[date, 'HY Corpo '+SelectData().zone] = np.sqrt(cov_matrix.loc['HY Corpo '+SelectData().zone, 'HY Corpo '+SelectData().zone]*252)
                estimated_vol.loc[date, 'Equity '+SelectData().zone] = np.sqrt(cov_matrix.loc['Equity '+SelectData().zone, 'Equity '+SelectData().zone]*252)
 
                #en connaisant les vrais YTM du mois suivant
                if(known_future_expected_returns):
                    ef = pypfopt.efficient_frontier.EfficientFrontier(YTM.loc[Timestamp(monthly_returns_aligned.index[i+1])], cov_matrix, weight_bounds=(0, 1), solver=None, verbose=False, solver_options=None)
                else:
                    ef = pypfopt.efficient_frontier.EfficientFrontier(expected_returns_aligned.loc[date], cov_matrix, weight_bounds=(0, 1), solver=None, verbose=False, solver_options=None)
                #ligne lorsque l'on ne connais pas les vrais expected YTM du prochain mois
                #ef = pypfopt.efficient_frontier.EfficientFrontier(expected_returns_aligned.loc[date], cov_matrix, weight_bounds=(0, 1), solver=None, verbose=False, solver_options=None)
                print(date)
                try :
                    riskfree = str_series2.loc[date,'ST index']
                    poids = ef.max_sharpe(risk_free_rate=riskfree/100)
                    weights.append({'date': date, 'weights': poids})
                except Exception as e :
                    print("error for date", date," ", e)
                    if(len(weights)==0):
                        poids = OrderedDict([('Treasuries CN', 0.0),('HG Corpo CN', 0.0),('HY Corpo CN', 0.0),('Equity CN', 0.0)])
                    else :
                        poids = weights[-1]['weights']  # Get the previous weights
                    weights.append({'date': date, 'weights': poids})
        data = {}

        # Parcours de chaque élément dans la liste weights et stocke dans le dataframe df
        for element in weights:
            date = element['date']  # extraction de la date
            poids = element['weights']  # extraction des poids
            data[date] = poids
        df = pd.DataFrame(data)
        df = df.transpose()
        df_weights = df.copy()
        df_weights.index = monthly_returns[datedeb:end_date].index
        weighted_returns = monthly_returns[datedeb:end_date]*df_weights 
        #les rendemments maximaux pr chaque dates parmis les 5 catégories d'actifs
        max_values = monthly_returns[datedeb:end_date].max(axis=1)
        strategy_returns = weighted_returns.sum(axis=1) #somme les poids * les returns
        strategy_returns_df = pd.DataFrame(data={'strategy_returns': strategy_returns})
        print(strategy_returns_df)
        cumulative_performance = (1 + strategy_returns_df['strategy_returns']).cumprod()
        print(cumulative_performance)
        years = (strategy_returns_df.index[-1] - strategy_returns_df.index[0]).days / 365
        annualized_return = (cumulative_performance[-1])**(1/years) - 1 
        print(annualized_return)
        print(df_weights)


        #cas LONG ONLY HG 
        copy = df
        copy['Treasuries '+ SelectData().zone] =0
        copy['HG Corpo ' + SelectData().zone] = 1
        copy['HY Corpo ' + SelectData().zone] = 0
        copy['Equity '+ SelectData().zone] = 0
        copy['Cash '+ SelectData().zone] = 0
        print(copy)
        copy.index = monthly_returns[datedeb:end_date].index
        weighted_returns_HG = monthly_returns[datedeb:end_date]*copy
        strategy_returns_HG = weighted_returns_HG.sum(axis=1)
        strategy_returns_df_HG = pd.DataFrame(data={'strategy_returns': strategy_returns_HG})
        print(strategy_returns_df_HG)
        cumulative_performance_HG = (1 + strategy_returns_df_HG['strategy_returns']).cumprod()


        #cas LONG la classe qui perf le mieux a chaque fois

        strategy_returns_df_maxV = pd.DataFrame(data={'strategy_returns':max_values})
        print(strategy_returns_df_maxV )
        cumulative_performance_maxV  = (1 + strategy_returns_df_maxV ['strategy_returns']).cumprod()



        #cas LONG ONLY HY
        copy_HY = df
        copy_HY['Treasuries '+ SelectData().zone] = 0
        copy_HY['HG Corpo '+ SelectData().zone] = 0
        copy_HY['HY Corpo '+ SelectData().zone] = 1
        copy_HY['Equity '+ SelectData().zone] = 0
        copy_HY['Cash '+ SelectData().zone] = 0
        copy_HY.index = monthly_returns[datedeb:end_date].index
        weighted_returns_HY = monthly_returns[datedeb:]*copy_HY
        strategy_returns_HY = weighted_returns_HY.sum(axis=1)
        strategy_returns_df_HY = pd.DataFrame(data={'strategy_returns': strategy_returns_HY})
        print(strategy_returns_df_HY)
        cumulative_performance_HY = (1 + strategy_returns_df_HY['strategy_returns']).cumprod()

        #cas LONG ONLY EQUITY
        copy_e = df
        copy_e['Treasuries '+ SelectData().zone] = 0
        copy_e['HG Corpo '+ SelectData().zone] = 0
        copy_e['HY Corpo '+ SelectData().zone] = 0
        copy_e['Equity '+ SelectData().zone] = 1
        copy_e['Cash '+ SelectData().zone] = 0
        copy_e.index = monthly_returns[datedeb:end_date].index
        weighted_returns_e = monthly_returns[datedeb:]*copy_e
        strategy_returns_e = weighted_returns_e.sum(axis=1)
        strategy_returns_df_e = pd.DataFrame(data={'strategy_returns': strategy_returns_e})
        print(strategy_returns_df_e)
        cumulative_performance_e = (1 + strategy_returns_df_e['strategy_returns']).cumprod()

        #cas LONG ONLY CASH EU
        copy_c = df
        copy_c['Treasuries '+SelectData().zone] = 0
        copy_c['HG Corpo '+SelectData().zone] = 0
        copy_c['HY Corpo '+SelectData().zone] = 0
        copy_c['Equity '+SelectData().zone] = 0
        copy_c['Cash '+SelectData().zone] = 1
        copy_c.index = monthly_returns[datedeb:end_date].index
        weighted_returns_c = monthly_returns[datedeb:]*copy_c
        strategy_returns_c = weighted_returns_c.sum(axis=1)
        strategy_returns_df_c = pd.DataFrame(data={'strategy_returns': strategy_returns_c})
        print(strategy_returns_df_c)
        cumulative_performance_c = (1 + strategy_returns_df_c['strategy_returns']).cumprod()

        #cas LONG ONLY TREASURIES
        copy_t = df
        copy_t['Treasuries '+SelectData().zone] = 1
        copy_t['HG Corpo '+SelectData().zone] = 0
        copy_t['HY Corpo '+SelectData().zone] = 0
        copy_t['Equity '+SelectData().zone] = 0
        copy_t['Cash '+SelectData().zone] = 0
        copy_t.index = monthly_returns[datedeb:end_date].index
        weighted_returns_t = monthly_returns[datedeb:]*copy_t
        strategy_returns_t = weighted_returns_t.sum(axis=1)
        strategy_returns_df_t = pd.DataFrame(data={'strategy_returns': strategy_returns_t})
        print(strategy_returns_df_t)
        cumulative_performance_t = (1 + strategy_returns_df_t['strategy_returns']).cumprod()

        #création du dossier pypfopt et du sous dossier correspond aux nom des paramètres de la fonction
        pypfopt_dossier = os.path.join(ClassGlobal().path_output,"Simulation\{}\Pypfopt".format(ClassGlobal().simulation,datetime.datetime.now().strftime("%Y_%m_%d"),"Pypfopt"))
        if not os.path.exists(pypfopt_dossier):
            os.makedirs(pypfopt_dossier)
            
        destination_folder = os.path.join(pypfopt_dossier,nom)
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)



        plt.figure(figsize=(12, 6))
        cumulative_performance_maxV.plot()
        plt.title('Rendement Cumulé maximal')
        plt.xlabel('Date')
        plt.ylabel('Rendement Cumulé maximal')
        plt.grid()
        destination_file = destination_folder + "/performance_cumulative_maximale"+nom+".png"
        plt.savefig(destination_file)
        plt.close()

        #création du graphe de performance cumulative
        plt.figure(figsize=(10, 6))  
        #plt.figure(figsize=(10, 6))
        plt.plot(cumulative_performance.index, cumulative_performance.values, label='Performance cumulative', color='brown')
        plt.plot(cumulative_performance_HG.index, cumulative_performance_HG.values, label='Performance cumulative HG', color='orange')
        plt.plot(cumulative_performance_HY.index, cumulative_performance_HY.values, label='Performance cumulative HY', color='green')
        plt.plot(cumulative_performance_e.index, cumulative_performance_e.values, label='Performance cumulative equities', color='red')
        plt.plot(cumulative_performance_t.index, cumulative_performance_t.values, label='Performance cumulative treasuries', color='lightskyblue')
        plt.plot(cumulative_performance_c.index, cumulative_performance_c.values, label='Cash '+ SelectData().zone, color='black')
        plt.title('Performance cumulative with '+nom)
        plt.xlabel('Date')
        plt.ylabel('Valeur cumulative')
        plt.legend()
        plt.grid(True)
        destination_file = destination_folder + "/performance_cumulative_"+nom+".png"
        plt.savefig(destination_file)
        plt.close()

        fig, ax = plt.subplots(figsize=(10, 6))
        colors = ['lightskyblue', 'orange', 'green','red']
        i = 0
        for col in df_weights.columns:
            ax.plot(df_weights.index, df_weights[col], label=col,color=colors[i] )
            i +=1
        ax.legend()
        ax.set_title('Évolution des poids au fil du temps with '+nom)
        ax.set_xlabel('Date')
        ax.set_ylabel('Poids')
        plt.savefig(os.path.join(destination_folder, "weight_matrix"+nom+".png"))
        print("La figure a été enregistrée sous :", destination_file)
        plt.close(fig)
        #affichons le graphe de la volatilité
        if (known_future_cov_matrix) : #correspond à la racine de la diagonale de la matrice covariance variance estimée en daily
            plt.figure(figsize=(10, 6))  
            i = 0
            colors = ['lightskyblue', 'orange', 'green','red']
            for column in estimated_vol.columns:
                plt.plot(estimated_vol.index, estimated_vol[column], label=column, color=colors[i])
                i +=1
            plt.xlabel('Date')
            plt.ylabel('Volatilité estimée par le modele')
            plt.title('Volatilité en fonction du time '+nom)
            plt.legend()
            plt.savefig(os.path.join(destination_folder, "volatility_estimated" +nom +".png"))
            print("La figure a été enregistrée sous :", destination_file)
            plt.close(fig)
        else : #sinon on utilise le fichier Vol déjà calculé dans l'algo précédemment
            vol = vol[start_date:]
            plt.figure(figsize=(10, 6))  
            i=0
            colors = ['lightskyblue', 'orange', 'green','red', 'black']
            for column in vol.columns:
                plt.plot(vol.index, vol[column], label=column, color=colors[i])
                i +=1
            plt.xlabel('Date')
            plt.ylabel('Volatilité estimée par notre algo')
            plt.title('Volatilité en fonction du time ' + nom)
            plt.legend()
            plt.savefig(os.path.join(destination_folder, "volatility_estimated"+nom+".png"))
            print("La figure a été enregistrée sous :", destination_file)
            plt.close(fig)
        #printons les expected returns 
        plt.figure(figsize=(10, 6))  
        i = 0
        colors = ['lightskyblue', 'orange', 'green','red']
        for column in expected_returns_aligned:
            plt.plot(expected_returns_aligned.index, expected_returns_aligned[column], label=column, color=colors[i])
            i+=1
        plt.xlabel('Date')
        plt.ylabel('Expected returns with '+nom)
        plt.title('Expected returns en fonction du time '+nom)
        plt.legend()
        plt.savefig(os.path.join(destination_folder, "expected_returns"+nom+".png"))
        print("La figure a été enregistrée sous :", destination_file)

        #printer le graphe de la théorie moderne du ptf markowitz
        from pypfopt import plotting
        S = cov_matrix
        mu = expected_returns_aligned.loc[end_date]
        n_samples = 1000
        w = np.random.dirichlet(np.ones(len(mu)), n_samples)
        rets = w.dot(mu)
        stds = np.sqrt((w.T * (S @ w.T)).sum(axis=0))
        sharpes = rets / stds
        ef =  pypfopt.efficient_frontier.EfficientFrontier(mu, S)
        fig, ax = plt.subplots()
        plotting.plot_efficient_frontier(ef, ax=ax, show_assets=False)
        ''# Find and plot the tangency portfolio
        ef2 =  pypfopt.efficient_frontier.EfficientFrontier(mu, S)
        ef2.max_sharpe()
        ret_tangent, std_tangent, _ = ef2.portfolio_performance()
        # Plot random portfolios
        ax.scatter(stds, rets, marker=".", c=sharpes, cmap="gray")
        print(mu)
        print(S.iloc[0,0])
        ax.set_title("Efficient Frontier with random portfolios with "+nom)
        #max_sharpe =monthly_returns[end_date]*df.loc[end_date]
        #ordonnee = max_sharp.sum(axis=1)
        ax.scatter(std_tangent, ret_tangent, marker="o", color='r', label='Portfolio Optimal')
        ax.legend()
        plt.tight_layout()
        plt.savefig(os.path.join(destination_folder, "Efficient Frontier with "+nom+".png"))
        plt.close(fig)

if __name__ == "__main__":
    luigi.build([ClassGlobal(), SelectData(), StrategyPypfopt()],local_scheduler = True)
    luigi.build([Strategy_Pypfopt()],local_scheduler = True, no_lock=True)
