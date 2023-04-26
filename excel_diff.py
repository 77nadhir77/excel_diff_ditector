# -*- coding: utf-8 -*-
"""
Created on Wed Apr 26 11:26:51 2023

@author: Nadhir
"""

import sys
import pandas as pd





'''
modifier les paramètres suivant selons vos données
'''

#l'idantifiant de chaque enregistrement dans le tableau
primary_key = "account number"
###########

#le nom du fichier obsolète
old_file_name = "sample-address-1.xlsx"
#########

#le numéro de sheet du fichier obsolète
old_file_sheet = "Sheet1"
###############


#le nom du fichier met à jour
new_file_name = "sample-address-1.xlsx"
##############

#le numéro de sheet du fichier met à jour
new_file_sheet = "Sheet2"
##############


#le nom du fichier ou le résultat s'affiche
result_file = "result.xlsx"
############

##############################################
#############################
################

















# Define the diff function to show the changes in each field
def report_diff(x):
    return x[0] if x[0] == x[1] else '{} ---> {}'.format(*x)

# We want to be able to easily tell which rows have changes
def has_change(row):
    if "--->" in row.to_string():
        return "Y"
    else:
        return "N"

#read the two excel files
old = pd.read_excel(old_file_name, old_file_sheet, na_values=['NA'])
new = pd.read_excel(new_file_name, new_file_sheet, na_values=['NA'])


old['version'] = "old"
new['version'] = "new"


old_accts_all = set(old[primary_key])
new_accts_all = set(new[primary_key])

dropped_accts = old_accts_all - new_accts_all
added_accts = new_accts_all - old_accts_all


#concatenate the two sets
concatenated_set = pd.concat([old,new],ignore_index=True)

changes = concatenated_set.drop_duplicates(subset=old.drop('version',axis=1).columns.tolist(),keep='last')

dupe_accts = changes.set_index(primary_key)[changes.set_index(primary_key).index.duplicated()].index.tolist()
dupes = changes[changes[primary_key].isin(dupe_accts)]


changes_old = dupes[(dupes['version'] == 'old')]
changes_new = dupes[(dupes['version'] == 'new')]


changes_old = changes_old.drop(['version'], axis=1)
changes_new = changes_new.drop(['version'], axis=1)

changes_old.set_index(primary_key, inplace=True)
changes_new.set_index(primary_key, inplace=True)


all_modifications = pd.concat([changes_old, changes_new],
                            axis='columns',
                            keys=['old', 'new'],
                            join='outer')
all_modifications = all_modifications.swaplevel(axis='columns')[changes_new.columns[0:]]


df_changed = all_modifications.groupby(level=0, axis=1).apply(lambda frame: frame.apply(report_diff, axis=1))
df_changed = df_changed.reset_index()



df_removed = changes[changes[primary_key].isin(dropped_accts)]
df_added = changes[changes[primary_key].isin(added_accts)]


output_columns = df_changed.columns.tolist()
writer = pd.ExcelWriter(result_file)
df_changed.to_excel(writer,"changés", index=False, columns=output_columns)
df_removed.to_excel(writer,"suprimés",index=False, columns=output_columns)
df_added.to_excel(writer,"ajoutés",index=False, columns=output_columns)
writer.save()
