import pandas as pd
# import portable_spreadsheet as ps
# import openpyxl
# from openpyxl import load_workbook
# from openpyxl import Workbook
# from openpyxl.styles import Alignment
# from openpyxl.utils import get_column_letter
import numpy as np
# from pathlib import Path
# import itertools
# from pandasql import *
# pysqldf = lambda q:sqldf(q,globals())
import os
# from re import match

# from ..excel.excel import Excel

class Spec:
    def __init__(self, file_path, spec_gen, append_message):
        self.file_path = file_path
        self.file_dir = os.path.dirname(file_path)
        self.spec_gen = spec_gen
        self.domains = []
        self.message = append_message
        self.df_srdm = None
        self.nextgen = None
        self.sdtm32 = None
        self.sdtm33 = None
        self.sdtm32_02 = None

    def get_domains(self):
        srdm_list = [filename for filename in os.listdir(self.file_dir) if filename.startswith(self.spec_gen) and filename.endswith('.xlsx')]
        df_srdm=[]
        cols = [17, 14, 19, 13]

        for i in range(len(srdm_list)):
            srdm01 = pd.read_excel(os.path.join(self.file_dir, srdm_list[i]), sheet_name=3, skiprows=1,
                              usecols=cols, names=['SType','SLength','SName','SLabel'])
            srdm01 = srdm01.loc[srdm01['SName'].str.contains('_')]
            srdm01['SType']= np.where(srdm01['SType'].str.upper() == 'VARCHAR2','Character',srdm01['SType'].str.title())
            n_col = (srdm01['SName'].str.split("_").apply(len)).max()
            collis = []
            for i in range(n_col):
                collis.append('V'+str(i))
            srdm01[collis] = srdm01['SName'].str.split('_', expand=True)
            df_srdm.append(srdm01)
        self.df_srdm = pd.concat(df_srdm)

        self.domains = np.unique([x[:2] for x in self.df_srdm['SName'] if isinstance(x, str) and x.count('_') == 1]).tolist()
        self.message("Domains Identified are: " + str(self.domains))

    def read_nextgen(self):
        self.message("Reading NextGEN SDTM Metadata.xlsx")
        file_path = os.path.join(self.file_dir, "NextGEN SDTM Metadata.xlsx")
        self.nextgen = pd.read_excel(file_path, sheet_name="Elements", usecols=['Order','Dataset','Name','Core'])
        self.nextgen = self.nextgen[self.nextgen['Dataset'].isin(self.domains)]
        self.nextgen.reset_index(drop=True)
        self.nextgen.head()
        self.message("Success Reading NextGEN SDTM Metadata.xlsx")

    def read_sdtm_ig(self):
        self.message("Reading SDTMIG.xlsx")
        file_path = os.path.join(self.file_dir, "SDTMIG.xlsx")
        col_names = ['Version','Order','Class','Dataset','Vname','Name','Label','Type','CodeListRef','Role','Description','Core']
        self.sdtm32 = pd.read_excel(file_path, sheet_name="SDTMIG v3.2",header=None, skiprows=1, names=col_names)
        self.sdtm33 = pd.read_excel(file_path, sheet_name="SDTMIG v3.3",header=None, skiprows=1, names=col_names)
        self.message("Success Reading SDTMIG.xlsx")

    def get_unique_domains_from_sdtm(self):
        self.message("Getting unique list of domains")
#         self.domain_32 = []
        self.domain_32 = self.sdtm32['Dataset'].unique()
        self.domain_32 = [x for x in self.domain_32 if str(x) != 'nan']

        #Get Unique list of domains in SDTM IG 3.3
#         domain_33=[]
        self.domain_33 = self.sdtm33['Dataset'].unique()
        self.domain_33 = [x for x in self.domain_33 if str(x) != 'nan']

        #List Domains in SDTM IG 3.3 which are not available in SDTM IG 3.2
#         self.in33_not_in32 = []
        self.in33_not_in32=[x for x in self.domain_33 if x not in set(self.domain_32)]

        #Keep Only those domains that do not exist in 3.2 and available in SDTM IG 3.3
        self.sdtm33 = self.sdtm33[self.sdtm33.Dataset.isin(self.in33_not_in32)]

        self.fval_in32 = [x for x in self.domains if x in set(self.domain_32)] #Domain in 3.2
        self.fval_in33 = [x for x in self.domains if x in set(self.in33_not_in32)] #Domain in 3.3
        self.neither_32_nor_33 = [x for x in self.domains if x not in set(self.domain_32 and self.domain_33)] #Domains in neither 3.2 nor 3.3

        #Consolidated SDTM IG 3.2 and 3.3 (Only those domains that do not exist in 3.2)
        self.sdtm_meta = self.sdtm32.append(self.sdtm33,ignore_index=True)
        self.message("Domains in 3.2: "+str(self.fval_in32))
        self.message("Domains in 3.3: "+str(self.fval_in33))
        self.message("Domains neither in 3.2 nor in 3.3 : "+str(self.neither_32_nor_33))

    def append_data_for_all_domains(self):
        self.message("Appending data for all domains")
        self.sdtm_00=[]
        for i in range(len(self.domains)):
            self.sdtm_01 = self.sdtm_meta[self.sdtm_meta['Dataset'] == self.domains[i]]   #To filter on domain
            self.sdtm_01['Type']=np.where(self.sdtm_01['Type']=='text','Character','Number')
            self.sdtm_00.append(self.sdtm_01)
        self.sdtm_00 = pd.concat(self.sdtm_00)
        self.message("Successfully appended data")

    def set_desired_columns(self):
        self.message("Filtering desired columns")
        self.sdtm_meta = self.sdtm_meta[['Class','Dataset','Name','Core','Label','Type', 'CodeListRef', 'Description']]
        self.df_00 = ['All Classes']    #Create a list for filteration criteria based on 'Class' column.
        self.sdtm32_02 = []
        print(self.sdtm32['Class'])
        print(self.sdtm32['Dataset'])
        print(self.sdtm32['Dataset'].isin(["CM"]))
        for i in range(len(self.domains)):
            domain = self.domains[i]
            self.sdtm32_00 = self.sdtm32['Class'][self.sdtm32['Dataset'].isin([domain])].unique()    #To get the 'Class' of respective domain from fval.
            if len(self.sdtm32_00) > 0:
                self.df_00.append([x+'-General' for x in self.sdtm32_00][0])    #To get the Class with corresponding '-General' for filteration
                # To filter the dataframe with required data and class in 'All Classes' and corresponding '-General' Class.
                self.sdtm32_01 = self.sdtm32[np.logical_and(self.sdtm32['Dataset'].isna(), self.sdtm32['Class'].isin(self.df_00))]
                self.sdtm32_01 =pd.concat([self.sdtm32_01,self.sdtm32[self.sdtm32['Dataset'] == domain]])   #To filter on domain
                self.sdtm32_01['Name']=self.sdtm32_01['Name'].str.replace('--',domain)    #replace'--' with Domain name in "Name" Column
                self.sdtm32_01['Dataset']=np.where(self.sdtm32_01['Dataset'].isnull(),domain,self.sdtm32_01['Dataset'])
                self.sdtm32_01['Core']=np.where(self.sdtm32_01['Core'].isnull(),'Perm',self.sdtm32_01['Core'])
                self.sdtm32_01['Type']=np.where(self.sdtm32_01['Type']=='Num','Number','Character')
                self.sdtm32_02.append(self.sdtm32_01)
                self.df_00=['All Classes']
        self.sdtm32_02 = pd.concat(self.sdtm32_02)
        #keep the last duplicate value per Dataset & Name
        self.sdtm32_02.drop_duplicates(subset=['Dataset','Name'], keep='last', inplace=True)
        #To Remove '()' from 'CodeListRef' Column.
        self.sdtm32_02['CodeListRef'] = self.sdtm32_02['CodeListRef'].str.replace(r'[^\w\s]+', '')
        self.sdtm32_02['CodeListRef']= np.where(self.sdtm32_02['Name'].str.upper() == 'EPOCH','EPOCH',self.sdtm32_02['CodeListRef'])
        self.sdtm32_02['CodeListRef']= np.where(self.sdtm32_02['CodeListRef'].str.upper().isin(['MEDDRA','ISO 8601','ISO 3166-1 Alpha-3']),np.nan,self.sdtm32_02['CodeListRef'])
        self.message("Successfully filtered data")
        print(self.sdtm32_02)

    def read_comm(self):
        self.message("Reading COMM file")
        comm_files = [filename for filename in os.listdir(self.file_dir) if filename.endswith('_COMM.xlsx')]
        if len(comm_files) > 0:
            file = os.path.join(self.file_dir, comm_files[0])
            self.df_vcom = pd.read_excel(file, sheet_name="COMM")
            for i in range(len(self.domains)):
                domain = self.domains[i]
                self.df_vcom['Domain'] = domain
                self.df_vcom['Name'] = self.df_vcom['Name'].str.replace('__', domain)
        self.message("Successfully read comm file")

    def merge_data(self):
        self.message("Merging")
#         print(self.sdtm32_02)
        print(self.sdtm_00)
        print(self.df_vcom)
        self.s_sdtm = pd.merge(self.sdtm_00, self.df_vcom, left_on = 'Name', right_on='V0', how = 'left')
        #df_002 = pysqldf("select a.*, b.* from sdtm32_02 as a left join srdm01 as b on a.Name = b.V0;")
        #with pd.ExcelWriter(path + 'spec.xlsx',engine='openpyxl', mode='a') as writer:
               #df_002.to_excel(writer, sheet_name="sdtm8",index=False)
        self.s_sdtm.head()
        self.message("Successfully merged data")
        print(self.s_sdtm)

    def process(self):
        self.get_domains()
        self.read_nextgen()
        self.read_sdtm_ig()
        self.get_unique_domains_from_sdtm()
        self.append_data_for_all_domains()
        self.set_desired_columns()
        self.read_comm()
        self.merge_data()
