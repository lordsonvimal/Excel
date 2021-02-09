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
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# from re import match

from ..excel.excel import ExcelSheet

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
        self.book = load_workbook(file_path)

    def get_domains(self):
        srdm_list = [filename for filename in os.listdir(self.file_dir) if filename.startswith(self.spec_gen) and filename.endswith('.xlsx')]
        df_srdm=[]
        cols = [17, 14, 19, 13]

        for i in range(len(srdm_list)):
            srdm00 = srdm_list[i]
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
        self.domainlist = np.unique([x[:2] for x in self.df_srdm['SName'] if isinstance(x, str) and x.count('_') == 1]).tolist()
        self.domains = np.unique([x[:2] for x in self.df_srdm['SName'] if isinstance(x, str) and x.count('_') == 1]).tolist()
        self.df_srdm1 = self.df_srdm[((self.df_srdm['SName'].str.count('_')==1) | (self.df_srdm['V2'].str.contains('DTS'))) & (self.df_srdm['SType']!='Date')]



        unique_chars = lambda x: ' '.join(x.unique())
        unique_max = lambda x: x.max()
        self.df_srdm2 = self.df_srdm1.groupby('V0').agg({'SName': unique_chars,'SType': unique_chars,'SLabel': unique_chars, 'SLength':unique_max}).reset_index()
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
        self.sdtm32 = pd.read_excel(file_path, engine="openpyxl", sheet_name="SDTMIG v3.2",skiprows=1, names=col_names, usecols="A:L")
        self.sdtm33 = pd.read_excel(file_path, sheet_name="SDTMIG v3.3",skiprows=1, names=col_names, usecols="A:L")
        self.message("Success Reading SDTMIG.xlsx")

    def get_unique_domains_from_sdtm(self):
        self.message("Getting unique list of domains")
        domain_32 = self.sdtm32['Dataset'].unique()
        self.domain_32 = [x[:2] for x in domain_32 if str(x) != 'nan']

        #Get Unique list of domains in SDTM IG 3.3
        self.domain_33 = self.sdtm33['Dataset'].unique()
        self.domain_33 = [x[:2] for x in self.domain_33 if str(x) != 'nan']

        #List Domains in SDTM IG 3.3 which are not available in SDTM IG 3.2
#         self.in33_not_in32 = []
        self.in33_not_in32=[x for x in self.domain_33 if x not in self.domain_32]

        #Keep Only those domains that do not exist in 3.2 and available in SDTM IG 3.3
        self.sdtm33 = self.sdtm33[self.sdtm33.Dataset.isin(self.in33_not_in32)]

        #Consolidated SDTM IG 3.2 and 3.3 (Only those domains that do not exist in 3.2)
        self.sdtm_meta = self.sdtm32.append(self.sdtm33,ignore_index=True)

        self.fval_in32 = [x for x in self.domains if x in self.domain_32] #Domain in 3.2
        self.fval_in33 = [x for x in self.domains if x in self.in33_not_in32] #Domain in 3.3
        self.neither_32_nor_33 = [x for x in self.domains if x not in self.domain_32 + self.domain_33] #Domains in neither 3.2 nor 3.3

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
        self.df_00 = ['All Classes']    #Create a list for filteration criteria based on 'Class' column.
        self.sdtm32_02=[]    #Create an empty list to append the data for all domains in fval_list
        for i in range(len(self.domains)):
            domain = self.domains[i]
            self.sdtm32_00 = self.sdtm_meta['Class'][self.sdtm_meta['Dataset'].isin([domain])].unique()    #To get the 'Class' of respective domain from fval.
            self.df_00.append([x+'-General' for x in self.sdtm32_00][0])    #To get the Class with corresponding '-General' for filteration
            # To filter the dataframe with required data and class in 'All Classes' and corresponding '-General' Class.
            self.sdtm32_01 = self.sdtm_meta[np.logical_and(self.sdtm_meta['Dataset'].isna(), self.sdtm_meta['Class'].isin(self.df_00))]
            self.sdtm32_01 =pd.concat([self.sdtm32_01,self.sdtm_meta[self.sdtm_meta['Dataset'] == domain]])   #To filter on domain
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
        self.s_sdtm = pd.merge(self.sdtm32_02, self.df_srdm2, left_on = 'Name', right_on='V0', how = 'left')
        self.s_sdtm.head()
        self.message("Successfully merged data")

    def f(self, row):
        if row['Name'] == row['V0']:
            if np.logical_or(row['Type'] == row['SType'], row['Label'] == row['SLabel']):
                val = "Rename " + row['SName'] + " as " + row['Name']+"|"+row['SName']
            else:
                val = "Direct Move from "+row['SName']+" to " + row['Name']+"|"+row['SName']
            return val

    def rule(self):
        self.message("Running rules engine")
        self.s_sdtm['PRuleSOrgn'] = self.s_sdtm.apply(self.f, axis=1)
        self.s_sdtm['PRule'] = np.where(self.s_sdtm['PRuleSOrgn'].isnull()
                                   ,self.s_sdtm['Name'].map(self.df_vcom.set_index('Name')['ProgrammerRule'])
                                   ,self.s_sdtm['PRuleSOrgn'].str.split('|', expand=True)[0])
        self.s_sdtm['SOrgn'] = np.where(self.s_sdtm['PRuleSOrgn'].isnull()
                                   ,self.s_sdtm['Name'].map(self.df_vcom.set_index('Name')['SRDMOrigin'])
                                   ,self.s_sdtm['PRuleSOrgn'].str.split('|', expand=True)[1])
        self.s_sdtm['SubCol'] = self.s_sdtm['Name'].map(self.df_vcom.set_index('Name')['Submission'])
        self.s_sdtm['Origin'] = self.s_sdtm['Name'].map(self.df_vcom.set_index('Name')['Origin'])
        self.s_sdtm['Length'] = self.s_sdtm['Name'].map(self.df_vcom.set_index('Name')['Length'])

        self.s_sdtm[~self.s_sdtm['SOrgn'].isnull()]
        self.message("Rules engine ran successfully")

    def save_sheet(self):
        self.message("Exporting. Please wait...")
        self.writer = pd.ExcelWriter(self.file_path, engine="openpyxl", mode="a")
        workbook = self.writer.book

        for i in range(len(self.domains)):
            domain = self.domains[i]
#             sheet = ExcelSheet(self.s_sdtm, domain, self.writer)
            self.df = pd.DataFrame(columns=['Name','Description','CodeListRef','Label','Length','Sequence',
                                                                  'Supplimentary','Comments', 'Type', 'Origin', 'Core',
                                                                  'ProgrammerRule', 'Submission', 'SRDMOrigin', 'Alias'])
            self.df1 = self.s_sdtm[self.s_sdtm['Dataset'] == domain]
            print("index: "+str(i))
            print(self.df1)
            self.df['Name'] = self.nextgen[self.nextgen['Dataset'] == domain]['Name'].to_numpy()
            self.df['NgCore'] = self.df['Name'].map(self.nextgen[self.nextgen['Dataset'] == domain].set_index('Name')['Core'])
            self.df['Sequence'] = self.df['Name'].map(self.nextgen[self.nextgen['Dataset'] == domain].set_index('Name')['Order'])
            self.df['Description'] = self.df['Name'].map(self.df1.set_index('Name')['Description'])
            self.df['Core'] = self.df['Name'].map(self.df1.set_index('Name')['Core'])
            self.df['CodeListRef'] = self.df['Name'].map(self.df1.set_index('Name')['CodeListRef'])
            self.df['Label'] = self.df['Name'].map(self.df1.set_index('Name')['Label'])
            self.df['Type'] = self.df['Name'].map(self.df1.set_index('Name')['Type'])
            self.df.assign(Supplimentary="0")
            self.df.assign(Comments="0")
            self.df['ProgrammerRule'] = self.df['Name'].map(self.df1.set_index('Name')['PRule'])
            self.df['SRDMOrigin'] = self.df['Name'].map(self.df1.set_index('Name')['SOrgn'])
            self.df['Origin'] = self.df['Name'].map(self.df1.set_index('Name')['Origin'])
            self.df['Length'] = self.df['Name'].map(self.df1.set_index('Name')['Length'])
            self.df.to_excel(self.writer, sheet_name=domain, index=False)
            worksheet = self.writer.sheets[domain]
#             worksheet.column_dimensions['Name'].width = 40
#             worksheet.column_dimensions['Description'].width = 40
#             worksheet.column_dimensions['CodeListRef'].width = 40
#             worksheet.column_dimensions['Label'].width = 40
#             worksheet.column_dimensions['Length'].width = 40
#             worksheet.column_dimensions['Sequence'].width = 40
#             worksheet.column_dimensions['Supplimentary'].width = 40
#             worksheet.column_dimensions['Comments'].width = 40
#             worksheet.column_dimensions['Type'].width = 40
#             worksheet.column_dimensions['Origin'].width = 40
#             worksheet.column_dimensions['Core'].width = 40
#             worksheet.column_dimensions['ProgrammerRule'].width = 40
#             worksheet.column_dimensions['Submission'].width = 40
#             worksheet.column_dimensions['SRDMOrigin'].width = 40
#             worksheet.column_dimensions['Alias'].width = 40
            print(worksheet)
#             for (col_name, columnData) in self.df.iteritems():
#                 for cell in self.writer.sheets[domain][col_name]:
#                     cell.alignment = Alignment(wrap_text=True)
#             format = workbook.add_format({'text_wrap': True})
#             worksheet.set_column('A:P', 40, format)

        self.writer.save()
        self.writer.close()
        self.message("Successfully exported sheet")

    def process(self):
        self.get_domains()
        self.read_nextgen()
        self.read_sdtm_ig()
        self.get_unique_domains_from_sdtm()
#         self.append_data_for_all_domains()
        self.set_desired_columns()
        self.read_comm()
        self.merge_data()
        self.rule()
        self.save_sheet()
