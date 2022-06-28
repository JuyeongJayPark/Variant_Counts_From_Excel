#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#Developer: Juyeong.Park
#Description: Calculating Variant Counts from Annotated VCF Excel File


# In[2]:


import os


# In[28]:


import pandas as pd 


# In[17]:


wd_path = "/TBI/People/tbi/jypark/busan_vcf_stat/Sample_Sheets_2021" # Working Directory Path


# In[3]:


# Listing Files From "Sample_Sheets_2021" Folder
Sample_Files = os.listdir("Sample_Sheets_2021")


# In[25]:


Sample_list = []
Sample_Panel_list = []
Variant_Count_list = []


# In[26]:


for Sample_File in Sample_Files: 
    Sample = Sample_File.split('.xlsx')[0]
    Sample_Panel = "_".join(Sample_File.split('_')[2:-1])
    Varinat_Count = pd.read_excel(os.path.join(wd_path,Sample_File), sheet_name = "SNV INDEL").shape[0]
    
    Sample_list.append(Sample)
    Sample_Panel_list.append(Sample_Panel)
    Variant_Count_list.append(Varinat_Count)


# In[29]:


Variants_Stat_df = pd.DataFrame({"Sample" : Sample_list, "Panel": Sample_Panel_list, "Variant_cnt" : Variant_Count_list})


# In[34]:


Variants_Stat_df.to_excel("Variants_Count_Table.xlsx", index =False, header = True)


# In[35]:


Variants_Stat_df.groupby("Panel").mean().to_excel("Variants_Count_Mean_Table.xlsx", index = True, header =True)


# In[ ]:




