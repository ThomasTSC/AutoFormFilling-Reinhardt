# -*- coding: utf-8 -*-
"""
Created on Tue Dec 19 15:50:52 2017

@author: chou
"""
import xlwings
import os
import distutils.dir_util

def Create_Folder_Files(self):
                
                if self.Option_Machine_var.get() == 'VG2':
                    
                    
                
                    #Create folder and documents#
                    try:
                        self.Target_Folder = os.makedirs(r"P:\Forschungs-Projekte\OLED-measurements Bruchsal\B 2018 VG2\B CW%d VG Device %d-%d Batch %d" %(self.var_CW.get(), self.Substrate_list[0], self.Substrate_list[-1], self.var_Batch_Number.get()))
                    except FileExistsError:
                        pass

                    distutils.dir_util.copy_tree(r"P:\Forschungs-Projekte\OLED-measurements Bruchsal\B 2018 VG2\Temporal template", r"P:\Forschungs-Projekte\OLED-measurements Bruchsal\B 2018 VG2\B CW%d VG Device %d-%d Batch %d" %(self.var_CW.get(), self.Substrate_list[0], self.Substrate_list[-1], self.var_Batch_Number.get()))

            
                    self.Batch_Folder = (r"P:\Forschungs-Projekte\OLED-measurements Bruchsal\B 2018 VG2\B CW%d VG Device %d-%d Batch %d" %(self.var_CW.get(), self.Substrate_list[0], self.Substrate_list[-1], self.var_Batch_Number.get()))

                    os.chdir(self.Batch_Folder)
            
                    os.rename('OLED-xxxx-yyyy-fabrication-sheet.xlsx', 'OLED-%d-%d-fabrication-sheet.xlsx' %(self.Substrate_list[0], self.Substrate_list[-1]))            
                    
                    
                    
                    #Put information into excel file#
            
                    self.FabForm = xlwings.Book(r"P:\Forschungs-Projekte\OLED-measurements Bruchsal\B 2018 VG2\B CW%d VG Device %d-%d Batch %d\OLED-%d-%d-fabrication-sheet.xlsx" %(self.var_CW.get(), self.Substrate_list[0], self.Substrate_list[-1], self.var_Batch_Number.get(),self.Substrate_list[0], self.Substrate_list[-1]))  # Creates a connection with workbook
                    self.FabSheet = self.FabForm.sheets['fabSheet']

                    
                    
                    self.FabSheet.range('D5').value = self.Batch_Info['Batch_Number']
                    self.FabSheet.range('D6').value = self.Batch_Info['Project_Name']
                    self.FabSheet.range('D10').value = self.Batch_Info['Aim']
                    
                    
                    if self.Substrate_Number == 4:
                        self.FabSheet.range('P23').value = self.Batch_Info['Substrate_List'][0]
                        self.FabSheet.range('P20').value = self.Batch_Info['Substrate_List'][1]
                        self.FabSheet.range('N23').value = self.Batch_Info['Substrate_List'][2]
                        self.FabSheet.range('N20').value = self.Batch_Info['Substrate_List'][3]
                    
                    
                        #Microscopic Check#
                        self.FabSheet.range('D149').value = self.Batch_Info['Substrate_List'][0]
                        self.FabSheet.range('D150').value = self.Batch_Info['Substrate_List'][1]
                        self.FabSheet.range('D151').value = self.Batch_Info['Substrate_List'][2]
                        self.FabSheet.range('D152').value = self.Batch_Info['Substrate_List'][3]
            
                    
                    
                    
                    self.Batch_Component_Index = 0
                    for L in range(15,15+self.Batch_Info['Layer_Number']+1):
                        for C in range(4,10):
                            if self.Batch_Component_Index < self.Batch_Info['Layer_Number']*6 :
                                self.FabSheet.range((L,C)).value = self.Batch_Info['Architecture'][self.Batch_Component_Index]
                                self.Batch_Component_Index = self.Batch_Component_Index + 1
                    
                    
                
                
                    #Evaporation step#
                    self.Evaporation_Step_count = 1
                    self.Evaporation_Step = []
                    self.Evaporation_Step_Material = []
                    self.Evaporation_Step_Variation = []
                    self.Evaporation_Step_ExpID = []
                    
                    
                    self.Cathode_Index = 15
                
                    for L in range(15,15+self.Batch_Info['Layer_Number']+1):
                        if self.FabSheet.range((L,6)).value == 'ITO':
                            self.HTL_Index = L-1

                    self.Deposited_Layer_count = self.HTL_Index
                
                    while self.Deposited_Layer_count >= self.Cathode_Index:
                        if self.FabSheet.range((self.Deposited_Layer_count,6)).value != '':
                            if ':' in self.FabSheet.range((self.Deposited_Layer_count,6)).value:
                                self.Co_Host = self.FabSheet.range((self.Deposited_Layer_count,6)).value
                                self.Co_Host = self.Co_Host.split(":")
                                self.Evaporation_Step.append(self.Evaporation_Step_count)
                                self.Evaporation_Step_Material.append(self.Co_Host[0])
                                self.Evaporation_Step.append(self.Evaporation_Step_count)
                                self.Evaporation_Step_Material.append(self.Co_Host[1])
                                self.Evaporation_Step_Variation.append(self.FabSheet.range((self.Deposited_Layer_count,4)).value)
                                self.Evaporation_Step_Variation.append(self.FabSheet.range((self.Deposited_Layer_count,4)).value)
                            else:
                                self.Evaporation_Step.append(self.Evaporation_Step_count)
                                self.Evaporation_Step_Material.append(self.FabSheet.range((self.Deposited_Layer_count,6)).value)
                                self.Evaporation_Step_Variation.append(self.FabSheet.range((self.Deposited_Layer_count,4)).value)
                    
                        if self.FabSheet.range((self.Deposited_Layer_count,7)).value != '':
                            if ':' in self.FabSheet.range((self.Deposited_Layer_count,7)).value:
                                self.Co_Emitter = self.FabSheet.range((self.Deposited_Layer_count,7)).value
                                self.Co_Emitter = self.Co_Emitter.split(":")
                                self.Evaporation_Step.append(self.Evaporation_Step_count)
                                self.Evaporation_Step_Material.append(self.Co_Emitter[0])
                                self.Evaporation_Step.append(self.Evaporation_Step_count)
                                self.Evaporation_Step_Material.append(self.Co_Emitter[1])
                                self.Evaporation_Step_Variation.append(self.FabSheet.range((self.Deposited_Layer_count,4)).value)
                                self.Evaporation_Step_Variation.append(self.FabSheet.range((self.Deposited_Layer_count,4)).value)                                
                            
                            else:    
                                self.Evaporation_Step.append(self.Evaporation_Step_count)
                                self.Evaporation_Step_Material.append(self.FabSheet.range((self.Deposited_Layer_count,7)).value)
                                self.Evaporation_Step_Variation.append(self.FabSheet.range((self.Deposited_Layer_count,4)).value)
                    
                        self.Evaporation_Step_count = self.Evaporation_Step_count + 1
                        self.Deposited_Layer_count = self.Deposited_Layer_count - 1
                    


                    for element in self.Evaporation_Step_Material:
                        if element in self.Tooling_List_Info['VG_Material_Name']:
                            self.Evaporation_Step_ExpID.append(self.Tooling_List_Info['VG_Material_Name'].index(element))
                    
                    
                            
                    self.Evaporation_Step_Source = [self.Tooling_List_Info['VG_Material_Source'][i] for i in self.Evaporation_Step_ExpID]
                    self.Evaporation_Step_Toolings = [self.Tooling_List_Info['VG_Material_Toolings'][i] for i in self.Evaporation_Step_ExpID]  
                    self.Evaporation_Step_Sensor = [self.Tooling_List_Info['VG_Material_Sensor'][i] for i in self.Evaporation_Step_ExpID]  
                    self.Evaporation_Step_ExpID = [self.Tooling_List_Info['VG_Material_ExpID'][i] for i in self.Evaporation_Step_ExpID]



            
                    self.Evaporation_Info = dict({'Evaporation_Step':self.Evaporation_Step, 
                                                  'Evaporation_Step_Material':self.Evaporation_Step_Material, 
                                                  'Evaporation_Step_Variation':self.Evaporation_Step_Variation,
                                                  'Evaporation_Step_ExpID':self.Evaporation_Step_ExpID, 
                                                  'Evaporation_Step_Source':self.Evaporation_Step_Source, 
                                                  'Evaporation_Step_Toolings':self.Evaporation_Step_Toolings,
                                                  'Evaporation_Step_Sensor':self.Evaporation_Step_Sensor})
            
                    
                    
                    
                    self.Fill_Index = 0
                    for ES in range (80, 80+len(self.Evaporation_Info['Evaporation_Step'])):
                        self.FabSheet.range((ES,1)).value = self.Evaporation_Info['Evaporation_Step'][self.Fill_Index]
                        self.FabSheet.range((ES,2)).value = self.Evaporation_Info['Evaporation_Step_Variation'][self.Fill_Index]
                        self.FabSheet.range((ES,4)).value = self.Evaporation_Info['Evaporation_Step_Material'][self.Fill_Index]
                        self.FabSheet.range((ES,5)).value = self.Evaporation_Info['Evaporation_Step_ExpID'][self.Fill_Index]
                        self.FabSheet.range((ES,8)).value = self.Evaporation_Info['Evaporation_Step_Source'][self.Fill_Index]
                        self.FabSheet.range((ES,9)).value = self.Evaporation_Info['Evaporation_Step_Toolings'][self.Fill_Index]
                        self.FabSheet.range((ES,10)).value = self.Evaporation_Info['Evaporation_Step_Sensor'][self.Fill_Index]

                        
                        self.Fill_Index = self.Fill_Index + 1                
                    
                    
                    
            
                if self.Option_Machine_var.get() == 'Lesker':
                    
                    #Create folder and documents#
                    try:
                        self.Target_Folder = os.makedirs(r"P:\Forschungs-Projekte\OLED-measurements Bruchsal\B 2018 Lesker\B CW%d Lesker Device %d-%d Batch %d" %(self.var_CW.get(), self.Substrate_list[0], self.Substrate_list[-1], self.var_Batch_Number.get()))
                    except FileExistsError:
                        pass

                    distutils.dir_util.copy_tree(r"P:\Forschungs-Projekte\OLED-measurements Bruchsal\B 2018 Lesker\Temporal template", r"P:\Forschungs-Projekte\OLED-measurements Bruchsal\B 2018 Lesker\B CW%d Lesker Device %d-%d Batch %d" %(self.var_CW.get(), self.Substrate_list[0], self.Substrate_list[-1], self.var_Batch_Number.get()))

            
                    self.Batch_Folder = (r"P:\Forschungs-Projekte\OLED-measurements Bruchsal\B 2018 Lesker\B CW%d Lesker Device %d-%d Batch %d" %(self.var_CW.get(), self.Substrate_list[0], self.Substrate_list[-1], self.var_Batch_Number.get()))

                    os.chdir(self.Batch_Folder)
            
                    os.rename('OLED-xxxx-yyyy-fabrication-sheet.xlsx', 'OLED-%d-%d-fabrication-sheet.xlsx' %(self.Substrate_list[0], self.Substrate_list[-1]))
                    
                    #Put information into excel file#
            
                    self.FabForm = xlwings.Book(r"P:\Forschungs-Projekte\OLED-measurements Bruchsal\B 2018 Lesker\B CW%d Lesker Device %d-%d Batch %d\OLED-%d-%d-fabrication-sheet.xlsx" %(self.var_CW.get(), self.Substrate_list[0], self.Substrate_list[-1], self.var_Batch_Number.get(),self.Substrate_list[0], self.Substrate_list[-1]))  # Creates a connection with workbook
                    self.FabSheet = self.FabForm.sheets['fabSheet']
             
                    
                    self.FabSheet.range('D5').value = self.Batch_Info['Batch_Number']
                    self.FabSheet.range('D6').value = self.Batch_Info['Project_Name']
                    self.FabSheet.range('D10').value = self.Batch_Info['Aim']


                    if self.Substrate_Number == 4:
                        self.FabSheet.range('P23').value = self.Batch_Info['Substrate_List'][0]
                        self.FabSheet.range('P20').value = self.Batch_Info['Substrate_List'][1]
                        self.FabSheet.range('N23').value = self.Batch_Info['Substrate_List'][2]
                        self.FabSheet.range('N20').value = self.Batch_Info['Substrate_List'][3]
            
                        
                        #Microscopic Check#
                        self.FabSheet.range('D149').value = self.Batch_Info['Substrate_List'][0]
                        self.FabSheet.range('D150').value = self.Batch_Info['Substrate_List'][1]
                        self.FabSheet.range('D151').value = self.Batch_Info['Substrate_List'][2]
                        self.FabSheet.range('D152').value = self.Batch_Info['Substrate_List'][3]
            
            
            
                    if self.Substrate_Number == 6:
                        self.FabSheet.range('P23').value = self.Batch_Info['Substrate_List'][0]
                        self.FabSheet.range('P20').value = self.Batch_Info['Substrate_List'][1]
                        self.FabSheet.range('N23').value = self.Batch_Info['Substrate_List'][2]
                        self.FabSheet.range('N20').value = self.Batch_Info['Substrate_List'][3]
                        self.FabSheet.range('L23').value = self.Batch_Info['Substrate_List'][4]
                        self.FabSheet.range('L20').value = self.Batch_Info['Substrate_List'][5]
            
            
                        #Microscopic Check#
                        self.FabSheet.range('D149').value = self.Batch_Info['Substrate_List'][0]
                        self.FabSheet.range('D150').value = self.Batch_Info['Substrate_List'][1]
                        self.FabSheet.range('D151').value = self.Batch_Info['Substrate_List'][2]
                        self.FabSheet.range('D152').value = self.Batch_Info['Substrate_List'][3]
                        self.FabSheet.range('D153').value = self.Batch_Info['Substrate_List'][4]
                        self.FabSheet.range('D154').value = self.Batch_Info['Substrate_List'][5]
            
            
            
                    if self.Substrate_Number == 9:
                        self.FabSheet.range('P23').value = self.Batch_Info['Substrate_List'][0]
                        self.FabSheet.range('P20').value = self.Batch_Info['Substrate_List'][1]
                        self.FabSheet.range('P17').value = self.Batch_Info['Substrate_List'][2]
                        self.FabSheet.range('N23').value = self.Batch_Info['Substrate_List'][3]
                        self.FabSheet.range('N20').value = self.Batch_Info['Substrate_List'][4]
                        self.FabSheet.range('N17').value = self.Batch_Info['Substrate_List'][5]
                        self.FabSheet.range('L23').value = self.Batch_Info['Substrate_List'][6]
                        self.FabSheet.range('L20').value = self.Batch_Info['Substrate_List'][7]
                        self.FabSheet.range('L17').value = self.Batch_Info['Substrate_List'][8]

                        #Microscopic Check#
                        self.FabSheet.range('D149').value = self.Batch_Info['Substrate_List'][0]
                        self.FabSheet.range('D150').value = self.Batch_Info['Substrate_List'][1]
                        self.FabSheet.range('D151').value = self.Batch_Info['Substrate_List'][2]
                        self.FabSheet.range('D152').value = self.Batch_Info['Substrate_List'][3]
                        self.FabSheet.range('D153').value = self.Batch_Info['Substrate_List'][4]
                        self.FabSheet.range('D154').value = self.Batch_Info['Substrate_List'][5]
                        self.FabSheet.range('D155').value = self.Batch_Info['Substrate_List'][6]
                        self.FabSheet.range('D156').value = self.Batch_Info['Substrate_List'][7]
                        self.FabSheet.range('D157').value = self.Batch_Info['Substrate_List'][8]


                    self.Batch_Component_Index = 0
                    for L in range(15,15+self.Batch_Info['Layer_Number']+1):
                        for C in range(4,10):
                            if self.Batch_Component_Index < self.Batch_Info['Layer_Number']*6 :
                                self.FabSheet.range((L,C)).value = self.Batch_Info['Architecture'][self.Batch_Component_Index]
                                self.Batch_Component_Index = self.Batch_Component_Index + 1



                    #Evaporation step#
                    self.Evaporation_Step_count = 1
                    self.Evaporation_Step = []
                    self.Evaporation_Step_Material = []
                    self.Evaporation_Step_Variation = []
                    self.Evaporation_Step_ExpID = []
                    self.Evaporation_Step_Source = []
                    self.Evaporation_Step_Toolings =[]
                    self.Evaporation_Step_Sensor= []
                        
                    self.Cathode_Index = 15
                
                    for L in range(15,15+self.Batch_Info['Layer_Number']+1):
                        if self.FabSheet.range((L,6)).value == 'ITO':
                            self.HTL_Index = L-1

                    self.Deposited_Layer_count = self.HTL_Index
                
                    while self.Deposited_Layer_count >= self.Cathode_Index:
                        if self.FabSheet.range((self.Deposited_Layer_count,6)).value != '':
                            if ':' in self.FabSheet.range((self.Deposited_Layer_count,6)).value:
                                self.Co_Host = self.FabSheet.range((self.Deposited_Layer_count,6)).value
                                self.Co_Host = self.Co_Host.split(":")
                                self.Evaporation_Step.append(self.Evaporation_Step_count)
                                self.Evaporation_Step_Material.append(self.Co_Host[0])
                                self.Evaporation_Step.append(self.Evaporation_Step_count)
                                self.Evaporation_Step_Material.append(self.Co_Host[1])
                                self.Evaporation_Step_Variation.append(self.FabSheet.range((self.Deposited_Layer_count,4)).value)
                                self.Evaporation_Step_Variation.append(self.FabSheet.range((self.Deposited_Layer_count,4)).value)
                            else:
                                self.Evaporation_Step.append(self.Evaporation_Step_count)
                                self.Evaporation_Step_Material.append(self.FabSheet.range((self.Deposited_Layer_count,6)).value)
                                self.Evaporation_Step_Variation.append(self.FabSheet.range((self.Deposited_Layer_count,4)).value)
                    
                        if self.FabSheet.range((self.Deposited_Layer_count,7)).value != '':
                            if ':' in self.FabSheet.range((self.Deposited_Layer_count,7)).value:
                                self.Co_Emitter = self.FabSheet.range((self.Deposited_Layer_count,7)).value
                                self.Co_Emitter = self.Co_Emitter.split(":")
                                self.Evaporation_Step.append(self.Evaporation_Step_count)
                                self.Evaporation_Step_Material.append(self.Co_Emitter[0])
                                self.Evaporation_Step.append(self.Evaporation_Step_count)
                                self.Evaporation_Step_Material.append(self.Co_Emitter[1])
                                self.Evaporation_Step_Variation.append(self.FabSheet.range((self.Deposited_Layer_count,4)).value)
                                self.Evaporation_Step_Variation.append(self.FabSheet.range((self.Deposited_Layer_count,4)).value)                                
                            
                            else:    
                                self.Evaporation_Step.append(self.Evaporation_Step_count)
                                self.Evaporation_Step_Material.append(self.FabSheet.range((self.Deposited_Layer_count,7)).value)
                                self.Evaporation_Step_Variation.append(self.FabSheet.range((self.Deposited_Layer_count,4)).value)
                        
                    
                        self.Evaporation_Step_count = self.Evaporation_Step_count + 1
                        self.Deposited_Layer_count = self.Deposited_Layer_count - 1


                    for element in self.Evaporation_Step_Material:
                        if element in self.Tooling_List_Info['Lesker_Material_Name']:
                            self.Evaporation_Step_ExpID.append(self.Tooling_List_Info['Lesker_Material_Name'].index(element))
                            
                    self.Evaporation_Step_Source = [self.Tooling_List_Info['Lesker_Material_Source'][i] for i in self.Evaporation_Step_ExpID]
                    self.Evaporation_Step_Toolings = [self.Tooling_List_Info['Lesker_Material_Toolings'][i] for i in self.Evaporation_Step_ExpID]  
                    self.Evaporation_Step_Sensor = [self.Tooling_List_Info['Lesker_Material_Sensor'][i] for i in self.Evaporation_Step_ExpID]  
                    self.Evaporation_Step_ExpID = [self.Tooling_List_Info['Lesker_Material_ExpID'][i] for i in self.Evaporation_Step_ExpID]


            
                    self.Evaporation_Info = dict({'Evaporation_Step':self.Evaporation_Step, 'Evaporation_Step_Material':self.Evaporation_Step_Material, 'Evaporation_Step_Variation':self.Evaporation_Step_Variation,'Evaporation_Step_ExpID':self.Evaporation_Step_ExpID, 'Evaporation_Step_Source':self.Evaporation_Step_Source, 'Evaporation_Step_Toolings':self.Evaporation_Step_Toolings,'Evaporation_Step_Sensor':self.Evaporation_Step_Sensor})
            
            
                    


                    self.Fill_Index = 0
                    for ES in range (80, 80+len(self.Evaporation_Info['Evaporation_Step'])):
                        self.FabSheet.range((ES,1)).value = self.Evaporation_Info['Evaporation_Step'][self.Fill_Index]
                        self.FabSheet.range((ES,2)).value = self.Evaporation_Info['Evaporation_Step_Variation'][self.Fill_Index]
                        self.FabSheet.range((ES,4)).value = self.Evaporation_Info['Evaporation_Step_Material'][self.Fill_Index]
                        self.FabSheet.range((ES,5)).value = self.Evaporation_Info['Evaporation_Step_ExpID'][self.Fill_Index]
                        self.FabSheet.range((ES,8)).value = self.Evaporation_Info['Evaporation_Step_Source'][self.Fill_Index]
                        self.FabSheet.range((ES,9)).value = self.Evaporation_Info['Evaporation_Step_Toolings'][self.Fill_Index]
                        self.FabSheet.range((ES,10)).value = self.Evaporation_Info['Evaporation_Step_Sensor'][self.Fill_Index]

                        
                        self.Fill_Index = self.Fill_Index + 1  
