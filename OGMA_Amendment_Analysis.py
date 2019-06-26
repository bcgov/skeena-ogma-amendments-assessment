'''

	OGMA Amendment Analysis	
	Created By: Jesse Fraser
	May 9th 2018
	
	Goal:
	
	An excel spreadsheet is your output - Not currently created
	Current output - GDB with tables and shapefiles that access the change in OGMAs from current ot new
'''

import sys, string, os, time, win32com.client, datetime, win32api, arcpy, arcpy.mapping , csv
#import  wml_library_arcpy_v3 as wml_library
from arcpy import env

arcpy.env.overwriteOutput = True

try:
    arcpy.CheckOutExtension("Spatial")
    from arcpy.sa import *
    from arcpy.da import *
except:
    arcpy.AddError("Spatial Extension could not be checked out")
    os.sys.exit(0)

print "Make sure that all your OGMAs proposed and current or a single feature"
#Original OGMA shape
Orig_OGMA = arcpy.GetParameterAsText(0)

#New OGMA Shape
New_OGMA = arcpy.GetParameterAsText(1)

#Project Name (Preferably OGMA ID) that you are analysing
proj_Name = arcpy.GetParameterAsText(2)

#Save location
save = arcpy.GetParameterAsText(3)

#Unique Field taken from Proposed OGMA feature (1)
og_ID = arcpy.GetParameterAsText(4)

#Location of BCGW w/Password embedded... You need to have a database called BCGW4Scripting.sde
BCGW = r'Database Connections\BCGW4Scripting.sde'

#VRI Path
VRI = 'WHSE_FOREST_VEGETATION.VEG_COMP_LYR_L1_POLY'
#LU Path
LU = 'WHSE_LAND_USE_PLANNING.RMP_LANDSCAPE_UNIT_SVW'
gdbname = proj_Name + "_OGMA_Analysis_" + time.strftime("%Y%m%d")
arcpy.CreateFileGDB_management(save, gdbname)
saveloc = save + '\\' + gdbname + '.gdb'

arcpy.env.overwriteOutput = True

''' Need to be rethought or figured out
#Create A spreadsheet? (From One Status thank you Mark McGirr)
def create_spreadsheet(self, output):
        arcpy.AddWarning("======================================================================")
        arcpy.AddWarning("Creating Spreadsheet, please wait...")
        arcpy.AddWarning("======================================================================")
    
        arcpy.AddMessage("Creating Spreadsheet, please wait...")
        ExcelApp=win32com.client.Dispatch("Excel.Application")
        ExcelApp.Visible = 0
        #ExcelApp.Visible = 1
         
        self.workbook = ExcelApp.Workbooks.Add()
        self.current_worksheet_number = 1 # changed to 1 as this seems to be where failing.
        self.set_xls_column_widths(self.current_worksheet_number)
         
          
        xls_to_save = os.path.join(saveloc+ "Analysis.csv")
        ExcelApp.DisplayAlerts = 0
        #self.this_workbook.rows.AutoFit()
        self.this_workbook.SaveAs(xls_to_save)
        ExcelApp.DisplayAlerts = 1
        ExcelApp.Quit()

'''		
		
#Removes the need to know if the area is shape or geometry (Thank you Carol Mahood)
def shape_v_geo (fc):
	desc = arcpy.Describe(fc)
	geomField = desc.shapeFieldName
	areaFieldName = str(geomField) + "_Area"

def tableToCSV(input_tbl, csv_filepath):
    fld_list = arcpy.ListFields(input_tbl)
    fld_names = [fld.name for fld in fld_list]
    with open(csv_filepath, 'wb') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(fld_names)
        with arcpy.da.SearchCursor(input_tbl, fld_names) as cursor:
            for row in cursor:
                writer.writerow(row)
				
	print csv_filepath + " CREATED"
    csv_file.close()
		
def CreateOGMA_AnalysisData(New, Current):
	#Data that will be pulled from BCGW
	Input_LU = os.path.join(BCGW,LU)
	Input_VRI = os.path.join(BCGW,VRI)
	
	
	
	''' Creating the BEC Variant Test '''
	
	#Variables for the clip of the LU and VRI to the respective OGMA
	new_clip_VRI = saveloc + '\\' + proj_Name + '_VRI_new_' + time.strftime("%Y%m%d")
	current_clip_VRI = saveloc + '\\' + proj_Name + '_VRI_current_' + time.strftime("%Y%m%d")
	new_clip_LU = saveloc + '\\' + proj_Name + '_LU_new_' + time.strftime("%Y%m%d")
	current_clip_LU = saveloc + '\\' + proj_Name + '_LU_current_' + time.strftime("%Y%m%d")
	#Clip the data	
	
	arcpy.Clip_analysis(Input_VRI, New, new_clip_VRI)
	arcpy.Clip_analysis(Input_LU, New, new_clip_LU)
	print "Proposed OGMA Data Created"
	
	arcpy.Clip_analysis(Input_VRI, Current, current_clip_VRI)
	arcpy.Clip_analysis(Input_LU, Current, current_clip_LU)	
	print "Current OGMA Data Created"
	
	
	
	#Union output variable
	new_union =  saveloc + '\\' + proj_Name + '_VRI_Union_new_' + time.strftime("%Y%m%d")
	current_union =  saveloc + '\\' + proj_Name + '_VRI_Union_current_' + time.strftime("%Y%m%d")
	
	
	#Union the LU and VRI clipped layers
	arcpy.Union_analysis([new_clip_LU, new_clip_VRI, New_OGMA], new_union)
	arcpy.Union_analysis([current_clip_LU, current_clip_VRI, Orig_OGMA], current_union)
	
	
		# I don't think this is necessary
	'''
	#get the areafield name to avoid geometry vs shape issue (Thanks you Carol Mahood)
	shape_v_geo(new_union)
	new_area = areaFieldName
	shape_v_geo(current_union)
	current_area = areaFieldName
	'''
	
	#Add a field and calculate geometry of new or current area
	new_area = arcpy.AddGeometryAttributes_management(new_union, "AREA", "", "HECTARES")
	current_area = arcpy.AddGeometryAttributes_management(current_union, "AREA", "", "HECTARES")
	
	
	
	#Frequency variable 
	new_freq = saveloc + '\\' + proj_Name + '_LU_VRI_freq_new_' + time.strftime("%Y%m%d")
	current_freq = saveloc + '\\' + proj_Name + '_LU_VRI_freq_current_' + time.strftime("%Y%m%d")

	#Frequency Test fields (BEC Zone/Subzone/Variant and LU)
	fields = ["BEC_ZONE_CODE", "BEC_SUBZONE", 'BEC_VARIANT', "LANDSCAPE_UNIT_NAME"]
	
	#Frequency analysis BEC Zone/Subzone/Variant by LU
	arcpy.Frequency_analysis(new_union, new_freq, fields, "POLY_AREA")
	arcpy.Frequency_analysis(current_union, current_freq, fields, "POLY_AREA")
	
	#Frequency analysis BEC Zone/Subzone/Variant by OGMA ID
	new_freq = saveloc + '\\' + proj_Name + '_VRI_freq_new_' + time.strftime("%Y%m%d")
	current_freq = saveloc + '\\' + proj_Name + '_VRI_freq_current_' + time.strftime("%Y%m%d")
	
	fields = ["BEC_ZONE_CODE", "BEC_SUBZONE", 'BEC_VARIANT', og_ID]
	arcpy.Frequency_analysis(new_union, new_freq, fields, "POLY_AREA")
	arcpy.Frequency_analysis(current_union, current_freq, fields, "POLY_AREA")
	
	
	print "Done with "
	
	'''
	Need to figure out a better way
	
	-Currently just merging which puts all the features into one tableToCSV
	-One option is the create a new 'join' column, which will combine all the fields that we want to combine together
	then delete all the fields we don't need and export to a excel sheet
		
	#Merge the two tables into single output tableToCSV
	arcpy.Merger([new_area,current_area],final_freq)
	
	
	
	#Need to add the Frequency to the spreadsheet that is created and 
	tableToCSV(freq, [saveloc+ "Analysis.csv"])
	'''
	
	'''	End '''
	
	''' Testing Interior Conditions '''
	
	#Create the output for buffer
	new_buffer = saveloc + '\\' + proj_Name + '_Interior_new_' + time.strftime("%Y%m%d")
	current_buffer = saveloc + '\\' + proj_Name + '_Interior_current_' + time.strftime("%Y%m%d")
	
	LU_new_buffer = saveloc + '\\' + proj_Name + '_LU_Interior_new_' + time.strftime("%Y%m%d")
	LU_current_buffer = saveloc + '\\' + proj_Name + '_LU_Interior_current_' + time.strftime("%Y%m%d")
	#Buffer the OGMA to see how much is interior
	arcpy.Buffer_analysis(New, new_buffer, "-100 Meters")
	arcpy.Buffer_analysis(Current, current_buffer, "-100 Meters")
	
	arcpy.Union_analysis([new_clip_LU, new_buffer], LU_new_buffer)
	arcpy.Union_analysis([current_clip_LU, current_buffer], LU_current_buffer)
	
	#Get the shape area name
	new_area = arcpy.AddGeometryAttributes_management(LU_new_buffer, "AREA", "", "HECTARES")
	current_area = arcpy.AddGeometryAttributes_management(LU_current_buffer, "AREA", "", "HECTARES")
	
	new_freq = saveloc + '\\' + proj_Name + '_Interior_new_table_' + time.strftime("%Y%m%d")
	current_freq = saveloc + '\\' + proj_Name + '_Interior_current_table_' + time.strftime("%Y%m%d")
	
	LU_new_freq = saveloc + '\\' + proj_Name + '_LU_Interior_new_table_' + time.strftime("%Y%m%d")
	LU_current_freq = saveloc + '\\' + proj_Name + '_LU_Interior_current_table_' + time.strftime("%Y%m%d")
	
	#Interior Forest by OGMA ID
	fields = og_ID	
	
	arcpy.Frequency_analysis(LU_current_buffer, current_freq, fields, "POLY_AREA")

	arcpy.Frequency_analysis(LU_new_buffer, new_freq, fields, "POLY_AREA")


	#Interior Forest by Landscape Unit
	fields = ["BUFF_DIST","LANDSCAPE_UNIT_NAME"]
	
	arcpy.Frequency_analysis(LU_current_buffer, LU_current_freq, fields, "POLY_AREA")

	arcpy.Frequency_analysis(LU_new_buffer, LU_new_freq, fields, "POLY_AREA")
	
	print 'Done with Interior Forest'
	
	#Frequency Test fields (Age and LU)
	fields = ["PROJ_AGE_CLASS_CD_1", "LANDSCAPE_UNIT_NAME"]
	
	#Frequency analysis Age by LU
	new_freq = saveloc + '\\' + proj_Name + '_LU_Age_freq_new_' + time.strftime("%Y%m%d")
	current_freq = saveloc + '\\' + proj_Name + '_LU_Age_freq_current_' + time.strftime("%Y%m%d")
		
	#Frequency Table creation
	arcpy.Frequency_analysis(new_union, new_freq, fields, "POLY_AREA")
	arcpy.Frequency_analysis(current_union, current_freq, fields, "POLY_AREA")
	
	
	
	#Frequency Test fields (Age and OGMA ID)
	fields = ["PROJ_AGE_CLASS_CD_1", og_ID]
	
	#Frequency analysis Age by OGMA ID
	new_freq = saveloc + '\\' + proj_Name + '_Age_freq_new_' + time.strftime("%Y%m%d")
	current_freq = saveloc + '\\' + proj_Name + '_Age_freq_current_' + time.strftime("%Y%m%d")
		
	#Frequency Table creation
	arcpy.Frequency_analysis(new_union, new_freq, fields, "POLY_AREA")
	arcpy.Frequency_analysis(current_union, current_freq, fields, "POLY_AREA")
	
	print "Done testing Age Class"
	
	''' End '''
	
	''' Testing Size Conditions 
	new_freq = saveloc + proj_Name + '_OGMA_freq_new_' + time.strftime("%Y%m%d")
	current_freq = saveloc + proj_Name + '_OGMA_freq_current_' + time.strftime("%Y%m%d")
	arcpy.Frequency_analysis(Current, new_freq, fields, new_area)
	arcpy.Frequency_analysis(New, current_freq, fields, current_area)
	
	End ''' 
	
#Run the definitions on each of the inputs

#See the above statement about spreadsheet
#self.create_spreadsheet()

CreateOGMA_AnalysisData(New_OGMA, Orig_OGMA)


#Figure out a way to export the frequency table values into a single spreadsheet comparing the current vs proposed