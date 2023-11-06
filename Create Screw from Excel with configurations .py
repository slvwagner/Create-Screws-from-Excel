#Author- florian.wagner@wagnius.ch
#Description- Erzeugen von Schrauben mit Innensechskant

import adsk.core, adsk.fusion, adsk.cam, traceback
import subprocess, sys
import math
import re #regex

def get_user_input(ui):
    # Create a dictionary to store user input
    user_input = {}

    # # Get user input for the filter criteria
    # user_input['Column1'] = (ui.inputBox('BN Nummer (BN7):'))[0]
    # user_input['Column2Min'] = (ui.inputBox('Länge min [mm]:'))[0]
    # user_input['Column2Max'] = (ui.inputBox('Länge max [mm]:'))[0]
    # user_input['Column3'] = (ui.inputBox('Gewinde (M3):'))[0]

    # Get user input for the filter criteria
    user_input['Column1'] = (ui.inputBox('BN Nummer (BN7):'))[0]
    user_input['Column2Min'] = (ui.inputBox('Länge min [mm]:'))[0]
    #user_input['Column2Max'] = (ui.inputBox('Länge max [mm]:'))[0]
    user_input['Column3'] = (ui.inputBox('Gewinde (M3):'))[0]
    
    return user_input

def run(context):
    ui = None
    try:
        ########################################################################################################
        # import packages needed for script
        ########################################################################################################
        PathToPyton = sys.path
        try:
            import numpy as np
            import pandas as pd

        # intall packages needed for script and import
        except:
            subprocess.check_call([PathToPyton[2] + '\\Python\\python.exe', "-m", "pip", "install", 'numpy'])
            subprocess.check_call([PathToPyton[2] + '\\Python\\python.exe', "-m", "pip", "install", 'pandas'])
            import numpy as np
            import pandas as pd

        ########################################################################################################
        # fusion application init
        ########################################################################################################
        
        app = adsk.core.Application.get()
        ui  = app.userInterface
        doc = app.activeDocument

        # Check that the active document has been saved. 
        if not doc.isSaved:
            ui.messageBox('The active document must be saved before running this script.')
            return

        
        ########################################################################################################
        # select file and load data 
        ########################################################################################################
        
        # fileDlg = ui.createFileDialog()
        # fileDlg.isMultiSelectEnabled = False
        # fileDlg.title = 'Fusion Open File Dialog'
        # fileDlg.filter = '*.*'
        
        # # Show file open dialog to let user deside what file needs to be processed
        # dlgResult = fileDlg.showOpen()
        # if dlgResult == adsk.core.DialogResults.DialogOK:
        #     PathToExcel = fileDlg.filename
        # else:
        #     # Screw data folder
        #     PathToExcel = "G:/Meine Ablage/020_Konstruktion/010_Normteile/Bossard/Innensechskantschrauben.xlsx"   

        # Screw data folder
        PathToExcel = "G:/Meine Ablage/020_Konstruktion/010_Normteile/Bossard/Innensechskantschrauben.xlsx"

        # get the parent folder to save 
        parentFolder = doc.dataFile.parentFolder       
        
        # Read data from excel file 
        try: 
            df = pd.read_excel(PathToExcel)
        except:
            ui.messageBox("Could not read excel:"+ PathToExcel)

        # Get data from data frame (Fusion360 base unit is cm but data is in mm)
        ScaleFactor = 10.0

        ########################################################################################################
        # Get user input for filter criteria and show items to be created to user
        ########################################################################################################
        filter_criteria = get_user_input(ui)

        # Convert the user input for Column2Min and Column2Max to numeric values
        user_column1 = filter_criteria['Column1']
        user_column2_min = float(filter_criteria['Column2Min'])
        # user_column2_max = float(filter_criteria['Column2Max'])
        user_column3 = filter_criteria['Column3']

        # Convert the user input for Column2Min and Column2Max to numeric values
        # user_column1 = "BN7"
        # user_column2_min = float(0)
        user_column2_max = float(5000)
        # user_column3 = "M1.6"

        # Modify the condition to use a regex with word boundaries
        filter_condition = (df['NormStandart'] == True) & \
                        (df['BN_nb'].str.contains(fr'\b{re.escape(user_column1)}\b', regex=True, case=False, na=False)) & \
                        (df['L_numeric'] > user_column2_min) & \
                        (df['L_numeric'] < user_column2_max) & \
                        (df['d1'].str.contains(fr'\b{re.escape(user_column3)}\b', regex=True, case=False, na=False))


        # # Define filtering conditions based on user input
        # filter_condition = (df['NormStandart'] == True) & \
        #                    (df['BN_nb'].str.contains(user_column1, case=False, na=False)) & \
        #                    (df['L_numeric'] > user_column2_min) & \
        #                    (df['L_numeric'] < user_column2_max) & \
        #                    (df['d1'].str.contains(user_column3, case=False, na=False))

        # Filter the DataFrame
        filtered_df = df[filter_condition]

        # Convert the filtered DataFrame to a string 
        filtered_df_str = filtered_df[["BN_nb", 'gewinde',"L_numeric"]].to_string(index=False)  # This converts the DataFrame to a string without the index

        # Display the filtered results in a message box
        ui.messageBox("Filtered Results:\n\n" + filtered_df_str)
        
        ########################################################################################################
        # create all screws according to filter
        ########################################################################################################

        for ii in range(0,len(filtered_df)): # select the rows to import, keep in mind that zero is an index and also the 
            # get the values of the screw
            ########################################################################################################
            d1 =    (filtered_df["d1_numeric"].iloc[ii]/2)/ScaleFactor # Gewindenormdurchmesser / Radius
            L =     (filtered_df["L_numeric"].iloc[ii])/ScaleFactor #Schraubenlänge
            t_min = (filtered_df["t_min"].iloc[ii])/ScaleFactor #Minimale Innensechskanttiefe
            d2 =    (filtered_df["d2"].iloc[ii]/2)/ScaleFactor #Schraubenkopfdurchmesser / Radius
            b =     (filtered_df["b"].iloc[ii])/ScaleFactor #Gewindelänge
            if np.isnan(b): 
                b = (filtered_df["L_numeric"].iloc[ii])/ScaleFactor #if screws with full lenght use lenght
            k =     (filtered_df["k"].iloc[ii])/ScaleFactor #Schraubenkopfhöhe
            s =     (filtered_df["s"].iloc[ii]/2)/ScaleFactor #Innensechskantschlüsselweite
            Part_NB = filtered_df["BN_nb"].iloc[ii]+ " " + filtered_df["d1"].iloc[ii] + "x" + filtered_df["L"].iloc[ii] + "," + str(filtered_df["Artikelnummer"].iloc[ii]) # Part number for fusion 360 file
            gewinde = filtered_df["gewinde"].iloc[ii] # Gewinde definition (Nennduchmesser und Steigung)
            steigung = filtered_df["steigung"].iloc[ii]/ScaleFactor # Gewinde definition (Nennduchmesser und Steigung)

            # Create document (new file in Fusion360)
            ########################################################################################################
            doc = app.documents.add(adsk.core.DocumentTypes.FusionDesignDocumentType)
            design = app.activeProduct

            # Get the root component of the active design.
            rootComp = design.rootComponent

            # Create a new sketch on the xy plane.
            sketches = rootComp.sketches;
            xyPlane = rootComp.xYConstructionPlane
            sketch = sketches.add(xyPlane)

            # create screw body 
            ########################################################################################################
            tadd = math.tan(31/180*math.pi)*t_min
            R1 = s*math.cos(30/180*math.pi)
            
            # Draw sketch for screw.
            lines = sketch.sketchCurves.sketchLines;
            line1 = lines.addByTwoPoints(adsk.core.Point3D.create(0, d1, 0), adsk.core.Point3D.create(L-(steigung/2), d1, 0))
            line11 = lines.addByTwoPoints(line1.endSketchPoint, adsk.core.Point3D.create( L,  d1-(steigung/2), 0))
            line2 = lines.addByTwoPoints(line11.endSketchPoint, adsk.core.Point3D.create( L,  0, 0))
            line30 = lines.addByTwoPoints(line2.endSketchPoint, adsk.core.Point3D.create(-(k-t_min-tadd),  0, 0))
            line31 = lines.addByTwoPoints(line30.endSketchPoint, adsk.core.Point3D.create(-(k-t_min),  s, 0))
            line32 = lines.addByTwoPoints(line31.endSketchPoint, adsk.core.Point3D.create(-(k+(R1-s)),  s, 0))
            line33 = lines.addByTwoPoints(line32.endSketchPoint, adsk.core.Point3D.create(-k,  s-(R1-s), 0))
            line4 = lines.addByTwoPoints(line33.endSketchPoint, adsk.core.Point3D.create(-k, d2, 0))
            line5 = lines.addByTwoPoints(line4.endSketchPoint, adsk.core.Point3D.create( 0, d2, 0))
            line6 = lines.addByTwoPoints(line5.endSketchPoint, adsk.core.Point3D.create( 0, d1, 0))

            # Add a fillet.
            arc = sketch.sketchCurves.sketchArcs.addFillet(line4, line4.endSketchPoint.geometry, line5, line5.startSketchPoint.geometry, steigung/2)

            # Get the first profile from the sketch
            prof = sketch.profiles.item(0)  
            
            # Get the RevolveFeatures collection.
            revolves = rootComp.features.revolveFeatures

            # Create a revolve input object that defines the input for a revolve feature.
            # When creating the input object, required settings are provided as arguments.
            revInput = revolves.createInput(prof, line30, adsk.fusion.FeatureOperations.NewBodyFeatureOperation)

            # Define a full revolve by specifying 2 pi as the revolve angle.
            angle = adsk.core.ValueInput.createByReal(math.pi * 2)
            revInput.setAngleExtent(False, angle)
                        
            # Create the revolve by calling the add method on the RevolveFeatures collection and passing it the RevolveInput object.
            screw = revolves.add(revInput)    

            # create hex drive
            #####################################################################################################
            # Get construction planes
            planes = rootComp.constructionPlanes
            # Create construction plane input
            planeInput = planes.createInput()
            
            # Add construction plane by offset
            offsetValue = adsk.core.ValueInput.createByReal(-(k-t_min))
            planeInput.setByOffset(rootComp.yZConstructionPlane, offsetValue)

            hex_plane = planes.add(planeInput)

            sketch_hex = sketches.add(hex_plane)

            s = math.pow(s,2)/R1
            sin_s = (math.sin(math.pi/3)*(s))
            cos_s = (math.cos(math.pi/3)*(s))

            # Draw 
            k_t = 0
            hex_lines = sketch_hex.sketchCurves.sketchLines
            hex_lines.addByTwoPoints(adsk.core.Point3D.create(0, s,k_t), adsk.core.Point3D.create(sin_s,cos_s,k_t))
            hex_lines.addByTwoPoints(adsk.core.Point3D.create(sin_s, cos_s,k_t), adsk.core.Point3D.create(sin_s,-cos_s,k_t))
            hex_lines.addByTwoPoints(adsk.core.Point3D.create(sin_s,-cos_s,k_t), adsk.core.Point3D.create(0,-s,k_t))
            hex_lines.addByTwoPoints(adsk.core.Point3D.create(0,-s,k_t), adsk.core.Point3D.create(-sin_s,-cos_s,k_t))
            hex_lines.addByTwoPoints(adsk.core.Point3D.create(-sin_s,-cos_s,k_t), adsk.core.Point3D.create(-sin_s, cos_s,k_t))
            hex_lines.addByTwoPoints(adsk.core.Point3D.create(-sin_s, cos_s,k_t), adsk.core.Point3D.create(0, s,k_t))

            # Get the first profile from the sketch
            prof_hex = sketch_hex.profiles.item(0)  

            # Get extrude features
            extrudes = rootComp.features.extrudeFeatures

            distance = adsk.core.ValueInput.createByReal(-k)
            hex = extrudes.addSimple(prof_hex, distance, adsk.fusion.FeatureOperations.CutFeatureOperation)

            # Gewinde feature erstellen
            #####################################################################################################

            # define all of the thread information.
            threadFeatures = rootComp.features.threadFeatures
                
            # create the thread information 
            threadInfo = threadFeatures.createThreadInfo(False, "ISO Metric profile", gewinde, "6g")

            index = 0  
            for ii in range(0,len(screw.sideFaces)):
                sideface = screw.sideFaces.item(ii)
                check_cylinder = (sideface.geometry.surfaceType == 1)
                #print("ii:",ii,sideface.geometry.objectType, sideface.geometry.surfaceType)
                if(check_cylinder): 
                    check_radius = (sideface.geometry.radius == d1)
                    #print("Radius:",sideface.geometry.radius, check_radius)
                    index = ii
                    break
            
            # get the face the thread will be applied to
            sideface = screw.sideFaces.item(index)                
                
            faces = adsk.core.ObjectCollection.create()
            faces.add(sideface)
                
            # define the thread input with the lenght of b
            threadInput = threadFeatures.createInput(faces, threadInfo)
            threadInput.isFullLength = False
            threadInput.threadLength = adsk.core.ValueInput.createByReal(b-(steigung/2))
                
            # create the final thread
            thread = threadFeatures.add(threadInput)

            # Save file 
            returnValue  = doc.saveAs(Part_NB,parentFolder ,"","")
            if not returnValue: ui.messageBox("Could save file:"+ Part_NB)
            
            # # close document
            # returnValue = doc.close(False)
            # if not returnValue: ui.messageBox("Could close the document:"+ Part_NB)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
