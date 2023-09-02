#Author- florian.wagner@wagnius.ch
#Description- Erzeugen von Schrauben mit Innensechskant

import adsk.core, adsk.fusion, adsk.cam, traceback
import subprocess, sys
import math

def get_user_input(ui):
    # Create a dictionary to store user input
    user_input = {}

    # Get user input for the filter criteria
    user_input['Column1'] = (ui.inputBox('BN Nummer (BN7):'))[0]
    user_input['Column2Min'] = (ui.inputBox('Länge min:'))[0]
    user_input['Column2Max'] = (ui.inputBox('Länge max:'))[0]
    user_input['Column3'] = (ui.inputBox('Gewinde (M3):'))[0]

    return user_input

def run(context):
    ui = None
    try:
        # import packages needed for script
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

        # fusion application init
        app = adsk.core.Application.get()
        ui  = app.userInterface
        doc = app.activeDocument

        # Check that the active document has been saved. 
        if not doc.isSaved:
            ui.messageBox('The active document must be saved before running this script.')
            return


        ########## select file with scew data 
        # Set styles of file dialog.
        fileDlg = ui.createFileDialog()
        fileDlg.isMultiSelectEnabled = False
        fileDlg.title = 'Fusion Open File Dialog'
        fileDlg.filter = '*.*'
        
        # Show file open dialog to let user deside what file needs to be processed
        dlgResult = fileDlg.showOpen()
        if dlgResult == adsk.core.DialogResults.DialogOK:
            PathToExcel = fileDlg.filename
        else:
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

        # Get user input for filter criteria
        filter_criteria = get_user_input(ui)

        # Convert the user input for Column2Min and Column2Max to numeric values
        user_column2_min = float(filter_criteria['Column2Min'])
        user_column2_max = float(filter_criteria['Column2Max'])

        # Define filtering conditions based on user input
        filter_condition = (df['NormStandart'] == True) & \
                           (df['BN_nb'].str.contains(filter_criteria['Column1'], case=False, na=False)) & \
                           (df['L_numeric'] > user_column2_min) & \
                           (df['L_numeric'] < user_column2_max) & \
                           (df['d1'].str.contains(filter_criteria['Column3'], case=False, na=False))

        # Filter the DataFrame
        filtered_df = df[filter_condition]

        # Convert the filtered DataFrame to a string for display
        filtered_df_str = filtered_df[["BN_nb", 'gewinde',"L_numeric"]].to_string(index=True)  # This converts the DataFrame to a string without the index

        # Display the filtered results in a message box
        ui.messageBox("Filtered Results:\n\n" + filtered_df_str)


        for ii in range(0,len(filtered_df)): # select the rows to import, keep in mind that zero is an index and also the 

            d1 =    (filtered_df["d1_numeric"].iloc[ii]/2)/ScaleFactor # Gewindenormdurchmesser
            l =     (filtered_df["L_numeric"].iloc[ii])/ScaleFactor #Schraubenlänge
            t_min = (filtered_df["t_min"].iloc[ii])/ScaleFactor #Minimale Innensechskanttiefe
            d2 =    (filtered_df["d2"].iloc[ii]/2)/ScaleFactor #Schraubenkopfdurchmesser
            b =     (filtered_df["b"].iloc[ii])/ScaleFactor #Gewindelänge
            if np.isnan(b): 
                b = (filtered_df["L_numeric"].iloc[ii])/ScaleFactor #if screws with full lenght use lenght
            k =     (filtered_df["k"].iloc[ii])/ScaleFactor #Schraubenkopfhöhe
            s =     (filtered_df["s"].iloc[ii]/2)/ScaleFactor #Innensechskantschlüsselweite
            Part_NB = filtered_df["BN_nb"].iloc[ii]+ " " + filtered_df["gewinde"].iloc[ii] + "," + str(filtered_df["Artikelnummer"].iloc[ii]) # Part number for fusion 360 file
            gewinde = filtered_df["gewinde"].iloc[ii] # Gewinde definition (Nennduchmesser und Steigung)
            steigung = filtered_df["steigung"].iloc[ii] # Gewinde definition (Nennduchmesser und Steigung)

            # Create document (new file in Fusion360)
            doc = app.documents.add(adsk.core.DocumentTypes.FusionDesignDocumentType)
            design = app.activeProduct

            # Get the root component of the active design.
            rootComp = design.rootComponent

            # Create a new sketch on the xy plane.
            sketches = rootComp.sketches;
            xyPlane = rootComp.xYConstructionPlane
            sketch = sketches.add(xyPlane)

            ########################################################################################################
            # create body 
            #####################################################################################################

            # Draw connected lines.
            lines = sketch.sketchCurves.sketchLines;
            line1 = lines.addByTwoPoints(adsk.core.Point3D.create(0, d1, 0), adsk.core.Point3D.create(l, d1, 0))
            line2 = lines.addByTwoPoints(line1.endSketchPoint, adsk.core.Point3D.create( l,  0, 0))
            line3 = lines.addByTwoPoints(line2.endSketchPoint, adsk.core.Point3D.create(-k,  0, 0))
            line4 = lines.addByTwoPoints(line3.endSketchPoint, adsk.core.Point3D.create(-k, d2, 0))
            line5 = lines.addByTwoPoints(line4.endSketchPoint, adsk.core.Point3D.create( 0, d2, 0))
            line6 = lines.addByTwoPoints(line5.endSketchPoint, adsk.core.Point3D.create( 0, d1, 0))

            # Get the first profile from the sketch
            prof = sketch.profiles.item(0)  

            # Get the RevolveFeatures collection.
            revolves = rootComp.features.revolveFeatures

            # Create a revolve input object that defines the input for a revolve feature.
            # When creating the input object, required settings are provided as arguments.
            revInput = revolves.createInput(prof, line3, adsk.fusion.FeatureOperations.NewBodyFeatureOperation)

            # Define a full revolve by specifying 2 pi as the revolve angle.
            angle = adsk.core.ValueInput.createByReal(math.pi * 2)
            revInput.setAngleExtent(False, angle)
                    
            # Create the revolve by calling the add method on the RevolveFeatures collection and passing it the RevolveInput object.
            screw = revolves.add(revInput)           

            #####################################################################################################
            # create hex drive
            #####################################################################################################
            hex_plane = rootComp.xYConstructionPlane
            sketch_hex = sketches.add(hex_plane)

            sin_s = math.sin(math.pi/3)*s
            cos_s = math.cos(math.pi/3)*s

            # Draw 
            k_t = -(k-t_min)
            hex_lines = sketch_hex.sketchCurves.sketchLines
            hLine1 = hex_lines.addByTwoPoints(adsk.core.Point3D.create(k_t, 0, s), adsk.core.Point3D.create(k_t,sin_s,cos_s))
            hLine2 = hex_lines.addByTwoPoints(adsk.core.Point3D.create(k_t, sin_s, cos_s), adsk.core.Point3D.create(k_t,sin_s,-cos_s))
            hLine3 = hex_lines.addByTwoPoints(adsk.core.Point3D.create(k_t,sin_s,-cos_s), adsk.core.Point3D.create(k_t,0,-s))
            hLine3 = hex_lines.addByTwoPoints(adsk.core.Point3D.create(k_t,0,-s), adsk.core.Point3D.create(k_t,-sin_s,-cos_s))
            hLine4 = hex_lines.addByTwoPoints(adsk.core.Point3D.create(k_t,-sin_s,-cos_s), adsk.core.Point3D.create(k_t,-sin_s, cos_s))
            hLine5 = hex_lines.addByTwoPoints(adsk.core.Point3D.create(k_t,-sin_s, cos_s), adsk.core.Point3D.create(k_t,0, s))

            # Get the first profile from the sketch
            prof_hex = sketch_hex.profiles.item(0)  

            # Get extrude features
            extrudes = rootComp.features.extrudeFeatures

            distance = adsk.core.ValueInput.createByReal(-k)
            hex = extrudes.addSimple(prof_hex, distance, adsk.fusion.FeatureOperations.CutFeatureOperation)
            
            #####################################################################################################
            # Gewinde feature erstellen
            #####################################################################################################
            # define all of the thread information.
            threadFeatures = rootComp.features.threadFeatures
            
            # query the thread table to get the thread information
            threadDataQuery = threadFeatures.threadDataQuery
            threadTypes = threadDataQuery.allThreadTypes
            threadType = threadTypes[10]
            
            # Gewinde Durchmesser und Steigung
            allDesignations = threadDataQuery.allDesignations(threadType, str(k*10))
            threadDesignation = allDesignations[0]
            
            # Toleranzklasse
            allClasses = threadDataQuery.allClasses(False, threadType, gewinde)
            threadClass = allClasses[1]
            
            # create the thread information 
            threadInfo = threadFeatures.createThreadInfo(False, threadType, gewinde, threadClass)
            
            # get the face the thread will be applied to
            sideface = screw.sideFaces.item(1)
            faces = adsk.core.ObjectCollection.create()
            faces.add(sideface)
            
            # define the thread input with the lenght of b
            threadInput = threadFeatures.createInput(faces, threadInfo)
            threadInput.isFullLength = False
            threadInput.threadLength = adsk.core.ValueInput.createByReal(b)
            
            # create the final thread
            thread = threadFeatures.add(threadInput)

            # Save file 
            returnValue  = doc.saveAs(Part_NB,parentFolder ,"","")

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
