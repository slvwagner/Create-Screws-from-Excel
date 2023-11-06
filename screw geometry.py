# Author-
# Description-
import math

import adsk.core, adsk.fusion, adsk.cam, traceback


def r(x, n):
    return math.ceil(x / n) * n

def get_top_face(component):
    body: adsk.fusion.BRepBody = component.bRepBodies[0]
    faces: adsk.fusion.BRepFaces = body.faces
    return max(faces, key=(lambda f: f.centroid.z))

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        
        #############################################################################################
        # Get the material libraries
        #############################################################################################
        material_libraries = app.materialLibraries

        # Create an empty list to store the different libraries
        material_libs = []

        for ii in material_libraries:
            material_libs.append(ii.name)
        print("\n***********************************************")
        for ii in material_libs: print(ii)

        # Create an empty list to store the different libraries
        all_materials = []

        # Iterate through the material libraries and gather all materials in the list
        for ii in material_libraries:
            for jj in ii.materials:
                all_materials.append(jj.name)

        # Get a reference to your user library named "wagnius"
        wagnius = app.materialLibraries.itemByName(material_libs[len(material_libs)-1])

        print("\n***********************************************")

        # Check if the library exists
        if wagnius:
            # Iterate through the materials in the "wagnius" library and list their names
            for ii in wagnius.materials:
                print(f"Material Name: {ii.name}")
        else:
            print("The 'wagnius' library was not found.")

        pcb = wagnius.materials.itemByName('1.0037')
        steel = wagnius.materials.itemByName('1.4301')
    

        # get the values of the screw
        ########################################################################################################
        ScaleFactor = 10
        d1 =    (3/2)/ScaleFactor # Gewindenormdurchmesser / Radius
        L =     (25)/ScaleFactor #Schraubenlänge
        t_min = (2.3)/ScaleFactor #Minimale Innensechskanttiefe
        d2 =    (6/2)/ScaleFactor #Schraubenkopfdurchmesser / Radius
        b =     (20)/ScaleFactor #Gewindelänge
        # if np.isnan(b): 
        #     b = (15)/ScaleFactor #if screws with full lenght use lenght
        k =     (4)/ScaleFactor #Schraubenkopfhöhe
        s =     (2.5/2)/ScaleFactor #Innensechskantschlüsselweite
        Part_NB = "BN4 M3x25, 114555222"# Part number for fusion 360 file

        gewinde = "M3x0.5" # Gewinde definition (Nennduchmesser und Steigung)
        steigung = (0.5)/ScaleFactor # Gewinde definition (Nennduchmesser und Steigung)

        # Create document (new file in Fusion360)
        ########################################################################################################
        doc = app.documents.add(adsk.core.DocumentTypes.FusionDesignDocumentType)
        design = app.activeProduct

        # Get the root component of the active design.
        rootComp = design.rootComponent


        ########################################################################################################
        # create screw body 
        ########################################################################################################
        # Create a new sketch on the xy plane.
        sketches = rootComp.sketches;
        xyPlane = rootComp.xYConstructionPlane
        sketch = sketches.add(xyPlane)
        sketch.name = "Body"

        R1 = s*math.cos(30/180*math.pi)
        tadd = math.cos((118/2)/180*math.pi)*R1

        # Draw sketch for screw.
        lines = sketch.sketchCurves.sketchLines;
        line1 = lines.addByTwoPoints(adsk.core.Point3D.create(0, d1, 0), adsk.core.Point3D.create(L-(steigung/2), d1, 0))
        line11 = lines.addByTwoPoints(line1.endSketchPoint, adsk.core.Point3D.create( L,  d1-(steigung/2), 0))
        line2 = lines.addByTwoPoints(line11.endSketchPoint, adsk.core.Point3D.create( L,  0, 0))
        line30 = lines.addByTwoPoints(line2.endSketchPoint, adsk.core.Point3D.create(-(k-(t_min+tadd)),  0, 0))
        line31 = lines.addByTwoPoints(line30.endSketchPoint, adsk.core.Point3D.create(-(k-t_min),  s, 0))
        line32 = lines.addByTwoPoints(line31.endSketchPoint, adsk.core.Point3D.create(-(k+(R1-s)),  s, 0))
        line33 = lines.addByTwoPoints(line32.endSketchPoint, adsk.core.Point3D.create(-k,  s-(R1-s), 0))
        line4 = lines.addByTwoPoints(line33.endSketchPoint, adsk.core.Point3D.create(-k, d2, 0))
        line5 = lines.addByTwoPoints(line4.endSketchPoint, adsk.core.Point3D.create( 0, d2, 0))
        line6 = lines.addByTwoPoints(line5.endSketchPoint, adsk.core.Point3D.create( 0, d1, 0))
        
        # Add a fillet.
        arc = sketch.sketchCurves.sketchArcs.addFillet(line4, line4.endSketchPoint.geometry, line5, line5.startSketchPoint.geometry, steigung/2)

        # Add dimensions
        #####################################################################################################
        sketch.sketchDimensions.addDistanceDimension(line2.endSketchPoint, sketch.originPoint, 1, adsk.core.Point3D.create(L/2, d2*2, 0))
        sketch.sketchDimensions.addDistanceDimension(line4.endSketchPoint, sketch.originPoint, 1, adsk.core.Point3D.create(-k/2, d2*2, 0))
        sketch.sketchDimensions.addDistanceDimension(line31.endSketchPoint, sketch.originPoint, 1, adsk.core.Point3D.create( -k/3, -d1/2, 0))
        
        sketch.sketchDimensions.addDistanceDimension(line5.endSketchPoint, sketch.originPoint, 2, adsk.core.Point3D.create(-k*1.5, d1/2, 0))
        sketch.sketchDimensions.addDistanceDimension(line11.endSketchPoint, sketch.originPoint, 2, adsk.core.Point3D.create( L*1.2, d1/2, 0))
        sketch.sketchDimensions.addDistanceDimension(line33.endSketchPoint, sketch.originPoint, 2, adsk.core.Point3D.create( -k*1.2, -d1, 0))
        
        sketch.sketchDimensions.addDistanceBetweenLineAndPlanarSurfaceDimension(line1, rootComp.xZConstructionPlane)
        sketch.sketchDimensions.addDistanceBetweenLineAndPlanarSurfaceDimension(line32, rootComp.xZConstructionPlane)

        sketch.sketchDimensions.addAngularDimension(line1, line11, adsk.core.Point3D.create(L, 0, 0))
        sketch.sketchDimensions.addAngularDimension(line31, line32, adsk.core.Point3D.create(-(k-t_min), R1, 0))
        sketch.sketchDimensions.addAngularDimension(line33, line4, adsk.core.Point3D.create(-k*0.8, d2/2, 0))

        sketch.sketchDimensions.addDiameterDimension(arc, adsk.core.Point3D.create(-k*1.2, d2*1.2, 0))

        # Add Constraints 
        #####################################################################################################
        sketch.geometricConstraints.addHorizontal(lines.item(0))
        sketch.geometricConstraints.addVertical(lines.item(2))
        sketch.geometricConstraints.addHorizontal(lines.item(3))
        sketch.geometricConstraints.addHorizontal(lines.item(5))
        sketch.geometricConstraints.addVertical(lines.item(7))
        sketch.geometricConstraints.addHorizontal(lines.item(8))
        sketch.geometricConstraints.addVertical(lines.item(9))

        # Add a coincident constraint between points 
        sketch.geometricConstraints.addCoincident(line6.endSketchPoint, line1.startSketchPoint)
        sketch.geometricConstraints.addCoincident(sketch.originPoint, lines.item(3))
        sketch.geometricConstraints.addCoincident(sketch.originPoint, lines.item(9))

        # resolfe the profile
        #####################################################################################################
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

        #####################################################################################################
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
        sketch_hex.name = "hex drive"

        s = math.pow(s,2)/R1
        sin_s = (math.sin(math.pi/3)*(s))
        cos_s = (math.cos(math.pi/3)*(s)) 

        # Draw polygon
        #####################################################################################################
        # Create the 3D points for the hex sketch
        p1 = sketch_hex.sketchPoints.add(adsk.core.Point3D.create(0, s,0))
        p2 = sketch_hex.sketchPoints.add(adsk.core.Point3D.create(sin_s,cos_s,0))
        p3 = sketch_hex.sketchPoints.add(adsk.core.Point3D.create(sin_s,-cos_s,0))
        p4 = sketch_hex.sketchPoints.add(adsk.core.Point3D.create(0,-s,0))
        p5 = sketch_hex.sketchPoints.add(adsk.core.Point3D.create(-sin_s,-cos_s,0))
        p6 = sketch_hex.sketchPoints.add(adsk.core.Point3D.create(-sin_s, cos_s,0))

        # Create lines 
        hex_lines = sketch_hex.sketchCurves.sketchLines
        l1 = hex_lines.addByTwoPoints(p1, p2)
        l2 = hex_lines.addByTwoPoints(p2, p3)
        l3 = hex_lines.addByTwoPoints(p3, p4)
        l4 = hex_lines.addByTwoPoints(p4, p5)
        l5 = hex_lines.addByTwoPoints(p5, p6)
        l6 = hex_lines.addByTwoPoints(p6, p1)

        # Add dimensions
        ######################################################################################################
        sketch_hex.sketchDimensions.addDistanceDimension(p1, sketch_hex.originPoint, 2, adsk.core.Point3D.create(d2, 2*s/4, 0))
        sketch_hex.sketchDimensions.addDistanceDimension(p2, sketch_hex.originPoint, 2, adsk.core.Point3D.create(d2*2/3, 1*s/4, 0))
        sketch_hex.sketchDimensions.addDistanceDimension(p3, sketch_hex.originPoint, 2, adsk.core.Point3D.create(d2*2/3, -1*s/4, 0))
        sketch_hex.sketchDimensions.addDistanceDimension(p4, sketch_hex.originPoint, 2, adsk.core.Point3D.create(d2, -2*s/4, 0))
        
        sketch_hex.sketchDimensions.addDistanceBetweenLineAndPlanarSurfaceDimension(l2, rootComp.xYConstructionPlane)
        sketch_hex.sketchDimensions.addDistanceBetweenLineAndPlanarSurfaceDimension(l5, rootComp.xYConstructionPlane)

        # Add constraints 
        #####################################################################################################
        sketch_hex.geometricConstraints.addLineOnPlanarSurface(l1, hex_plane)
        sketch_hex.geometricConstraints.addLineOnPlanarSurface(l3, hex_plane)
        sketch_hex.geometricConstraints.addLineOnPlanarSurface(l5, hex_plane)


        sketch_hex.sketchDimensions.addAngularDimension(l1, l2, adsk.core.Point3D.create(0, 0, 0))
        sketch_hex.sketchDimensions.addAngularDimension(l2, l3, adsk.core.Point3D.create(0, 0, 0))
        sketch_hex.sketchDimensions.addAngularDimension(l3, l4, adsk.core.Point3D.create(0, 0, 0))
        sketch_hex.sketchDimensions.addAngularDimension(l5, l6, adsk.core.Point3D.create(0, 0, 0))

        # Get the first profile from the sketch
        prof_hex = sketch_hex.profiles.item(0)  

        # Cut feature
        #####################################################################################################
        extrudes = rootComp.features.extrudeFeatures
        hex = extrudes.addSimple(prof_hex, adsk.core.ValueInput.createByReal(-k), adsk.fusion.FeatureOperations.CutFeatureOperation)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
 


