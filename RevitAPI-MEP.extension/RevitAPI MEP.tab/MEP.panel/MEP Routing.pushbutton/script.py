__author__ = "Mohsin Sarwar"
__DOC__ = """
Create A MEP Network in Revit by selected revit line"""

# System Imports
import traceback
import sys


# Revit API Imports
import Autodesk.Revit.DB as DB

# Pyrevit Imports
from pyrevit.forms import alert
from pyrevit.forms import CommandSwitchWindow

# Custom Imports
from utils import Commands
from utils import (get_MEPcurve_elementtypes_by_category,
                   create_MEPcurve_element,
                   group_MEPcuve_element_connectors_by_location,
                   filter_MEPcurve_elements_using_connectors,
                   create_fitting,
                   flexform,
                   )

DOC = __revit__.ActiveUIDocument.Document
UIDOC = __revit__.ActiveUIDocument

# Get the selected model lines
selection_ids = UIDOC.Selection.GetElementIds()
if not selection_ids:
    alert("Please select a line to create a ducts network.")
    sys.exit()
selected_model_lines = []
for sel_id in selection_ids:
    element = DOC.GetElement(sel_id)
    if isinstance(element, DB.ModelLine):
        selected_model_lines.append(element)
    else:
        alert("Please select only model lines to create a ducts network.")
        sys.exit()

# Ask user to select a command to create MEP network
MEP_NETWORKS = [
    Commands.CREATE_DUCT_NETWORK,
    Commands.CREATE_PIPE_NETWORK,
    Commands.CREATE_CABLE_TRAY_NETWORK,
    Commands.CREATE_CONDUITS_NETWORK
]
PICKED_COMMAND = CommandSwitchWindow.show(MEP_NETWORKS, message='Select Option')
if not PICKED_COMMAND:
    alert("Please select a MEP network type to create.")
    sys.exit()

# Get MEP network types and systems based on the selected command
mep_network_types = None
mep_network_systems = None
if PICKED_COMMAND == Commands.CREATE_DUCT_NETWORK:
    mep_network_types = get_MEPcurve_elementtypes_by_category(DB.BuiltInCategory.OST_DuctCurves)
    mep_network_systems = get_MEPcurve_elementtypes_by_category(DB.BuiltInCategory.OST_DuctSystem)

elif PICKED_COMMAND == Commands.CREATE_PIPE_NETWORK:
    mep_network_types = get_MEPcurve_elementtypes_by_category(DB.BuiltInCategory.OST_PipeCurves)
    mep_network_systems = get_MEPcurve_elementtypes_by_category(DB.BuiltInCategory.OST_PipingSystem)

elif PICKED_COMMAND == Commands.CREATE_CABLE_TRAY_NETWORK:
    mep_network_types = get_MEPcurve_elementtypes_by_category(DB.BuiltInCategory.OST_CableTray)
    

elif PICKED_COMMAND == Commands.CREATE_CONDUITS_NETWORK:
    mep_network_types = get_MEPcurve_elementtypes_by_category(DB.BuiltInCategory.OST_Conduit)

# Show a flexform to get the MEP network type, system and level
# If the user cancels the operation, exit the script
try:
    flexform_data = flexform(PICKED_COMMAND, mep_network_types, mep_network_systems)
except:
    alert("User cancelled the operation.")
    sys.exit()
if not any(flexform_data):
    alert("Please select all required fields to create a MEP network.")
    sys.exit()
    
# Unpack the flexform data
mep_network_type_id, mep_network_system_id, level_id = flexform_data

# Start creating the MEP network
mep_elements = [] # List to store created MEP elements
mep_elements_connectors = [] # List to store all connectors from created MEP elements

# Start a transaction group to create the PICKED_COMMAND
tg = DB.TransactionGroup(DOC, "{}".format(PICKED_COMMAND))
tg.Start()

t = DB.Transaction(DOC, "{}".format(PICKED_COMMAND)) # Transaction to create MEP network
t.Start()

for mline in selected_model_lines:
    curve = mline.GeometryCurve
    start = curve.GetEndPoint(0)
    end = curve.GetEndPoint(1)
    duct = create_MEPcurve_element(DOC, PICKED_COMMAND,
                                mep_network_type_id,
                                level_id,
                                start,
                                end,
                                system_type_id=mep_network_system_id)
    mep_elements.append(duct)
    mep_elements_connectors.extend([c for c in duct.ConnectorManager.Connectors])
t.Commit()

connector_groups = group_MEPcuve_element_connectors_by_location(mep_elements_connectors)

t = DB.Transaction(DOC, "{} Fittings".format(PICKED_COMMAND)) # Transaction to create fittings
t.Start()
for group in connector_groups.values():
    try:
        connected_ducts = filter_MEPcurve_elements_using_connectors(group, mep_elements)
        create_fitting(DOC, connected_ducts)
    except:
        print("Error\n")
        print(traceback.format_exc())
        print("\n")
        
t.Commit()

tg.Assimilate()

    








