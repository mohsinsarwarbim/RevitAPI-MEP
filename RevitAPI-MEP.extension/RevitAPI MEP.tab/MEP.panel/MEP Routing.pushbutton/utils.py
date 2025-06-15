# System Imports
import sys
import math
from collections import defaultdict

# Revit API Imports
import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Mechanical import Duct
from Autodesk.Revit.DB.Electrical import CableTray, Conduit

# Pyrevit Imports
from rpw.ui.forms import ComboBox, FlexForm, Button
from pyrevit.forms import alert
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class Commands:
    CREATE_DUCT_NETWORK = "Create Duct Network"
    CREATE_PIPE_NETWORK = "Create Pipe Network"
    CREATE_CABLE_TRAY_NETWORK = "Create Cable Tray Network"
    CREATE_CONDUITS_NETWORK = "Create Conduits Network"
    
def group_MEPcuve_element_connectors_by_location(MEPcuve_element_connectors):
    """
    Groups MEP curve element connectors by their spatial location.

    Args:
        MEPcuve_element_connectors (List[DB.Connector]): A list of MEP curve element connectors.

    Returns:
        defaultdict: A dictionary where the keys are tuples representing
        the rounded (to 3 decimal places) X, Y, Z coordinatesof each connector's origin,
        and the values are lists of connectors at those locations.

    """
    grouped = defaultdict(list)
    for c in MEPcuve_element_connectors:
        p = c.Origin
        key = (round(p.X, 3), round(p.Y, 3), round(p.Z, 3))
        grouped[key].append(c)
    return grouped

def filter_MEPcurve_elements_using_connectors(connectors, all_MEPcurve_elements):
    """
    Filters a list of MEP curve elements to include only those that are associated with the given connectors.

    Args:
        MEPcuve_element_connectors (List[DB.Connector]): A list of MEP curve element connectors.
        all_MEPcurve_elements (List): An iterable of MEP curve elements (e.g., ducts, pipes, cable trays, conduits).

    Returns:
        list: A list of MEP curve elements whose Id matches the Owner.Id of any connector in the connectors list.
    """
    connector_owner_ids = set(c.Owner.Id for c in connectors)
    return [duct for duct in all_MEPcurve_elements if duct.Id in connector_owner_ids]

def find_shared_point_between_curves(c1, c2):
    """
    Finds a shared (coincident) endpoint between two curve objects.

    Given two curve objects c1 and c2, this function checks if any of their endpoints are
    within a small distance threshold (0.1 units) of each other, indicating they share a point.
    If such a point is found, it returns the midpoint between the two coincident endpoints.
    If no shared point is found, returns None.

    Args:
        c1: The first curve object, expected to have a GetEndPoint(index) method.
        c2: The second curve object, expected to have a GetEndPoint(index) method.

    Returns:
        The midpoint (as a point object) between the shared endpoints if found, otherwise None.
    """
    if not c1 or not c2: return None
    points1 = [c1.GetEndPoint(0), c1.GetEndPoint(1)]
    points2 = [c2.GetEndPoint(0), c2.GetEndPoint(1)]
    for p1 in points1:
        for p2 in points2:
            if p1.DistanceTo(p2) < 0.01:
                return (p1 + p2) * 0.5
    return None

def get_MEPcurve_element_direction(MEPcurve_element):
    """ Retrieves the direction of a MEP curve element.
    This function checks if the MEP curve element has a Location property of type LocationCurve.
    If so, it retrieves the curve and calculates the direction vector by subtracting the start point from the end point.
    

    Args:
        MEPcurve_element (DB.MEPcurve_element): The MEP curve element whose direction is to be determined.

    Returns:
        XYZ: A normalized direction vector of the MEP curve element, or None if the element does not have a valid curve.
    """
    curve = MEPcurve_element.Location.Curve if isinstance(MEPcurve_element.Location, LocationCurve) else None
    return (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize() if curve else None

def are_directions_parallel(v1, v2, tolerance=0.01):
    """ Checks if two direction vectors are parallel within a specified tolerance.
    This function calculates the sine of the angle between two vectors using the AngleTo method.
    If the sine of the angle is less than the specified tolerance, the vectors are considered parallel.
    This is useful for determining if two MEP curve elements are aligned in the same or opposite directions.

    Args:
        v1 (DB.XYZ): The first direction vector.
        v2 (DB.XYZ): The second direction vector.
        tolerance (float, optional): The tolerance level for determining parallelism. Defaults to 0.01.

    Returns:
        bool: True if the vectors are parallel within the specified tolerance, False otherwise.
    """
    return math.sin(v1.AngleTo(v2)) < tolerance

def MEPcurve_element_nearest_connector_to_point(MEPcurve_element, point):
    """ Finds the nearest connector to a given point on a MEP curve element.
    This function retrieves all connectors from the MEP curve element's ConnectorManager,
    sorts them based on their distance to the specified point, and returns the closest connector.
    If the MEP curve element has no connectors, it returns None.
    If there are multiple connectors at the same distance, it returns the first one found.

    Args:
        MEPcurve_element (DB.MEPcurve_element): The MEP curve element from which to find the nearest connector.
        point (DB.XYZ): The point to which the nearest connector is to be found.

    Returns:
        DB.Connector: The nearest connector to the specified point, or None if no connectors are found.
    """
    connectors = MEPcurve_element.ConnectorManager.Connectors
    return sorted(connectors, key=lambda c: c.Origin.DistanceTo(point))[0]

def create_fitting(doc, ducts):
    """ Creates a fitting for MEP curve elements based on their connectors and directions.
    This function checks the number of ducts provided and creates a fitting based on their connectors.
    It supports creating union, elbow, tee, and cross fittings depending on the number of ducts and their directions.
    It first finds a shared intersection point between the first two ducts' curves,
    then retrieves the nearest connectors to that point for each duct.
    It checks the directions of the ducts to determine if they are parallel or not,
    and creates the appropriate fitting type accordingly.

    Args:
        doc (DB.Document): The Revit document where the fitting will be created.
        ducts (List[DB.MEPcurve_element]): A list of MEP curve elements (ducts) to create a fitting for.
    """
    count = len(ducts)
    if count < 2 or count > 4:
        return

    c1 = (ducts[0].Location).Curve if isinstance(ducts[0].Location, LocationCurve) else None
    c2 = (ducts[1].Location).Curve if isinstance(ducts[1].Location, LocationCurve) else None
    intersection = find_shared_point_between_curves(c1, c2)
    if not intersection: return

    conn1 = MEPcurve_element_nearest_connector_to_point(ducts[0], intersection)
    conn2 = MEPcurve_element_nearest_connector_to_point(ducts[1], intersection)

    dir1 = get_MEPcurve_element_direction(ducts[0])
    dir2 = get_MEPcurve_element_direction(ducts[1])

    if count == 2:
        if are_directions_parallel(dir1, dir2):
            doc.Create.NewUnionFitting(conn1, conn2)
        else:
            doc.Create.NewElbowFitting(conn1, conn2)

    elif count == 3:
        duct3 = ducts[2]
        conn3 = MEPcurve_element_nearest_connector_to_point(duct3, intersection)
        dir3 = get_MEPcurve_element_direction(duct3)

        if are_directions_parallel(dir1, dir2):
            doc.Create.NewTeeFitting(conn1, conn2, conn3)
        elif are_directions_parallel(dir1, dir3):
            doc.Create.NewTeeFitting(conn3, conn1, conn2)
        else:
            doc.Create.NewTeeFitting(conn3, conn2, conn1)

    elif count == 4:
        duct3 = ducts[2]
        duct4 = ducts[3]
        conn3 = MEPcurve_element_nearest_connector_to_point(duct3, intersection)
        conn4 = MEPcurve_element_nearest_connector_to_point(duct4, intersection)

        dir3 = get_MEPcurve_element_direction(duct3)
        dir4 = get_MEPcurve_element_direction(duct4)

        if are_directions_parallel(dir1, dir2):
            doc.Create.NewCrossFitting(conn1, conn2, conn3, conn4)
        elif are_directions_parallel(dir1, dir3):
            doc.Create.NewCrossFitting(conn1, conn3, conn2, conn4)
        elif are_directions_parallel(dir1, dir4):
            doc.Create.NewCrossFitting(conn1, conn4, conn2, conn3)
       
def get_MEPcurve_elementtypes_by_category(builtin_category):
    """
    Retrieves MEP curve element types by built-in category.

    Args:
        category (BuiltInCategory): The category of MEP curve elements to filter by.

    Returns:
        dict: A dictionary mapping element names to their corresponding element types.
    """
    element_types = DB.FilteredElementCollector(doc).\
                    OfCategory(builtin_category).\
                    WhereElementIsElementType()
    names = [x.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() for x in element_types]
    return dict(zip(names, element_types))

def create_MEPcurve_element(doc,command, type_id, level_id, start, end, system_type_id=None):
    """
    Creates a MEP curve element (Duct, Pipe, Cable Tray, or Conduit) in the Revit document.

    Args:
        doc (Document): The Revit document.
        command (str): The type of MEP curve element to create.
        type_id (ElementId): The ID of the MEP curve element type.
        level_id (ElementId): The ID of the level where the element will be created.
        start (XYZ): The start point of the MEP curve element.
        end (XYZ): The end point of the MEP curve element.
        system_type_id (ElementId, optional): The ID of the system type for Ducts or Pipes.

    Returns:
        Element: The created MEP curve element.
    """
    if command == Commands.CREATE_DUCT_NETWORK:
        return Duct.Create(doc, system_type_id, type_id, level_id, start, end)
    elif command == Commands.CREATE_PIPE_NETWORK:
        return Pipe.Create(doc, system_type_id, type_id, level_id, start, end)
    elif command == Commands.CREATE_CABLE_TRAY_NETWORK:
        return CableTray.Create(doc, type_id, start, end, level_id)
    elif command == Commands.CREATE_CONDUITS_NETWORK:
        return Conduit.Create(doc, type_id, start, end, level_id)
    else:
        raise ValueError("Invalid command for creating MEP curve element.")

def get_levels_data():
    """ Retrieves a dictionary of levels in the Revit document.
    This function collects all levels in the document, extracts their names,
    and creates a dictionary mapping level names to their corresponding level elements.
    It uses the FilteredElementCollector to gather elements of the BuiltInCategory OST_Levels,
    ensuring that only non-element type levels are included.
    It is useful for populating a dropdown list in a user interface for selecting levels.
    """
    levels = DB.FilteredElementCollector(doc).\
                OfCategory(DB.BuiltInCategory.OST_Levels).\
                WhereElementIsNotElementType().\
                ToElements()
    level_names = [level.Name for level in levels]
    levelsdata = dict(zip(level_names, levels))
    return levelsdata

def flexform(commad, mep_network_types, mep_network_systems):
    """ Creates a FlexForm dialog for selecting MEP network types and systems.
    This function constructs a FlexForm with ComboBox components for selecting MEP network types,
    MEP network systems, and levels. It returns the selected type ID, system type ID, and level ID.
    The form is displayed to the user, and upon selection, the values are returned as IDs.

    Args:
        commad (Commands): The command to determine the type of MEP network to create.
        mep_network_types (List): A list of MEP network types to choose from.
        mep_network_systems (List): A list of MEP network systems to choose from.

    Returns:
        tuple: A tuple containing the selected type ID, system type ID, and level ID.
    """
    type_id = None
    mepsystem = None
    level_id = None
    
    if commad == Commands.CREATE_DUCT_NETWORK or \
         commad == Commands.CREATE_PIPE_NETWORK:
        component = [
            ComboBox('combobox1', mep_network_types),
            ComboBox('combobox2', mep_network_systems),
            ComboBox('combobox3', get_levels_data()),
            Button('Select')
        ]
    elif commad == Commands.CREATE_CABLE_TRAY_NETWORK or \
            commad == Commands.CREATE_CONDUITS_NETWORK:
                component = [
            ComboBox('combobox1', mep_network_types),
            ComboBox('combobox3', get_levels_data()),
            Button('Select')
        ]        
    form = FlexForm("Create Duct Network", component)
    form.show()

    type_id = form.values.get('combobox1', None)
    mepsystem = form.values.get('combobox2', None)
    level_id = form.values.get('combobox3', None)
    
    return (type_id.Id if type_id else None,
            mepsystem.Id if mepsystem else None,
            level_id.Id if level_id else None)