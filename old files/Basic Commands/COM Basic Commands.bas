'declaration of variables
Public Vissim As Object            ' Variable Vissim

Public Sub COM_Basic_Commands()

'==========================================================================
' VBA-Script for Vissim 6+
' Copyright (C) PTV AG, Jochen Lohmiller
' All rights reserved.
' -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -
' Example of basic syntax - DO NOT MODIFY
'==========================================================================

' This script demonstrates how to use the COM interface in VBA.
' Basic commands for loading a network and layout, reading and setting
' attributes of network objects, running a simulation and retrieving
' evaluations are shown. This example is also available for the programming
' languages VBS, Python, Matlab, C#, C++ and Java.
'
' If you start using COM, please see also our COM introduction document:
' C:\Program Files\PTV Vision\PTV Vissim 10\Doc\Eng\Vissim 10 - COM Intro.pdf
'
' For information about the attributes and methods of PTV Vissim objects, see
' the COM Help, which is located in the PTV Vissim menu: Help > COM Help.
'
' Hint: You can easily see all attributes of the PTV Vissim objects in the lists in
' Vissim. In case your PTV Vissim language is set to English, the name in the
' headline of the list correspond to the command to access via COM.
' Example: If you want to access the number of lanes of a link, go to PTV Vissim
' and see the headline of "Number of lanes" in the list, which is
' "NumLanes". To access this attribute via COM use:
' Vissim.Net.Links.ItemByKey(1).AttValue("NumLanes")

' Connecting the COM Server => Opens a new Vissim Window:
Set Vissim = CreateObject("Vissim.Vissim")
' If you have installed multiple Vissim Versions, you can open a specific Vissim version adding the bit Version (32 or 64bit) and Version number
' Set Vissim = CreateObject("Vissim.Vissim-32.9") ' Vissim 9 - 32 bit
' Set Vissim = CreateObject("Vissim.Vissim-64.9") ' Vissim 9 - 64 bit
' Set Vissim = CreateObject("Vissim.Vissim-32.10") ' Vissim 10 - 32 bit
' Set Vissim = CreateObject("Vissim.Vissim-64.10") ' Vissim 10 - 64 bit
Dim Path_of_COM_Basic_Commands_network As String
Dim Filename                    As String
Dim flag_read_additionally      As Boolean

Path_of_COM_Basic_Commands_network = "C:\Users\Public\Documents\PTV Vision\PTV Vissim 10\Examples Training\COM\Basic Commands"

' Load a Vissim Network:
Filename = Path_of_COM_Basic_Commands_network & "\COM Basic Commands.inpx"
flag_read_additionally = False   ' you can read network(elements) additionally, in this case set "flag_read_additionally" to true
Vissim.LoadNet Filename, flag_read_additionally

' Load a Layout:
Filename = Path_of_COM_Basic_Commands_network & "\COM Basic Commands.layx"
Vissim.LoadLayout Filename

' Counting output row to Excel
cnt_row = 4

' ========================================================================
' Read and Set attributes
'==========================================================================
' Note: All of the following commands can also be executed during a
' simulation.
Dim link_number                 As Integer
Dim Name_of_Link                As String
Dim new_Name_of_Link            As String

' Read Link Name:
link_number = 1
Name_of_Link = Vissim.Net.Links.ItemByKey(link_number).AttValue("Name")
cnt_row = cnt_row + 1
Cells(cnt_row, 1) = "Name of Link(" & link_number & "): " & Name_of_Link
' Set Link Name:
new_Name_of_Link = "New Link Name"
Vissim.Net.Links.ItemByKey(link_number).AttValue("Name") = new_Name_of_Link

' Set a signal controller program:
Dim SC_number                   As Integer
Dim new_signal_programm_number  As Integer
Dim SignalController            As Object

SC_number = 1 ' SC = SignalController
Set SignalController = Vissim.Net.SignalControllers.ItemByKey(SC_number)
new_signal_programm_number = 1
SignalController.AttValue("ProgNo") = new_signal_programm_number

' Set relative flow of a static vehicle route of a static vehicle routing decision:
Dim SVRD_number         As Integer
Dim SVR_number          As Integer
Dim new_relativ_flow    As Double
SVRD_number = 1         ' SVRD = Static Vehicle Routing Decision
SVR_number = 1          ' SVR = Static Vehicle Route (of a specific Static Vehicle Routing Decision)
new_relativ_flow = 0.6
Vissim.Net.VehicleRoutingDecisionsStatic.ItemByKey(SVRD_number).VehRoutSta.ItemByKey(SVR_number).AttValue("RelFlow(1)") = new_relativ_flow
' 'RelFlow(1)' means the first defined time interval; to access the third defined time interval: 'RelFlow(3)'

' Set vehicle input:
Dim VI_number   As Integer
Dim new_volume  As Double
VI_number = 1   ' VI = Vehicle Input
new_volume = 600  ' vehicles per hour
Vissim.Net.VehicleInputs.ItemByKey(VI_number).AttValue("Volume(1)") = new_volume
' 'Volume(1)' means the first defined time interval
' Hint: The Volumes of following intervals Volume(i) i = 2...n can only be
' edited, if continuous is deactivated: (otherwise error: "AttValue failed: Object 2: Attribute Volume (300) is no subject to changes.")
Vissim.Net.VehicleInputs.ItemByKey(VI_number).AttValue("Cont(2)") = False
Vissim.Net.VehicleInputs.ItemByKey(VI_number).AttValue("Volume(2)") = 400

' Set vehicle composition:
Dim Veh_composition_number  As Integer
Dim Rel_Flows               As Variant
Veh_composition_number = 1
Rel_Flows = Vissim.Net.VehicleCompositions.ItemByKey(Veh_composition_number).VehCompRelFlows.GetAll
Rel_Flows(0).AttValue("VehType") = 100 ' Changing the vehicle type
Rel_Flows(0).AttValue("DesSpeedDistr") = 50 ' Changing the desired speed distribution
Rel_Flows(0).AttValue("RelFlow") = 0.9 ' Changing the relative flow
Rel_Flows(1).AttValue("RelFlow") = 0.1 ' Changing the relative flow of the 2nd Relative Flow.

'' ========================================================================
' Accessing Multiple Attributes:
'========================================================================

' GetMultiAttValues         Read one attribute of all objects:
Dim Attribute1 As String
Dim NameOfLinks As Variant
Attribute1 = "Name"
NameOfLinks = Vissim.Net.Links.GetMultiAttValues(Attribute1)

' SetMultiAttValues         Set one attribute of multiple (not necessarily all) objects
NameOfLinks(0, 1) = "New Link Name of Link #1"
NameOfLinks(1, 1) = "New Link Name of Link #2"
NameOfLinks(3, 1) = "New Link Name of Link #4"
Vissim.Net.Links.SetMultiAttValues Attribute1, NameOfLinks ' 1st input is the consecutively number of links (not the ID), 2nd the link name
' Please note: The first column of GetMultiAttValues or SetMultiAttValues is not the ID of an object, it is consecutively numbered all available objects.

' GetMultipleAttributes     Read multiple attributes of all objects:
Dim Attributes1(1) As Variant
Attributes1(0) = "Name"
Attributes1(1) = "Length2D"
Name_Length_of_Links = Vissim.Net.Links.GetMultipleAttributes(Attributes1)

' SetMultipleAttributes     Set multiple attribute of multiple (always the first x) objects:
Dim Attributes2(1) As Variant
Dim Link_Name_Cost(3, 1) As Variant
Attributes2(0) = "Name"
Attributes2(1) = "CostPerKm"
Link_Name_Cost(0, 0) = "Name1"
Link_Name_Cost(0, 1) = 12
Link_Name_Cost(1, 0) = "Name2"
Link_Name_Cost(1, 1) = 7
Link_Name_Cost(2, 0) = "Name3"
Link_Name_Cost(2, 1) = 5
Link_Name_Cost(3, 0) = "Name4"
Link_Name_Cost(3, 1) = 3
Vissim.Net.Links.SetMultipleAttributes Attributes2, Link_Name_Cost

' SetAllAttValues           Set all attributes of one object to one value:
Dim Link_Name As String
Dim Cost As Double
Attribute1 = "Name"
Link_Name = "All Links have the same Name"
Vissim.Net.Links.SetAllAttValues Attribute1, Link_Name
Attribute1 = "CostPerKm"
Cost = 5.5
Vissim.Net.Links.SetAllAttValues Attribute1, Cost
' Note the method SetAllAttValues has a 3rd optional input: Optional ByVal add As Boolean = False; Use only for numbers!
Vissim.Net.Links.SetAllAttValues Attribute1, Cost, True ' setting the 3rd input to true, will add 5.5 to all previous costs!



'' ========================================================================
' Simulation
'==========================================================================

' Chose Random Seed
Dim Random_Seed         As Integer
Dim End_of_simulation   As Double
Dim Sim_break_at        As Double
Random_Seed = 42
Vissim.Simulation.AttValue("RandSeed") = Random_Seed

' To start a simulation you can run a single step:
Vissim.Simulation.RunSingleStep
' Or run the simulation continuous (it stops at breakpoint or end of simulation)
End_of_simulation = 600 ' simulation second [s]
Vissim.Simulation.AttValue("SimPeriod") = End_of_simulation
Sim_break_at = 200 ' simulation second [s]
Vissim.Simulation.AttValue("SimBreakAt") = Sim_break_at
' Set maximum speed:
Vissim.Simulation.AttValue("UseMaxSimSpeed") = True
' Hint: to change the speed use: Vissim.Simulation.AttValue("SimSpeed") = 10 ' 10 => 10 Sim. sec. / s
Vissim.Simulation.RunContinuous

' To stop the simulation:
Vissim.Simulation.Stop

'' ========================================================================
' Access during simulation
'==========================================================================
' Note: All of commands of "Read and Set attributes (vehicles)" can also be executed during a
' simulation (e.g. changing signal controller program, setting relative flow of a static vehicle route,
' changing the vehicle input, changing the vehicle composition).

Sim_break_at = 198 ' simulation second [s]
Vissim.Simulation.AttValue("SimBreakAt") = Sim_break_at
Vissim.Simulation.RunContinuous ' Start the simulation until SimBreakAt (198s)

' Get the state of a signal head:
Dim SH_number       As Integer
Dim State_of_SH     As String
SH_number = 1 ' SH = SignalHead
State_of_SH = Vissim.Net.SignalHeads.ItemByKey(SH_number).AttValue("SigState") ' possible output e.g. 'GREEN', 'RED', 'AMBER', 'REDAMBER'
cnt_row = cnt_row + 1
Cells(cnt_row, 1) = "Actual state of SignalHead(" & SH_number & ") is: " & State_of_SH

' Set the state of a signal controller:
' Note: Once a state of a signal group is set, the attribute "ContrByCOM" is automatically set to True. Meaning the signal group will keep this state until another state is set by COM or the end of the simulation
' To switch back to the defined signal controller, set the attribute signal "ContrByCOM" to False (example see below).
Dim SignalGroup As Object
Dim SG_number As Integer
Dim new_state As String
SC_number = 1 ' SC = SignalController
SG_number = 1 ' SG = SignalGroup
Set SignalController = Vissim.Net.SignalControllers.ItemByKey(SC_number)
Set SignalGroup = SignalController.SGs.ItemByKey(SG_number)
new_state = "GREEN" 'possible values e.g. "GREEN", "RED", "AMBER", "REDAMBER"
SignalGroup.AttValue("SigState") = new_state
' Note: The signal controller can only be called at whole simulation seconds, so the state will be set in Vissim at the next whole simulation second, here 199s
' Simulate so that the new state is active in the Vissim simulation:
Sim_break_at = 200 ' simulation second [s]
Vissim.Simulation.AttValue("SimBreakAt") = Sim_break_at
Vissim.Simulation.RunContinuous ' Start the simulation until SimBreakAt (200s)
' Give the control back:
SignalGroup.AttValue("ContrByCOM") = False


' Information about all vehicles in the network (in the current simulation second):
' In the following, 4 different methods to access attributes are shown:
' Method #1: Loop over all Vehicles using "GetAll"
' Method #2: Loop over all Vehicles using Object Enumeration
' Method #3: Using the Iterator
' Method #4: Accessing all attributes directly using "GetMultiAttValues" (fast way if you want the attributes of all vehicles)
' Method #5: Accessing all attributes directly using "GetMultipleAttributes" (even more faster)
' The result of the four methods is the same (except the format).

' Method #1: Loop over all Vehicles:
Dim All_Vehicles    As Variant
Dim cnt_Veh         As Integer
Dim veh_number      As Integer
Dim veh_type        As Integer
Dim veh_speed       As Double
Dim veh_position    As Double
Dim veh_linklane    As String
All_Vehicles = Vissim.Net.Vehicles.GetAll ' get all vehicles in the network at the actual simulation second
For cnt_Veh = 0 To UBound(All_Vehicles)
    veh_number = All_Vehicles(cnt_Veh).AttValue("No")
    veh_type = All_Vehicles(cnt_Veh).AttValue("VehType")
    veh_speed = All_Vehicles(cnt_Veh).AttValue("Speed")
    veh_position = All_Vehicles(cnt_Veh).AttValue("Pos")
    veh_linklane = All_Vehicles(cnt_Veh).AttValue("Lane")
    cnt_row = cnt_row + 1
    Cells(cnt_row, 1) = veh_number
    Cells(cnt_row, 2) = veh_type
    Cells(cnt_row, 3) = veh_speed
    Cells(cnt_row, 4) = veh_position
    Cells(cnt_row, 5) = veh_linklane

Next cnt_Veh

' Method #2: Loop over all Vehicles using Object Enumeration
Dim Vehicle As Object
For Each Vehicle In Vissim.Net.Vehicles
    veh_number = Vehicle.AttValue("No")
    veh_type = Vehicle.AttValue("VehType")
    veh_speed = Vehicle.AttValue("Speed")
    veh_position = Vehicle.AttValue("Pos")
    veh_linklane = Vehicle.AttValue("Lane")
    cnt_row = cnt_row + 1
    Cells(cnt_row, 1) = veh_number
    Cells(cnt_row, 2) = veh_type
    Cells(cnt_row, 3) = veh_speed
    Cells(cnt_row, 4) = veh_position
    Cells(cnt_row, 5) = veh_linklane
Next Vehicle


' Method #3: Using the Iterator (this method is a little bit slower)
Dim Vehicles_Iterator As Object
Set Vehicles_Iterator = Vissim.Net.Vehicles.Iterator
While Vehicles_Iterator.Valid
    Set Vehicle = Vehicles_Iterator.Item
    veh_number = Vehicle.AttValue("No")
    veh_type = Vehicle.AttValue("VehType")
    veh_speed = Vehicle.AttValue("Speed")
    veh_position = Vehicle.AttValue("Pos")
    veh_linklane = Vehicle.AttValue("Lane")
    cnt_row = cnt_row + 1
    Cells(cnt_row, 1) = veh_number
    Cells(cnt_row, 2) = veh_type
    Cells(cnt_row, 3) = veh_speed
    Cells(cnt_row, 4) = veh_position
    Cells(cnt_row, 5) = veh_linklane
    Vehicles_Iterator.Next
Wend

' Method #4: Accessing all Attributes directly using "GetMultiAttValues" (fast)
Dim veh_numbers     As Variant
Dim veh_types       As Variant
Dim veh_speeds      As Variant
Dim veh_positions   As Variant
Dim veh_linklanes   As Variant
veh_numbers = Vissim.Net.Vehicles.GetMultiAttValues("No")          ' Output 1. column:consecutive number; 2. column: AttValue
veh_types = Vissim.Net.Vehicles.GetMultiAttValues("VehType")       ' Output 1. column:consecutive number; 2. column: AttValue
veh_speeds = Vissim.Net.Vehicles.GetMultiAttValues("Speed")        ' Output 1. column:consecutive number; 2. column: AttValue
veh_positions = Vissim.Net.Vehicles.GetMultiAttValues("Pos")       ' Output 1. column:consecutive number; 2. column: AttValue
veh_linklanes = Vissim.Net.Vehicles.GetMultiAttValues("Lane")      ' Output 1. column:consecutive number; 2. column: AttValue
For i = 0 To UBound(veh_numbers)
    cnt_row = cnt_row + 1
    Cells(cnt_row, 1) = veh_numbers(i, 1)    ' only display the 2nd column
    Cells(cnt_row, 2) = veh_types(i, 1)      ' only display the 2nd column
    Cells(cnt_row, 3) = veh_speeds(i, 1)     ' only display the 2nd column
    Cells(cnt_row, 4) = veh_positions(i, 1)  ' only display the 2nd column
    Cells(cnt_row, 5) = veh_linklanes(i, 1)  ' only display the 2nd column
Next i

' Method #5: Accessing all attributes directly using "GetMultipleAttributes" (even more faster)
Dim all_veh_attributes As Variant
all_veh_attributes = Vissim.Net.Vehicles.GetMultipleAttributes(Array("No", "VehType", "Speed", "Pos", "Lane"))
For i = 0 To UBound(all_veh_attributes)
    cnt_row = cnt_row + 1
    Cells(cnt_row, 1) = all_veh_attributes(i, 0)
    Cells(cnt_row, 2) = all_veh_attributes(i, 1)
    Cells(cnt_row, 3) = all_veh_attributes(i, 2)
    Cells(cnt_row, 4) = all_veh_attributes(i, 3)
    Cells(cnt_row, 5) = all_veh_attributes(i, 4)
Next i


' Operations at one specific vehicle:
All_Vehicles = Vissim.Net.Vehicles.GetAll    ' get all vehicles in the network at the actual simulation second
Set Vehicle = All_Vehicles(1)
' alternatively with ItemByKey:
' veh_number = 66 ' the same as: get(All_Vehicles{1}, 'AttValue', 'No')
' Vehicle = Vissim.Net.Vehicles.ItemByKey(veh_number)

' Set desired speed to a vehicle:
Dim new_desspeed As Double
new_desspeed = 30
Vehicle.AttValue("DesSpeed") = new_desspeed

' Move a vehicle:
Dim lane_number         As Integer
Dim link_coordinate     As Double
link_number = 1
lane_number = 1
link_coordinate = 70
Vehicle.MoveToLinkPosition link_number, lane_number, link_coordinate ' This function will operate during the next simulation step
'Vehicle = Vehicle.MoveToLinkPosition(link_number, lane_number, link_coordinate) ' This function will operate during the next simulation step

' Hint: In earlier Vissim releases, the name of the function was: MoveToLinkCoordinate

Vissim.Simulation.RunSingleStep ' Next Step, so that the vehicle gets moved.

' Remove a vehicle:
veh_number = Vehicle.AttValue("No")
Vissim.Net.Vehicles.RemoveVehicle veh_number

' Putting a new vehicle to the network:
Dim vehicle_type    As Integer
Dim link            As Integer
Dim lane            As Integer
Dim desired_speed   As Double
Dim xcoordinate     As Double
Dim Interaction     As Boolean
vehicle_type = 100
desired_speed = 53 ' unit according to the user setting in Vissim [km/h or mph]
link = 1
lane = 1
xcoordinate = 15 ' unit according to the user setting in Vissim [m or ft]
Interaction = True ' optional boolean
Vissim.Net.Vehicles.AddVehicleAtLinkPosition vehicle_type, link, lane, xcoordinate, desired_speed, Interaction
' Note: In earlier Vissim releases, the name of the function was: AddVehicleAtLinkCoordinate

' Make Screenshots of the intersection 2D and 3D:
' ZoomTo:
' Zooms the view to the rectangle defined by the two points  (x1, y1) and (x2,y2),  which  are  given  in  world coordinates.  If  the  rectangle  proportions  differ
' from  the  proportions  of  the  network  window,  the  specified  rectangle  will  be centred in the network editor window.
Dim X1    As Integer
Dim Y1    As Integer
Dim X2    As Integer
Dim Y2    As Integer
X1 = 250
Y1 = 30
X2 = 350
Y2 = 135
Vissim.Graphics.CurrentNetworkWindow.ZoomTo X1, Y1, X2, Y2

' Make a Screenshot in 2D:
' It  creates  a  graphic  file  of  the  VISSIM  main  window  formatted  according  to its extension: PNG, TIFF, GIF, JPG, JPEG or BMP. A BMP file will be written if the extension can not be recognized.
Dim Filename_screenshot As String
Dim sizeFactor                  As Integer
Filename_screenshot = "screenshot2D.jpg" ' to set to a specific path: "C:\Screenshots\screenshot2D.jpg"
sizeFactor = 1 ' 1: original size, 2: doubles size
Vissim.Graphics.CurrentNetworkWindow.Screenshot Filename_screenshot, sizeFactor

' Make a Screenshot in 3D:
' Set 3D Mode:
Vissim.Graphics.CurrentNetworkWindow.AttValue("3D") = 1
' Set the camera position (viewing angle):
Dim xPos        As Integer
Dim yPos        As Integer
Dim zPos        As Integer
Dim yawAngle    As Integer
Dim pitchAngle  As Integer
xPos = 270
yPos = 30
zPos = 15
yawAngle = 45
pitchAngle = 10
Vissim.Graphics.CurrentNetworkWindow.SetCameraPositionAndAngle xPos, yPos, zPos, yawAngle, pitchAngle
Filename_screenshot = "screenshot3D.jpg" ' to set to a specific path: "C:\Screenshots\screenshot3D.jpg"
Vissim.Graphics.CurrentNetworkWindow.Screenshot Filename_screenshot, sizeFactor

' Set 2D Mode and old Network position:
Vissim.Graphics.CurrentNetworkWindow.AttValue("3D") = 0
X1 = -10
Y1 = -10
X2 = 600
Y2 = 300
Vissim.Graphics.CurrentNetworkWindow.ZoomTo X1, Y1, X2, Y2

' Continue the simulation until end of simulation (get(Vissim.Simulation, 'AttValue', 'SimPeriod'))
Vissim.Simulation.RunContinuous


'' ========================================================================
' Results of Simulations:
'==========================================================================

' Run 3 Simulations at maximum speed:
' Ensure that all previous results are deleted in order to start clean
Dim simRun As Object
For Each simRun in Vissim.Net.SimulationRuns
  Call Vissim.Net.SimulationRuns.RemoveSimulationRun(simRun)
Next simRun

Dim cnt_Sim     As Integer
' Activate QuickMode:
Vissim.Graphics.CurrentNetworkWindow.AttValue("QuickMode") = 1
Vissim.SuspendUpdateGUI ' stop updating of the complete Vissim workspace (network editor, list, chart and signal time table windows)
' Alternatively, load a layout (*.laxy) where dynamic elements (vehicles and pedestrians) are not visible:
' Vissim.LoadLayout Path_of_COM_Basic_Commands_network & "\COM Basic Commands - Hide vehicles.layx" ' loading a layout where vehicles are not displayed => much faster simulation
End_of_simulation = 600
Vissim.Simulation.AttValue("SimPeriod") = End_of_simulation
Sim_break_at = 0 ' simulation second [s] => 0 means no break!
Vissim.Simulation.AttValue("SimBreakAt") = Sim_break_at
' Set maximum speed:
Vissim.Simulation.AttValue("UseMaxSimSpeed") = True
For cnt_Sim = 1 To 3
        Vissim.Simulation.AttValue("RandSeed") = cnt_Sim
    Vissim.Simulation.RunContinuous
Next cnt_Sim
Vissim.ResumeUpdateGUI(true) ' allow updating of the entire Vissim workspace (network editor, list, chart and signal time table windows)
Vissim.Graphics.CurrentNetworkWindow.AttValue("QuickMode") = 0 ' deactivate QuickMode

' Vissim.LoadLayout Path_of_COM_Basic_Commands_network & "\COM Basic Commands.layx"

' List of all Simulation runs:
Dim number_of_sim_runs  As Integer
Dim cnt_A               As Integer
Dim cnt_S               As Integer
Dim Attributes()
Dim List_Sim_Runs       As Variant

number_of_sim_runs = Vissim.Net.SimulationRuns.Count()
Attributes = Array("Timestamp", "RandSeed", "SimEnd")
List_Sim_Runs = Vissim.Net.SimulationRuns.GetMultipleAttributes(Attributes)
For cnt_S = 0 To number_of_sim_runs - 1
    cnt_row = cnt_row + 1
    For cnt_A = 0 To UBound(Attributes)
        Cells(cnt_row, cnt_A + 1) = List_Sim_Runs(cnt_S, cnt_A)    ' output to Excel
    Next cnt_A
Next cnt_S


' Get the results of Vehicle Travel Time Measurements:
Dim Veh_TT_measurement_number   As Integer
Dim No_Veh                      As Integer
Dim Veh_TT_measurement          As Object
Dim TT                          As Double
Veh_TT_measurement_number = 2
Set Veh_TT_measurement = Vissim.Net.VehicleTravelTimeMeasurements.ItemByKey(Veh_TT_measurement_number)
' Syntax to get the travel times:
'   Veh_TT_measurement.AttValue("TravTm(sub_attribut_1, sub_attribut_2, sub_attribut_3)"
'
' sub_attribut_1: SimulationRun
'       1, 2, 3, ... Current:     the value of one specific simulation (number according to the attribute "No" of Simulation Runs (see List of Simulation Runs))
'       Avg, StdDev, Min, Max:    aggregated value of all simulations runs: Avg, StdDev, Min, Max
' sub_attribut_2: TimeInterval
'       1, 2, 3, ... Last:        the value of one specific time interval (number of time interval always starts at 1 (first time interval), 2 (2nd TI), 3 (3rd TI), ...)
'       Avg, StdDev, Min, Max:    aggregated value of all time interval of one simulation: Avg, StdDev, Min, Max
'       Total:                    sum of all time interval of one simulation
' sub_attribut_3: VehicleClass
'       10, 20 or All             values only from vehicles of the defined vehicle class number (according to the attribute "No" of Vehicle Classes)
'                                 Note: You can only access the results of specific vehicle classes if you set it in Evaluation > Configuration > Result Attributes
'
' The value of on time interval is the arithmetic mean of all single travel times of the vehicles.

' Example #1:
' Average of all simulations (1. input = Avg)
'   of the average of all time intervals  (2. input = Avg)
'   of all vehicle classes (3. input = All)
TT = Veh_TT_measurement.AttValue("TravTm(Avg,Avg,All)")
No_Veh = Veh_TT_measurement.AttValue("Vehs(Avg,Avg,All)")
cnt_row = cnt_row + 1
Cells(cnt_row, 1) = "Average travel time all time intervals of all simulation of all vehicle classes: " & TT & " number of vehicles: " & No_Veh & ")"    ' output to Excel

' Example #2:
' Value of the Current simulation (1. input = Current)
'   of the maximum of all time intervals (2. input = Max)
'   of vehicle class HGV (3. input = 20)
TT = Veh_TT_measurement.AttValue("TravTm(Current,Max,20)")
No_Veh = Veh_TT_measurement.AttValue("Vehs(Current,Max,20)")
cnt_row = cnt_row + 1
Cells(cnt_row, 1) = "Maximum travel time of all time intervals of the current simulation of vehicle class HGV: " & TT & " number of vehicles: " & No_Veh & ")"    ' output to Excel

' Example #3: Note: A Travel times from 2nd simulation run must be available
' Value of the 2nd simulation (1. input = 2)
'   of the 1st time interval (2. input = 1)
'   of all vehicle classes (3. input = All)
' TT = Veh_TT_measurement.AttValue("TravTm(2,1,All)")
' No_Veh = Veh_TT_measurement.AttValue("Vehs(2,1,All)")
' cnt_row = cnt_row + 1
' Cells(cnt_row, 1) = "Travel time of the 1st time interval of the 2nd simulation of all vehicle classes: " & TT & " number of vehicles: " & No_Veh & ")"    ' output to Excel

' Data Collection
Dim DC_measurement_number   As Integer
Dim DC_measurement          As Object
Dim Speed                   As Double
Dim Acceleration            As Double
Dim Length                  As Double
DC_measurement_number = 1
Set DC_measurement = Vissim.Net.DataCollectionMeasurements.ItemByKey(DC_measurement_number)
' Syntax to get the data:
'   DC_measurement.AttValue("Vehs(sub_attribut_1, sub_attribut_2, sub_attribut_3)")
'
' sub_attribut_1: SimulationRun (same as described at Vehicle Travel Time Measurements)
' sub_attribut_2: TimeInterval  (same as described at Vehicle Travel Time Measurements)
' sub_attribut_3: VehicleClass  (same as described at Vehicle Travel Time Measurements)
'
' The value of on time interval is the arithmetic mean of all single values of the vehicles.

' Example #1:
' Average value of all simulations (1. input = Avg)
'   of the 1st time interval (2. input = 1)
'   of all vehicle classes (3. input = All)
No_Veh = DC_measurement.AttValue("Vehs(Avg,1,All)")                 ' number of vehicles
Speed = DC_measurement.AttValue("Speed(Avg,1,All)")                 ' Speed of vehicles
Acceleration = DC_measurement.AttValue("Acceleration(Avg,1,All)")   ' Acceleration of vehicles
Length = DC_measurement.AttValue("Length(Avg,1,All)")               ' Length of vehicles
cnt_row = cnt_row + 1
Cells(cnt_row, 1) = "Data Collection #" & DC_measurement_number & ": Average values of all Simulations runs of 1st time interval of all vehicle classes"    ' output to Excel
Cells(cnt_row, 2) = "#vehicles: " & No_Veh & " ; Speed:" & Speed & " ; Acceleration: " & Acceleration & " Length: " & Length    ' output to Excel

' Queue length
' Syntax to get the data:
'   QueueCounter.AttValue("QLen(sub_attribut_1, sub_attribut_2)")
'
' sub_attribut_1: SimulationRun (same as described at Vehicle Travel Time Measurements)
' sub_attribut_2: TimeInterval  (same as described at Vehicle Travel Time Measurements)
'

' Example #1:
' Average value of all simulations (1. input = Avg)
'   of the average of all time intervals (2. input = Avg)
Dim QC_number   As Integer
QC_number = 1
maxQ = Vissim.Net.QueueCounters.ItemByKey(QC_number).AttValue("QLenMax(Avg, Avg)")
cnt_row = cnt_row + 1
Cells(cnt_row, 1) = "Average maximum Queue length of all simulations and time intervals of Queue Counter #" & QC_number & ": " & maxQ    ' output to Excel


'' ========================================================================
' Saving
'==========================================================================
Filename = Path_of_COM_Basic_Commands_network & "\COM Basic Commands saved.inpx"
Vissim.SaveNetAs (Filename)
Filename = Path_of_COM_Basic_Commands_network & "\COM Basic Commands saved.layx"
Vissim.SaveLayout (Filename)


'' ========================================================================
' End Vissim
'==========================================================================
Set Vissim = Nothing



End Sub
