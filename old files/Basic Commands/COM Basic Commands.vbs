'==========================================================================
' VBS-Script for Vissim 8+
' Copyright (C) PTV AG, Sven Beller
' All rights reserved.
' -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -
' Examples of basic COM syntax
'==========================================================================

' This script demonstrates how to use the Vissim COM interface in Visual Basic Script.
' Basic commands for loading a layout, reading and setting
' attributes of network objects, running a simulation and retrieving
' evaluations are shown. This example is also available for other programming
' languages such as Matlab, Python, C# and C++ etc.
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
'
' For adding network objects with COM, please see the example Network Objects Adding

Option Explicit

'==========================================================================
' Load a layout file
'==========================================================================
Sub LoadLayout()
	Dim filename
	filename = "COM Basic Commands VBS.layx"

	Vissim.LoadLayout Filename
End Sub

'==========================================================================
' Read and set single attribute values
'==========================================================================
Sub SingleAttributes()
	' Note: All of the following commands can also be executed during a
	' simulation.
	Dim linkNo
	Dim linkName
	Dim linkNameNew

	' Read Link Name:
	linkNo = 1
	linkName = Vissim.Net.Links.ItemByKey(linkNo).AttValue("Name")
	MsgBox "Name of Link " + CStr(linkNo) +" is '" + CStr(linkName) + "'."

	' Set Link Name:
	linkNameNew = "New Link Name"
	Vissim.Net.Links.ItemByKey(linkNo).AttValue("Name") = linkNameNew
	MsgBox "Name of Link " + CStr(linkNo) +" is now set to '" + CStr(linkNameNew) + "'." + vbcrlf + vbcrlf _
         + "About to change the signal program number..."

	' Set a signal controller program:
	Dim controllerNo 
	Dim signalController
	Dim programNoNew
	controllerNo = 1 
	Set signalController = Vissim.Net.signalControllers.ItemByKey(controllerNo)
	programNoNew = 1
	signalController.AttValue("ProgNo") = programNoNew
	MsgBox "The program number of signal controller " + CStr(controllerNo) +" is now set to " + CStr(programNoNew) + "."  + vbcrlf + vbcrlf _
         + "About to change a route relative flow..."
	
	' Set relative flow of a static vehicle route of a static vehicle routing decision:
	Dim rouDecNo
	Dim routeNo
	Dim relativeFlowNew
	rouDecNo = 1
	routeNo = 1		' specific to the static vehicle routing decision
	relativeFlowNew = 0.6
	Vissim.Net.VehicleRoutingDecisionsStatic.ItemByKey(rouDecNo).VehRoutSta.ItemByKey(routeNo).AttValue("RelFlow(1)") = relativeFlowNew
	' In RelFlow(1) the index 1 refers to the first time interval defined in Vissim (and not the start time in seconds). 
	' Index 1 refers to the first time interval defined; to access the third time interval use RelFlow(3)
	Dim msgString
  msgString = "The relative flow of Route " + CStr(rouDecNo) + "-" + CStr(routeNo) + " is now set to " + CStr(relativeFlowNew) + "."

	' Set vehicle input:
	Dim vehInputNo
	Dim volumeNew
	vehInputNo = 1   

  MsgBox msgString + vbcrlf + vbcrlf +  "About to change the volume of vehicle input no. " + CStr(vehInputNo) + "..."
  
	volumeNew = 600  ' in vehicles per hour
	Vissim.Net.VehicleInputs.ItemByKey(vehInputNo).AttValue("Volume(1)") = volumeNew
	MsgBox "The volume of the first time interval of vehicle input " + CStr(vehInputNo) + " is now set to " + CStr(volumeNew) + "."

	' In Volume(1) the index 1 refers to the first time interval defined in Vissim (and not the start time in seconds). 
	' Please note: The volumes of the subsequent time intervals (Volume(i) with i = 2...n) can only be
	'              edited, if the input attribute 'Continued' is deactivated. Otherwise an error simliar like this is triggered:
  '	             "AttValue failed: Object 2: Attribute Volume (300) is no subject to changes.")
	Vissim.Net.VehicleInputs.ItemByKey(vehInputNo).AttValue("Cont(2)") = False
	volumeNew = 400  ' in vehicles per hour
	Vissim.Net.VehicleInputs.ItemByKey(vehInputNo).AttValue("Volume(2)") = volumeNew
	MsgBox "The volume of the 2nd time interval of vehicle input " + CStr(vehInputNo) + " is now set to " + CStr(volumeNew) + "." + vbcrlf + vbcrlf _
           + "About to change a vehicle composition..."
	
	' Set vehicle composition:
	Dim vehCompositionNo
	Dim relFlows
	vehCompositionNo = 1
	relFlows = Vissim.Net.VehicleCompositions.ItemByKey(vehCompositionNo).VehCompRelFlows.GetAll
	relFlows(0).AttValue("VehType") = 100 			' Setting the vehicle type
	relFlows(0).AttValue("DesSpeedDistr") = 50 	' Changing the desired speed distribution
	relFlows(0).AttValue("RelFlow") = 0.9 			' Changing the flow value for the first vehicle type in the composition
	relFlows(1).AttValue("RelFlow") = 0.1 			' Changing the flow value for the second vehicle type in the composition
	MsgBox "Vehicle composition " + CStr(vehCompositionNo) + " is now changed."
	
End Sub

'==========================================================================
' Accessing Multiple Attributes
'==========================================================================
Sub MultiAttributes()
	MsgBox "To see the effect of the script we recommend to open the links list."

	' GetMultiAttValues: Read one attribute of all objects
	Dim attribute1
	Dim linkNames
	attribute1 = "Name"
	linkNames = Vissim.Net.Links.GetMultiAttValues(attribute1)

	' SetMultiAttValues: Set one attribute for multiple (not necessarily all) objects
	Dim i
	For i = 0 to UBound(linkNames)
		if linkNames(i, 0) = 1 then linkNames(i, 1) = "New name of link #1"
		if linkNames(i, 0) = 2 then linkNames(i, 1) = "New name of link #2"
		if linkNames(i, 0) = 4 then linkNames(i, 1) = "New name of link #4"
	Next
	Vissim.Net.Links.SetMultiAttValues attribute1, linkNames
	MsgBox "The names of links 1, 2 and 4 were changed."
	
	' GetMultipleAttributes: Read multiple attributes of all objects
	Dim attributeList
	Dim link_Name_Length
	attributeList = Array ("Name", "Length2D")
	link_Name_Length = Vissim.Net.Links.GetMultipleAttributes(attributeList)

	' SetMultipleAttributes: Set multiple attribute of multiple objects
	Dim link_Name_Cost
	attributeList = Array ("Name", "CostPerKm")
	link_Name_Cost = Vissim.Net.Links.GetMultipleAttributes(attributeList)
	
	link_Name_Cost(0, 0) = "Name1"
	link_Name_Cost(0, 1) = 12
	link_Name_Cost(1, 0) = "Name2"
	link_Name_Cost(1, 1) = 7
	link_Name_Cost(2, 0) = "Name3"
	link_Name_Cost(2, 1) = 5
	link_Name_Cost(3, 0) = "Name4"
	link_Name_Cost(3, 1) = 3
	
	Vissim.Net.Links.SetMultipleAttributes attributeList, link_Name_Cost
	MsgBox "The name and link cost of the first 4 links were changed."

	' SetAllAttValues: Set the same attribute value to all objects of a specific type
	attribute1 = "Name"
	Dim linkName
	linkName = "All Links have the same name"
	Vissim.Net.Links.SetAllAttValues attribute1, linkName
	MsgBox "All links now have the same name."
	
	attribute1 = "costPerKm"
	Dim cost
	cost = 5.5
	Vissim.Net.Links.SetAllAttValues attribute1, cost
	MsgBox "All links now have a cost of " + CStr(cost)
	
	' Note: SetAllAttValues has a 3rd optional parameter 'add' which by default is false. Use it only for numerical values.
	cost = 7.5
	Vissim.Net.Links.SetAllAttValues attribute1, cost, True  ' this command ADDs 5.5 to all previous cost values.
	MsgBox "To the cost of all links an additional value of " + CStr(cost) + " was added."
	
End Sub

'==========================================================================
' Simulation
'==========================================================================
Sub RunSimulation()
	
  ' Set random seed
	Dim randomSeed
	randomSeed = 42
	Vissim.Simulation.AttValue("RandSeed") = randomSeed

  ' Set the simulation end
	Dim endOfSimulation
	endOfSimulation = 350 	' [simulation second]
	Vissim.Simulation.AttValue("SimPeriod") = endOfSimulation

	' To start a simulation you can run a single step ...
	Vissim.Simulation.RunSingleStep

	Dim simBreakAt
	simBreakAt = 200 				' [simulation second]
	Vissim.Simulation.AttValue("SimBreakAt") = simBreakAt
	' Set maximum simulation speed:
	Vissim.Simulation.AttValue("UseMaxSimSpeed") = True
	' Hint: to change the speed use: Vissim.Simulation.AttValue("SimSpeed") = 10 ' 10 => 10 Sim. sec. / s
	
	' ... or run the simulation continuous: it then stops at the breakpoint (if defined) or at the end of the simulation
	MsgBox "The simulation will now run until the 'Break at' time at simulation second " + CStr(simBreakAt) + ". From there you can either continue the simulation (F5) or stop it (Escape).", vbInformation+vbOKOnly ,"Run Simulation" 

	Vissim.Simulation.RunContinuous

End Sub
	
'==========================================================================
' Access during simulation
'==========================================================================
' Note: All of the commands mentioned above in the section "Read and Set attributes (vehicles)"
' can also be executed during a simulation (e.g. changing signal controller program,
' setting relative flow of a route, changing the vehicle input, changing the vehicle composition).
Sub AccessDuringSimulation()

	Vissim.Simulation.Stop ' Stop the simulation in case its running

  ' Set the simulation end
	Dim endOfSimulation
	endOfSimulation = 70 	' [simulation second]
	Vissim.Simulation.AttValue("SimPeriod") = endOfSimulation

  ' Set simulation break at
	Dim simBreakAt
	simBreakAt = 28    ' [Simulation second]
	Vissim.Simulation.AttValue("SimBreakAt") = simBreakAt

 	MsgBox "The simulation will now run until simulation second " + CStr(simBreakAt) + ". Then various attributes are accessed while the simulation is active. Afterwards the simulation will run until the end of " + CStr(endOfSimulation) + " s." , vbInformation+vbOKOnly ,"Access Data During Simulation" 
	Vissim.Simulation.RunContinuous ' Start the simulation until SimBreakAt

	' Set the state of a signal group
	' Note: Once a state of a signal group is set by COM, the attribute 'ContrByCOM' is automatically set to 'True'.
	' 	    Hence the signal group will keep this state until a different state is set by a COM command.
	'       The signal group will NOT react to any calls by the signal controller any more.
	'       To hand back control to the signal controller, set the attribute 'ContrByCOM' to 'False' (see example below).
	Dim controllerNo
	Dim signalGroupNo
	Dim signalGroup
	controllerNo = 1 
	signalGroupNo = 3 
	Set signalGroup = Vissim.Net.signalControllers.ItemByKey(controllerNo).SGs.ItemByKey(signalGroupNo)

 	Dim stateNew
	stateNew = "RED" 	' possible values e.g.: "GREEN", "RED", "AMBER", "REDAMBER"

	' Get the state of a signal group:
	Dim signalGroupState
  signalGroupState = signalGroup.AttValue("SigState")    ' possible return values e.g.: 'GREEN', 'RED', 'AMBER', 'REDAMBER'
	MsgBox "Current state of signal group " + CStr(signalGroupNo) + " is " + signalGroupState + vbcrlf _
          + "It will be controlled by COM now and set to " + stateNew + "."

	' Set the new state of a signal group
  signalGroup.AttValue("SigState") = stateNew

	' Note: A signal controller in Vissim is called only at simulation seconds that are multiples of the controller frequency
  '      (every simulation second by default). Hence when changing the state of a signal group through COM, the visible
	'      change in the Vissim network might be delayed. In this example, the next controller actuation is at simulation second 199.

 	' To do something in every time step:
	Do 
		' do something
		Vissim.Simulation.RunSingleStep
	Loop While Vissim.Simulation.AttValue("SimSec") < 29 

	MsgBox "The control of signal controller " + CStr(controllerNo) + " will now " + vbcrlf _
       + "be handed back from COM to the signal controller - " + vbcrlf _
       + "hence the state of signal group " + CStr(signalGroupNo) + " will change again."
	' Return control back to the signal controller:
	signalGroup.AttValue("ContrByCOM") = False

	Do 
		' do something
		Vissim.Simulation.RunSingleStep
	Loop While Vissim.Simulation.AttValue("SimSec") < 30

	' --------------------------------------------------------------------------------------
	' Fetch information about all vehicles in the network (in the current simulation second)
	' --------------------------------------------------------------------------------------
	' Below various methods are shown to access attributes:
	' Method 1: Loop over all vehicles using "GetAll"
	' Method 2: Loop over all vehicles using object enumeration
	' Method 3: Using the iterator (this option is a little slower)
	' Method 4: Accessing all attributes directly using "GetMultiAttValues" (faster option if you want to access the attributes of all vehicles)
	' Method 5: Accessing all attributes directly using "GetMultipleAttributes" (fastest option, especially for access of several attributes)
	' The result of each method is identical (except the format).
	' For simplicity, for each method only the data of the last object is printed to a message dialogue.

	' .......................................................................................
	' Method 1: Loop over all vehicles:
	Dim allVehicles
	Dim vehCount
	Dim vehNo
	Dim vehType
	Dim vehSpeed
	Dim vehPos
	Dim vehLinkLane
	allVehicles = Vissim.Net.Vehicles.GetAll 		' get all vehicles in the network at the current simulation second
	For vehCount = 0 To UBound(allVehicles)
			vehNo = allVehicles(vehCount).AttValue("No")
			vehType = allVehicles(vehCount).AttValue("VehType")
			vehSpeed = allVehicles(vehCount).AttValue("Speed")
			vehPos = allVehicles(vehCount).AttValue("Pos")
			vehLinkLane = allVehicles(vehCount).AttValue("Lane")
	Next
	Dim msgText
  msgText = "Number of vehicles in network: " + CStr(Vissim.Net.Vehicles.Count) + vbcrlf + vbcrlf _
          + "Information of last vehicle:" + vbCrlf _
          + "vehicle number = " + CStr(vehNo) + vbCrlf _
          + "Type = " + CStr(vehType) + vbCrlf _
          + "Speed = " + CStr(FormatNumber(vehSpeed,1)) + vbCrlf _
          + "Pos = " + CStr(FormatNumber(vehPos,2)) + vbCrlf _
          + "Link & lane = " + CStr(vehLinkLane)
  MsgBox msgText,,"Loop over all vehicles"

	' .......................................................................................
	' Method 2: Loop over all vehicles using object enumeration
	Dim vehicle
	For Each vehicle In Vissim.Net.Vehicles
			vehNo = vehicle.AttValue("No")
			vehType = vehicle.AttValue("VehType")
			vehSpeed = vehicle.AttValue("Speed")
			vehPos = vehicle.AttValue("Pos")
			vehLinkLane = vehicle.AttValue("Lane")
	Next
	msgText = "Number of vehicles in network: " + CStr(Vissim.Net.Vehicles.Count) + vbcrlf + vbcrlf _
          + "Information of last vehicle:" + vbCrlf _
          + "vehicle number = " + CStr(vehNo) + vbCrlf _
          + "Type = " + CStr(vehType) + vbCrlf _
          + "Speed = " + CStr(FormatNumber(vehSpeed,1)) + vbCrlf _
          + "Pos = " + CStr(FormatNumber(vehPos,2)) + vbCrlf _
          + "Link & lane = " + CStr(vehLinkLane)
  MsgBox msgText,, "Loop (object enumeration)"

	' .......................................................................................
	' Method 3: Using the iterator
	Dim iteratorVehicles
	Set iteratorVehicles = Vissim.Net.Vehicles.Iterator
	While iteratorVehicles.Valid
			Set vehicle = iteratorVehicles.Item
			vehNo = vehicle.AttValue("No")
			vehType = vehicle.AttValue("VehType")
			vehSpeed = vehicle.AttValue("Speed")
			vehPos = vehicle.AttValue("Pos")
			vehLinkLane = vehicle.AttValue("Lane")
			iteratorVehicles.Next
	Wend
	msgText = "Number of vehicles in network: " + CStr(Vissim.Net.Vehicles.Count) + vbcrlf + vbcrlf _
          + "Information of last vehicle:" + vbCrlf _
          + "vehicle number = " + CStr(vehNo) + vbCrlf _
          + "Type = " + CStr(vehType) + vbCrlf _
          + "Speed = " + CStr(FormatNumber(vehSpeed,1)) + vbCrlf _
          + "Pos = " + CStr(FormatNumber(vehPos,2)) + vbCrlf _
          + "Link & lane = " + CStr(vehLinkLane)
  MsgBox msgText,, "Loop (iterator)"

	' .......................................................................................
	' Method 4: Accessing all attributes directly using "GetMultiAttValues"
	Dim vehNumbers
	Dim vehTypes
	Dim vehSpeeds
	Dim vehPositions
	Dim vehLinkLanes

	vehNumbers = Vissim.Net.Vehicles.GetMultiAttValues("No")          ' Output 1. column: consecutive no; 2. column: AttValue
	vehTypes = Vissim.Net.Vehicles.GetMultiAttValues("VehType")       ' Output 1. column: consecutive no; 2. column: AttValue
	vehSpeeds = Vissim.Net.Vehicles.GetMultiAttValues("Speed")        ' Output 1. column: consecutive no; 2. column: AttValue
	vehPositions = Vissim.Net.Vehicles.GetMultiAttValues("Pos")       ' Output 1. column: consecutive no; 2. column: AttValue
	vehLinkLanes = Vissim.Net.Vehicles.GetMultiAttValues("Lane")      ' Output 1. column: consecutive no; 2. column: AttValue
	Dim i
	For i = 0 To UBound(vehNumbers)
			vehNo = vehNumbers(i, 1)    
			vehType = vehTypes(i, 1)
			vehSpeed = vehSpeeds(i, 1)
			vehPos = vehPositions(i, 1)  
			vehLinkLane = vehLinkLanes(i, 1)
	Next
	msgText = "Number of vehicles in network: " + CStr(Vissim.Net.Vehicles.Count) + vbcrlf + vbcrlf _
          + "Information of last vehicle:" + vbCrlf _
          + "vehicle number = " + CStr(vehNo) + vbCrlf _
          + "Type = " + CStr(vehType) + vbCrlf _
          + "Speed = " + CStr(FormatNumber(vehSpeed,1)) + vbCrlf _
          + "Pos = " + CStr(FormatNumber(vehPos,2)) + vbCrlf _
          + "Link & lane = " + CStr(vehLinkLane)
  MsgBox msgText,, "GetMultiAttValues"

	' .......................................................................................
	' Method 5: Accessing all attributes directly using "GetMultipleAttributes"
	Dim allVehAttributes
	allVehAttributes = Vissim.Net.Vehicles.GetMultipleAttributes(Array("No", "VehType", "Speed", "Pos", "Lane"))
	For i = 0 To UBound(allVehAttributes)
			vehNo = allVehAttributes(i, 0)
			vehType = allVehAttributes(i, 1)
			vehSpeed = allVehAttributes(i, 2)
			vehPos = allVehAttributes(i, 3)
			vehLinkLane = allVehAttributes(i, 4)
	Next
	msgText = "Number of vehicles in network: " + CStr(Vissim.Net.Vehicles.Count) + vbcrlf + vbcrlf _
          + "Information of last vehicle:" + vbCrlf _
          + "vehicle number = " + CStr(vehNo) + vbCrlf _
          + "Type = " + CStr(vehType) + vbCrlf _
          + "Speed = " + CStr(FormatNumber(vehSpeed,1)) + vbCrlf _
          + "Pos = " + CStr(FormatNumber(vehPos,2)) + vbCrlf _
          + "Link & lane = " + CStr(vehLinkLane)
  MsgBox msgText,, "GetMultipleAttributes"

	' -----------------------------------------------------------------------------------------
	' Operations on a specific vehicle
	' -----------------------------------------------------------------------------------------
	vehNo = 1
	Set vehicle = Vissim.Net.Vehicles.ItemByKey(vehNo)
	' The following option will retrieve the second vehicle of the current list of all vehicles -
	' NOT the vehicle with number 1:
	'    allVehicles = Vissim.Net.Vehicles.GetAll    ' get all vehicles in the network at the current simulation second
	'    Set vehicle = allVehicles(1)
	MsgBox "Desired speed of vehicle " + CStr(vehNo) + " is currently " + CStr(FormatNumber(vehicle.AttValue("DesSpeed"),1)) + "."

	' Set desired speed
	Dim desSpeedNew
	desSpeedNew = 30
	vehicle.AttValue("DesSpeed") = desSpeedNew
	MsgBox "Desired speed of vehicle " + CStr(vehNo) + " is now set to " + CStr(vehicle.AttValue("DesSpeed")) + "."
	
	' Move to specific link position
	Dim linkNo, laneNo, linkCoordinate
	linkNo = 1
	laneNo = 1
	linkCoordinate = 70
	MsgBox "Vehicle " + CStr(vehNo) + " is currently travelling on link " + vehicle.AttValue("Lane") + "." + vbCrlf _
	      + "It will now be moved to link " + CStr(linkNo) + ", coordinate " + CStr(FormatNumber(linkCoordinate,2)) + "."
	vehicle.MoveToLinkPosition linkNo, laneNo, linkCoordinate 	' Result of this command is reflected in the subsequent simulation time step
	' Run a single simulation step to see the vehicle being moved
	Vissim.Simulation.RunSingleStep 

	' Remove vehicle from the network
	vehNo = vehicle.AttValue("No")
	MsgBox "Vehicle " + CStr(vehNo) + " will shortly be removed from the network..."
	Vissim.Net.Vehicles.RemoveVehicle vehNo

	' Generate a new vehicle and place it into the network
	vehType = 100
	linkNo = 1
	laneNo = 1
	linkCoordinate = 15 ' unit according to the user setting in Vissim [m or ft]
	desSpeedNew = 53 ' unit according to the user setting in Vissim [km/h or mph]
	Dim interaction
	interaction = True 	' optional parameter

	MsgBox "A new vehicle will be added shortly on link " + CStr(linkNo) + ", coordinate " + CStr(FormatNumber(linkCoordinate,2)) + " with desired speed of " + CStr(desSpeedNew) + "."
	Vissim.Net.Vehicles.AddVehicleAtLinkPosition vehType, linkNo, laneNo, linkCoordinate, desSpeedNew, interaction
	' Note: In earlier Vissim releases, the name of the function was: AddVehicleAtLinkCoordinate

  Vissim.Simulation.RunContinuous 	' Run the simulation until simulation end
	
End Sub

'==========================================================================
Sub Screenshot2D()	
	' Generate a screenshot of the 2D view of the current Vissim network editor

	' Check if 2D mode is currently active. If not, enable it.
	Dim restoreGraphicsMode
	restoreGraphicsMode = false
	if Vissim.Graphics.CurrentNetworkWindow.AttValue("3D") <> 0 then
		Vissim.Graphics.CurrentNetworkWindow.AttValue("3D") = 0		' Enable 2D Mode
		restoreGraphicsMode = true
	end if

	' Define the desired network section for the screenshot by zooming within the current network editor:
	' 'ZoomTo' maximizes the network in the current network editor within the bounding rectangle 
	' defined by Vissim world coordinates of the lower left (x1, y1) and the upper right (x2, y2) corner.
	' If the rectangle proportions differ from the proportions of the current network editor, 
	' the specified bounding rectangle will be centred within the network editor.
	Dim x1, y1, x2, y2
	x1 = 250
	y1 = 30
	x2 = 350
	y2 = 135
	Vissim.Graphics.CurrentNetworkWindow.ZoomTo x1, y1, x2, y2

	' Generate screenshot from the current 2D view of the Vissim network editor
	' and save it to an image file. The image file format is determined by the file extension.
	' Possible options are: PNG, TIFF, GIF, JPG, JPEG or BMP.
	Dim filenameScreenshot
	filenameScreenshot = "Screenshot 2D.jpg"        ' you may optionally include also a path
	Dim sizeFactor
	sizeFactor = 1 			' 1.0 = original size, 0.5 half the size. 
											' A factor > 1 does not increase the screen resolution but does a 'digital zoom'
	Vissim.Graphics.CurrentNetworkWindow.Screenshot filenameScreenshot, sizeFactor
  MsgBox "The screenshot was saved to " + filenameScreenshot,, "Screenshot 2D"
  
	if restoreGraphicsMode then 
	' restore the original graphics mode
		Vissim.Graphics.CurrentNetworkWindow.AttValue("3D") = 1
	end if
	
End Sub

'==========================================================================
Sub Screenshot3D()
	' Generate screenshot from a 3D view of the Vissim network editor
	' Hint: Best results are achieved if "Presentation - 3D Anti-Aliasing" is enabled
	
	' Check if 3D mode is already active. If not, enable it.
	Dim restoreGraphicsMode
	restoreGraphicsMode = false
	if Vissim.Graphics.CurrentNetworkWindow.AttValue("3D") <> 1 then
		Vissim.Graphics.CurrentNetworkWindow.AttValue("3D") = 1		' Enable 3D Mode
		restoreGraphicsMode = true
	end if
	
	
	' Set camera position and viewing angle:
	Dim camPosX, camPosY, camPosZ
	Dim yawAngle, pitchAngle
	camPosX = 270
	camPosY = 30
	camPosZ = 15
	yawAngle = 45
	pitchAngle = 10
	Vissim.Graphics.CurrentNetworkWindow.SetCameraPositionAndAngle camPosX, camPosY, camPosZ, yawAngle, pitchAngle

	Dim filenameScreenshot
	filenameScreenshot = "Screenshot 3D.jpg" 		' you may optionally include also a path
	Dim sizeFactor
	sizeFactor = 1.0 			' 1.0 = original size, 0.5 half the size. 
												' A factor > 1 does not increase the screen resolution but only does a 'digital zoom' of the resulting image file. 
	Vissim.Graphics.CurrentNetworkWindow.Screenshot filenameScreenshot, sizeFactor
  MsgBox "The screenshot was saved to " + filenameScreenshot,, "Screenshot 3D"

	if restoreGraphicsMode then 
	' restore the original graphics mode
		Vissim.Graphics.CurrentNetworkWindow.AttValue("3D") = 0
	end if
	
End Sub

'==========================================================================
' Simulation Results
'==========================================================================
Sub Run3Simulations()
	' Run several short simulation runs and return some evaluations

  ' Ensure that all previous results are deleted in order to start clean
  Dim simRun
  for each simRun in Vissim.Net.SimulationRuns
    Vissim.Net.SimulationRuns.RemoveSimulationRun(simRun)
  next
  
	' To maximize simulation speed ...
	' ... activate QuickMode:
	Vissim.Graphics.CurrentNetworkWindow.AttValue("QuickMode") = 1
	
	' ... stop updating the Vissim workspace (network editors, lists, charts and other Vissim windows) 
	'     while COM script controls the simulation:
	Vissim.SuspendUpdateGUI
	
	' Set the simulation time bounds
	Vissim.Simulation.AttValue("SimPeriod") = 600	'[Simulation second]
	Vissim.Simulation.AttValue("SimBreakAt") = 0  '[Simulation second]. If set to 0 then the simulation is not paused.

	' Run the simulation three times
	Dim simRunIndex
	For simRunIndex = 1 To 3
		Vissim.Simulation.AttValue("RandSeed") = simRunIndex		' ensure that each simulation run is done with a different random seed
		Vissim.Simulation.RunContinuous
	Next
	
	' restore the Vissim workspace back to normal operation and deactivate quick mode
	Vissim.ResumeUpdateGUI (true)
	Vissim.Graphics.CurrentNetworkWindow.AttValue("QuickMode") = 0
End Sub

'==========================================================================
Sub GetSimulationRunResults()
	' Fetch a list of all simulation runs:

	Dim attributes
	attributes = Array("Timestamp", "RandSeed", "SimEnd")
	Dim allSimulationRuns
	allSimulationRuns = Vissim.Net.SimulationRuns.GetMultipleAttributes(attributes)

	Dim resultString
	resultString = ""
	Dim simRun, attIndex
	For simRun = 0 To UBound(allSimulationRuns)
			For attIndex = 0 To UBound(attributes)
					resultString = resultString + CStr(allSimulationRuns(simRun, attIndex)) + ", "
			Next
			resultString = resultString + vbCrlf
	Next
	MsgBox resultString, vbOKOnly+vbInformation ,"All simulation runs"		' send the results to a message box
End Sub

'==========================================================================
Sub GetTravelTimeResults()
	' Get the results of a specific vehicle travel time measurement

	' Syntax to get the travel time data:
	' vehTravelTimeMeasurement.AttValue("TravTm( subAttribute1, subAttribute2, subAttribute3 )"
	'
	' 	subAttribute1 = simulation run. Possible values:
	'   	1, 2, 3, ..., Current:    run index: get data from a specific simulation run (run index according 
	'                               to the attribute "No" of simulation runs (see list of simulation runs))
	'     Avg, StdDev, Min, Max:    aggregated data over all simulations runs
	'
	'   subAttribute2 = time interval. Possible values:
	'   	1, 2, 3, ..., Last:       time interval index (always starts at 1 (first time interval)): get data from a specific time interval
	'     Avg, StdDev, Min, Max:    aggregated data over all time intervals of a single simulation run
	'     Total:                    sum over all time interval of a single simulation run
	' 
	' 	subAttribute3 = vehicle class. Possible values:
	'     10, 20, ... or All:				Vehicle class number: get data only for the specified vehicle class
	'                               Note: Result data for specific vehicle classes is available only if it was configured 
	' 																		in Evaluation - Configuration - Result Attributes PRIOR to running the simulation(s) 
	
	Dim vehTravelTimeMeasurement, vehTravTimeMeasurementNo
	vehTravTimeMeasurementNo = 2
	Set vehTravelTimeMeasurement = Vissim.Net.VehicleTravelTimeMeasurements.ItemByKey(vehTravTimeMeasurementNo)

	Dim travTime, vehCount
	
	' ............................................................................
	' Example 1:	Get for all vehicle classes (subAttribute 3 = All)
	'             the average data over all time intervals (subAttribute 2 = Avg)
	'             and over all simulation runs (subAttribute 1 = Avg)
	' ............................................................................
	travTime = vehTravelTimeMeasurement.AttValue("TravTm(Avg,Avg,All)")
	vehCount = vehTravelTimeMeasurement.AttValue("Vehs(Avg,Avg,All)")
	Dim msgText
  msgText = "Vehicle travel time measurement no. " + CStr(vehTravTimeMeasurementNo) + ":" + vbcrlf _
				+ "Average travel time" + vbcrlf _
        + "� of all vehicle classes" + vbcrlf _
				+ "� over all time intervals" + vbcrlf _
				+ "� and over all simulation runs" + vbcrlf _ 
				+ "= " + CStr(FormatNumber(travTime,1)) + " s for " + CStr(vehCount) + " vehicles."
  MsgBox msgText,,"Travel Times"
        
	' ............................................................................
	' Example 2:	Get for all HGVs (subAttribute3 = 20) 
  '             the maximum data over all time intervals (subAttribute2 = Max) 
	' 						of the current simulation run (subAttribute1 = Current)
	' ............................................................................
	travTime = vehTravelTimeMeasurement.AttValue("TravTm(Current,Max,20)")
	vehCount = vehTravelTimeMeasurement.AttValue("Vehs(Current,Max,20)")
	msgText = "Vehicle travel time measurement no. " + CStr(vehTravTimeMeasurementNo) + ":" + vbcrlf _
				+ "Maximum travel time" + vbcrlf _
				+ "� of all HGVs" + vbcrlf _
				+ "� over all time intervals" + vbcrlf _
				+ "� of the current simulation run" + vbcrlf _ 
				+ "= " + CStr(FormatNumber(travTime,1)) + " s for " + CStr(vehCount) + " vehicles."
  MsgBox msgText,,"Travel Times"

	' ............................................................................
	' Example 3:	Get for all vehicle classes (subAttribute3 = All) 
	'							the data of the 1st time interval (subAttribute2 = 1) 
	' 						of the 3rd simulation run (subAttribute1 = 3)
	' ............................................................................
	travTime = vehTravelTimeMeasurement.AttValue("TravTm(3,1,All)")
	vehCount = vehTravelTimeMeasurement.AttValue("Vehs(3,1,All)")
	msgText = "Vehicle travel time measurement no. " + CStr(vehTravTimeMeasurementNo) + ":" + vbcrlf _
				+ "Travel time of " + vbcrlf _
				+ "� all vehicle classes" + vbcrlf _
				+ "� of the 1st time interval" + vbcrlf _
				+ "� of the 2nd simulation run" + vbcrlf _ 
				+ "= " + CStr(FormatNumber(travTime,1)) + " s for " + CStr(vehCount) + " vehicles."
  MsgBox msgText,,"Travel Times"
        
End Sub
	
'==========================================================================
Sub GetDataCollectionResults()	
	' Get the results of a specific data collection measurement

	' Syntax to get the data collection data:
	' dataCollMeasurement.AttValue("Vehs(subAttribute1, subAttribute2, subAttribute3)")
	'
	' 	subAttribute1 = simulation run (same as described for vehicle travel time measurements)
	' 	subAttribute2 = time interval (same as described for vehicle travel time measurements)
	'   subAttribute3 = vehicle class (same as described for vehicle travel time measurements)

	Dim dataCollMeasurement, dataCollMeasurementNo
	dataCollMeasurementNo = 1
	Set dataCollMeasurement = Vissim.Net.DataCollectionMeasurements.ItemByKey(dataCollMeasurementNo)

	Dim vehCount, speed, acceleration, length

	' ............................................................................
	' Example 1:	Get for all vehicle classes (subAttribute3 = All)
	'							the data of the 1st time interval (subAttribute2 = 1)
	'							averaged over all simulation runs (subAttribute1 = Avg)
	' ............................................................................
	vehCount = dataCollMeasurement.AttValue("Vehs(Avg,1,All)")
	speed = dataCollMeasurement.AttValue("Speed(Avg,1,All)")
	acceleration = dataCollMeasurement.AttValue("Acceleration(Avg,1,All)")
	length = dataCollMeasurement.AttValue("Length(Avg,1,All)")
	
  Dim msgText
  msgText = "Vehicle data of data collection measurement no. " + CStr(dataCollMeasurementNo) + vbcrlf _
				+ "� of all vehicle classes" + vbcrlf _
				+ "� of the 1st time interval" + vbcrlf _
				+ "� averaged over all simulation runs:" + vbcrlf + vbcrlf _ 
				+ "Vehicle count = " + CStr(vehCount) + vbcrlf _ 
				+ "Speed = " + CStr(FormatNumber(Speed,1)) + vbcrlf _ 
				+ "Acceleration = " + CStr(FormatNumber(Acceleration,1)) + vbcrlf _ 
				+ "Vehicle length = " + CStr(FormatNumber(Length,1))
  MsgBox msgText,,"Data Collection"

End Sub	

'==========================================================================
Sub GetQueueLengthResults()
	' Get the results of a specific queue counter measurement

	' Syntax to get the data:
	' QueueCounter.AttValue("QLen(subAttribute1, subAttribute2)")
	'
	' 	subAttribute1: simulation run (same as described for vehicle travel time measurements)
	' 	subAttribute2: time interval (same as described for vehicle travel time measurements)

	' ............................................................................
	' Example 1:	Get the average data over all time intervals (subAttribute2 = Avg)
	'							and over all simulation runs (subAttribute1 = Avg)
	' ............................................................................
	Dim queueCounterNo, queueLenAvg
	queueCounterNo = 1
	queueLenAvg = Vissim.Net.QueueCounters.ItemByKey(queueCounterNo).AttValue("QLen(Avg, Avg)")
	
  Dim msgText
  msgText = "Queue counter no. " + CStr(queueCounterNo) + ":" + vbcrlf _
				+ "Average queue length of all time intervals" + vbcrlf _
				+ "and over all simulation runs = " + CStr(FormatNumber(queueLenAvg,1))
  MsgBox msgText,,"Queue Length"
        
End Sub

'==========================================================================
' Saving
'==========================================================================
Sub SaveNetworkAndLayout()
  Dim path
  Dim filename
  path = "C:\Users\Public\Documents\PTV Vision\PTV Vissim 9\Examples Training\COM\Basic Commands\"
  filename = "COM Basic Commands VBS - Saved"

	Vissim.SaveNetAs (path + filename + ".inpx")
	Vissim.SaveLayout (path + filename + ".layx")
  MsgBox "The network and layout files were saved to "+ vbcrlf + vbcrlf _
         + path + vbcrlf + vbcrlf _
         + "as '" + filename + "'" ,vbInformation,"Save Network and Layout"
  
  End Sub

'==========================================================================
' End of Script
'==========================================================================