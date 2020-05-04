//= == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == =
// C++ - Script for Vissim 6 +
// Copyright (C) PTV AG, Jochen Lohmiller
// All rights reserved.
// -----------------
// Example of basic syntax - DO NOT MODIFY
//= == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == =

// This script demonstrates how to use the COM interface in C++.
// Basic commands for loading a network and layout, reading and setting
// attributes of network objects, running a simulation and retrieving
// evaluations are shown. This example is also available for the programming
// languages VBA, VBS, Python, C#, Matlab and Java.
//
// If you start using COM, please see also our COM introduction document: 
// C:\Program Files\PTV Vision\PTV Vissim 10\Doc\Eng\Vissim 10 - COM Intro.pdf
//
// For information about the attributes and methods of PTV Vissim objects, see
// the COM Help, which is located in the PTV Vissim menu: Help > COM Help. 
//
// Hint: You can easily see all attributes of the PTV Vissim objects in the lists in
// Vissim. In case your PTV Vissim language is set to English, the name in the
// headline of the list correspond to the command to access via COM.
// Example: If you want to access the number of lanes of a link, go to PTV Vissim
// and see the headline of "Number of lanes" in the list, which is
// "NumLanes". To access this attribute via COM use:
// Vissim->Net->Links->GetItemByKey(1)->GetAttValue("NumLanes");

// import the *.exe file providing the COM interface
// use your Vissim installation path here as the example shows
#import "C:\\Program Files\\PTV Vision\\PTV Vissim 10\\Exe\\Vissim10.exe" rename_namespace ("VISSIMLIB")

// #include <string>
#include <iostream>
#include <iomanip>

// using namespace std;
using std::cout;
using std::endl;


int main(int argc, char* argv[])
{
	// initialize COM
	CoInitialize(NULL);
	{
		// create Vissim object
		VISSIMLIB::IVissimPtr Vissim;
		Vissim.CreateInstance("Vissim.Vissim");
		// If you have installed multiple Vissim Versions, you can open a specific Vissim version adding the bit Version(32 or 64bit) and Version number
		// Vissim.CreateInstance('Vissim.Vissim-32.9'); //Vissim 9 - 32 bit
		// Vissim.CreateInstance('Vissim.Vissim-64.9'); //Vissim 9 - 64 bit
		// Vissim.CreateInstance('Vissim.Vissim-32.10'); //Vissim 10 - 32 bit
		// Vissim.CreateInstance('Vissim.Vissim-64.10'); //Vissim 10 - 64 bit
		
		bstr_t Path_of_COM_Basic_Commands_network = "C:\\Users\\Public\\Documents\\PTV Vision\\PTV Vissim 10\\Examples Training\\COM\\Basic Commands\\"; // always use \\ at the end

		// Load a Vissim Network :
		bstr_t Filename = Path_of_COM_Basic_Commands_network + "COM Basic Commands.inpx";
		cout << Filename << endl;
		bool flag_read_additionally = false; // you can read network(elements) additionally, in this case set "flag_read_additionally" to true
		Vissim->LoadNet(Filename, flag_read_additionally);

		// Load a Layout :
		Filename = Path_of_COM_Basic_Commands_network + "COM Basic Commands.layx";
		Vissim->LoadLayout(Filename);

		//== == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == ==
		// Read and Set attributes(vehicles)
		//= == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == =
		// Note: All of the following commands can also be executed during a simulation.

		// Read Link Name :
		int Link_number = 1;
		bstr_t Name_of_Link = Vissim->Net->Links->GetItemByKey(1)->GetAttValue("Name");
		cout << "Name of Link(" << Link_number << "): " << Name_of_Link << endl;

		// Set Link Name :
		bstr_t new_Name_of_Link = "New Link Name";
		Vissim->Net->Links->GetItemByKey(Link_number)->PutAttValue("Name", new_Name_of_Link);

		// Set a signal controller program :
		int SC_number = 1; //SC = SignalController
		VISSIMLIB::ISignalControllerPtr SignalController = Vissim->Net->SignalControllers->GetItemByKey(SC_number);
		int new_signal_programm_number = 2;
		SignalController->PutAttValue("ProgNo", new_signal_programm_number);

		// Set relative flow of a static vehicle route of a static vehicle routing decision :
		int SVRD_number = 1; //SVRD = Static Vehicle Routing Decision
		int SVR_number = 1; //SVR = Static Vehicle Route(of a specific Static Vehicle Routing Decision)
		double new_relativ_flow = 0.6;
		Vissim->Net->VehicleRoutingDecisionsStatic->GetItemByKey(SVRD_number)->VehRoutSta->GetItemByKey(SVR_number)->PutAttValue("RelFlow(1)", new_relativ_flow);
		//'RelFlow(1)' means the first defined time interval; to access the third defined time interval : 'RelFlow(3)'

		// Set vehicle input :
		int VI_number = 1; //VI = Vehicle Input
		double new_volume = 600; //vehicles per hour
		Vissim->Net->VehicleInputs->GetItemByKey(VI_number)->PutAttValue("Volume(1)", new_volume);
		// "Volume(1)" means the first defined time interval
		// Hint: The Volumes of following intervals Volume(i) i = 2...n can only be
		// edited, if continuous is deactivated : (otherwise error : "AttValue failed: Object 2: Attribute Volume (300) is no subject to changes.")
		Vissim->Net->VehicleInputs->GetItemByKey(VI_number)->PutAttValue("Cont(2)", false);
		Vissim->Net->VehicleInputs->GetItemByKey(VI_number)->PutAttValue("Volume(2)", 400);

		// Set vehicle composition :
		int Veh_composition_number = 1;
		int number_Rel_Flows = Vissim->Net->VehicleCompositions->GetItemByKey(Veh_composition_number)->VehCompRelFlows->GetCount();
		// Direct method also possible (casting Rel_Flows):
		variant_t Rel_Flows = Vissim->Net->VehicleCompositions->GetItemByKey(Veh_composition_number)->VehCompRelFlows->GetAll();
		VARIANT VehCompRelFlowVar;
		LONG Index;
		Index = 0;
		HRESULT hr = SafeArrayGetElement(Rel_Flows.parray, &Index, &VehCompRelFlowVar); assert(SUCCEEDED(hr));
		VISSIMLIB::IVehicleCompositionRelativeFlowPtr VehCompRelFlow = _variant_t(VehCompRelFlowVar);

		VehCompRelFlow->PutAttValue("VehType", 100); // Changing the vehicle type
		VehCompRelFlow->PutAttValue("DesSpeedDistr", 30); // Changing the desired speed distribution
		VehCompRelFlow->PutAttValue("RelFlow", 0.9); // Changing the relative flow
		Index = 1;
		hr = SafeArrayGetElement(Rel_Flows.parray, &Index, &VehCompRelFlowVar); assert(SUCCEEDED(hr));
		VehCompRelFlow = _variant_t(VehCompRelFlowVar);
		VehCompRelFlow->PutAttValue("RelFlow", 0.123); // Changing the relative flow of the 2nd Relative Flow.

		//// Direct method also possible (without safearray):
		//VISSIMLIB::IIteratorPtr VehCompRelFlow_Iterator = Vissim->Net->VehicleCompositions->GetItemByKey(Veh_composition_number)->VehCompRelFlows->GetIterator();
		//int veh_type;
		//while (VehCompRelFlow_Iterator->GetValid())
		//{
		//	VISSIMLIB::IVehicleCompositionRelativeFlowPtr VehCompRelFlow = VehCompRelFlow_Iterator->GetItem();
		//	veh_type = VehCompRelFlow->GetAttValue("VehType");
		//	if (veh_type == 100)
		//	{
		//		VehCompRelFlow->PutAttValue("DesSpeedDistr", 30);
		//		VehCompRelFlow->PutAttValue("RelFlow", 0.9);
		//	}
		//	else if (veh_type == 200)
		//	{
		//		VehCompRelFlow->PutAttValue("RelFlow", 0.123);
		//	}
		//	VehCompRelFlow_Iterator->Next();
		//}

		// ========================================================================
		// Accessing Multiple Attributes:
		//========================================================================

		// GetMultiAttValues         Read one attribute of all objects:
		bstr_t Attribute = "Name";
		variant_t Name_of_Links = Vissim->Net->Links->GetMultiAttValues(Attribute);

		// SetAllAttValues           Set all attributes of one object to one value:
		Attribute = "Name";
		bstr_t Link_Name = "All Links have the same Name";
		Vissim->Net->Links->SetAllAttValues(Attribute, Link_Name, false); // Note the method SetAllAttValues has a 3rd optional input: Optional ByVal add As Boolean = False; Use only for numbers!
		Attribute = "CostPerKm";
		double Cost = 5.5;
		Vissim->Net->Links->SetAllAttValues(Attribute, Cost, false);
		// Note the method SetAllAttValues has a 3rd optional input: Optional ByVal add As Boolean = False; Use only for numbers!
		Vissim->Net->Links->SetAllAttValues(Attribute, Cost, true); // setting the 3rd input to true, will add 5.5 to all previous costs!


		//== == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == ==
		// Simulation
		//= == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == =

		// Chose Random Seed
		int Random_Seed = 42;
		Vissim->Simulation->PutAttValue("RandSeed", Random_Seed);

		// To start a simulation you can run a single step :
		Vissim->Simulation->RunSingleStep();
		// Or run the simulation continuous(it stops at breakpoint or end of simulation)
		double End_of_simulation = 600; // simulation second[s]
		Vissim->Simulation->PutAttValue("SimPeriod", End_of_simulation);
		double Sim_break_at = 200; // simulation second[s]
		Vissim->Simulation->PutAttValue("SimBreakAt", Sim_break_at);
		// Set maximum speed :
		Vissim->Simulation->PutAttValue("UseMaxSimSpeed", true);
		// Hint: to change the speed use : Vissim->Simulation->PutAttValue("SimSpeed", 10); // 10 = > 10 Sim. sec. / s
		Vissim->Simulation->RunContinuous();

		// To stop the simulation:
		Vissim->Simulation->Stop();

		//== == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == ==
		// Access during simulation
		//= == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == =
		// Note: All of commands of "Read and Set attributes (vehicles)" can also be executed during a
		// simulation(e.g.changing signal controller program, setting relative flow of a static vehicle route,
		// changing the vehicle input, changing the vehicle composition).

		Sim_break_at = 198; // simulation second[s]
		Vissim->Simulation->PutAttValue("SimBreakAt", Sim_break_at);
		Vissim->Simulation->RunContinuous(); // Start the simulation until SimBreakAt(198s)

		//Get the state of a signal head :
		int SH_number = 1; // SH = SignalHead
		bstr_t State_of_SH = Vissim->Net->SignalHeads->GetItemByKey(SH_number)->GetAttValue("SigState"); // possible output e.g. 'GREEN', 'RED', 'AMBER', 'REDAMBER'
		cout << "Actual state of SignalHead(" << SH_number << ") is: " << State_of_SH << endl;

		// Set the state of a signal controller:
		// Note: Once a state of a signal group is set, the attribute "ContrByCOM" is automatically set to True. Meaning the signal group will keep this state until another state is set by COM or the end of the simulation
		// To switch back to the defined signal controller, set the attribute signal "ContrByCOM" to False (example see below).
		SC_number = 1; // SC = SignalController
		int SG_number = 1; // SG = SignalGroup
		SignalController = Vissim->Net->SignalControllers->GetItemByKey(SC_number);
		VISSIMLIB::ISignalGroupPtr SignalGroup = SignalController->SGs->GetItemByKey(SG_number);
		bstr_t new_state = "GREEN"; //possible values e.g. "GREEN", "RED", "AMBER", "REDAMBER"
		SignalGroup->PutAttValue("SigState", new_state);
		// Note: The signal controller can only be called at whole simulation seconds, so the state will be set in Vissim at the next whole simulation second, here 199s
		// Simulate so that the new state is active in the Vissim simulation:
		Sim_break_at = 200; // simulation second [s]
		Vissim->Simulation->PutAttValue("SimBreakAt", Sim_break_at);
		Vissim->Simulation->RunContinuous(); // Start the simulation until SimBreakAt(200s)
		// Give the control back:
		SignalGroup->PutAttValue("ContrByCOM", false);

		// Information about all vehicles in the network(in the current simulation second) :
		// In the following, 4 different methods to access attributes are shown :
		// Method #1: Loop over all Vehicles using "GetAll"
		// Method #2: Loop over all Vehicles using Object Enumeration (not supported in combination with C++)
		// Method #3: Using the Iterator
		// Method #4: Accessing all attributes directly using "GetMultiAttValues"
		// The result of the four methods is the same (except the format).

		VARIANT VehicleVar;
		VISSIMLIB::IVehiclePtr Vehicle;
		int veh_number;
		int veh_type;
		double veh_speed;
		double veh_position;
		bstr_t veh_linklane;

		// Method #1: Loop over all Vehicles:
		//number_vehicles = Vissim->Net->Vehicles->Count;
		variant_t All_Vehicles = Vissim->Net->Vehicles->GetAll(); //get all vehicles in the network at the actual simulation second
		LONG Count = 0;
		hr = SafeArrayGetUBound(All_Vehicles.parray, 1, &Count);
		assert(SUCCEEDED(hr));
		for (LONG Index = 0; Index < Count; ++Index) {
			hr = SafeArrayGetElement(All_Vehicles.parray, &Index, &VehicleVar);
			assert(SUCCEEDED(hr));
			Vehicle = _variant_t(VehicleVar);
			veh_number = Vehicle->GetAttValue("No");
			veh_type = Vehicle->GetAttValue("VehType");
			veh_speed = Vehicle->GetAttValue("Speed");
			veh_position = Vehicle->GetAttValue("Pos");
			veh_linklane = Vehicle->GetAttValue("Lane");
			cout << std::setprecision(4) << veh_number << "  |  " << veh_type << "  |  " << veh_speed << "  |  " << veh_position << "  |  " << veh_linklane << endl;
		}


		//Method #2: Loop over all Vehicles using Object Enumeration (for each loop is not supported with VISSIMLIB::IVehiclePtr)
		//for each (VISSIMLIB::IVehiclePtr Vehicle in Vissim->Net->GetVehicles())


		// Method #3: Using the Iterator
		VISSIMLIB::IIteratorPtr Vehicles_Iterator = Vissim->Net->Vehicles->GetIterator();
		while (Vehicles_Iterator->GetValid())
		{
			Vehicle = Vehicles_Iterator->GetItem();
			veh_number = Vehicle->GetAttValue("No");
			veh_type = Vehicle->GetAttValue("VehType");
			veh_speed = Vehicle->GetAttValue("Speed");
			veh_position = Vehicle->GetAttValue("Pos");
			veh_linklane = Vehicle->GetAttValue("Lane");
			cout << std::setprecision(4) << veh_number << "  |  " << veh_type << "  |  " << veh_speed << "  |  " << veh_position << "  |  " << veh_linklane << endl;
			Vehicles_Iterator->Next();
		}

		// Method #4: Accessing all Attributes directly using "GetMultiAttValues"
		// Safearrays from GetMultiAttValues have two dimensions 1. column: consecutive number; 2. column: AttValue
		VARIANT veh_numberVar;
		VARIANT veh_typesVar;
		VARIANT veh_speedsVar;
		VARIANT veh_positionsVar;
		VARIANT veh_linklanesVar;
		variant_t veh_numbers = Vissim->Net->Vehicles->GetMultiAttValues("No");
		variant_t veh_types = Vissim->Net->Vehicles->GetMultiAttValues("VehType");
		variant_t veh_speeds = Vissim->Net->Vehicles->GetMultiAttValues("Speed");
		variant_t veh_positions = Vissim->Net->Vehicles->GetMultiAttValues("Pos");
		variant_t veh_linklanes = Vissim->Net->Vehicles->GetMultiAttValues("Lane");
		Count = 0;
		hr = SafeArrayGetUBound(veh_numbers.parray, 1, &Count);
		assert(SUCCEEDED(hr));
		for (LONG Index = 0; Index < Count; ++Index)
		{
			LONG Indexx[] = { Index, 1 }; // Safearrays from GetMultiAttValues have two dimensions 1. column: consecutive number; 2. column: AttValue; we want only the 2nd column.
			hr = SafeArrayGetElement(veh_numbers.parray, Indexx, &veh_numberVar); assert(SUCCEEDED(hr));
			veh_number = _variant_t(veh_numberVar);
			hr = SafeArrayGetElement(veh_types.parray, Indexx, &veh_typesVar); assert(SUCCEEDED(hr));
			veh_type = _variant_t(veh_typesVar);
			hr = SafeArrayGetElement(veh_speeds.parray, Indexx, &veh_speedsVar); assert(SUCCEEDED(hr));
			veh_speed = _variant_t(veh_speedsVar);
			hr = SafeArrayGetElement(veh_positions.parray, Indexx, &veh_positionsVar); assert(SUCCEEDED(hr));
			veh_position = _variant_t(veh_positionsVar);
			hr = SafeArrayGetElement(veh_linklanes.parray, Indexx, &veh_linklanesVar); assert(SUCCEEDED(hr));
			veh_linklane = _variant_t(veh_linklanesVar);
			cout << std::setprecision(4) << veh_number << "  |  " << veh_type << "  |  " << veh_speed << "  |  " << veh_position << "  |  " << veh_linklane << endl;
		}

		// Operations at one specific vehicle :
		veh_number = 66; // Note: The vehicle number has to exists !!!
		Vehicle = Vissim->Net->Vehicles->GetItemByKey(veh_number);

		//	Set desired speed to a vehicle :
		double new_desspeed = 30;
		Vehicle->PutAttValue("DesSpeed", new_desspeed);

		//Move a vehicle :
		int link_number = 1;
		int lane_number = 1;
		double link_coordinate = 70;
		Vehicle->MoveToLinkPosition(link_number, lane_number, link_coordinate); // This function will operate during the next simulation step
		// Hint: In earlier Vissim releases, the name of the function was : MoveToLinkCoordinate

		Vissim->Simulation->RunSingleStep(); //Next Step, so that the vehicle gets moved.

		// Remove a vehicle :
		veh_number = Vehicle->GetAttValue("No");
		Vissim->Net->Vehicles->RemoveVehicle(veh_number);

		// Putting a new vehicle to the network :
		int vehicle_type = 100;
		double desired_speed = 53; //unit according to the user setting in Vissim[km / h or mph]
		int link = 1;
		int lane = 1;
		double xcoordinate = 15; //unit according to the user setting in Vissim[m or ft]
		bool interaction = true; //optional boolean
		VISSIMLIB::IVehiclePtr new_Vehicle = Vissim->Net->Vehicles->AddVehicleAtLinkPosition(vehicle_type, link, lane, xcoordinate, desired_speed, interaction);
		// Note: In earlier Vissim releases, the name of the function was : AddVehicleAtLinkCoordinate

		// Make Screenshots of the intersection 2D and 3D:
		// ZoomTo:
		// Zooms the view to the rectangle defined by the two points  (x1, y1) and (x2,y2),  which  are  given  in  world coordinates.  If  the  rectangle  proportions  differ
		// from  the  proportions  of  the  network  window,  the  specified  rectangle  will  be centred in the network editor window.
		int X1 = 250;
		int Y1 = 30;
		int X2 = 350;
		int Y2 = 135;
		Vissim->Graphics->CurrentNetworkWindow->ZoomTo(X1, Y1, X2, Y2);

		// Make a Screenshot in 2D:
		// It  creates  a  graphic  file  of  the  VISSIM  main  window  formatted  according  to its extension: PNG, TIFF, GIF, JPG, JPEG or BMP. A BMP file will be written if the extension can not be recognized.
		bstr_t Filename_screenshot = "screenshot2D.jpg"; // to set to a specific path: "C:\\Screenshots\\screenshot2D.jpg"
		int sizeFactor = 1; // 1: original size, 2: doubles size
		Vissim->Graphics->CurrentNetworkWindow->Screenshot(Filename_screenshot, sizeFactor);

		// Make a Screenshot in 3D:
		// Set 3D Mode:
		Vissim->Graphics->CurrentNetworkWindow->PutAttValue("3D", 1);
		// Set the camera position (viewing angle):
		int xPos = 270;
		int yPos = 30;
		int zPos = 15;
		int yawAngle = 45;
		int pitchAngle = 10;
		Vissim->Graphics->CurrentNetworkWindow->SetCameraPositionAndAngle(xPos, yPos, zPos, yawAngle, pitchAngle);
		Filename_screenshot = "screenshot3D.jpg"; // to set to a specific path: "C:\\Screenshots\\screenshot3D.jpg"
		Vissim->Graphics->CurrentNetworkWindow->Screenshot(Filename_screenshot, sizeFactor);

		// Set 2D Mode and old Network position:
		Vissim->Graphics->CurrentNetworkWindow->PutAttValue("3D", 0);
		X1 = -10;
		Y1 = -10;
		X2 = 600;
		Y2 = 300;
		Vissim->Graphics->CurrentNetworkWindow->ZoomTo(X1, Y1, X2, Y2);

		// Continue the simulation until end of simulation ( Vissim->Simulation->GetAttValue("SimPeriod"))
		Vissim->Simulation->RunContinuous();


		//== == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == ==
		// Results of Simulations :
		//= == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == =

		// Run 3 Simulations at maximum speed :
		// Activate QuickMode:
		Vissim->Graphics->CurrentNetworkWindow->PutAttValue("QuickMode", 1);
		Vissim->SuspendUpdateGUI(); // stop updating of the complete Vissim workspace (network editor, list, chart and signal time table windows)
		// Alternatively, load a layout (*.inpx) where dynamic elements (vehicles and pedestrians) are not visible:
		// Vissim->LoadLayout(Path_of_COM_Basic_Commands_network + "COM Basic Commands - Hide vehicles.layx"); // loading a layout where vehicles are not displayed = > much faster simulation
		End_of_simulation = 600;
		Vissim->Simulation->PutAttValue("SimPeriod", End_of_simulation);
		Sim_break_at = 0; // simulation second[s] = > 0 means no break!
		Vissim->Simulation->PutAttValue("SimBreakAt", Sim_break_at);
		// Set maximum speed:
		Vissim->Simulation->PutAttValue("UseMaxSimSpeed", true);

		for (int cnt_Sim = 1; cnt_Sim <= 3; cnt_Sim = cnt_Sim + 1)
		{
			Vissim->Simulation->PutAttValue("RandSeed", cnt_Sim);
			Vissim->Simulation->RunContinuous();
		}
		Vissim->ResumeUpdateGUI(false); // allow updating of the complete Vissim workspace (network editor, list, chart and signal time table windows)
		Vissim->Graphics->CurrentNetworkWindow->PutAttValue("QuickMode", 0); // deactivate QuickMode
		// Vissim->LoadLayout(Path_of_COM_Basic_Commands_network + "COM Basic Commands.layx"); // loading a layout to display vehicles again


		// List of all Simulation runs:
		bstr_t Attributes[3] = { "Timestamp", "RandSeed", "SimEnd" };
		bstr_t AttValue[3];
		VISSIMLIB::IIteratorPtr SimulationRuns_Iterator = Vissim->Net->SimulationRuns->GetIterator();
		while (SimulationRuns_Iterator->GetValid())
		{
			VISSIMLIB::ISimulationRunPtr SimulationRun = SimulationRuns_Iterator->GetItem();
			for (int cnt_A = 0; cnt_A <= 2; cnt_A = cnt_A + 1)
			{
				AttValue[cnt_A] = SimulationRun->GetAttValue(Attributes[cnt_A]);
			}
			cout << AttValue[0] << "  |  " << AttValue[1] << "  |  " << AttValue[2] << endl;
			SimulationRuns_Iterator->Next();
		}


		// Get the results of Vehicle Travel Time Measurements:
		double TT;	// travel time
		int No_Veh; // Number of Vehicles
		int Veh_TT_measurement_number = 2;
		VISSIMLIB::IVehicleTravelTimeMeasurementPtr Veh_TT_measurement = Vissim->Net->VehicleTravelTimeMeasurements->GetItemByKey(Veh_TT_measurement_number);
		// Syntax to get the travel times :
		//		Veh_TT_measurement->GetAttValue("TravTm(sub_attribut_1, sub_attribut_2, sub_attribut_3)");
		// sub_attribut_1: SimulationRun
		// 1, 2, 3, ... Current : the value of one specific simulation(number according to the attribute "No" of Simulation Runs(see List of Simulation Runs))
		// Avg, StdDev, Min, Max : aggregated value of all simulations runs : Avg, StdDev, Min, Max
		// sub_attribut_2 : TimeInterval
		// 1, 2, 3, ... Last	: the value of one specific time interval (number of time interval always starts at 1 (first time interval), 2 (2nd TI), 3 (3rd TI), ...)
		// Avg, StdDev, Min, Max : aggregated value of all time interval of one simulation : Avg, StdDev, Min, Max
		//       Total : sum of all time interval of one simulation
		// sub_attribut_3 : VehicleClass
		// 10, 20 or All  values only from vehicles of the defined vehicle class number(according to the attribute "No" of Vehicle Classes)
		// Note : You can only access the results of specific vehicle classes if you set it in Evaluation > Configuration > Result Attributes
		//
		// The value of on time interval is the arithmetic mean of all single travel times of the vehicles.

		// Example #1:
		// Average of all simulations (1. input = Avg)
		// of the average of all time intervals (2. input = Avg)
		// of all vehicle classes (3. input = All)
		TT = Veh_TT_measurement->GetAttValue("TravTm(Avg,Avg,All)");
		No_Veh = Veh_TT_measurement->GetAttValue("Vehs(Avg, Avg, All)");
		cout << "Average travel time all time intervals of all simulation of all vehicle classes: " << TT << "(number of vehicles: " << No_Veh << ")" << endl;

		// Example #2:
		// Value of the Current simulation (1. input = Current)
		// of the maximum of all time intervals (2. input = Max)
		// of vehicle class HGV(3. input = 20)
		TT = Veh_TT_measurement->GetAttValue("TravTm(Current,Max,20)");
		No_Veh = Veh_TT_measurement->GetAttValue("Vehs(Current,Max,20)");
		cout << "Maximum travel time of all time intervals of the current simulation of vehicle class HGV: " << TT << "(number of vehicles: " << No_Veh << ")" << endl;

		//Example #3:
		//Value of the 2nd simulation(1. input = 2)
		//of the 1st time interval(2. input = 1)
		//of all vehicle classes(3. input = All)
		// TT = Veh_TT_measurement->GetAttValue("TravTm(2,1,All)");
		// No_Veh = Veh_TT_measurement->GetAttValue("Vehs(2,1,All)");
		// cout << "Travel time of the 1st time interval of the 2nd simulation of all vehicle classes: " << TT << "(number of vehicles: " << No_Veh << ")" << endl;

		//Data Collection
		int DC_measurement_number = 1;
		VISSIMLIB::IDataCollectionMeasurementPtr DC_measurement = Vissim->Net->DataCollectionMeasurements->GetItemByKey(DC_measurement_number);
		// Syntax to get the data :
		//		DC_measurement->GetAttValue("Vehs(sub_attribut_1, sub_attribut_2, sub_attribut_3)");
		//
		// sub_attribut_1: SimulationRun(same as described at Vehicle Travel Time Measurements)
		// sub_attribut_2 : TimeInterval(same as described at Vehicle Travel Time Measurements)
		// sub_attribut_3 : VehicleClass(same as described at Vehicle Travel Time Measurements)
		//
		// The value of on time interval is the arithmetic mean of all single values of the vehicles.

		// Example #1:
		// Average value of all simulations (1. input = Avg)
		// of the 1st time interval (2. input = 1)
		// of all vehicle classes (3. input = All)
		No_Veh = DC_measurement->GetAttValue("Vehs(Avg,1,All)"); //number of vehicles
		double Speed = DC_measurement->GetAttValue("Speed(Avg,1,All)"); //Speed of vehicles
		double Acceleration = DC_measurement->GetAttValue("Acceleration(Avg,1,All)"); //Acceleration of vehicles
		double Length = DC_measurement->GetAttValue("Length(Avg,1,All)"); //Length of vehicles
		cout << "Data Collection #" << DC_measurement_number << ": Average values of all Simulations runs of 1st time interval of all vehicle classes:" << endl;
		cout << "#vehicles: " << No_Veh << "; Speed: " << Speed << "; Acceleration: " << Acceleration << "; Length: " << Length << endl;

		// Queue length
		// Syntax to get the data :
		//		QueueCounter->GetAttValue("QLen(sub_attribut_1, sub_attribut_2)");
		//
		// sub_attribut_1: SimulationRun(same as described at Vehicle Travel Time Measurements)
		// sub_attribut_2: TimeInterval(same as described at Vehicle Travel Time Measurements)

		// Example #1:
		// Average value of all simulations (1. input = Avg)
		// of the average of all time intervals (2. input = Avg)
		int	QC_number = 1;
		double maxQ = Vissim->Net->QueueCounters->GetItemByKey(QC_number)->GetAttValue("QLenMax(Avg, Avg)");
		cout << "Average maximum Queue length of all simulations and time intervals of Queue Counter #" << QC_number << ": " << maxQ << endl;

		//== == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == ==
		// Saving
		//= == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == == =
		Filename = Path_of_COM_Basic_Commands_network + "COM Basic Commands saved.inpx";
		Vissim->SaveNetAs(Filename);
		Filename = Path_of_COM_Basic_Commands_network + "COM Basic Commands saved.layx";
		Vissim->SaveLayout(Filename);


	} // Vissim closes without brackets use: Vissim->Release();

	// uninitialize COM
	CoUninitialize();

	return 0;
}
