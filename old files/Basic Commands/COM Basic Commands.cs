//==========================================================================
// C#-Script for Vissim 6+
// Copyright (C) PTV AG, Jochen Lohmiller
// All rights reserved.
// -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -
// Example of basic syntax - DO NOT MODIFY
//==========================================================================

// This script demonstrates how to use the COM interface in C#.
// Basic commands for loading a network and layout, reading and setting
// attributes of network objects, running a simulation and retrieving
// evaluations are shown. This example is also available for the programming
// languages VBA, VBS, Python, Matlab, C++ and Java.
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
// (Int32)Vissim.Net.Links.get_ItemByKey(1).AttValue["NumLanes"];
//
// Important:
// In the development environment used, set a reference to the Vissim executable
// file installed (currently this is Vissim100.exe) in order to use Vissim COM objects.
// Then the objects and methods of the Vissim COM object model can be used in the
// C# project. When using the development framework Microsoft Visual C# .NET,
// the object model can be browsed using the Object Browser similar to the object
// catalogue in VBA. Note: Embedded Interop types are not supported. Make sure to
// set the Embed Interop Types property of VISSIMLIB to false, otherwise your code
// can not be compiled.


using System;
using System.IO;
using VISSIMLIB;

namespace Example
{
    /// <summary>
    /// Class Syntax_Example shows the basic Syntax to read and write attributes to VISSIM via COM
    /// <summary>
    class Syntax_Example
    {
        /// <summary>
        /// The main entry point for the application.
        /// <summary>
        static void Main()
        {

            // Connecting the COM Server => Open a new Vissim Window:
            Vissim Vissim = new Vissim();
            // If you have installed multiple Vissim Versions, you have to set the reference to the Vissim Version you want to open.

            string Path_of_COM_Basic_Commands_network = Directory.GetCurrentDirectory();
            Path_of_COM_Basic_Commands_network = "C:\\Users\\Public\\Documents\\PTV Vision\\PTV Vissim 10\\Examples Training\\COM\\Basic Commands\\"; // always use \\ at the end !!
            string Filename = Path_of_COM_Basic_Commands_network + "COM Basic Commands.inpx";
            Vissim.LoadNet(Filename, false);

            // Load a Layout:
            Filename = Path_of_COM_Basic_Commands_network + "COM Basic Commands.layx";
            Vissim.LoadLayout(Filename);

            //// ========================================================================
            // Read and Set attributes
            //==========================================================================
            // Note: All of the following commands can also be executed during a
            // simulation.

            // Read Link Name:
            uint Link_number = 1;
            string Name_of_Link = (string)Vissim.Net.Links.get_ItemByKey(Link_number).AttValue["Name"];
            Console.WriteLine("Name of Link(" + Link_number + "): " + Name_of_Link);

            // Set Link Name:
            string new_Name_of_Link = "New Link Name";
            Vissim.Net.Links.get_ItemByKey(Link_number).set_AttValue("Name", new_Name_of_Link);

            // Set a signal controller program:
            uint SC_number = 1; // SC = SignalController
            ISignalController SignalController = Vissim.Net.SignalControllers.get_ItemByKey(SC_number);
            int new_signal_programm_number = 2;
            SignalController.set_AttValue("ProgNo", new_signal_programm_number);


            // Set relative flow of a static vehicle route of a static vehicle routing decision:
            uint SVRD_number = 1; // SVRD = Static Vehicle Routing Decision
            uint SVR_number = 1; // SVR = Static Vehicle Route (of a specific Static Vehicle Routing Decision)
            double new_relativ_flow = 0.6;
            Vissim.Net.VehicleRoutingDecisionsStatic.get_ItemByKey(SVRD_number).VehRoutSta.get_ItemByKey(SVR_number).set_AttValue("RelFlow(1)", new_relativ_flow);
            // "RelFlow(1)" means the first defined time interval; to access the third defined time interval: "RelFlow(3)"

            // Set vehicle input:
            uint VI_number = 1; // VI = Vehicle Input
            double new_volume = 600; // vehicles per hour
            Vissim.Net.VehicleInputs.get_ItemByKey(VI_number).set_AttValue("Volume(1)", new_volume);
            // "Volume(1)" means the first defined time interval
            // Hint: The Volumes of following intervals Volume(i) i = 2...n can only be
            // edited, if continuous is deactivated: (otherwise error: "AttValue failed: Object 2: Attribute Volume (300) is no subject to changes.")
            Vissim.Net.VehicleInputs.get_ItemByKey(VI_number).set_AttValue("Cont(2)", false);
            Vissim.Net.VehicleInputs.get_ItemByKey(VI_number).set_AttValue("Volume(2)", 400);

            // Set vehicle composition:
            uint Veh_composition_number = 1;
            object[] Rel_Flows = (object[])Vissim.Net.VehicleCompositions.get_ItemByKey(Veh_composition_number).VehCompRelFlows.GetAll();
            IVehicleCompositionRelativeFlow Rel_Flow0 = (IVehicleCompositionRelativeFlow)Rel_Flows[0];
            IVehicleCompositionRelativeFlow Rel_Flow1 = (IVehicleCompositionRelativeFlow)Rel_Flows[1];
            Rel_Flow0.set_AttValue("VehType", 100);   // Changing the vehicle type
            Rel_Flow0.set_AttValue("DesSpeedDistr", 50);    // Changing the desired speed distribution
            Rel_Flow0.set_AttValue("RelFlow", 0.9);   // Changing the relative flow
            Rel_Flow1.set_AttValue("RelFlow", 0.1);   // Changing the relative flow of the 2nd Relative Flow.


            //// ========================================================================
            // Accessing Multiple Attributes:
            //========================================================================

            // GetMultiAttValues         Read one attribute of all objects:
            string Attribute = "Name";
            object[,] NameOfLinks = (object[,])Vissim.Net.Links.GetMultiAttValues(Attribute);

            //// SetMultiAttValues         Set one attribute of multiple (not necessarily all) objects
            //object[,] Link_No_Name = { { 1, "New Link Name of Link #1" }, { 2, "New Link Name of Link #2" }, { 4, "New Link Name of Link #4" } };
            //Vissim.Net.Links.SetMultiAttValues(Attribute, Link_No_Name); // 1st input is the Link number, 2nd the link name

            // SetMultiAttValues         Set one attribute of multiple (not necessarily all) objects
            NameOfLinks[0, 1] = "New Link Name of Link #1";
            NameOfLinks[1, 1] = "New Link Name of Link #2";
            NameOfLinks[3, 1] = "New Link Name of Link #4";
            Vissim.Net.Links.SetMultiAttValues(Attribute, NameOfLinks); // 1st input is the consecutively number of links (not the ID), 2nd the link name
            // Please note: The first column of GetMultiAttValues or SetMultiAttValues is not the ID of an object, it is consecutively numbered all available objects.

            // GetMultipleAttributes     Read multiple attributes of all objects:
            object[] Attributes1 = { "Name", "Length2D" };
            object[,] Name_Length_of_Links = (object[,])Vissim.Net.Links.GetMultipleAttributes(Attributes1);
            //
            // SetMultipleAttributes     Set multiple attribute of multiple (always the first x) objects:
            object[] Attributes2 = { "Name", "CostPerKm" };
            object[,] Link_Name_Cost = { { "Name1", 12 }, { "Name2", 7 }, { "Name3", 5 }, { "Name4", 3 } };
            Vissim.Net.Links.SetMultipleAttributes(Attributes2, Link_Name_Cost);

            // SetAllAttValues           Set all attributes of one object to one value:
            Attribute = "Name";
            string Link_Name = "All Links have the same Name";
            Vissim.Net.Links.SetAllAttValues(Attribute, Link_Name);
            Attribute = "CostPerKm";
            double Cost = 5.5;
            Vissim.Net.Links.SetAllAttValues(Attribute, Cost);
            // Note the method SetAllAttValues has a 3rd optional input: Optional ByVal add As Boolean = False; Use only for numbers!
            Vissim.Net.Links.SetAllAttValues(Attribute, Cost, true); // setting the 3rd input to true, will add 5.5 to all previous costs!


            //// ========================================================================
            // Simulation
            //==========================================================================

            // Chose Random Seed
            int Random_Seed = 42;
            Vissim.Simulation.set_AttValue("RandSeed", Random_Seed);

            // To start a simulation you can run a single step:
            Vissim.Simulation.RunSingleStep();
            // Or run the simulation continuous (it stops at breakpoint or end of simulation)
            double End_of_simulation = 600; // simulation second [s]
            Vissim.Simulation.set_AttValue("SimPeriod", End_of_simulation);
            double Sim_break_at = 200; // simulation second [s]
            Vissim.Simulation.set_AttValue("SimBreakAt", Sim_break_at);
            // Set maximum speed:
            Vissim.Simulation.set_AttValue("UseMaxSimSpeed", true);
            // Hint: to change the speed use: Vissim.Simulation.set_AttValue("SimSpeed", 10); // 10 => 10 Sim. sec. / s
            Vissim.Simulation.RunContinuous();

            // To stop the simulation:
            Vissim.Simulation.Stop();

            //// ========================================================================
            // Access during simulation
            //==========================================================================
            // Note: All of commands of "Read and Set attributes (vehicles)" can also be executed during a
            // simulation (e.g. changing signal controller program, setting relative flow of a static vehicle route,
            // changing the vehicle input, changing the vehicle composition).

            Sim_break_at = 198; // simulation second [s]
            Vissim.Simulation.set_AttValue("SimBreakAt", Sim_break_at);
            Vissim.Simulation.RunContinuous(); // Start the simulation until SimBreakAt (198s)

            // Get the state of a signal head:
            uint SH_number = 1; // SH = SignalHead
            string State_of_SH = (string)Vissim.Net.SignalHeads.get_ItemByKey(SH_number).get_AttValue("SigState"); // possible output e.g. 'GREEN', 'RED', 'AMBER', 'REDAMBER'
            Console.WriteLine("Actual state of SignalHead(" + SH_number + ") is: " + State_of_SH);

            // Set the state of a signal controller:
            // Note: Once a state of a signal group is set, the attribute "ContrByCOM" is automatically set to True. Meaning the signal group will keep this state until another state is set by COM or the end of the simulation
            // To switch back to the defined signal controller, set the attribute signal "ContrByCOM" to False (example see below).
            SC_number = 1; // SC = SignalController
            uint SG_number = 1; // SG = SignalGroup
            SignalController = Vissim.Net.SignalControllers.get_ItemByKey(SC_number);
            ISignalGroup SignalGroup = SignalController.SGs.get_ItemByKey(SG_number);
            string new_state = "GREEN"; //possible values e.g. "GREEN", "RED", "AMBER", "REDAMBER"
            SignalGroup.set_AttValue("SigState", new_state);
            // Note: The signal controller can only be called at whole simulation seconds, so the state will be set in Vissim at the next whole simulation second, here 199s
            // Simulate so that the new state is active in the Vissim simulation:
            Sim_break_at = 200; // simulation second [s]
            Vissim.Simulation.set_AttValue("SimBreakAt", Sim_break_at);
            Vissim.Simulation.RunContinuous(); // Start the simulation until SimBreakAt (200s)
            // Give the control back:
            SignalGroup.set_AttValue("ContrByCOM", false);

            int veh_number;
            string veh_type;
            double veh_speed;
            double veh_position;
            string veh_linklane;
            // Information about all vehicles in the network (in the current simulation second):
            // In the following, 4 different methods to access attributes are shown:
            // Method #1: Loop over all Vehicles using "GetAll"
            // Method #2: Loop over all Vehicles using Object Enumeration
            // Method #3: Using the Iterator
            // Method #4: Accessing all attributes directly using "GetMultiAttValues" (fast way if you want the attributes of all vehicles)
            // Method #5: Accessing all attributes directly using "GetMultipleAttributes" (even more faster)
            // The result of the four methods is the same (except the format).

            // Method #1: Loop over all Vehicles:
            object[] All_Vehicles = (object[])Vissim.Net.Vehicles.GetAll(); // get all vehicles in the network at the actual simulation second
            for (int cnt_Veh = 0; cnt_Veh < Vissim.Net.Vehicles.Count; cnt_Veh++)
            {
                IVehicle Vehicle = (IVehicle)All_Vehicles[cnt_Veh];
                veh_number = (int)Vehicle.get_AttValue("No");
                veh_type = (string)Vehicle.get_AttValue("VehType");
                veh_speed = (double)Vehicle.get_AttValue("Speed");
                veh_position = (double)Vehicle.get_AttValue("Pos");
                veh_linklane = (string)Vehicle.get_AttValue("Lane");
                Console.WriteLine("{0}  |  {1}  |  {2:F}  |  {3:F}  |  {4}", veh_number, veh_type, veh_speed, veh_position, veh_linklane);
            }

            // Method #2: Loop over all Vehicles using Object Enumeration
            foreach (IVehicle Vehicle in Vissim.Net.Vehicles)
            {
                veh_number = (int)Vehicle.get_AttValue("No");
                veh_type = (string)Vehicle.get_AttValue("VehType");
                veh_speed = (double)Vehicle.get_AttValue("Speed");
                veh_position = (double)Vehicle.get_AttValue("Pos");
                veh_linklane = (string)Vehicle.get_AttValue("Lane");
                Console.WriteLine("{0}  |  {1}  |  {2:F}  |  {3:F}  |  {4}", veh_number, veh_type, veh_speed, veh_position, veh_linklane);
            }

            // Method #3: Using the Iterator (this method is a little bit slower)
            IIterator Vehicles_Iterator = Vissim.Net.Vehicles.Iterator;
            while (Vehicles_Iterator.Valid)
            {
                IVehicle Vehicle = (IVehicle)Vehicles_Iterator.Item;
                veh_number = (int)Vehicle.get_AttValue("No");
                veh_type = (string)Vehicle.get_AttValue("VehType");
                veh_speed = (double)Vehicle.get_AttValue("Speed");
                veh_position = (double)Vehicle.get_AttValue("Pos");
                veh_linklane = (string)Vehicle.get_AttValue("Lane");
                Console.WriteLine("{0}  |  {1}  |  {2:F}  |  {3:F}  |  {4}", veh_number, veh_type, veh_speed, veh_position, veh_linklane);
                Vehicles_Iterator.Next();
            }

            // Method #4: Accessing all Attributes directly using "GetMultiAttValues" (fast)
            object[,] veh_numbers = (object[,])Vissim.Net.Vehicles.GetMultiAttValues("No");      // Output 1. column:consecutive number; 2. column: AttValue
            object[,] veh_types = (object[,])Vissim.Net.Vehicles.GetMultiAttValues("VehType"); // Output 1. column:consecutive number; 2. column: AttValue
            object[,] veh_speeds = (object[,])Vissim.Net.Vehicles.GetMultiAttValues("Speed");   // Output 1. column:consecutive number; 2. column: AttValue
            object[,] veh_positions = (object[,])Vissim.Net.Vehicles.GetMultiAttValues("Pos");     // Output 1. column:consecutive number; 2. column: AttValue
            object[,] veh_linklanes = (object[,])Vissim.Net.Vehicles.GetMultiAttValues("Lane");    // Output 1. column:consecutive number; 2. column: AttValue
            for (int cnt_Veh = 0; cnt_Veh < veh_numbers.GetLength(0); cnt_Veh++)
            {
                Console.WriteLine("{0}  |  {1}  |  {2:F}  |  {3:F}  |  {4}", veh_numbers[cnt_Veh, 1], veh_types[cnt_Veh, 1], veh_speeds[cnt_Veh, 1], veh_positions[cnt_Veh, 1], veh_linklanes[cnt_Veh, 1]);
            }

            // Method #5: Accessing all attributes directly using "GetMultipleAttributes" (even more faster)
            object[] Attributes_veh = new object[5];
            Attributes_veh[0] = "No";
            Attributes_veh[1] = "VehType";
            Attributes_veh[2] = "Speed";
            Attributes_veh[3] = "Pos";
            Attributes_veh[4] = "Lane";
            object[,] all_veh_attributes = (object[,])Vissim.Net.Vehicles.GetMultipleAttributes(Attributes_veh);
            for (int cnt_Veh = 0; cnt_Veh < all_veh_attributes.GetLength(0); cnt_Veh++)
            {
                Console.WriteLine("{0}  |  {1}  |  {2:F}  |  {3:F}  |  {4}", all_veh_attributes[cnt_Veh, 0], all_veh_attributes[cnt_Veh, 1], all_veh_attributes[cnt_Veh, 2], all_veh_attributes[cnt_Veh, 3], all_veh_attributes[cnt_Veh, 4]);
            }

            //// Operations at one specific vehicle:
            All_Vehicles = (object[])Vissim.Net.Vehicles.GetAll(); // get all vehicles in the network at the actual simulation second
            IVehicle Vehicle1 = (IVehicle)All_Vehicles[1];
            // alternatively with ItemByKey:
            // veh_number = 66; // the same as: All_Vehicles[1].get_AttValue("No");
            // IVehicle Vehicle1 = Vissim.Net.Vehicles.get_ItemByKey(veh_number);

            // Set desired speed to a vehicle:
            double new_desspeed = 30;
            Vehicle1.set_AttValue("DesSpeed", new_desspeed);

            // Move a vehicle:
            int link_number = 1;
            int lane_number = 1;
            double link_coordinate = 70;
            Vehicle1.MoveToLinkPosition(link_number, lane_number, link_coordinate); // This function will operate during the next simulation step
            // Hint: In earlier Vissim releases, the name of the function was: MoveToLinkCoordinate

            Vissim.Simulation.RunSingleStep(); // Next Step, so that the vehicle gets moved.

            // Remove a vehicle:
            veh_number = (int)Vehicle1.get_AttValue("No");
            Vissim.Net.Vehicles.RemoveVehicle(veh_number);

            // Putting a new vehicle to the network:
            int vehicle_type = 100;
            double desired_speed = 53; // unit according to the user setting in Vissim [km/h or mph]
            int link = 1;
            int lane = 1;
            double xcoordinate = 15; // unit according to the user setting in Vissim [m or ft]
            bool interaction = true; // optional boolean
            IVehicle new_Vehicle = Vissim.Net.Vehicles.AddVehicleAtLinkPosition(vehicle_type, link, lane, xcoordinate, desired_speed, interaction);
            // Note: In earlier Vissim releases, the name of the function was: AddVehicleAtLinkCoordinate

            // Make Screenshots of the intersection 2D and 3D:
            // ZoomTo:
            // Zooms the view to the rectangle defined by the two points  (x1, y1) and (x2,y2),  which  are  given  in  world coordinates.  If  the  rectangle  proportions  differ
            // from  the  proportions  of  the  network  window,  the  specified  rectangle  will  be centred in the network editor window.
            int X1 = 250;
            int Y1 = 30;
            int X2 = 350;
            int Y2 = 135;
            Vissim.Graphics.CurrentNetworkWindow.ZoomTo(X1, Y1, X2, Y2);

            // Make a Screenshot in 2D:
            // It  creates  a  graphic  file  of  the  VISSIM  main  window  formatted  according  to its extension: PNG, TIFF, GIF, JPG, JPEG or BMP. A BMP file will be written if the extension can not be recognized.
            string Filename_screenshot = "screenshot2D.jpg"; // to set to a specific path: "C:\\Screenshots\\screenshot2D.jpg"
            int sizeFactor = 1; // 1: original Size, 2: doubles size
            Vissim.Graphics.CurrentNetworkWindow.Screenshot(Filename_screenshot, sizeFactor);

            // Make a Screenshot in 3D:
            // Set 3D Mode:
            Vissim.Graphics.CurrentNetworkWindow.set_AttValue("3D", 1);
            // Set the camera position (viewing angle):
            int xPos = 270;
            int yPos = 30;
            int zPos = 15;
            int yawAngle = 45;
            int pitchAngle = 10;
            Vissim.Graphics.CurrentNetworkWindow.SetCameraPositionAndAngle(xPos, yPos, zPos, yawAngle, pitchAngle);
            Filename_screenshot = "screenshot3D.jpg"; // to set to a specific path: "C:\\Screenshots\\screenshot3D.jpg"
            Vissim.Graphics.CurrentNetworkWindow.Screenshot(Filename_screenshot, sizeFactor);

            // Set 2D Mode and old Network position:
            Vissim.Graphics.CurrentNetworkWindow.set_AttValue("3D", 0);
            X1 = -10;
            Y1 = -10;
            X2 = 600;
            Y2 = 300;
            Vissim.Graphics.CurrentNetworkWindow.ZoomTo(X1, Y1, X2, Y2);

            // Continue the simulation until end of simulation (get(Vissim.Simulation, 'AttValue', 'SimPeriod'))
            Vissim.Simulation.RunContinuous();


            //// ========================================================================
            // Results of Simulations:
            //==========================================================================

            // Delete all previous simulation runs first:
            foreach (ISimulationRun simRun in Vissim.Net.SimulationRuns)
            {
                Vissim.Net.SimulationRuns.RemoveSimulationRun(simRun);
            }

            // Run 3 Simulations at maximum speed:
            // Activate QuickMode:
            Vissim.Graphics.CurrentNetworkWindow.set_AttValue("QuickMode", 1);
            Vissim.SuspendUpdateGUI(); // stop updating of the complete Vissim workspace (network editor, list, chart and signal time table windows)
            // Alternatively, load a layout (*.laxy) where dynamic elements (vehicles and pedestrians) are not visible:
            // Vissim.LoadLayout(Path_of_COM_Basic_Commands_network + "COM Basic Commands - Hide vehicles.layx"); // loading a layout where vehicles are not displayed => much faster simulation
            End_of_simulation = 600;
            Vissim.Simulation.set_AttValue("SimPeriod", End_of_simulation);
            Sim_break_at = 0; // simulation second [s] => 0 means no break!
            Vissim.Simulation.set_AttValue("SimBreakAt", Sim_break_at);
            // Set maximum speed:
            Vissim.Simulation.set_AttValue("UseMaxSimSpeed", true);
            for (int cnt_Sim = 1; cnt_Sim <= 3; cnt_Sim++)
            {
                Vissim.Simulation.set_AttValue("RandSeed", cnt_Sim);
                Vissim.Simulation.RunContinuous();
            }
            Vissim.ResumeUpdateGUI(false); // allow updating of the complete Vissim workspace (network editor, list, chart and signal time table windows)
            Vissim.Graphics.CurrentNetworkWindow.set_AttValue("QuickMode", 0); // deactivate QuickMode
            // Vissim.LoadLayout(Path_of_COM_Basic_Commands_network + "COM Basic Commands.layx"); // loading a layout to display vehicles again


            // List of all Simulation runs:
            object[] Attributes = new object[3];
            Attributes[0] = "Timestamp";
            Attributes[1] = "RandSeed";
            Attributes[2] = "SimEnd";
            object[,] List_Sim_Runs = (object[,])Vissim.Net.SimulationRuns.GetMultipleAttributes(Attributes);
            int number_of_Sim_Runs = Vissim.Net.SimulationRuns.Count;
            // Write the List:
            for (int cnt_C = 0; cnt_C < number_of_Sim_Runs; cnt_C++)
            {
                Console.WriteLine("{0} | {1} | {2}", List_Sim_Runs[cnt_C, 0], List_Sim_Runs[cnt_C, 1], List_Sim_Runs[cnt_C, 2]); // show the List
            }


            // Get the results of Vehicle Travel Time Measurements:
            double TT;
            int No_Veh;
            uint Veh_TT_measurement_number = 2;
            IVehicleTravelTimeMeasurement Veh_TT_measurement = Vissim.Net.VehicleTravelTimeMeasurements.get_ItemByKey(Veh_TT_measurement_number);
            // Syntax to get the travel times:
            //   Veh_TT_measurement.get_AttValue("TravTm(sub_attribut_1, sub_attribut_2, sub_attribut_3)");
            //
            // sub_attribut_1: SimulationRun
            //       1, 2, 3, ... Current:     the value of one specific simulation (number according to the attribute "No" of Simulation Runs (see List of Simulation Runs))
            //       Avg, StdDev, Min, Max:    aggregated value of all simulations runs: Avg, StdDev, Min, Max
            // sub_attribut_2: TimeInterval
            //       1, 2, 3, ... Last:        the value of one specific time interval (number of time interval always starts at 1 (first time interval), 2 (2nd TI), 3 (3rd TI), ...)
            //       Avg, StdDev, Min, Max:    aggregated value of all time interval of one simulation: Avg, StdDev, Min, Max
            //       Total:                    sum of all time interval of one simulation
            // sub_attribut_3: VehicleClass
            //       10, 20 or All             values only from vehicles of the defined vehicle class number (according to the attribute "No" of Vehicle Classes)
            //                                 Note: You can only access the results of specific vehicle classes if you set it in Evaluation > Configuration > Result Attributes
            //
            // The value of on time interval is the arithmetic mean of all single travel times of the vehicles.

            // Example #1:
            // Average of all simulations (1. input = Avg)
            // 	 of the average of all time intervals  (2. input = Avg)
            //   of all vehicle classes (3. input = All)
            TT = (double)Veh_TT_measurement.get_AttValue("TravTm(Avg,Avg,All)");
            No_Veh = (int)Veh_TT_measurement.get_AttValue("Vehs  (Avg,Avg,All)");
            Console.WriteLine("Average travel time all time intervals of all simulation of all vehicle classes: {0:F} (number of vehicles: {1})", TT, No_Veh);

            // Example #2:
            // Value of the Current simulation (1. input = Current)
            // 	 of the maximum of all time intervals (2. input = Max)
            //   of vehicle class HGV (3. input = 20)
            TT = (double)Veh_TT_measurement.get_AttValue("TravTm(Current,Max,20)");
            No_Veh = (int)Veh_TT_measurement.get_AttValue("Vehs  (Current,Max,20)");
            Console.WriteLine("Maximum travel time of all time intervals of the current simulation of vehicle class HGV: {0:F} (number of vehicles: {1})", TT, No_Veh);

            // Example #3: Note: A Travel times from 2nd simulation run must be available
            // Value of the 2nd simulation (1. input = 2)
            // 	 of the 1st time interval (2. input = 1)
            //   of all vehicle classes (3. input = All)
            // TT = (double)Veh_TT_measurement.get_AttValue("TravTm(2,1,All)");
            // No_Veh = (int)Veh_TT_measurement.get_AttValue("Vehs  (2,1,All)");
            // Console.WriteLine("Travel time of the 1st time interval of the 2nd simulation of all vehicle classes: {0:F} (number of vehicles: {1})", TT, No_Veh);


            // Data Collection
            uint DC_measurement_number = 1;
            IDataCollectionMeasurement DC_measurement = Vissim.Net.DataCollectionMeasurements.get_ItemByKey(DC_measurement_number);
            // Syntax to get the data:
            //   DC_measurement.get_AttValue("Vehs(sub_attribut_1, sub_attribut_2, sub_attribut_3)");
            //
            // sub_attribut_1: SimulationRun (same as described at Vehicle Travel Time Measurements)
            // sub_attribut_2: TimeInterval  (same as described at Vehicle Travel Time Measurements)
            // sub_attribut_3: VehicleClass  (same as described at Vehicle Travel Time Measurements)
            //
            // The value of on time interval is the arithmetic mean of all single values of the vehicles.

            // Example #1:
            // Average value of all simulations (1. input = Avg)
            //   of the 1st time interval (2. input = 1)
            //   of all vehicle classes (3. input = All)
            No_Veh = (int)DC_measurement.get_AttValue("Vehs        (Avg,1,All)"); // number of vehicles
            double Speed = (double)DC_measurement.get_AttValue("Speed       (Avg,1,All)"); // Speed of vehicles
            double Acceleration = (double)DC_measurement.get_AttValue("Acceleration(Avg,1,All)"); // Acceleration of vehicles
            double Length = (double)DC_measurement.get_AttValue("Length      (Avg,1,All)"); // Length of vehicles
            Console.WriteLine("Data Collection #" + DC_measurement_number + ": Average values of all Simulations runs of 1st time interval of all vehicle classes:");
            Console.WriteLine("#vehicles: {0}; Speed: {1:F}; Acceleration: {2:F}; Length: {3:F}", No_Veh, Speed, Acceleration, Length);


            // Queue length
            // Syntax to get the data:
            //   get(QueueCounter.get_AttValue("QLen(sub_attribut_1, sub_attribut_2)");
            //
            // sub_attribut_1: SimulationRun (same as described at Vehicle Travel Time Measurements)
            // sub_attribut_2: TimeInterval  (same as described at Vehicle Travel Time Measurements)
            //

            // Example #1:
            // Average value of all simulations (1. input = Avg)
            // 	of the average of all time intervals (2. input = Avg)
            uint QC_number = 1;
            double maxQ = (double)Vissim.Net.QueueCounters.get_ItemByKey(QC_number).get_AttValue("QLenMax(Avg, Avg)");
            Console.WriteLine("Average maximum Queue length of all simulations and time intervals of Queue Counter #{0}: {1:F}", QC_number, maxQ);

            //// ========================================================================
            // Saving
            //==========================================================================
            Filename = Path_of_COM_Basic_Commands_network + "COM Basic Commands saved.inpx";
            Vissim.SaveNetAs(Filename);
            Filename = Path_of_COM_Basic_Commands_network + "COM Basic Commands saved.layx";
            Vissim.SaveLayout(Filename);


            //// ========================================================================
            // End Vissim
            //==========================================================================
            // Vissim.Exit();


        }
    }
}