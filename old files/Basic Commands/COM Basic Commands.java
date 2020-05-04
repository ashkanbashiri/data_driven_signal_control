
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.EnumVariant;
import com.jacob.com.SafeArray;
import com.jacob.com.Variant;

// ==========================================================================
// Java-Script for Vissim 6+
// Copyright (C) PTV AG. All rights reserved.
// Roland Osterrieter 2015
// - - - - - - - - - - - - - - - - -
// Example of basic syntax - DO NOT MODIFY
// ==========================================================================

// This script demonstrates how to use the COM interface in Java.
// Basic commands for loading a network and layout, reading and setting
// attributes of network objects, running a simulation and retrieving
// evaluations are shown. This example is also available for the programming
// languages VBA, VBS Matlab, C#, C++ and Java.
//
// If you start using COM, please see also our COM introduction document: 
// C:\Program Files\PTV Vision\PTV Vissim 10\Doc\Eng\Vissim 10 - COM Intro.pdf
//
// For information about the attributes and methods of PTV Vissim objects, see
// the COM Help, which is located in the PTV Vissim menu: Help > COM Help. 

public class COM_Basic_Commands
{



    public static void main(String[] args) throws Exception
    {
        // Connecting the COM Server => Open a new Vissim Window:
        ActiveXComponent vissim = new ActiveXComponent("VISSIM.Vissim");
        // Load example net
        String path_of_COM_Basic_Commands_network = "C:\\Users\\Public\\Documents\\PTV Vision\\PTV Vissim 10\\Examples Training\\COM\\Basic Commands\\";
        vissim.invoke("LoadNet", path_of_COM_Basic_Commands_network + "COM Basic Commands.inpx");
        // Load layout
        vissim.invoke("LoadLayout", path_of_COM_Basic_Commands_network + "COM Basic Commands.layx");

        //// ========================================================================
        // Read and Set attributes (vehicles)
        //==========================================================================
        // Note: All of the following commands can also be executed during a
        // simulation.

        // Read Link Name:
        int linkNumber = 1;
        ActiveXComponent net = vissim.invokeGetComponent("Net");
        ActiveXComponent linkContainer = net.invokeGetComponent("Links");
        ActiveXComponent link = linkContainer.invokeGetComponent("ItemByKey", new Variant(linkNumber));
        String linkName = link.invoke("AttValue", "Name").getString();
        System.out.println("Name of Link(" + linkNumber + "): " + linkName);

        // Set Link Name:
        String newLinkName = "New Link Name";
        Dispatch.invoke(link, "AttValue", Dispatch.Put, new Object[]{"Name", newLinkName}, new int[1]);

        // Set a signal controller program:
        int scNumber = 1; // SC = SignalController
        ActiveXComponent signalController = net.invokeGetComponent("SignalControllers").invokeGetComponent("ItemByKey",
                                                                                                           new Variant(scNumber));
        int newSignalProgrammNumber = 2;
        Dispatch.invoke(signalController, "AttValue", Dispatch.Put, new Object[]{"ProgNo", newSignalProgrammNumber}, new int[1]);

        // Set relative flow of a static vehicle route of a static vehicle routing decision:
        int svrdNumber = 1; // SVRD = Static Vehicle Routing Decision
        int svrNumber = 1; // SVR = Static Vehicle Route (of a specific Static Vehicle Routing Decision)
        double newRelativFlow = 0.6;
        ActiveXComponent svrd = net.invokeGetComponent("VehicleRoutingDecisionsStatic")
                .invokeGetComponent("ItemByKey", new Variant(svrdNumber));
        ActiveXComponent svr = svrd.invokeGetComponent("VehRoutSta").invokeGetComponent("ItemByKey", new Variant(svrNumber));
        Dispatch.invoke(svr, "AttValue", Dispatch.Put, new Object[]{"RelFlow(1)", newRelativFlow}, new int[1]);
        // "RelFlow(1)" means the first defined time interval; to access the third defined time interval: "RelFlow(3)"

        // Set vehicle input:
        int viNumber = 1; // VI = Vehicle Input
        double newVolume = 600; // vehicles per hour
        ActiveXComponent vehicleInputs = net.invokeGetComponent("VehicleInputs");
        ActiveXComponent vehicleInput = vehicleInputs.invokeGetComponent("ItemByKey", new Variant(viNumber));
        Dispatch.invoke(vehicleInput, "AttValue", Dispatch.Put, new Object[]{"Volume(1)", newVolume}, new int[1]);
        // "Volume(1)" means the first defined time interval
        // Hint: The Volumes of following intervals Volume(i) i = 2...n can only be
        // edited, if continuous is deactivated: (otherwise error: "AttValue failed: Object 2: Attribute Volume (300) is no subject to changes.")
        Dispatch.invoke(vehicleInput, "AttValue", Dispatch.Put, new Object[]{"Cont(2)", false}, new int[1]);
        Dispatch.invoke(vehicleInput, "AttValue", Dispatch.Put, new Object[]{"Volume(2)", 400}, new int[1]);

        // Set vehicle composition:
        int vehCompositionNumber = 1;
        ActiveXComponent vehicleCompositions = net.invokeGetComponent("VehicleCompositions");
        ActiveXComponent vehicleComposition = vehicleCompositions.invokeGetComponent("ItemByKey",
                                                                                     new Variant(vehCompositionNumber));
        ActiveXComponent vehCompRelFlows = vehicleComposition.invokeGetComponent("VehCompRelFlows");
        Variant[] vehCompRelFlowsArray = Dispatch.invoke(vehCompRelFlows, "GetAll", Dispatch.Method, new Object[0], new int[1])
                .toSafeArray().toVariantArray();
        Dispatch relFlow0 = vehCompRelFlowsArray[0].getDispatch();
        Dispatch relFlow1 = vehCompRelFlowsArray[1].getDispatch();
        // Changing the vehicle type
        Dispatch.invoke(relFlow0, "AttValue", Dispatch.Put, new Object[]{"VehType", 100}, new int[1]);
        // Changing the desired speed distribution
        Dispatch.invoke(relFlow0, "AttValue", Dispatch.Put, new Object[]{"DesSpeedDistr", 50}, new int[1]);
        // Changing the relative flow
        Dispatch.invoke(relFlow0, "AttValue", Dispatch.Put, new Object[]{"RelFlow", 0.9}, new int[1]);
        // Changing the relative flow of the 2nd Relative Flow.
        Dispatch.invoke(relFlow1, "AttValue", Dispatch.Put, new Object[]{"RelFlow", 0.1}, new int[1]);

        //// ========================================================================
        // Simulation
        //==========================================================================

        // Chose Random Seed
        int randomSeed = 42;
        ActiveXComponent simulation = vissim.invokeGetComponent("Simulation");
        Dispatch.invoke(simulation, "AttValue", Dispatch.Put, new Object[]{"RandSeed", randomSeed}, new int[1]);

        // To start a simulation you can run a single step:
        simulation.invoke("RunSingleStep");
        // Or run the simulation continuous (it stops at breakpoint or end of simulation)
        double endOfSimulation = 600; // simulation second [s]
        Dispatch.invoke(simulation, "AttValue", Dispatch.Put, new Object[]{"SimPeriod", endOfSimulation}, new int[1]);
        double simBreakAt = 200; // simulation second [s]
        Dispatch.invoke(simulation, "AttValue", Dispatch.Put, new Object[]{"SimBreakAt", simBreakAt}, new int[1]);
        // Set maximum speed:
        Dispatch.invoke(simulation, "AttValue", Dispatch.Put, new Object[]{"UseMaxSimSpeed", true}, new int[1]);
        // Hint: to change the speed use:  Dispatch.invoke(simulation, "AttValue", Dispatch.Put, new Object[]{"SimSpeed", 10}, new int[1]); // 10 => 10 Sim. sec. / s
        simulation.invoke("RunContinuous");
        // To stop the simulation:
        simulation.invoke("Stop");

        //// ========================================================================
        // Access during simulation
        //==========================================================================
        // Note: All of commands of "Read and Set attributes (vehicles)" can also be executed during a
        // simulation (e.g. changing signal controller program, setting relative flow of a static vehicle route,
        // changing the vehicle input, changing the vehicle composition).


        simBreakAt = 198; // simulation second [s]
        Dispatch.invoke(simulation, "AttValue", Dispatch.Put, new Object[]{"SimBreakAt", simBreakAt}, new int[1]);
        simulation.invoke("RunContinuous");// Start the simulation until SimBreakAt (198s)


        // Get the state of a signal head:
        int shNumber = 1; // SH = SignalHead
        ActiveXComponent signalHeads = net.invokeGetComponent("SignalHeads");
        ActiveXComponent signalHead = signalHeads.invokeGetComponent("ItemByKey", new Variant(shNumber));
        String stateOfSh = signalHead.invoke("AttValue", "SigState").getString();

        System.out.println("Actual state of SignalHead(" + shNumber + ") is: " + stateOfSh);// possible output e.g. 'GREEN', 'RED', 'AMBER', 'REDAMBER'

        // Set the state of a signal controller:
        // Note: Once a state of a signal group is set, the attribute "ContrByCOM" is automatically set to True. Meaning the signal group will keep this state until another state is set by COM or the end of the simulation
        // To switch back to the defined signal controller, set the attribute signal "ContrByCOM" to False (example see below).
        scNumber = 1; // SC = SignalController
        int sgNumber = 1; // SG = SignalGroup
        signalController = net.invokeGetComponent("SignalControllers").invokeGetComponent("ItemByKey", new Variant(scNumber));
        ActiveXComponent signalGroup = signalController.invokeGetComponent("SGs").invokeGetComponent("ItemByKey",
                                                                                                     new Variant(sgNumber));
        String newState = "GREEN"; //possible values e.g. "GREEN", "RED", "AMBER", "REDAMBER"
        Dispatch.invoke(signalGroup, "AttValue", Dispatch.Put, new Object[]{"SigState", newState}, new int[1]);
        // Note: The signal controller can only be called at whole simulation seconds, so the state will be set in Vissim at the next whole simulation second, here 199s
        // Simulate so that the new state is active in the Vissim simulation:
        simBreakAt = 200; // simulation second [s]
        Dispatch.invoke(simulation, "AttValue", Dispatch.Put, new Object[]{"SimBreakAt", simBreakAt}, new int[1]);
        simulation.invoke("RunContinuous"); // Start the simulation until SimBreakAt (200s)
        // Give the control back:
        Dispatch.invoke(signalGroup, "AttValue", Dispatch.Put, new Object[]{"ContrByCOM", false}, new int[1]);


        int vehNumber;
        String vehType;
        double vehSpeed;
        double vehPosition;
        String vehLinklane;

        // Information about all vehicles in the network (in the current simulation second):
        // In the following, 4 different methods to access attributes are shown:
        // Method #1: Loop over all Vehicles using "GetAll"
        // Method #2: Loop over all Vehicles using Object Enumeration
        // Method #3: Using the Iterator
        // Method #4: Accessing all attributes directly using "GetMultiAttValues" (fast way if you want the attributes of all vehicles)
        // The result of the four methods is the same (except the format).

        // Method #1: Loop over all Vehicles:
        ActiveXComponent vehicles = net.invokeGetComponent("Vehicles");
        // get all vehicles in the network at the actual simulation second
        Variant[] vehicleArray = Dispatch.invoke(vehicles, "GetAll", Dispatch.Method, new Object[0], new int[1]).toSafeArray()
                .toVariantArray();
        System.out.println("Loop over all Vehicles:");
        for (int i = 0; i < vehicleArray.length; i++)
        {
            Dispatch vehicle = vehicleArray[i].getDispatch();
            vehNumber = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"No"}, new int[1]).getInt();
            vehType = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"VehType"}, new int[1]).getString();
            vehSpeed = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"Speed"}, new int[1]).getDouble();
            vehPosition = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"Pos"}, new int[1]).getDouble();
            vehLinklane = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"Lane"}, new int[1]).getString();
            System.out.println(" --" + vehNumber + " " + vehType + " " + vehSpeed + " " + vehPosition + " " + vehLinklane);
        }

        // Method #2: Loop over all Vehicles using Object Enumeration
        System.out.println("Loop over all Vehicles (Enumeration):");
        for (Variant vehicleVariant : vehicleArray)
        {
            Dispatch vehicle = vehicleVariant.getDispatch();
            vehNumber = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"No"}, new int[1]).getInt();
            vehType = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"VehType"}, new int[1]).getString();
            vehSpeed = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"Speed"}, new int[1]).getDouble();
            vehPosition = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"Pos"}, new int[1]).getDouble();
            vehLinklane = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"Lane"}, new int[1]).getString();
            System.out.println(" --" + vehNumber + " " + vehType + " " + vehSpeed + " " + vehPosition + " " + vehLinklane);
        }

        // Method #3: Using the Iterator (this method is a little bit slower)
        System.out.println("Loop over all Vehicles (Iterator):");
        ActiveXComponent iterator = vehicles.invokeGetComponent("Iterator");
        while (iterator.invoke("Valid").getBoolean())
        {
            Dispatch vehicle = iterator.invoke("Item").getDispatch();
            vehNumber = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"No"}, new int[1]).getInt();
            vehType = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"VehType"}, new int[1]).getString();
            vehSpeed = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"Speed"}, new int[1]).getDouble();
            vehPosition = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"Pos"}, new int[1]).getDouble();
            vehLinklane = Dispatch.invoke(vehicle, "AttValue", Dispatch.Get, new Object[]{"Lane"}, new int[1]).getString();
            System.out.println(" --" + vehNumber + " " + vehType + " " + vehSpeed + " " + vehPosition + " " + vehLinklane);
            iterator.invoke("Next");
        }

        System.out.println("Get Vehicle-Attributes (MultiAttValues):");
        // Method #4: Accessing all Attributes directly using "GetMultiAttValues" (fastest way)
        /*
         * The multiple values(MultiAttValues) are committed as a 2d array.
         * JACOBs SafeArray format allows to handle 2d or n-dimensional arrays,
         * but there is no straight template or documentation. The
         * COM_Basic_Commands_utils class offers a simple approach to read complete
         * 2D-Arrays from SafeArrays.
         */
        String mthd = "GetMultiAttValues";
        //Get all Attribute SafeArrays
        SafeArray vnsarray = vehicles.invoke(mthd, "No").toSafeArray();
        SafeArray vtsarray = vehicles.invoke(mthd, "VehType").toSafeArray();
        SafeArray vssarray = vehicles.invoke(mthd, "Speed").toSafeArray();
        SafeArray vpsarray = vehicles.invoke(mthd, "Pos").toSafeArray();
        SafeArray vlsarray = vehicles.invoke(mthd, "Lane").toSafeArray();
        //Read data from SafeArray into simple ArrayFormat
        int[][] vehicleNumbers = COM_Basic_Commands_utils.get2DIntArrayFromSafeArray(vnsarray);
        String[][] vehicleTypes = COM_Basic_Commands_utils.get2DStringArrayFromSafeArray(vtsarray);
        double[][] vehicleSpeeds = COM_Basic_Commands_utils.get2DDoubleArrayFromSafeArray(vssarray);
        double[][] vehiclePositions = COM_Basic_Commands_utils.get2DDoubleArrayFromSafeArray(vpsarray);
        String[][] vehicleLinklanes = COM_Basic_Commands_utils.get2DStringArrayFromSafeArray(vlsarray);
        for (int i = 0; i < vehicleNumbers.length; i++)
        {
            vehNumber = vehicleNumbers[i][1];
            vehType = vehicleTypes[i][1];
            vehSpeed = vehicleSpeeds[i][1];
            vehPosition = vehiclePositions[i][1];
            vehLinklane = vehicleLinklanes[i][1];
            System.out.println(" --" + vehNumber + " " + vehType + " " + vehSpeed + " " + vehPosition + " " + vehLinklane);
        }

        // Operations at one specific vehicle:
        Dispatch vehicle1 = vehicleArray[1].getDispatch();
        // alternatively with ItemByKey

        // Set desired speed to a vehicle:
        double newDesspeed = 30;
        Dispatch.invoke(vehicle1, "AttValue", Dispatch.Put, new Object[]{"DesSpeed", newDesspeed}, new int[1]);

        // Move a vehicle:
        linkNumber = 1;
        int laneNumber = 1;
        double linkCoordinate = 70;
        // This function will operate during the next simulation step
        // Hint: In earlier Vissim releases, the name of the function was: MoveToLinkCoordinate
        Dispatch.invoke(vehicle1, "MoveToLinkPosition", Dispatch.Method, new Object[]{linkNumber, laneNumber, linkCoordinate},
                        new int[1]);

        simulation.invoke("RunSingleStep"); // Next Step, so that the vehicle gets moved.


        // Remove a vehicle:
        vehNumber = Dispatch.invoke(vehicle1, "AttValue", Dispatch.Get, new Object[]{"No"}, new int[1]).getInt();
        vehicles.invoke("RemoveVehicle", vehNumber);

        // Putting a new vehicle to the network:
        int vehicleType = 100;
        double desiredSpeed = 53; // unit according to the user setting in Vissim [km/h or mph]
        int atLink = 1;
        int atLane = 1;
        double xcoordinate = 15; // unit according to the user setting in Vissim [m or ft]
        boolean interaction = true; // optional boolean
        Object[] oArg = {vehicleType, atLink, atLane, xcoordinate, desiredSpeed, interaction};
        Dispatch.invoke(vehicles, "AddVehicleAtLinkPosition", Dispatch.Method, oArg, new int[1]);
        // Note: In earlier Vissim releases, the name of the function was: AddVehicleAtLinkCoordinate

        // Make Screenshots of the intersection 2D and 3D:
        // ZoomTo:
        // Zooms the view to the rectangle defined by the two points  (x1, y1) and (x2,y2),  which  are  given  in  world coordinates.  If  the  rectangle  proportions  differ
        // from  the  proportions  of  the  network  window,  the  specified  rectangle  will  be centred in the network editor window.
        int x1 = 250;
        int y1 = 30;
        int x2 = 350;
        int y2 = 135;
        ActiveXComponent graphics = vissim.invokeGetComponent("Graphics");
        ActiveXComponent window = graphics.invokeGetComponent("CurrentNetworkWindow");

        Object[] oArg2 = {x1, y1, x2, y2};
        Dispatch.invoke(window, "ZoomTo", Dispatch.Method, oArg2, new int[1]);

        // Make a Screenshot in 2D:
        // It  creates  a  graphic  file  of  the  VISSIM  main  window  formatted  according  to its extension: PNG, TIFF, GIF, JPG, JPEG or BMP. A BMP file will be written if the extension can not be recognized.
        String filenameScreenshot = "screenshot2D.jpg"; // to set to a specific path: "C:\\Screenshots\\screenshot2D.jpg"
        int sizeFactor = 1; // 1: original size, 2: doubles size

        Object[] oArg3 = {filenameScreenshot, sizeFactor};
        Dispatch.invoke(window, "Screenshot", Dispatch.Method, oArg3, new int[1]);

        // Make a Screenshot in 3D:
        // Set 3D Mode:
        Object[] oArg4 = {"3D", 1};
        Dispatch.invoke(window, "AttValue", Dispatch.Put, oArg4, new int[1]);

        // Set the camera position (viewing angle):
        int xPos = 270;
        int yPos = 30;
        int zPos = 15;
        int yawAngle = 45;
        int pitchAngle = 10;
        Object[] oArg5 = {xPos, yPos, zPos, yawAngle, pitchAngle};
        Dispatch.invoke(window, "SetCameraPositionAndAngle", Dispatch.Method, oArg5, new int[1]);

        filenameScreenshot = "screenshot3D.jpg"; // to set to a specific path: "C:\\Screenshots\\screenshot3D.jpg"
        Object[] oArg6 = {filenameScreenshot, sizeFactor};
        Dispatch.invoke(window, "Screenshot", Dispatch.Method, oArg6, new int[1]);

        // Set 2D Mode and old Network position:
        Dispatch.invoke(window, "AttValue", Dispatch.Put, new Object[]{"3D", 0}, new int[1]);
        x1 = -10;
        y1 = -10;
        x2 = 600;
        y2 = 300;
        Object[] oArg7 = {x1, y1, x2, y2};
        Dispatch.invoke(window, "ZoomTo", Dispatch.Method, oArg7, new int[1]);


        // Continue the simulation until end of simulation (get(Vissim.Simulation, 'AttValue', 'SimPeriod'))
        simulation.invoke("RunContinuous");

        //// ========================================================================
        // Results of Simulations:
        //==========================================================================

        // Run 3 Simulations at maximum speed:
        // Activate QuickMode:
        Dispatch.invoke(window, "AttValue", Dispatch.Put, new Object[]{"QuickMode", 1}, new int[1]);
		vissim.invoke("SuspendUpdateGUI"); // stop updating of the complete Vissim workspace (network editor, list, chart and signal time table windows)

        // Alternatively, load a layout (*.layx) where dynamic elements (vehicles and pedestrians) are not visible:
        endOfSimulation = 600;
        Dispatch.invoke(simulation, "AttValue", Dispatch.Put, new Object[]{"SimPeriod", endOfSimulation}, new int[1]);
        simBreakAt = 0; // simulation second [s] => 0 means no break!
        Dispatch.invoke(simulation, "AttValue", Dispatch.Put, new Object[]{"SimBreakAt", simBreakAt}, new int[1]);


        // Set maximum speed:
        Dispatch.invoke(simulation, "AttValue", Dispatch.Put, new Object[]{"UseMaxSimSpeed", true}, new int[1]);

        for (int count = 1; count <= 3; count++)
        {
            Dispatch.invoke(simulation, "AttValue", Dispatch.Put, new Object[]{"RandSeed", count}, new int[1]);
            simulation.invoke("RunContinuous");

        }
		vissim.invoke("ResumeUpdateGUI"); // allow updating of the complete Vissim workspace (network editor, list, chart and signal time table windows)
        // deactivate QuickMode
        Dispatch.invoke(window, "AttValue", Dispatch.Put, new Object[]{"QuickMode", 0}, new int[1]);

        // List of all Simulation runs:
        String[] attributes = {"Timestamp", "RandSeed", "SimEnd"};
        ActiveXComponent simulationRuns = net.invokeGetComponent("SimulationRuns");
        iterator = simulationRuns.invokeGetComponent("Iterator");
        while (iterator.invoke("Valid").getBoolean())
        {
            Dispatch simulationRun = iterator.invoke("Item").getDispatch();
            String output = "";
            for (String attr : attributes)
            {
                output += " " + Dispatch.invoke(simulationRun, "AttValue", Dispatch.Get, new Object[]{attr}, new int[1]);
            }
            System.out.println(output);
            iterator.invoke("Next");
        }


        // Get the results of Vehicle Travel Time Measurements:
        double traveltime;
        int vehicleNo;
        int vehTtMeasurementNumber = 2;
        ActiveXComponent vehTtMeasurements = net.invokeGetComponent("VehicleTravelTimeMeasurements");
        ActiveXComponent vehTtMeasurement = vehTtMeasurements
                .invokeGetComponent("ItemByKey", new Variant(vehTtMeasurementNumber));
        // Syntax to get the travel times:
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
        //   of the average of all time intervals  (2. input = Avg)
        //   of all vehicle classes (3. input = All)
        Object[] arg = {"TravTm(Avg,Avg,All)"};
        traveltime = Dispatch.invoke(vehTtMeasurement, "AttValue", Dispatch.Get, arg, new int[1]).getDouble();
        Object[] arg2 = {"Vehs  (Avg,Avg,All)"};
        vehicleNo = Dispatch.invoke(vehTtMeasurement, "AttValue", Dispatch.Get, arg2, new int[1]).getInt();
        System.out.println("Average travel time all time intervals of all simulation of all vehicle classes: " + traveltime
                + " (number of vehicles: " + vehicleNo + ")");

        // Example #2:
        // Value of the Current simulation (1. input = Current)
        //   of the maximum of all time intervals (2. input = Max)
        //   of vehicle class HGV (3. input = 20)
        Object[] arg3 = {"TravTm(Current,Max,20)"};
        traveltime = Dispatch.invoke(vehTtMeasurement, "AttValue", Dispatch.Get, arg3, new int[1]).getDouble();
        Object[] arg4 = {"Vehs  (Current,Max,20)"};
        vehicleNo = Dispatch.invoke(vehTtMeasurement, "AttValue", Dispatch.Get, arg4, new int[1]).getInt();
        System.out.println("Maximum travel time of all time intervals of the current simulation of vehicle class HGV: "
                + traveltime + " (number of vehicles: " + vehicleNo + ")");

        // Example #3: Note: A Travel times from 2nd simulation run must be available
        // Value of the 2nd simulation (1. input = 2)
        //   of the 1st time interval (2. input = 1)
        //   of all vehicle classes (3. input = All)
        // Object[] arg5 = {"TravTm(2,1,All)"};
        // traveltime = Dispatch.invoke(vehTtMeasurement, "AttValue", Dispatch.Get, arg5, new int[1]).getDouble();
        // Object[] arg6 = {"Vehs  (2,1,All)"};
        //vehicleNo = Dispatch.invoke(vehTtMeasurement, "AttValue", Dispatch.Get, arg6, new int[1]).getInt();
        // System.out.println("Travel time of the 1st time interval of the 2nd simulation of all vehicle classes:" + traveltime
        //    + " (number of vehicles: " + vehicleNo + ")");

        // Data Collection
        int dcMeasurementNr = 1;
        ActiveXComponent dcMeasurements = net.invokeGetComponent("DataCollectionMeasurements");
        ActiveXComponent dcMeasurement = dcMeasurements.invokeGetComponent("ItemByKey", new Variant(dcMeasurementNr));
        // The value of on time interval is the arithmetic mean of all single values of the vehicles.

        // Example #1:
        // Average value of all simulations (1. input = Avg)
        //   of the 1st time interval (2. input = 1)
        //   of all vehicle classes (3. input = All)
        Object[] arg7 = {"Vehs        (Avg,1,All)"};
        // number of vehicles
        int noVeh = Dispatch.invoke(dcMeasurement, "AttValue", Dispatch.Get, arg7, new int[1]).getInt();
        Object[] arg8 = {"Speed        (Avg,1,All)"};
        // Speed of vehicles
        double speed = Dispatch.invoke(dcMeasurement, "AttValue", Dispatch.Get, arg8, new int[1]).getDouble();
        Object[] arg9 = {"Acceleration        (Avg,1,All)"};
        // Acceleration of vehicles
        double acceleration = Dispatch.invoke(dcMeasurement, "AttValue", Dispatch.Get, arg9, new int[1]).getDouble();
        Object[] arg10 = {"Length        (Avg,1,All)"};
        // Length of vehicles
        double length = Dispatch.invoke(dcMeasurement, "AttValue", Dispatch.Get, arg10, new int[1]).getDouble();
        System.out.println("Data Collection #" + dcMeasurementNr
                + ": Average values of all Simulations runs of 1st time interval of all vehicle classes:");
        System.out.println("#vehicles: "+noVeh+" Speed: "+speed+" Acceleration: "+acceleration+" Length: "+length);

        // Queue length
        //
        // sub_attribut_1: SimulationRun (same as described at Vehicle Travel Time Measurements)
        // sub_attribut_2: TimeInterval  (same as described at Vehicle Travel Time Measurements)
        //

        // Example #1:
        //  Average value of all simulations (1. input = Avg)
        //  of the average of all time intervals (2. input = Avg)
        int qcNumber = 1;
        ActiveXComponent queueCounters = net.invokeGetComponent("QueueCounters");
        ActiveXComponent queueCounter = queueCounters.invokeGetComponent("ItemByKey", new Variant(qcNumber));
        Object[] arg11 = {"QLenMax(Avg, Avg)"};
        double maxQ =  Dispatch.invoke(queueCounter, "AttValue", Dispatch.Get, arg11, new int[1]).getDouble();
        System.out.println("Average maximum Queue length of all simulations and time intervals of Queue Counter #"+qcNumber+" "+maxQ);


        //// ========================================================================
        // Saving
        //==========================================================================
        String filename = path_of_COM_Basic_Commands_network + "COM Basic Commands saved.inpx";
        vissim.invoke("SaveNetAs", filename);

        filename = path_of_COM_Basic_Commands_network + "COM Basic Commands saved.layx";
        vissim.invoke("SaveLayout", filename);


        //// ========================================================================
        // End Vissim
        //==========================================================================
        // vissim.invoke("Exit");

    }


}
