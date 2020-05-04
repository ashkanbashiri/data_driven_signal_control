%==========================================================================
% MATLAB-Script for Vissim 6+
% Copyright (C) PTV AG, Jochen Lohmiller
% All rights reserved.
% -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -
% Example of basic syntax - DO NOT MODIFY
%==========================================================================

% This script demonstrates how to use the COM interface in MATLAB.
% Basic commands for loading a network and layout, reading and setting
% attributes of network objects, running a simulation and retrieving
% evaluations are shown. This example is also available for the programming
% languages VBA, VBS, Python, C#, C++ and Java.
%
% If you start using COM, please see also our COM introduction document: 
% C:\Program Files\PTV Vision\PTV Vissim 10\Doc\Eng\Vissim 10 - COM Intro.pdf
%
% For information about the attributes and methods of PTV Vissim objects, see
% the COM Help, which is located in the PTV Vissim menu: Help > COM Help. 
%
% Hint: You can easily see all attributes of the PTV Vissim objects in the lists in
% Vissim. In case your PTV Vissim language is set to English, the name in the
% headline of the list correspond to the command to access via COM.
% Example: If you want to access the number of lanes of a link, go to PTV Vissim
% and see the headline of "Number of lanes" in the list, which is
% "NumLanes". To access this attribute via COM use:
% get(Vissim.Net.Links.ItemByKey(1), 'AttValue', 'NumLanes')


feature('COM_SafeArraySingleDim', 1); % Matlab should only pass one-dimensional array to COM

%% Connecting the COM Server => Open a new Vissim Window:
Vissim = actxserver('Vissim.Vissim');
% If you have installed multiple Vissim Versions, you can open a specific Vissim version adding the bit Version (32 or 64bit) and Version number
% Vissim = actxserver('Vissim.Vissim-32.9'); % Vissim 9 - 32 bit
% Vissim = actxserver('Vissim.Vissim-64.9'); % Vissim 9 - 64 bit
% Vissim = actxserver('Vissim.Vissim-32.10'); % Vissim 10 - 32 bit
% Vissim = actxserver('Vissim.Vissim-64.10'); % Vissim 10 - 64 bit
Path_of_COM_Basic_Commands_network = cd; %'C:\Users\Public\Documents\PTV Vision\PTV Vissim 10\Examples Training\COM\Basic Commands';

%% Load a Vissim Network:
Filename                = fullfile(Path_of_COM_Basic_Commands_network, 'COM Basic Commands.inpx');
flag_read_additionally  = false; % you can read network(elements) additionally, in this case set "flag_read_additionally" to true
Vissim.LoadNet(Filename, flag_read_additionally);

%% Load a Layout:
Filename = fullfile(Path_of_COM_Basic_Commands_network, 'COM Basic Commands.layx');
Vissim.LoadLayout(Filename);


%% ========================================================================
% Read and Set attributes
%==========================================================================
% Note: All of the following commands can also be executed during a
% simulation.

% Read Link Name:
Link_number = 1;
Name_of_Link = get(Vissim.Net.Links.ItemByKey(Link_number), 'AttValue', 'Name');
% Name_of_Link = Vissim.Net.Links.ItemByKey(Link_number).AttValue('Name'); % alternative (using the method "AttValue")
disp(['Name of Link(',num2str(Link_number),'):',32,Name_of_Link]) % char(32) is whitespace

% Set Link Name:
new_Name_of_Link = 'New Link Name';
set(Vissim.Net.Links.ItemByKey(Link_number), 'AttValue', 'Name', new_Name_of_Link);

% Set a signal controller program:
SC_number = 1; % SC = SignalController
SignalController = Vissim.Net.SignalControllers.ItemByKey(SC_number);
new_signal_programm_number = 2;
set(SignalController, 'AttValue', 'ProgNo', new_signal_programm_number);

% Set relative flow of a static vehicle route of a static vehicle routing decision:
SVRD_number         = 1; % SVRD = Static Vehicle Routing Decision
SVR_number          = 1; % SVR = Static Vehicle Route (of a specific Static Vehicle Routing Decision)
new_relativ_flow    = 0.6;
set(Vissim.Net.VehicleRoutingDecisionsStatic.ItemByKey(SVRD_number).VehRoutSta.ItemByKey(SVR_number), 'AttValue', 'RelFlow(1)', new_relativ_flow);
% 'RelFlow(1)' means the first defined time interval; to access the third defined time interval: 'RelFlow(3)'

% Set vehicle input:
VI_number   = 1; % VI = Vehicle Input
new_volume  = 600; % vehicles per hour
set(Vissim.Net.VehicleInputs.ItemByKey(VI_number), 'AttValue', 'Volume(1)', new_volume);
% 'Volume(1)' means the first defined time interval
% Hint: The Volumes of following intervals Volume(i) i = 2...n can only be
% edited, if continuous is deactivated: (otherwise error: "AttValue failed: Object 2: Attribute Volume (300) is no subject to changes.")
set(Vissim.Net.VehicleInputs.ItemByKey(VI_number), 'AttValue', 'Cont(2)',     false);
set(Vissim.Net.VehicleInputs.ItemByKey(VI_number), 'AttValue', 'Volume(2)',   400  );

% Set vehicle composition:
Veh_composition_number = 1;
Rel_Flows = Vissim.Net.VehicleCompositions.ItemByKey(Veh_composition_number).VehCompRelFlows.GetAll;
set(Rel_Flows{1}, 'AttValue', 'VehType',        100); % Changing the vehicle type
set(Rel_Flows{1}, 'AttValue', 'DesSpeedDistr',   50); % Changing the desired speed distribution
set(Rel_Flows{1}, 'AttValue', 'RelFlow',        0.9); % Changing the relative flow
set(Rel_Flows{2}, 'AttValue', 'RelFlow',        0.1); % Changing the relative flow of the 2nd Relative Flow.


%% ========================================================================
% Accessing Multiple Attributes:
%========================================================================

% GetMultiAttValues         Read one attribute of all objects:
Attribute = 'Name';
NameOfLinks = Vissim.Net.Links.GetMultiAttValues(Attribute);

% SetMultiAttValues         Set one attribute of multiple (not necessarily all) objects
NameOfLinks{1, 2} = 'New Link Name of Link #1';
NameOfLinks{2, 2} = 'New Link Name of Link #2'; 
NameOfLinks{4, 2} = 'New Link Name of Link #4';
Vissim.Net.Links.SetMultiAttValues(Attribute, NameOfLinks); % 1st input is the consecutively number of links (not the ID), 2nd the link name
% Please note: The first column of GetMultiAttValues or SetMultiAttValues is not the ID of an object, it is consecutively numbered all available objects.

% GetMultipleAttributes     Read multiple attributes of all objects:
Attributes1 = {'Name'; 'Length2D'};
Name_Length_of_Links = Vissim.Net.Links.GetMultipleAttributes(Attributes1);

% SetMultipleAttributes     Set multiple attribute of multiple (always the first x) objects:
Attributes2 = {'Name'; 'CostPerKm'};
Link_Name_Cost = {'Name1', 12; 'Name2', 7; 'Name3', 5; 'Name4', 3};
Vissim.Net.Links.SetMultipleAttributes(Attributes2, Link_Name_Cost);

% SetAllAttValues           Set all attributes of one object to one value:
Attribute = 'Name';
Link_Name = 'All Links have the same Name';
Vissim.Net.Links.SetAllAttValues(Attribute, Link_Name);
Attribute = 'CostPerKm';
Cost = 5.5;
Vissim.Net.Links.SetAllAttValues(Attribute, Cost);
% Note the method SetAllAttValues has a 3rd optional input: Optional ByVal add As Boolean = False; Use only for numbers!
Vissim.Net.Links.SetAllAttValues(Attribute, Cost, true); % setting the 3rd input to true, will add 5.5 to all previous costs!


%% ========================================================================
% Simulation
%==========================================================================

% Chose Random Seed
Random_Seed = 42;
set(Vissim.Simulation, 'AttValue', 'RandSeed', Random_Seed);

% To start a simulation you can run a single step:
Vissim.Simulation.RunSingleStep;
% Or run the simulation continuous (it stops at breakpoint or end of simulation)
End_of_simulation = 600; % simulation second [s]
set(Vissim.Simulation, 'AttValue', 'SimPeriod', End_of_simulation);
Sim_break_at = 200; % simulation second [s]
set(Vissim.Simulation, 'AttValue', 'SimBreakAt', Sim_break_at);
% Set maximum speed:
set(Vissim.Simulation, 'AttValue', 'UseMaxSimSpeed', true);
% Hint: to change the speed use: set(Vissim.Simulation, 'AttValue', 'SimSpeed', 10); % 10 => 10 Sim. sec. / s
Vissim.Simulation.RunContinuous;

% To stop the simulation:
Vissim.Simulation.Stop;

%% ========================================================================
% Access during simulation
%==========================================================================
% Note: All of commands of "Read and Set attributes (vehicles)" can also be executed during a
% simulation (e.g. changing signal controller program, setting relative flow of a static vehicle route,
% changing the vehicle input, changing the vehicle composition).

Sim_break_at = 198; % simulation second [s]
set(Vissim.Simulation, 'AttValue', 'SimBreakAt', Sim_break_at);
Vissim.Simulation.RunContinuous; % Start the simulation until SimBreakAt (198s)

% Get the state of a signal head:
SH_number = 1; % SH = SignalHead
State_of_SH = get(Vissim.Net.SignalHeads.ItemByKey(SH_number), 'AttValue', 'SigState'); % possible output e.g. 'GREEN', 'RED', 'AMBER', 'REDAMBER'
disp(['Actual state of SignalHead(',num2str(SH_number),') is:',32,State_of_SH]) % char(32) is whitespace

% Set the state of a signal controller:
% Note: Once a state of a signal group is set, the attribute "ContrByCOM" is automatically set to True. Meaning the signal group will keep this state until another state is set by COM or the end of the simulation
% To switch back to the defined signal controller, set the attribute signal "ContrByCOM" to False (example see below).
SC_number = 1; % SC = SignalController
SG_number = 1; % SG = SignalGroup
SignalController = Vissim.Net.SignalControllers.ItemByKey(SC_number);
SignalGroup = SignalController.SGs.ItemByKey(SG_number);
new_state = 'GREEN'; %possible values e.g. 'GREEN', 'RED', 'AMBER', 'REDAMBER'
set(SignalGroup, 'AttValue', 'SigState', new_state);
% Note: The signal controller can only be called at whole simulation seconds, so the state will be set in Vissim at the next whole simulation second, here 199s
% Simulate so that the new state is active in the Vissim simulation:
Sim_break_at = 200; % simulation second [s]
set(Vissim.Simulation, 'AttValue', 'SimBreakAt', Sim_break_at);
Vissim.Simulation.RunContinuous; % Start the simulation until SimBreakAt (200s)
% Give the control back:
set(SignalGroup, 'AttValue', 'ContrByCOM', false);

% Information about all vehicles in the network (in the current simulation second):
% In the following, 4 different methods to access attributes are shown:
% Method #1: Loop over all Vehicles using "GetAll"
% Method #2: Loop over all Vehicles using Object Enumeration (not supported in combination with Matlab)
% Method #3: Using the Iterator
% Method #4: Accessing all attributes directly using "GetMultiAttValues" (fast way if you want the attributes of all vehicles)
% Method #5: Accessing all attributes directly using "GetMultipleAttributes" (even more faster)
% The result of the four methods is the same (except the format).

% Method #1: Loop over all Vehicles:
All_Vehicles = Vissim.Net.Vehicles.GetAll; % get all vehicles in the network at the actual simulation second
for cnt_Veh = 1 : length(All_Vehicles),
    veh_number      = get(All_Vehicles{cnt_Veh}, 'AttValue', 'No');
    veh_type        = get(All_Vehicles{cnt_Veh}, 'AttValue', 'VehType');
    veh_speed       = get(All_Vehicles{cnt_Veh}, 'AttValue', 'Speed');
    veh_position    = get(All_Vehicles{cnt_Veh}, 'AttValue', 'Pos');
    veh_linklane    = get(All_Vehicles{cnt_Veh}, 'AttValue', 'Lane');
    disp([  num2str(veh_number)     ,9, ... # char(9) is tabulator
            num2str(veh_type)       ,9, ...
            num2str(veh_speed)      ,9, ...
            num2str(veh_position)   ,9, ...
            veh_linklane            ...
         ]);
end

% Method #2: Loop over all Vehicles using Object Enumeration (not supported in combination with Matlab)

% Method #3: Using the Iterator (this method is a little bit slower)
Vehicles_Iterator = Vissim.Net.Vehicles.Iterator;
while Vehicles_Iterator.Valid
    Vehicle = Vehicles_Iterator.Item;
    veh_number      = get(Vehicle, 'AttValue', 'No');
    veh_type        = get(Vehicle, 'AttValue', 'VehType');
    veh_speed       = get(Vehicle, 'AttValue', 'Speed');
    veh_position    = get(Vehicle, 'AttValue', 'Pos');
    veh_linklane    = get(Vehicle, 'AttValue', 'Lane');
    disp([  num2str(veh_number)     ,9, ... % char(9) is tabulator
            num2str(veh_type)       ,9, ...
            num2str(veh_speed)      ,9, ...
            num2str(veh_position)   ,9, ...
            veh_linklane            ...
         ]);
     Vehicles_Iterator.Next;
end

% Method #4: Accessing all Attributes directly using "GetMultiAttValues" (fast)
veh_numbers     = Vissim.Net.Vehicles.GetMultiAttValues('No');      % Output 1. column:consecutive number; 2. column: AttValue
veh_types       = Vissim.Net.Vehicles.GetMultiAttValues('VehType'); % Output 1. column:consecutive number; 2. column: AttValue
veh_speeds      = Vissim.Net.Vehicles.GetMultiAttValues('Speed');   % Output 1. column:consecutive number; 2. column: AttValue
veh_positions   = Vissim.Net.Vehicles.GetMultiAttValues('Pos');     % Output 1. column:consecutive number; 2. column: AttValue
veh_linklanes   = Vissim.Net.Vehicles.GetMultiAttValues('Lane');    % Output 1. column:consecutive number; 2. column: AttValue
disp([veh_numbers(:,2), veh_types(:,2), veh_speeds(:,2), veh_positions(:,2), veh_linklanes(:,2)]) % only display the 2nd column

% Method #5: Accessing all attributes directly using "GetMultipleAttributes" (fastest)
all_veh_attributes = Vissim.Net.Vehicles.GetMultipleAttributes({'No'; 'VehType'; 'Speed'; 'Pos'; 'Lane'});
disp(all_veh_attributes)


%% Operations at one specific vehicle:
All_Vehicles    = Vissim.Net.Vehicles.GetAll; % get all vehicles in the network at the actual simulation second
Vehicle         = All_Vehicles{1};
% alternatively with ItemByKey:
% veh_number = 66; % the same as: get(All_Vehicles{1}, 'AttValue', 'No')
% Vehicle = Vissim.Net.Vehicles.ItemByKey(veh_number);

% Set desired speed to a vehicle:
new_desspeed = 30;
set(Vehicle, 'AttValue', 'DesSpeed', new_desspeed);

% Move a vehicle:
link_number     = 1;
lane_number     = 1;
link_coordinate = 70;
Vehicle.MoveToLinkPosition(link_number, lane_number, link_coordinate); % This function will operate during the next simulation step
% Hint: In earlier Vissim releases, the name of the function was: MoveToLinkCoordinate

Vissim.Simulation.RunSingleStep; % Next Step, so that the vehicle gets moved.

% Remove a vehicle:
veh_number      = get(Vehicle, 'AttValue', 'No');
Vissim.Net.Vehicles.RemoveVehicle(veh_number);

% Putting a new vehicle to the network:
vehicle_type = 100;
desired_speed = 53; % unit according to the user setting in Vissim [km/h or mph]
link = 1;
lane = 1;
xcoordinate = 15; % unit according to the user setting in Vissim [m or ft]
interaction = true; % optional boolean
new_Vehicle = Vissim.Net.Vehicles.AddVehicleAtLinkPosition(vehicle_type, link, lane, xcoordinate, desired_speed, interaction);
% Note: In earlier Vissim releases, the name of the function was: AddVehicleAtLinkCoordinate

% Make Screenshots of the intersection 2D and 3D:
% ZoomTo:
% Zooms the view to the rectangle defined by the two points  (x1, y1) and (x2,y2),  which  are  given  in  world coordinates.  If  the  rectangle  proportions  differ
% from  the  proportions  of  the  network  window,  the  specified  rectangle  will  be centred in the network editor window.
X1 = 250;
Y1 = 30;
X2 = 350;
Y2 = 135;
Vissim.Graphics.CurrentNetworkWindow.ZoomTo(X1, Y1, X2, Y2);

% Make a Screenshot in 2D:
% It  creates  a  graphic  file  of  the  VISSIM  main  window  formatted  according  to its extension: PNG, TIFF, GIF, JPG, JPEG or BMP. A BMP file will be written if the extension can not be recognized.
Filename_screenshot = 'screenshot2D.jpg'; % to set to a specific path: "C:\Screenshots\screenshot2D.jpg"
sizeFactor = 1; % 1: original size, 2: doubles size
Vissim.Graphics.CurrentNetworkWindow.Screenshot(Filename_screenshot, sizeFactor);

% Make a Screenshot in 3D:
% Set 3D Mode:
set(Vissim.Graphics.CurrentNetworkWindow, 'AttValue', '3D', 1)
% Set the camera position (viewing angle):
xPos = 270;
yPos = 30;
zPos = 15;
yawAngle = 45;
pitchAngle = 10;
Vissim.Graphics.CurrentNetworkWindow.SetCameraPositionAndAngle(xPos, yPos, zPos, yawAngle, pitchAngle);
Filename_screenshot = 'screenshot3D.jpg'; % to set to a specific path: "C:\Screenshots\screenshot3D.jpg"
Vissim.Graphics.CurrentNetworkWindow.Screenshot(Filename_screenshot, sizeFactor);

% Set 2D Mode and old Network position:
set(Vissim.Graphics.CurrentNetworkWindow, 'AttValue', '3D', 0)
X1 = -10;
Y1 = -10;
X2 = 600;
Y2 = 300;
Vissim.Graphics.CurrentNetworkWindow.ZoomTo(X1, Y1, X2, Y2);

% Continue the simulation until end of simulation (get(Vissim.Simulation, 'AttValue', 'SimPeriod'))
Vissim.Simulation.RunContinuous;


%% ========================================================================
% Results of Simulations:
%==========================================================================

% Delete all previous simulation runs first:
simRuns = Vissim.Net.SimulationRuns.GetAll; 
for simRunNo = 1 : length(simRuns),
    Vissim.Net.SimulationRuns.RemoveSimulationRun( simRuns{simRunNo} );
end

% Run 3 Simulations at maximum speed:
% Activate QuickMode:
set(Vissim.Graphics.CurrentNetworkWindow, 'AttValue', 'QuickMode', 1)
Vissim.SuspendUpdateGUI; %  stop updating of the complete Vissim workspace (network editor, list, chart and signal time table windows)
% Alternatively, load a layout (*.laxy) where dynamic elements (vehicles and pedestrians) are not visible:
% Vissim.LoadLayout(fullfile(Path_of_COM_Basic_Commands_network, 'COM Basic Commands - Hide vehicles.layx')); % loading a layout where vehicles are not displayed => much faster simulation
End_of_simulation = 600;
set(Vissim.Simulation, 'AttValue', 'SimPeriod', End_of_simulation);
Sim_break_at = 0; % simulation second [s] => 0 means no break!
set(Vissim.Simulation, 'AttValue', 'SimBreakAt', Sim_break_at);
% Set maximum speed:
set(Vissim.Simulation, 'AttValue', 'UseMaxSimSpeed', true);
for cnt_Sim = 1 : 3,
    set(Vissim.Simulation, 'AttValue', 'RandSeed', cnt_Sim);
    Vissim.Simulation.RunContinuous;
end
Vissim.ResumeUpdateGUI; % allow updating of the complete Vissim workspace (network editor, list, chart and signal time table windows)
set(Vissim.Graphics.CurrentNetworkWindow, 'AttValue', 'QuickMode', 0) % deactivate QuickMode
% Vissim.LoadLayout(fullfile(Path_of_COM_Basic_Commands_network, 'COM Basic Commands.layx')); % loading a layout to display vehicles again


% List of all Simulation runs:
Attributes      = {'Timestamp'; 'RandSeed'; 'SimEnd'};
List_Sim_Runs = Vissim.Net.SimulationRuns.GetMultipleAttributes(Attributes);
disp(List_Sim_Runs) % show the List


% Get the results of Vehicle Travel Time Measurements:
Veh_TT_measurement_number = 2;
Veh_TT_measurement = Vissim.Net.VehicleTravelTimeMeasurements.ItemByKey(Veh_TT_measurement_number);
% Syntax to get the travel times:
%   get(Veh_TT_measurement, 'AttValue', 'TravTm(sub_attribut_1, sub_attribut_2, sub_attribut_3)');
%
% sub_attribut_1: SimulationRun
%       1, 2, 3, ... Current:     the value of one specific simulation (number according to the attribute "No" of Simulation Runs (see List of Simulation Runs))
%       Avg, StdDev, Min, Max:    aggregated value of all simulations runs: Avg, StdDev, Min, Max
% sub_attribut_2: TimeInterval
%       1, 2, 3, ... Last:        the value of one specific time interval (number of time interval always starts at 1 (first time interval), 2 (2nd TI), 3 (3rd TI), ...)
%       Avg, StdDev, Min, Max:    aggregated value of all time interval of one simulation: Avg, StdDev, Min, Max
%       Total:                    sum of all time interval of one simulation
% sub_attribut_3: VehicleClass
%       10, 20 or All             values only from vehicles of the defined vehicle class number (according to the attribute "No" of Vehicle Classes)
%                                 Note: You can only access the results of specific vehicle classes if you set it in Evaluation > Configuration > Result Attributes
%
% The value of on time interval is the arithmetic mean of all single travel times of the vehicles.

% Example #1:
% Average of all simulations (1. input = Avg)
% 	of the average of all time intervals  (2. input = Avg)
%   of all vehicle classes (3. input = All)
TT      = get(Veh_TT_measurement, 'AttValue', 'TravTm(Avg,Avg,All)');
No_Veh  = get(Veh_TT_measurement, 'AttValue', 'Vehs  (Avg,Avg,All)');
disp(['Average travel time all time intervals of all simulation of all vehicle classes:',32,num2str(TT),32,'(number of vehicles:',32,num2str(No_Veh),')']) % char(32) is whitespace

% Example #2:
% Value of the Current simulation (1. input = Current)
% 	of the maximum of all time intervals (2. input = Max)
%   of vehicle class HGV (3. input = 20)
TT      = get(Veh_TT_measurement, 'AttValue', 'TravTm(Current,Max,20)');
No_Veh  = get(Veh_TT_measurement, 'AttValue', 'Vehs  (Current,Max,20)');
disp(['Maximum travel time of all time intervals of the current simulation of vehicle class HGV:',32,num2str(TT),32,'(number of vehicles:',32,num2str(No_Veh),')']) % char(32) is whitespace

% Example #3: Note: A Travel times from 2nd simulation run must be available
% Value of the 2nd simulation (1. input = 2)
% 	of the 1st time interval (2. input = 1)
%   of all vehicle classes (3. input = All)
% TT      = get(Veh_TT_measurement, 'AttValue', 'TravTm(2,1,All)');
% No_Veh  = get(Veh_TT_measurement, 'AttValue', 'Vehs  (2,1,All)');
% disp(['Travel time of the 1st time interval of the 2nd simulation of all vehicle classes:',32,num2str(TT),32,'(number of vehicles:',32,num2str(No_Veh),')']) % char(32) is whitespace


% Data Collection
DC_measurement_number = 1;
DC_measurement = Vissim.Net.DataCollectionMeasurements.ItemByKey(DC_measurement_number);
% Syntax to get the data:
%   get(DC_measurement, 'AttValue', 'Vehs(sub_attribut_1, sub_attribut_2, sub_attribut_3)');
%
% sub_attribut_1: SimulationRun (same as described at Vehicle Travel Time Measurements)
% sub_attribut_2: TimeInterval  (same as described at Vehicle Travel Time Measurements)
% sub_attribut_3: VehicleClass  (same as described at Vehicle Travel Time Measurements)
%
% The value of on time interval is the arithmetic mean of all single values of the vehicles.

% Example #1:
% Average value of all simulations (1. input = Avg)
% 	of the 1st time interval (2. input = 1)
%   of all vehicle classes (3. input = All)
No_Veh          = get(DC_measurement, 'AttValue', 'Vehs        (Avg,1,All)'); % number of vehicles
Speed           = get(DC_measurement, 'AttValue', 'Speed       (Avg,1,All)'); % Speed of vehicles
Acceleration    = get(DC_measurement, 'AttValue', 'Acceleration(Avg,1,All)'); % Acceleration of vehicles
Length          = get(DC_measurement, 'AttValue', 'Length      (Avg,1,All)'); % Length of vehicles
disp(['Data Collection #',num2str(DC_measurement_number),': Average values of all Simulations runs of 1st time interval of all vehicle classes:'])
disp(['#vehicles:',32,num2str(No_Veh),'; Speed:',32,num2str(Speed),'; Acceleration:',32,num2str(Acceleration),'; Length:',32,num2str(Length)]) % char(32) is whitespace



% Queue length
% Syntax to get the data:
%   get(QueueCounter, 'AttValue', 'QLen(sub_attribut_1, sub_attribut_2)');
%
% sub_attribut_1: SimulationRun (same as described at Vehicle Travel Time Measurements)
% sub_attribut_2: TimeInterval  (same as described at Vehicle Travel Time Measurements)
%

% Example #1:
% Average value of all simulations (1. input = Avg)
% 	of the average of all time intervals (2. input = Avg)
QC_number = 1;
maxQ = get(Vissim.Net.QueueCounters.ItemByKey(QC_number),'AttValue', 'QLenMax(Avg, Avg)');
disp(['Average maximum Queue length of all simulations and time intervals of Queue Counter #',num2str(QC_number),':',32,num2str(maxQ)]) % char(32) is whitespace



%% ========================================================================
% Saving
%==========================================================================
Filename = fullfile(Path_of_COM_Basic_Commands_network, 'COM Basic Commands saved.inpx');
Vissim.SaveNetAs(Filename)
Filename = fullfile(Path_of_COM_Basic_Commands_network, 'COM Basic Commands saved.layx');
Vissim.SaveLayout(Filename)


%% ========================================================================
% End Vissim
%==========================================================================
Vissim.release
