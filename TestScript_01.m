%====================================================================
% Script for Running Fit in SimBiology
% Model: 2 compartmental
% Dosing: Only IV
% Author: J Jeevan Roy
%====================================================================

%% OPERATIONS ON WORKSPACE & COMMAND WINDOW

% Clear Workspace
clear %Removes all variables from the workspace.
close all % Closes all open figure windows.
clc %Clears the command window.
warning('off') %Turns off warnings (use with caution, as this may suppress important messages).
disp('1. Workspace cleared, ready for a fresh run.✅');

%% HARDCODING INFO EXCEL FILE

hardcodingDataFile = 'Hardcoding - info.xlsx'; % Replace with the actual file name
hardcodingDataTable = readtable(hardcodingDataFile);

params = struct();
for i = 1:height(hardcodingDataTable)
    cleanValue = strrep(hardcodingDataTable.value{i}, '''', ''); % Remove extra quotes
    params.(hardcodingDataTable.variableName{i}) = cleanValue;
end

disp('2. Hardcoded info is stored in variable params in key-value pairs format where keys are the variable names.✅')


%% LOAD PROJECT FILE

% Project folder path
project_folder_path = params.projectFolderPath;
%project_folder_path = strcat('asdf',project_folder_path); %This line is
%added to check the if condition on folder path

project_name = params.projectFileName;
project_full_path = fullfile(project_folder_path, project_name);


% Load project only after validation
try
    proj = sbioloadproject(project_full_path);
    disp(['3. Project loaded successfully! ✅  [', project_full_path, ']']);
catch ME
    error('❌ Failed to load project: %s\nError: %s', project_full_path, ME.message);
end

%% SELECT MODEL

% Access all models within the project dynamically
modelFields = fieldnames(proj);  % Get all fields in the project object
modelFields = modelFields(startsWith(modelFields, 'm'));  % Filter for model fields ('m1', 'm2', etc.)

% Select a specific model by NAME
% Define the name of the model you want to select
specificModelName = params.modelName;  % Replace with the name of the model you want to use------------------------------------------------->

% Find the model by name
modelFound = false;
for i = 1:numel(modelFields)
    if strcmp(proj.(modelFields{i}).Name, specificModelName)
        model = proj.(modelFields{i});
        modelFound = true;
        break;
    end
end

% Check if the model was found
if modelFound
    disp(['4. Currently, the model being used is [', specificModelName, '].']);
else
    error(['4. Model with name "', specificModelName, '" not found.']);
end



%% CHANGE MODEL UNITS 

% Change model time units to days
configset = getconfigset(model);
X_axis_units = params.xAxisUnits; %------------------------------------------------------------------------------------------->

set(configset, 'TimeUnits', X_axis_units); % Set model's time units to "day"
disp(['6. Time units(x axis units) are changed to [',X_axis_units,'].']);

%% ENABLE UNIT CONVERSION

% Enable unit conversion
configset.CompileOptions.UnitConversion = true;

%% GROUPED DATA %% LOAD DATA (NOTE: CONSIDERING THE DIRECTORY IS THE SAME AS THE PROJECT FILE DIRECTORY)

% Load experimental data
dataFile = params.pkDataFilePath; % Replace with the actual file name----------------------------------------------->

try
    allData = readcell(dataFile);%readcell reads the entire Excel sheet as a raw cell array, preserving different data types within the same array.
    %Unlike readtable, which returns a table with named columns
    disp('5. pk Data Excel file is loaded ✅')
catch ME
    error('5. ❌ PK data read failed: %s\n-> %s', dataFile, ME.message);
end

headers = allData(1,:);
disp('The headers are extracted.');
disp(headers);
units = allData(2,:);
disp('The units are extracted');
disp(units);
numericData = allData(3:end,:);
disp('The units are extracted.');
disp(numericData);

% Convert to table with proper headers
try
    dataTable = cell2table(numericData, 'VariableNames', headers);
catch
    error('5. ❌ Column count mismatch between headers (%d) and data rows', numel(headers));
end

columnNames = dataTable.Properties.VariableNames;
drugConcCol = columnNames(contains(columnNames, 'DrugConc')); 
TimeCol = columnNames(contains(columnNames, 'Time'));
DoseCol = columnNames(contains(columnNames,'Dose'));


groupedDataObj = groupedData(dataTable, 'ID', TimeCol{1});
disp('7. Data converted to groupedData ✅');

groupedDataObj.Properties.VariableUnits = units;

% groupDataObj =
% magic_function(params.pkDataFilePath)----------------------------------------->------------------------------
%% 


disp('Checking Variable Names');
disp(groupedDataObj.Properties.VariableNames); %column names

disp('Checking Assigned Units to the variables');
disp(groupedDataObj.Properties.VariableUnits); %column units

disp('Checking GroupedDataObj properties')
disp(groupedDataObj.Properties); %
%%

dependent_variable_name = 'DrugConc_ug_mL_'; %----------------------------------------------------------->

%% RESPONSE MAP (AUTOMATED) -  RESPONSE MAP (SPECIES = COLUMN_NAME)

response_species_compartment = params.responseSpeciesCompartment;
response_species = params.responseSpecies;

% if ~ismember(response_species_compartment, get(model.Compartments, 'Name'))
%      error('8. ❌ Invalid compartment: %s\n-> Available compartments: %s',response_species_compartment, strjoin(get(model.Compartments, 'Name'), ', '));
%  end
% 
%  if ~ismember(response_species, get(model.Species, 'Name'))
%      error('8. ❌ Invalid species: %s\n-> Available species: %s', response_species, strjoin(get(model.Species, 'Name'), ', '));

% Make the relation
the_relation = strcat(response_species_compartment , '.' , response_species ,' ', ' = ', ' '  , dependent_variable_name); %----------------->
% Display it
%disp(the_relation); 
%the_relation = response_species_compartment + "." + response_species + " = " + independent_variable_name;
%responseMap = {'Cen.Drug_conc_cen = DrugConc_ug_mL_'};

% Map the response
responseMap = {the_relation};
disp(['8. Response Mapping is Done ✅. The relation is [',the_relation,'].']);



%% PARAMETERS TO BE ESTIMATED

paramsToBeEstimatedDataFilePath = params.parametersToBeEstimatedExcelFilePath;
paramsToBeEstimatedDataTable = readtable(paramsToBeEstimatedDataFilePath);


% Extract relevant columns
paramNames = paramsToBeEstimatedDataTable.Parameter;
transformations = paramsToBeEstimatedDataTable.Transformation
% ------------------------------------------>
initialValues = paramsToBeEstimatedDataTable.InitialValue;
lowerBounds = paramsToBeEstimatedDataTable.LowerBound;
upperBounds = paramsToBeEstimatedDataTable.UpperBound;

% Create estimatedInfo object dynamically
estimatedInfoObj = estimatedInfo(strcat(transformations,'(', paramNames, ')')); %transformation should be handled here - Handled. Done!!

% Assign initial values and bounds
for i = 1:length(paramNames)
    estimatedInfoObj(i).InitialValue = initialValues(i);
    estimatedInfoObj(i).Bounds = [lowerBounds(i), upperBounds(i)];
end

disp('9. The parameters to be estimated data is successfully assigned in the estimatedInfoObj ✅.')
disp(estimatedInfoObj);

%% DOSES

% DOCUMENTATION : if the input groupedData object has dosing information, 
% you can use the 'createDoses' method to construct doses.



% Create the data dose for Cen.Dose_Cen
doses1 = sbiodose(DoseCol{1});
doses1.TargetName = params.doseTargetName;  %-----------------------------> Has to be included
doses1 = createDoses(groupedDataObj, DoseCol{1}, '', doses1);

% Convert doses1 to a cell array (still needed if you plan to process doses as cells).
doses1 = num2cell(doses1);

% If no SC dosing, dosesForFit is directly built from doses1.
dosesForFit = cell(size(doses1, 1), 1);
for i = 1:numel(dosesForFit)
    dosesForFit{i} = doses1{i}; % Directly assign doses from doses1
end

%disp(dosesForFit);
disp('10. Doses: Defined using sbiodose ✅ (No SC dosing)');

%% DOSES INFO IN TABULAR FORMAT

%Displaying Doses in Tabular Format
% Assuming dosesForFit is already defined and has the properties

% Initialize variables to store the data
numRows = size(dosesForFit, 1);
Names = cell(numRows, 1);
Amounts = zeros(numRows, 1);
AmountUnits = cell(numRows, 1);
TargetNames = cell(numRows, 1);
TimeUnits = cell(numRows, 1);

% Extract data from dosesForFit
for i = 1:numRows
    Names{i} = dosesForFit{i,1}.Name;
    Amounts(i) = dosesForFit{i,1}.Amount;
    AmountUnits{i} = dosesForFit{i,1}.AmountUnits;
    TargetNames{i} = dosesForFit{i,1}.TargetName;
    TimeUnits{i} = dosesForFit{i,1}.TimeUnits;
end

% Create the table
doseTable = table(Names, Amounts, AmountUnits, TargetNames, TimeUnits, ...
    'VariableNames', {'Name', 'Amount', 'AmountUnits', 'TargetName', 'TimeUnits'});

% Display the table
disp(doseTable);

%% OPTIONS

% Define Algorithm options.
options                   = optimoptions('particleswarm');
options.FunctionTolerance = 1e-08;
options.MaxIterations     = 400;

disp('No issues until Options')

%% FIT 

% Define fit problem.
f              = fitproblem('FitFunction', 'sbiofit');
f.Data         = groupedDataObj;
f.Model        = model;
f.Estimated    = estimatedInfoObj;
f.ResponseMap  = responseMap;
f.ErrorModel   = 'exponential';
f.Doses        = dosesForFit;
f.FunctionName = 'particleswarm';
f.Options      = options;
f.ProgressPlot = true;
f.UseParallel  = true;
f.Pooled       = true;

% Estimate parameter values.
[results, simdataI] = f.fit;

% Configure the names of the resulting simulation data.
for i = 1:length(simdataI)
    simdataI(i).Name = 'SimData Individual';
end

% Assign output arguments.
args.output.results  = results;
args.output.simdataI = simdataI;


%% CREATING CV COLUMN IN RESULTS

% Assuming 'results.ParameterEstimates' is the name of your table
ParameterEstimates = results.ParameterEstimates;

% Calculate CV as StandardError / Estimate
ParameterEstimates.CV = (ParameterEstimates.StandardError ./ ParameterEstimates.Estimate)*100;

% Display the updated table
disp(ParameterEstimates);

%% PLOTTING FIT
plot(results);

%% TILED LAYOUT

% Initialize figure for plotting
figure;
hold on; % Allow multiple curves on the same plot

% Set y-axis to log scale
set(gca, 'YScale', 'log'); % This is the key line
%%

t = tiledlayout('flow');

for i = 1:size(simdataI, 1)

    % Check dose amount and set plot title
    if isempty(dosesForFit{i,1}(1,1).Amount)
        % SC dose: use the second column if the first column is empty
        doseAmount = dosesForFit{i,1}(1,2).Amount;
        plot_title = [num2str(doseAmount) 'mg (SC Dose)'];
    else
        % IV dose: use the first column if it's not empty
        doseAmount = dosesForFit{i,1}(1,1).Amount;
        plot_title = [num2str(doseAmount) 'mg (IV Dose)'];
    end

    % Extract time and output data
    time = simdataI(i,1).Time; % Simulation time points
    drugConc = simdataI(i,1).Data(:, 2); % 'Drug_conc_Cen' data is in 3rd column
    
    % Plot the data
    nexttile;
    hold on;

    plot(time, drugConc, 'r', 'LineWidth', 1.25);


    % Set y-axis to log scale
    set(gca, 'YScale', 'log');

    % Set axis limits
    xlim([-0.1, 2.1]); % X-axis range
    ylim([0.001, 10000]); % Y-axis range

    title(plot_title); % Plot title

    % Match experimental data using ID column
    matchingRows = dataTable.ID == i; % Assuming 'ID' column matches dose index
    if any(matchingRows)
        expTime = dataTable.Time_Day_(matchingRows); % Replace 'Time' with actual column name for time
        expConc = dataTable.DrugConc_ug_mL_(matchingRows); % Replace 'Concentration' with actual column name
        scatter(expTime, expConc, 20, 'MarkerEdgeColor','r','MarkerFaceColor','r','LineWidth',1.5); % Scatter plot
        % Connect the dots with a dotted line
        %plot(expTime, expConc,'HandleVisibility', 'off'); % 'HandleVisibility', 'off' prevents duplicate legend entries
    end
    hold off;

end

title(t,'PK Profiles');

xlabel(t,'Time (Day)');
ylabel(t,['Drug conc (ugmL)']);


%%