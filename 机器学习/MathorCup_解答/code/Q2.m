%% Step 1: Load Data

% Load typhoon data from Excel file

filename = '贝碧嘉数据.xlsx';

data = readtable(filename, 'PreserveVariableNames', true);

% Display column names in the data table to confirm longitude and latitude column names

disp('Column names in data table:');

disp(data.Properties.VariableNames);

%% Step 2: Extract Relevant Data

% Extract longitude and latitude data

% Assuming longitude column is 'LON' and latitude column is 'LAT'

_longitude = data.('LON');

_latitude = data.('LAT');

% Save extracted longitude and latitude data to a new table

path_data = table(_longitude, _latitude);

% Save the path changes to a new Excel file

output_filename = 'Typhoon_Path_Changes_LAT_LON.xlsx';

writetable(path_data, output_filename);

disp(['Path changes saved to file: ', output_filename]);

%% Step 3: Plot Typhoon Path

% Plot typhoon path changes

figure;

plot(_longitude, _latitude, 'o-', 'Color', 'b', 'LineWidth', 1.5);

xlabel('Longitude');

ylabel('Latitude');

title('Typhoon Bebinca Path Changes');

grid on;

% Add legend

legend('Typhoon Path');
