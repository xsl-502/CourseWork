%支撑
%使用软件为：MATLAB R2024a
%运行文件命令：输入对应的文件名即可


% Clear workspace and command window

clear;

clc;

%% 1. Data Reading and Column Renaming

% Read Excel file and preserve original column headers

filename = '1945-2023.xlsx';

data = readtable(filename, 'VariableNamingRule', 'preserve');

% Display the first few rows of the original data to confirm it was read correctly

disp('First few rows of the original data:');

disp(head(data));

% Rename columns to English names

% Assuming original column names are:

% '台风编号', '台风中文名称', '台风英文名称', '台风起始时间', '台风结束时间',

% '当前台风时间', '经度', '纬度', '台风强度', '台风等级', '风速', '气压',

% '移动方向', '移动速度'

% Create a mapping of column names

original_names = {'台风编号', '台风中文名称', '台风英文名称', '台风起始时间', '台风结束时间', ...

'当前台风时间', '经度', '纬度', '台风强度', '台风等级', '_风速', '气压', ...

'移动方向', '移动速度'};

new_names = {'_typhoon_id', '_typhoon_chinese_name', '_typhoon_english_name', '_typhoon_start_time', '_typhoon_end_time', ...

'_current_typhoon_time', '_longitude', '_latitude', '_typhoon_intensity', '_typhoon_grade', '_wind_speed', '_pressure', ...

'_movement_direction', '_movement_speed'};

% Check if original column names exist and rename them

for i = 1:length(original_names)

    if any(strcmp(data.Properties.VariableNames, original_names{i}))

        data.Properties.VariableNames{strcmp(data.Properties.VariableNames, original_names{i})} = new_names{i};

    else

        error('Column name "%s" does not exist in the data table. Please check the column names in the Excel file.', original_names{i});

    end

end

% Display the first few rows of the renamed data table

disp('First few rows of the renamed data:');

disp(head(data));

%% 2. Time Format Conversion

% Convert '_current_typhoon_time' from string to datetime format

time_data = data._current_typhoon_time;

% Check the data type of '_current_typhoon_time'

if iscell(time_data)

    % If it is a cell array, convert to string array

    time_data = string(time_data);

end

if isstring(time_data) || ischar(time_data)

    % If it is a string or character array, convert using the specified format

    try

        data._current_typhoon_time = datetime(time_data, 'InputFormat', 'yyyy-MM-dd''T''HH:mm:ss');

    catch ME

        disp('Error in time format conversion. Please check if the input format is ''yyyy-MM-ddTHH:mm:ss''.');

        rethrow(ME);

    end

elseif isnumeric(time_data)

    % If it is numeric, assume it is an Excel serial date number and convert

    try

        data._current_typhoon_time = datetime(time_data, 'ConvertFrom', 'excel');

    catch ME

        disp('Error in time format conversion. Please confirm if the numeric input is an Excel serial date number.');

        rethrow(ME);

    end

else

    error('Unknown time data type: %s', class(time_data));

end

% Create a new numeric time format 'YYYYMMDDHH'

data._time_num = year(data._current_typhoon_time) * 1e6 + month(data._current_typhoon_time) * 1e4 + ...

    day(data._current_typhoon_time) * 1e2 + hour(data._current_typhoon_time);

% Display the results of the time conversion

disp('First few rows of the converted time format:');

disp(head(data(:, {'_current_typhoon_time', '_time_num'})));

%% 3. Handling Missing Values

% Remove rows with missing values in '_pressure', '_typhoon_grade', '_typhoon_intensity'

data = rmmissing(data, 'DataVariables', {'_pressure', '_typhoon_grade', '_typhoon_intensity'});

% Calculate '_movement_direction'

for i = 1:(height(data) - 1)

    if data._typhoon_id(i) == data._typhoon_id(i + 1)

        lat1 = data._latitude(i);

        lon1 = data._longitude(i);

        lat2 = data._latitude(i + 1);

        lon2 = data._longitude(i + 1);

        delta_lon = lon2 - lon1;

        y = sind(delta_lon) * cosd(lat2);

        x = cosd(lat1) * sind(lat2) - sind(lat1) * cosd(lat2) * cosd(delta_lon);

        angle = atan2d(y, x);

        if angle < 0

            angle = angle + 360;

        end

        if isnan(data._movement_direction(i))

            data._movement_direction(i) = angle;

        end

    else

        data._movement_direction(i) = NaN;

    end

end

% The last row's '_movement_direction' cannot be calculated, fill it with NaN

data._movement_direction(end) = NaN;

% Calculate '_movement_speed' (unit: m/s)

radius_earth = 6371000; % Earth radius in meters

for i = 1:(height(data) - 1)

    if data._typhoon_id(i) == data._typhoon_id(i + 1) && isnan(data._movement_speed(i))

        lat1 = deg2rad(data._latitude(i));

        lon1 = deg2rad(data._longitude(i));

        lat2 = deg2rad(data._latitude(i + 1));

        lon2 = deg2rad(data._longitude(i + 1));

        delta_lat = lat2 - lat1;

        delta_lon_rad = lon2 - lon1;

        a = sin(delta_lat / 2)^2 + cos(lat1) * cos(lat2) * sin(delta_lon_rad / 2)^2;

        c = 2 * atan2(sqrt(a), sqrt(1 - a));

        distance = radius_earth * c; % Distance between two points in meters

        time_diff_seconds = seconds(data._current_typhoon_time(i + 1) - data._current_typhoon_time(i)); % Time difference in seconds

        if time_diff_seconds > 0

            data._movement_speed(i) = distance / time_diff_seconds; % Movement speed in m/s

        else

            data._movement_speed(i) = NaN;

        end

    end

end

% The last row's '_movement_speed' cannot be calculated, fill it with NaN

data._movement_speed(end) = NaN;

%% 4. Outlier Handling

% Define outlier thresholds

extreme_wind_speed = 110; % Wind speed threshold in m/s

low_pressure = 800; % Pressure threshold in hPa

% Mark extreme wind speed and low pressure

is_wind_speed_extreme = data._wind_speed > extreme_wind_speed;

low_pressure_extreme = data._pressure < low_pressure;

% Set outliers to NaN

data._wind_speed(is_wind_speed_extreme) = NaN;

data._pressure(low_pressure_extreme) = NaN;

% Perform linear interpolation for outliers

data._wind_speed = fillmissing(data._wind_speed, 'linear');

data._pressure = fillmissing(data._pressure, 'linear');

%% 5. Remove Remaining Missing Values

% Remove all rows containing NaN

data = rmmissing(data);

% Display a summary of the processed data

disp('Summary of processed data:');

summary(data);

%% 6. Data Encoding

% Typhoon grade and intensity have been manually encoded and do not require automatic encoding

% Handle missing values in '_typhoon_chinese_name' and '_typhoon_english_name'

% Fill with empty strings

data._typhoon_chinese_name = fillmissing(data._typhoon_chinese_name, 'constant', '');

data._typhoon_english_name = fillmissing(data._typhoon_english_name, 'constant', '');

%% 7. Data Normalization

% Select numerical variables to normalize

numerical_vars = {'_wind_speed', '_pressure', '_movement_speed', '_typhoon_intensity', '_typhoon_grade'};

% Calculate mean and standard deviation

mu = mean(data{:, numerical_vars});

sigma = std(data{:, numerical_vars});

% Normalize the data

data_standardized = data;

data_standardized{:, numerical_vars} = (data{:, numerical_vars} - mu) ./ sigma;

%% 8. Statistical Description

% Calculate basic statistics

disp('Basic statistics:');

summary(data_standardized(:, numerical_vars));

%% 9. Normality Test (K-S Test)

% Perform K-S test on numerical variables to check for normal distribution

for var = numerical_vars

    [h, p] = kstest(data_standardized{:, var{1}}); % K-S test

    fprintf('K-S test result for variable %s: h = %d, p = %.4f\n', var{1}, h, p);

end

%% 10. Save Data

% Save processed data to a new Excel file

output_filename = 'processed_typhoon_data.xlsx';

writetable(data_standardized, output_filename);

disp(['Processed data has been saved to ', output_filename]);