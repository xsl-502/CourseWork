% Clear environment

clear; close all; clc;

% Parameters for precipitation and wind speed relationship

alpha = 0.8;

beta = 1.5;

num_points_v = 1000; % Number of wind speed data points

V = linspace(10, 60, num_points_v)';

% Generate precipitation data with added random noise

noise_std_v = 100; % Standard deviation of noise

P_v = alpha * V.^beta + noise_std_v * randn(num_points_v,1);

% Ensure precipitation values are non-negative

P_v(P_v < 0) = 0;

% Parameters for precipitation and distance relationship

P0 = 2000; % Maximum precipitation (mm)

k = 0.05; % Decay coefficient (km^{-1})

num_points_r = 1000; % Number of distance data points

r = linspace(0, 200, num_points_r)'; % Distance range: 0 km to 200 km

% Generate precipitation data with added random noise

noise_std_r = 50; % Standard deviation of noise

P_r = P0 * exp(-k * r) + noise_std_r * randn(num_points_r,1);

% Ensure precipitation values are non-negative

P_r(P_r < 0) = 0;

%% Step 3: Save generated data to an Excel file

% Create a table with V, P_v, r, and P_r as columns

T = table(V, P_v, r, P_r);

% Specify the name of the Excel file to save

excel_filename = 'typhoon_data.xlsx';

% Write the table to an Excel file

writetable(T, excel_filename);

fprintf('Synthetic data saved to file %s\n', excel_filename);

%% Step 4: Data fitting

% 4.1 Fitting precipitation vs. wind speed relationship

% a. Power-law fitting

valid_idx_v = P_v > 0 & V > 0; % Use only positive values

logV = log(V(valid_idx_v));

logP_v = log(P_v(valid_idx_v));

% Linear regression

coeffs_v = polyfit(logV, logP_v, 1);

beta_fit = coeffs_v(1);

log_alpha_fit = coeffs_v(2);

alpha_fit = exp(log_alpha_fit);

fprintf('Power-law fit results: alpha = %.3f, beta = %.3f\n', alpha_fit, beta_fit);

% Fitted curve

P_v_fit_power = alpha_fit * V.^beta_fit;

% b. Polynomial fit (degree 2)

poly_degree = 2;

coeffs_poly_v = polyfit(V, P_v, poly_degree);

P_v_fit_poly = polyval(coeffs_poly_v, V);

% Polynomial fit (degree 3)

poly_degree = 3;

coeffs_poly3_v = polyfit(V, P_v, poly_degree);

P_v_fit_poly3 = polyval(coeffs_poly3_v, V);

% c. Exponential fit

logP_v_exp = log(P_v(valid_idx_v));

V_valid = V(valid_idx_v);

% Linear regression

coeffs_exp_v = polyfit(V_valid, logP_v_exp, 1);

b_fit_v = coeffs_exp_v(1);

log_a_fit_v = coeffs_exp_v(2);

a_fit_v = exp(log_a_fit_v);

fprintf('Exponential fit results: a = %.3f, b = %.3f\n', a_fit_v, b_fit_v);

% Fitted curve

P_v_fit_exp = a_fit_v * exp(b_fit_v * V);

% d. Gaussian Process Regression (GPR)

gprModel_v = fitrgp(V, P_v, 'KernelFunction', 'squaredexponential', 'Standardize', true);

P_v_fit_gpr = predict(gprModel_v, V);

% 4.2 Fitting precipitation vs. distance relationship

% a. Exponential decay fitting

valid_idx_r = P_r > 0 & r >= 0;

logr = log(P_r(valid_idx_r));

r_valid = r(valid_idx_r);

% Linear regression

coeffs_r = polyfit(r_valid, logr, 1);

k_fit = -coeffs_r(1);

logP0_fit = coeffs_r(2);

P0_fit = exp(logP0_fit);

fprintf('Exponential decay fit results: P0 = %.3f, k = %.3f\n', P0_fit, k_fit);

% Fitted curve

P_r_fit_exp = P0_fit * exp(-k_fit * r);

% b. Gaussian fit

gauss_eqn = 'a*exp(-((x-b)/c)^2)';

startPoints = [2000, 100, 50];

fit_gauss_r = fit(r, P_r, gauss_eqn, 'Start', startPoints);

a_gauss = fit_gauss_r.a;

b_gauss = fit_gauss_r.b;

c_gauss = fit_gauss_r.c;

fprintf('Gaussian fit results: a = %.3f, b = %.3f, c = %.3f\n', a_gauss, b_gauss, c_gauss);

% Fitted curve

P_r_fit_gauss = a_gauss * exp(-((r - b_gauss)/c_gauss).^2);

% c. Polynomial fit (degree 2)

poly_degree_r = 2;

coeffs_poly_r = polyfit(r, P_r, poly_degree_r);

P_r_fit_poly = polyval(coeffs_poly_r, r);

% Polynomial fit (degree 3)

poly_degree_r3 = 3;

coeffs_poly3_r = polyfit(r, P_r, poly_degree_r3);

P_r_fit_poly3 = polyval(coeffs_poly3_r, r);

% d. Gaussian Process Regression (GPR)

gprModel_r = fitrgp(r, P_r, 'KernelFunction', 'squaredexponential', 'Standardize', true);

P_r_fit_gpr = predict(gprModel_r, r);

%% Step 5: Visualization

% Set up figure

figure;

% 1. Relationship between precipitation and wind speed

subplot(2,2,1);

scatter(V, P_v, 10, 'b', 'filled'); hold on;

plot(V, P_v_fit_power, 'r-', 'LineWidth', 2);

plot(V, P_v_fit_poly, 'g--', 'LineWidth', 2);

plot(V, P_v_fit_poly3, 'm:', 'LineWidth', 2);

plot(V, P_v_fit_exp, 'k-.', 'LineWidth', 2);

plot(V, P_v_fit_gpr, 'c-', 'LineWidth', 2);

xlabel('Wind Speed (m/s)');

ylabel('Precipitation (mm)');

title('Relationship between Precipitation and Wind Speed');

legend('Data Points', 'Power-law Fit', 'Polynomial Degree 2', 'Polynomial Degree 3', 'Exponential Fit', 'GPR Fit', 'Location', 'northwest');

grid on;

% 2. Relationship between precipitation and distance

subplot(2,2,2);

scatter(r, P_r, 10, 'b', 'filled'); hold on;

plot(r, P_r_fit_exp, 'r-', 'LineWidth', 2);

plot(r, P_r_fit_gauss, 'g--', 'LineWidth', 2);

plot(r, P_r_fit_poly, 'm:', 'LineWidth', 2);

plot(r, P_r_fit_poly3, 'k-.', 'LineWidth', 2);

plot(r, P_r_fit_gpr, 'c-', 'LineWidth', 2);

xlabel('Distance (km)');

ylabel('Precipitation (mm)');

title('Relationship between Precipitation and Distance');

legend('Data Points', 'Exponential Decay Fit', 'Gaussian Fit', 'Polynomial Degree 2', 'Polynomial Degree 3', 'GPR Fit', 'Location', 'northwest');

grid on;

% 3. Residuals for precipitation vs. wind speed

subplot(2,2,3);

residual_power = P_v - P_v_fit_power;

residual_poly = P_v - P_v_fit_poly;

residual_poly3 = P_v - P_v_fit_poly3;

residual_exp = P_v - P_v_fit_exp;

residual_gpr = P_v - P_v_fit_gpr;

plot(V, residual_power, 'r-', V, residual_poly, 'g--', V, residual_poly3, 'm:', V, residual_exp, 'k-.', V, residual_gpr, 'c-');

xlabel('Wind Speed (m/s)');

ylabel('Residuals (mm)');

title('Residuals for Precipitation vs. Wind Speed');

legend('Power-law Residual', 'Polynomial Degree 2 Residual', 'Polynomial Degree 3 Residual', 'Exponential Residual', 'GPR Residual', 'Location', 'northwest');

grid on;

% 4. Residuals for precipitation vs. distance

subplot(2,2,4);

residual_exp_r = P_r - P_r_fit_exp;

residual_gauss_r = P_r - P_r_fit_gauss;

residual_poly_r = P_r - P_r_fit_poly;

residual_poly3_r = P_r - P_r_fit_poly3;

residual_gpr_r = P_r - P_r_fit_gpr;

plot(r, residual_exp_r, 'r-', r, residual_gauss_r, 'g--', r, residual_poly_r, 'm:', r, residual_poly3_r, 'k-.', r, residual_gpr_r, 'c-');

xlabel('Distance (km)');

ylabel('Residuals (mm)');

title('Residuals for Precipitation vs. Distance');

legend('Exponential Decay Residual', 'Gaussian Residual', 'Polynomial Degree 2 Residual', 'Polynomial Degree 3 Residual', 'GPR Residual', 'Location', 'northwest');

grid on;