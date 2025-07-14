clc;
clear; 
close all; 

filename = "C:\00_Asenjo\00_Project\Ketupa\simulation\Ansys_HFSS\sim_results\Edge_Fed_Rectangular_Patch_Antenna.s1p";
[~, name, ~] = fileparts(filename);
disp(name);

S_scattering=sparameters(filename');
figure(1)
rfplot(S_scattering)
title(name)