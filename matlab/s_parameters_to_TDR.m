clc;
clear; 
close all; 

filename = "C:\00_Asenjo\00_Project\Ketupa\simulation\Ansys_HFSS_from_si9000\sim_results\xx.s4p";
backplane = sparameters(filename);
data = backplane.Parameters;
freq = backplane.Frequencies;
z0 = backplane.Impedance;

diffdata = s2sdd(data);
diffsparams = sparameters(diffdata,freq,2*z0);
z0differential = diffsparams.Impedance;

s11 = rfparam(diffsparams,1,1);
s11fit = rational(freq,s11);
Ts = 5e-12; 
N = 5000;
Trise = 5e-11;
[tdr,tdrT] = stepresp(s11fit,Ts,N,Trise);

zLt = gamma2z(tdr, z0differential);
figure
plot(tdrT*1e9,zLt,'r','LineWidth',2)
ylabel('Differential TDR (Ω)')
xlabel('Time (ns)')
legend('Calculated TDR')