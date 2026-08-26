%  moduliloop.m
%
% This program converts compliance measurements into storage and loss
% moduli, using the method described by 
% R M L Evans, Manlio Tassieri, Dietmar Auhl and Thomas A Waigh in 
% Phys. Rev. E 80, 012501 (2009).
% 
% Coded by Stephen R. Baker, Leeds Institute for Cardiovascular and Medabolic 
% Medicine, University of Leeds, U.K.
%
% Modified from code by Chirag Kalelkar, Complex Fluids and Polymer Engineering Group,
% National Chemical Laboratory, Pune, India,
% Modified from code by David Pearce, University of Manchester, U.K.
% Edited by R M L Evans, University of Leeds, U.K.
% 
% Inputs:
%   J = compliance data (column vector)
%   t = time data (column vector)
%   eta = the long-time viscosity = reciprocal of gradient of asymptote of J_b1(t).
%   J0 = value of compliance extrapolated to time t=0.
%   dpoints = number of output points required.
%   maxfreq = upper limit of frequency range to output.
%             Should usually set this to 1/t(1).
% 
% Output:
%   Data = frequency, G' and G'' for each bead

close all;

% Get the initital input files and rename them so that they can run through
% the program

% file1 = dlmread(uigetfile('file1.txt', 'Initial Values (Smaller File)')); 
% file2 = dlmread(uigetfile('file2.txt', 'Data (Larger File)'));
if ~exist('file1_path', 'var')
    file1_path = uigetfile('file1.txt', 'Initial Values (Smaller File)');
end

if ~exist('file2_path', 'var')
    file2_path = uigetfile('file2.txt', 'Data (Larger File)');
end

file1 = dlmread(file1_path);
file2 = dlmread(file2_path);

% *************************************
% *** Edit the maxfreq relating to the maximum frequency of your camera ***
maxfreq = 25;
[nr, nc] = size(file2);
dpoints = nr;
M = floor(nc/2);
if nc ~= 2*M
    error(['file2 has an odd number of columns (' num2str(nc) ') — expected pairs of (time, J). ' ...
           'Check the "Data Output to Matlab 1" sheet for a stray/partial column, ' ...
           'or a bead that failed after Prism.']);
end

% Establish empty data structure to speed up computation
GData = zeros(M,dpoints,3);
J0 = cell(1,M);
eta = cell(1,M);
t = cell(1, M);
J = cell(1,M);
GStar = cell(1,M);
Data = zeros(dpoints,3*M);

for k=1:M
  
  J0{k}=file1(:,2*k);
  eta{k}=file1(:,(2*k-1));
  t{k}=file2(:, (2*k-1));
  J{k}=file2(:, 2*k);
  
  frange = logspace(log10(1/t{k}(length(t{k}))),log10(maxfreq),dpoints);
  
  for ww = 1:dpoints
      w = frange(ww);
      GStar{k} = 1i*w./(1i*w*J0{k} + (1-exp(-1i*w*t{k}(1)))*(J{k}(1)-J0{k})/t{k}(1) + exp(-1i*w*t{k}(size(t{k},1)))/eta{k} + sum(diff(J{k})./diff(t{k}).*(exp(-1i*w*(t{k}(1:(numel(t{k})-1))))-exp(-1i*w*(t{k}(2:numel(t{k})))))));
      GData(k,ww,:) = [w real(GStar{k}) imag(GStar{k})];

  end

% Print the Figure for each bead (skipped in batch mode: figure windows disabled)
% figure;
% loglog(GData(k,:,1),GData(k,:,2),'b');
% hold on
% loglog(GData(k,:,1),GData(k,:,3),'r');
% legend('G''','G''''','Location','NorthWest');
% xlabel('\omega [s^{-1}]');
% ylabel('Moduli [Pa]');

end

wf=GData(:,:,1)';
Gp=GData(:,:,2)';
Gpp=GData(:,:,3)';

% This loop puts everything in the right format for the output file, mainly
% wf1, Gp1, Gpp1, wf2, Gp2, Gpp2, ... 
for ii=1:M
    Data(:,3*ii-2)=wf(:,ii);
    Data(:,3*ii-1)=Gp(:,ii);
    Data(:,3*ii)=Gpp(:,ii);
end

% Write output to the same directory as the input files (set by Python)
[analysis_dir, ~, ~] = fileparts(file1_path);
outputfile = fullfile(analysis_dir, 'AllBeads.csv');
csvwrite(outputfile,Data);