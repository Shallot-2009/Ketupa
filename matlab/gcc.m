% % clc ;
% % clear variables ; 
% % close all ; 
fprintf('--------------------------------------------------------------------------------------------------------------------------------------------\n')
main_file_name = mfilename("fullpath");
[main_pathstr, ~] = fileparts(main_file_name);
[main_pathstr, ~] = fileparts(main_pathstr);
addpath(genpath(main_pathstr));

Time_now = datetime('now');
Timestamp = datestr(Time_now, 'yyyymmdd_HHMMSS');

gcc_buildpath = struct();
%%%----------------------Create build directory-------------------------%%%
path_build = fullfile(main_pathstr,'build');  % 当前工作目录
mkdir(path_build);
  if exist(path_build, 'dir') == 7
    disp(['The build directory has been created：', main_pathstr]);
  else
    error('Failed to create build directory.');
  end
gcc_buildpath.path_build = path_build;
%%%-------------------Create COM analysis case directory-------------------%%%
caseName = ['build_' char(Timestamp)];
path_case = fullfile(path_build,caseName);  % 当前工作目录
mkdir(path_case);
  if exist(path_case, 'dir') == 7
    disp(['The case directory has been created：', path_build]);
  else
    error('Failed to create the case directory.');
  end
gcc_buildpath.path_case = path_case;
%%%------------------------Create vbsScript directory-------------------------%%%
path_vbsScript = fullfile(path_case,'vbsScript'); % 当前工作目录
mkdir(path_vbsScript);
  if exist(path_vbsScript, 'dir') == 7
    disp(['The vbsScript directory has been created：', path_vbsScript]);
  else
    error('Failed to create vbsScript directory.');
  end
gcc_buildpath.path_vbsScript = path_vbsScript;
%%%------------------------Create results directory-------------------------%%%
path_results = fullfile(path_case,'results'); % 当前工作目录
mkdir(path_results);
  if exist(path_results, 'dir') == 7
    disp(['The results directory has been created：', path_results]);
  else
    error('Failed to create the results directory.');
  end
gcc_buildpath.path_results = path_results;
%%%---------------------------Create logs directory-------------------------%%%
path_logs = fullfile(path_case,'logs'); % 当前工作目录
mkdir(path_logs);
  if exist(path_logs, 'dir') == 7
    disp(['The logs directory has been created：', path_logs]);
  else
    error('Failed to the logs directory.');
  end
gcc_buildpath.path_logs = path_logs;
fprintf('--------------------------------------------------------------------------------------------------------------------------------------------\n')
