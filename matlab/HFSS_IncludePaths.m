% ----------------------------------------------------------------------------%
function HFSS_IncludePaths(relPath)

if (nargin < 1)
	relPath = '';
end

addpath([relPath, '/boundary/']);
addpath([relPath, '/3dmodeler/']);
addpath([relPath, '/analysis/']);
addpath([relPath, '/general/']);
addpath([relPath, '/radiation/']);
addpath([relPath, '/reporter/']);
addpath([relPath, '/fieldsCalculator/']);
addpath([relPath, '/mesh/']);
addpath([relPath, '/materiallibrary/']);
addpath([relPath, '/stackupbuild/']);
addpath([relPath, '/display/']);
addpath([relPath, '/matlab_python/']);
addpath([relPath, '/customize/']);
addpath([relPath, '/package/']);
addpath([relPath, '/pcb_isi/']);

end