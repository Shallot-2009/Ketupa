% ----------------------------------------------------------------------------%
function HFSS_RemovePaths(relPath)%

if (nargin < 1)
	relPath = '';
end

rmpath([relPath, '/boundary/']);
rmpath([relPath, '/3dmodeler/']);
rmpath([relPath, '/analysis/']);
rmpath([relPath, '/general/']);
rmpath([relPath, '/radiation/']);
rmpath([relPath, '/reporter/']);
rmpath([relPath, '/fieldsCalculator/']);
rmpath([relPath, '/mesh/']);
rmpath([relPath, '/materiallibrary/']);
rmpath([relPath, '/stackupbuild/']);
rmpath([relPath, '/display/']);
rmpath([relPath, '/matlab_python/']);
rmpath([relPath, '/customize/']);
rmpath([relPath, '/package/']);
rmpath([relPath, '/pcb_isi/']);

end