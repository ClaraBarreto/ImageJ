path = getDirectory("Choose file");
dir = getFileList(path);
folder = File.directory;
folder_name = File.getName(folder);

for (filenr=0; filenr < dir.length; filenr++) {
	if (endsWith(dir[filenr],"czi")) {
		file_name = File.getName(dir[filenr]);
		file_name = replace(file_name, ".czi", "");
		file_name = replace(file_name, ".CZI", "");
		run("Bio-Formats Importer", "open=["+path+dir[filenr]+"] autoscale color_mode=Default open_all_series rois_import=[ROI manager] view=Hyperstack stack_order=XYCZT");
		run("Enhance Contrast", "saturated=0.35");
		run("Time Stamper", "starting=0 interval=1 x=2 y=10 font=100 '00 decimal=0 anti-aliased or=sec");
		saveAs("AVI", path+File.separator+file_name+".avi");
	};
};



//Developed by Clara Barreto =)