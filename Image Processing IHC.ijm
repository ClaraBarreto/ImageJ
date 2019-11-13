path = getDirectory("Choose file");
dir=getFileList(path);
folder = File.getName(path);
	for (i=0; i < dir.length; i++) {
		if (endsWith(dir[i],"czi")) {
		open(path+dir[i]);
		name = getTitle();
		series_name = replace(name, ".czi", "");
		rename("IMAGE");
		run("Stack to RGB");
		selectWindow("IMAGE");
		close();
		selectWindow("IMAGE (RGB)");
		run("Set Measurements...", "area mean standard area_fraction display redirect=None decimal=3");
		run("Enhance Contrast", "saturated=0.35");
		run("Duplicate...", " ");
		run("Duplicate...", " ");
		run("Scale Bar...", "width=40 height=10 font=42 color=Black background=None location=[Lower Right] bold hide");
		saveAs("tiff", path+File.separator+series_name+"_original.tiff");
		close();
		makeRectangle(1587, 357, 600, 600);
		waitForUser;
		run("Crop");
		run("Scale Bar...", "width=40 height=10 font=42 color=Black background=None location=[Lower Right] bold hide");
		saveAs("tiff", path+File.separator+series_name+"_Crop.tiff");
		run("Close All");
		};
	};