path = getDirectory("Choose file");
dir = getFileList(path);
folder = File.directory;
folder_name = File.getName(folder);

for (filenr=0; filenr < dir.length; filenr++) {
	if (endsWith(dir[filenr],"czi")) {
		run("Bio-Formats Importer", "open=["+path+dir[filenr]+"] autoscale color_mode=Colorized open_all_series rois_import=[ROI manager] view=Hyperstack stack_order=XYCZT");
		file_name = File.getName(dir[filenr]);
		images = getList("image.titles");
		image_name = getTitle();
		name = replace(image_name, ".CZI", "");
		name = replace(name, ".czi", "");
		name = replace(name, " ", "_");
		
		rename(name);
		run("Split Channels");
		selectWindow("C1-"+name);
		rename("DONOR");
		selectWindow("C2-"+name);
		rename("FRET");
		selectWindow("C3-"+name);
		close();
		selectWindow("C4-"+name);
		close();
		selectWindow("FRET");
		makeRectangle(50, 50, 200, 200);
		waitForUser("Move the square to a cell-free zone.");
		run("Measure");
		run("Select None");
		back_f = getResult("Mean", 0);
		run("Subtract...", "value="+back_f+" stack");
		run("Enhance Contrast", "saturated=0.35");
		run("Duplicate...", "duplicate");
		saveAs("Tiff", path+File.separator+name+"-FRET.tif");
		close();
		selectWindow("DONOR");
		makeRectangle(50, 50, 200, 200);
		waitForUser("Move the square to a cell-free zone.");
		run("Measure");
		run("Select None");
		back_d = getResult("Mean", 1);
		run("Subtract...", "value="+back_d+" stack");
		run("Clear Results");
		run("Enhance Contrast", "saturated=0.35");
		run("Duplicate...", "duplicate");
		saveAs("Tiff", path+File.separator+name+"-DONOR.tif");
		close();

		//FRET staining.
		selectWindow("FRET");
		run("Duplicate...", "duplicate");
		run("8-bit");
		run("Auto Local Threshold", "method=Phansalkar radius=15 parameter_1=0 parameter_2=0 white stack");
		run("Convert to Mask", "method=Yen background=Dark calculate black");
		rename("MASK");
		imageCalculator("Multiply create 32-bit stack", "DONOR","MASK");
		rename("MASKED_DONOR");
		imageCalculator("Multiply create 32-bit stack", "FRET","MASK");
		rename("MASKED_FRET");
		imageCalculator("Divide create 32-bit stack", "MASKED_FRET","MASKED_DONOR");
		rename("FRET_RESULT");
		selectWindow("FRET_RESULT");
		rename("FRET_RESULT"+name);
		run("16-bit");
		run("Rainbow RGB");
		setMinAndMax(0, 7295);
		setTool("polygon");
		waitForUser("Draw the region for quantification.");
		setBackgroundColor(0, 0, 0);
		run("Clear Outside", "stack");
		roiManager("Add");
		run("Set Measurements...", "mean redirect=None decimal=2");
		roiManager("Multi Measure");
		close();
		selectWindow("ROI Manager");
		run("Close");
		selectWindow("MASK");
		close();
		selectWindow("MASKED_DONOR");
		close();
		selectWindow("MASKED_FRET");
		close();
		
		if (filenr == 0) {
			selectWindow("Results");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] dataset_label="+name+" file_mode = read_and_open");
			run("Close");
		} else if (filenr == dir.length-1) {
				selectWindow("Results");
				run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] dataset_label="+name+" file_mode =  write_and_close");
				run("Close");
		} else {
				selectWindow("Results");
				run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] dataset_label="+name+" file_mode = queue_write");
				run("Close");
		}
		selectWindow("DONOR");
		run("Duplicate...", "duplicate");

			var numbimgs=nImages;
			if (numbimgs<3)exit("This macro requires three independent images open: donor, FRET and acceptor (or other image for intensity modulation)");
		
			setBatchMode(true);
		
			//	set variables
			var nume="numerator.tif";
			var denom="denominator.tif";
			var intmod="intensitymodulator.tif";
			var fretmin=1;
			var fretmax=2;
			var iaamin=0;
			var iaamax=65535;
		
		
			arrayimgs=newArray(numbimgs);
			for(i=1;i<=numbimgs;i++){selectImage(i);arrayimgs[i-1]=getTitle;}
			Dialog.create("IMD Fret");
			Dialog.addChoice("Numerator",arrayimgs,arrayimgs[0]);
			Dialog.addChoice("Denominator",arrayimgs,arrayimgs[1]);
			Dialog.addChoice("Intensity modulator",arrayimgs,arrayimgs[2]);
			Dialog.addNumber("ratio min",fretmin);
			Dialog.addNumber("ratio max",fretmax);
			Dialog.addNumber("intensity min",iaamin);
			Dialog.addNumber("intensity max",iaamax);
			Dialog.show;
			nume=Dialog.getChoice;
			denom=Dialog.getChoice;
			intmod=Dialog.getChoice;
			fretmin=Dialog.getNumber;
			fretmax=Dialog.getNumber;
			iaamin=Dialog.getNumber;
			iaamax=Dialog.getNumber;
		
			// add double slashes untill here to disable dialog and use automatically the values set above
		
			//	create ratio image
			selectWindow(nume);
			selectWindow(denom);
			w=getWidth;
			h=getHeight;
			f=nSlices;
			imageCalculator("Divide create 32-bit stack",nume,denom);
			selectWindow(nume);
			close();
			selectWindow(denom);
			close();
		
		
			//	insert intensity modulator image in brightness slice
			selectWindow(intmod);
			//resetMinAndMax();
			setMinAndMax(iaamin, iaamax);
			run("8-bit");
			setPasteMode("Copy");
			for (n=1; n<=nSlices; n++) {
				setSlice(n);
				run("Copy");
				newImage("hsv"+n, "RGB White", w, h, 1);
				run("HSB Stack");
				selectWindow("hsv"+n);
				setSlice(3);
				run("Paste");
				selectWindow(intmod);
			}
			close();
		
		
			//	set max level of saturation
		
			for (n=1; n<=f; n++) {
				selectWindow("hsv"+n);
				setSlice(2);
				run("Set...", "value=255 slice");
			}
		
		
			//	process and insert ratioimage as hue
		
			//	rearrange scale
			fretmin=fretmin-(fretmax-fretmin)/2;
		
			selectWindow("Result of "+nume);
			setMinAndMax(fretmin, fretmax);
			run("8-bit");
			run("Invert", "stack");
			for (n=1; n<=nSlices; n++) {
				setSlice(n);
		
		
				run("Copy");
				selectWindow("hsv"+n);
				setSlice(1);
				run("Paste");
				selectWindow("Result of "+nume);
			}
			close();
		
		
		
			//	generate RGB
			for (n=1; n<=f; n++) {
				selectWindow("hsv"+n);
				run("RGB Color");
			}
			if (f>1)
			run("Images to Stack", "name=IMDfret title=hsv use");
			else
			rename("IMDfret");
			setBatchMode(false);

		getDimensions(width, height, channels, slices, frames);
		for (s = 1; s <= slices; s++) {
			setSlice(s);
			setMinAndMax(1, 1);
		}
		saveAs("Tiff", path+File.separator+name+"_IMDfret.tif");
		run("Close All");
	}
}

			
//Developed by Clara Barreto =)