path = getDirectory("Choose file");
dir = getFileList(path);
folder = File.directory;
folder_name = File.getName(folder);

Dialog.create("Parameters");
Dialog.addMessage("Determine foci size limits (um^2) :");
Dialog.addNumber("lower :", 0.1);
Dialog.addNumber("higher :", 0.3);
Dialog.addMessage("Determine nuclei size limits (um^2) :");
Dialog.addNumber("lower :", 30.00);
Dialog.addNumber("higher :", 500.00);
Dialog.show();
L=Dialog.getNumber();
H=Dialog.getNumber();
LN=Dialog.getNumber();
HN=Dialog.getNumber();

Dialog.create("Parameters");
Dialog.addMessage("Do you wish to check your results at the end of every image?");
Dialog.addString("Y/N :", "Y");
Dialog.show();
CK=Dialog.getString();

lower_DAPI = 0;
upper_DAPI = 0;
lower_FOCI_GREEN = 0;
upper_FOCI_GREEN = 0;
lower_FOCI_RED = 0;
upper_FOCI_RED = 0;

for (filenr=0; filenr < dir.length; filenr++) {
	if ((endsWith(dir[filenr],"czi")) || (endsWith(dir[filenr],"C0.tiff"))) {
		file_name = File.getName(dir[filenr]);
		run("Bio-Formats Importer", "open=["+path+dir[filenr]+"] autoscale color_mode=Colorized open_all_series rois_import=[ROI manager] view=Hyperstack stack_order=XYCZT");
		images = getList("image.titles");
		image_name = getTitle();
		name = replace(image_name, "\\.czi", "");
		name = replace(name, "\\.tiff", "");
		name = replace(name, " ", "_");
		getDimensions(width, height, channels, slices, frames);
					
		if (frames > 1) {
			run("Z Project...", "projection=[Max Intensity]");
			selectWindow(image_name);
			close();
		};
	
		rename(name);
		run("8-bit");
		run("Split Channels");
		selectWindow("C1-"+name);
		rename("FOCI_GREEN");
		run("Green");
		selectWindow("C2-"+name);
		rename("FOCI_RED");
		run("Red");
		selectWindow("C3-"+name);
		rename("DAPI");
								
		if (lower_DAPI > 0 || upper_DAPI > 0 || lower_FOCI_GREEN > 0 || upper_FOCI_GREEN > 0 || lower_FOCI_RED > 0 || upper_FOCI_RED >0 ) {
			//DAPI staining.
			selectWindow("DAPI");
			run("Duplicate...", " ");
			rename("DAPI - "+name);
			run("Threshold...");
			setThreshold(lower_DAPI, upper_DAPI);
			setOption("BlackBackground", true);
			run("Convert to Mask");
			run("Fill Holes");
			run("Analyze Particles...", "size="+LN+"-"+HN+" circularity=0.0-1.00 show=Outlines exclude include summarize add in_situ");
			roiManager("Show All");
			if (CK == "Y") {
			Dialog.create("Quality Check");
			Dialog.addMessage("Proceed if all nuclei are selected. Otherwise repeat.");
			Dialog.show();
			};
			roiManager("deselect");
			number = roiManager("count");
			roiManager("save", path+"roi_"+name+".zip");
			NumberofRows = Table.size("Summary");
			Table.deleteRows(0, NumberofRows-1, "Summary");
				
			//FOCI_GREEN staining.
			selectWindow("DAPI - "+name);
			close();
			selectWindow("FOCI_GREEN");
			run("Duplicate...", " ");
			selectWindow("FOCI_GREEN");
			for (roi = 0; roi < number; roi++) {
				run("Duplicate...", " ");
				rename("nucleus_"+roi+1);
				roiManager("deselect");
				roiManager("select", roi);
				roiName = Roi.getName();
				rename(roiName+"_green_nucleus"+roi+1);
				run("Set Measurements...", "mean redirect=None decimal=2");
				run("Measure");
				close();
			};
			run("Threshold...");
			setThreshold(lower_FOCI_GREEN, upper_FOCI_GREEN);
			setOption("BlackBackground", true);
			run("Convert to Mask");
			selectWindow("Threshold");
			run("Close");
			selectWindow("FOCI_GREEN-1");
			close();
			imageCalculator("AND create", "FOCI_GREEN","DAPI");
			rename(name+" (pos)");
			run("Convert to Mask");
			for (roi = 0; roi < number; roi++) {
				run("Duplicate...", " ");
				rename("nucleus_"+roi+1);
				roiManager("deselect");
				roiManager("select", roi);
				roiName = Roi.getName();
				run("Watershed");
				rename(roiName+"_green_nucleus"+roi+1);
				run("Analyze Particles...", "size="+L+"-"+H+" circularity=0.50-1.00 show=Outlines include summarize in_situ"); 
				wait(500);
				close();
			};
			
			if (CK == "Y") {
				waitForUser("Confirm your results.");		
			};
			
			selectWindow("Results");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] sheet="+name+" dataset_label=foci_green_"+1+" file_mode = queue_write");
			run("Close");
			Table.rename("Summary", "Results");			
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] dataset_label=foci_green_"+2+" file_mode = queue_write");
			run("Close");

			//FOCI_RED staining.
			selectWindow(name+" (pos)");
			close();
			selectWindow("FOCI_GREEN");
			close();
			selectWindow("FOCI_RED");
			run("Duplicate...", " ");
			selectWindow("FOCI_RED");
			for (roi = 0; roi < number; roi++) {
				run("Duplicate...", " ");
				rename("nucleus_"+roi+1);
				roiManager("deselect");
				roiManager("select", roi);
				roiName = Roi.getName();
				rename(roiName+"_red_nucleus"+roi+1);
				run("Set Measurements...", "mean redirect=None decimal=2");
				run("Measure");
				close();
			};
			
			run("Threshold...");
			setThreshold(lower_FOCI_RED, upper_FOCI_RED);
			setOption("BlackBackground", true);
			run("Convert to Mask");
			selectWindow("Threshold");
			run("Close");
			selectWindow("FOCI_RED-1");
			close();
			imageCalculator("AND create", "FOCI_RED","DAPI");
			rename(name+" (pos)");
			run("Convert to Mask");				
			for (roi = 0; roi < number; roi++) {
				run("Duplicate...", " ");
				rename("nucleus_"+roi+1);
				roiManager("deselect");
				roiManager("select", roi);
				roiName = Roi.getName();
				run("Watershed");
				rename(roiName+"_red_nucleus"+roi+1);
				run("Analyze Particles...", "size="+L+"-"+H+" circularity=0.50-1.00 show=Outlines include summarize in_situ");
				wait(500);
				close();				
			};
			
			if (CK == "Y") {
				waitForUser("Confirm your results.");		
			};
			
			selectWindow("Results");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] sheet="+name+" dataset_label=foci_red_"+1+" file_mode = queue_write");
			run("Close");
			Table.rename("Summary", "Results");
			if (filenr+3 == dir.length) {
				run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] dataset_label=foci_red_"+2+" file_mode =  write_and_close");
				run("Close");
				selectWindow("ROI Manager");
				run("Close");
				run("Close All");
			};
			
			else {
				run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] dataset_label=foci_red_"+2+" file_mode =  queue_write");
				run("Close");
				selectWindow("ROI Manager");
				run("Close");
				run("Close All");
			};
		};	
		
		else{
			//DAPI staining.
			selectWindow("DAPI");
			run("Duplicate...", " ");
			rename("DAPI - "+name);
			run("Threshold...");
			waitForUser("Set the threshold values to select the DAPI staining.");
			getThreshold(lower_DAPI, upper_DAPI);
			setOption("BlackBackground", true);
			run("Convert to Mask");
			run("Fill Holes");
			run("Analyze Particles...", "size="+LN+"-"+HN+" circularity=0.0-1.00 show=Outlines exclude include summarize add in_situ");
			roiManager("Show All");
			Dialog.create("Quality Check");
			Dialog.addMessage("Proceed if all nuclei are selected. Otherwise repeat.");
			Dialog.show();
			roiManager("deselect");
			number = roiManager("count");
			roiManager("save", path+"roi_"+name+".zip");
			NumberofRows = Table.size("Summary");
			Table.deleteRows(0, NumberofRows-1, "Summary");
							
			//FOCI_GREEN staining.
			selectWindow("DAPI - "+name);
			close();
			selectWindow("FOCI_GREEN");
			run("Duplicate...", " ");
			selectWindow("FOCI_GREEN");
			for (roi = 0; roi < number; roi++) {
				run("Duplicate...", " ");
				rename("nucleus_"+roi+1);
				roiManager("deselect");
				roiManager("select", roi);
				roiName = Roi.getName();
				rename(roiName+"_green_nucleus"+roi+1);
				run("Set Measurements...", "mean redirect=None decimal=2");
				run("Measure");
				close();
			};
			run("Threshold...");
			waitForUser("Set the threshold values to select the FOCI staining.");
			getThreshold(lower_FOCI_GREEN, upper_FOCI_GREEN);
			setOption("BlackBackground", true);
			run("Convert to Mask");
			selectWindow("Threshold");
			run("Close");
			selectWindow("FOCI_GREEN-1");
			close();
			imageCalculator("AND create", "FOCI_GREEN","DAPI");
			rename(name+" (pos)");
			run("Convert to Mask");				
			for (roi = 0; roi < number; roi++) {
				run("Duplicate...", " ");
				rename("nucleus_"+roi+1);
				roiManager("deselect");
				roiManager("select", roi);
				roiName = Roi.getName();
				run("Watershed");
				rename(roiName+"_green_nucleus"+roi+1);
				run("Analyze Particles...", "size="+L+"-"+H+" circularity=0.50-1.00 show=Outlines include summarize in_situ");
				wait(500);
				close();				
			};
			waitForUser("Confirm your results.");
			selectWindow("Results");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] sheet="+name+" dataset_label=foci_green"+1+" file_mode = read_and_open");
			run("Close");
			Table.rename("Summary", "Results");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] dataset_label=foci_green_"+2+" file_mode = queue_write");
			run("Close");

			//FOCI_RED staining.
			selectWindow(name+" (pos)");
			close();
			selectWindow("FOCI_GREEN");
			close();
			selectWindow("FOCI_RED");
			run("Duplicate...", " ");
			selectWindow("FOCI_RED");
			for (roi = 0; roi < number; roi++) {
				run("Duplicate...", " ");
				rename("nucleus_"+roi+1);
				roiManager("deselect");
				roiManager("select", roi);
				roiName = Roi.getName();
				rename(roiName+"_green_nucleus"+roi+1);
				run("Set Measurements...", "mean redirect=None decimal=2");
				run("Measure");
				close();
			};
			run("Threshold...");
			waitForUser("Set the threshold values to select the FOCI staining.");
			getThreshold(lower_FOCI_RED, upper_FOCI_RED);
			setOption("BlackBackground", true);
			run("Convert to Mask");
			selectWindow("Threshold");
			run("Close");
			selectWindow("FOCI_RED-1");
			close();
			imageCalculator("AND create", "FOCI_RED","DAPI");
			rename(name+" (pos)");
			run("Convert to Mask");				
			for (roi = 0; roi < number; roi++) {
				run("Duplicate...", " ");
				rename("nucleus_"+roi+1);
				roiManager("deselect");
				roiManager("select", roi);
				roiName = Roi.getName();
				run("Watershed");
				rename(roiName+"_red_nucleus"+roi+1);
				run("Analyze Particles...", "size="+L+"-"+H+" circularity=0.50-1.00 show=Outlines include summarize in_situ");
				wait(500);
				close();				
			};
			waitForUser("Confirm your results.");
			selectWindow("Results");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] sheet="+name+" dataset_label=foci_red_"+1+" file_mode = queue_write");
			run("Close");
			Table.rename("Summary", "Results");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] dataset_label=foci_red_"+2+" file_mode = queue_write");
			run("Close");
			selectWindow("ROI Manager");
			run("Close");
			run("Close All");
		};
	};
};

//Developed by Clara Barreto =)