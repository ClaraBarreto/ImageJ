path = getDirectory("Choose file");
dir = getFileList(path);
folder = File.directory;
folder_name = File.getName(folder);

Dialog.create("Parameters");
Dialog.addMessage("Determine foci size limits (um^2) :");
Dialog.addNumber("lower :", 0.1);
Dialog.addNumber("higher :", 0.3);
Dialog.show();
L=Dialog.getNumber();
H=Dialog.getNumber();

Dialog.create("Parameters");
Dialog.addMessage("Do you wish to check your results at the end of every image?");
Dialog.addString("Y/N :", "Y");
Dialog.show();
CK=Dialog.getString();

lower_DAPI = 0;
upper_DAPI = 0;
lower_FOCI = 0;
upper_FOCI = 0;

for (filenr=0; filenr < dir.length; filenr++) {
	if ((endsWith(dir[filenr],"czi")) || (endsWith(dir[filenr],"tiff"))) {
		file_name = File.getName(dir[filenr]);
		run("Bio-Formats Importer", "open=["+path+dir[filenr]+"] autoscale color_mode=Colorized open_all_series rois_import=[ROI manager] view=Hyperstack stack_order=XYCZT");
		images = getList("image.titles");
		image_name = getTitle();
		name = replace(image_name, "\\.CZI", "");
		name = replace(name, "\\.czi", ""); 
		name = replace(name, "\\.tiff", "");
		name = replace(name, "\\.TIFF", "");
		name = replace(name, " ", "_");
		getDimensions(width, height, channels, slices, frames);

		if (slices > 1) {
			run("Z Project...", "projection=[Max Intensity]");
			selectWindow(image_name);
			close();
		};
			
		rename(name);
		run("Split Channels");
		waitForUser("Check which channels are for quantification.");
		Dialog.create("Parameters");
		Dialog.addMessage("Channels to quantify.");
		Dialog.addString("DAPI :", "C1");
		Dialog.addString("FOCI :", "C2");
		Dialog.show();
		CD=Dialog.getString();
		CF=Dialog.getString();
			
		if (channels > 2) {
			for (ch = 1; ch < channels; ch++) {
				selectWindow("C"+ch+"-"+name);
				if ((getTitle() != CD+"-"+name) && (getTitle() != CF+"-"+name)) {
					close();	
				};
			};
		};
						
		selectWindow(CD+"-"+name);
		rename("DAPI");
		selectWindow(CF+"-"+name);
		rename("FOCI");
			
		if (lower_DAPI > 0 || upper_DAPI > 0 || lower_FOCI > 0 || upper_FOCI > 0) {
			//DAPI staining.
			selectWindow("DAPI");
			run("Duplicate...", " ");
			rename("DAPI - "+name);
			run("Threshold...");
			setThreshold(lower_DAPI, upper_DAPI);
			setOption("BlackBackground", true);
			run("Convert to Mask");
			run("Fill Holes");
			run("Analyze Particles...", "size=80.00-900.00 circularity=0.20-1.00 show=Outlines exclude include summarize add in_situ");
			roiManager("Show All");
			Dialog.create("Quality Check");
			Dialog.addMessage("Proceed if all nuclei are selected. Otherwise repeat.");
			Dialog.show();
			roiManager("deselect");
			number = roiManager("count");
			roiManager("save", path+"roi_"+name+".zip");	
			NumberofRows = Table.size("Summary");
			Table.deleteRows(0, NumberofRows-1, "Summary");
			
			//FOCI staining.
			selectWindow("DAPI - "+name);
			close();
			selectWindow("FOCI");
			run("Duplicate...", " ");
			selectWindow("FOCI");
			for (roi = 0; roi < number; roi++) {
				run("Duplicate...", " ");
				rename("nucleus_"+roi+1);
				roiManager("deselect");
				roiManager("select", roi);
				roiName = Roi.getName();
				rename(roiName+"_nucleus"+roi+1);
				run("Set Measurements...", "mean redirect=None decimal=2");
				run("Measure");
				close();
			};
			run("Threshold...");
			setThreshold(lower_FOCI, upper_FOCI);
			setOption("BlackBackground", true);
			run("Convert to Mask");
			selectWindow("Threshold");
			run("Close");
			selectWindow("FOCI-1");
			close();
			imageCalculator("AND create", "FOCI","DAPI");
			rename(name+" (pos)");
			run("Convert to Mask");
			

			for (roi = 0; roi < number; roi++) {
				run("Duplicate...", " ");
				rename("nucleus_"+roi+1);
				roiManager("deselect");
				roiManager("select", roi);
				roiName = Roi.getName();
				run("Watershed");
				rename(roiName+"_nucleus"+roi+1);
				run("Analyze Particles...", "size="+L+"-"+H+" circularity=0.50-1.00 show=Outlines include summarize in_situ"); 
				wait(500);
				close();
			};
			if (CK == "Y") {
				waitForUser("Confirm your results.");		
			};
			selectWindow("Results");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] sheet="+name+" dataset_label=result_"+1+" file_mode = queue_write");
			run("Close");
			Table.rename("Summary", "Results");
			if (filenr+1 == dir.length) {
				run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] dataset_label=result_"+2+" file_mode = write_and_close");
				run("Close");
				selectWindow("ROI Manager");
				run("Close");
				run("Close All");
			};
			
			else{
				run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] dataset_label=result_"+2+" file_mode = queue_write");
				run("Close");
				selectWindow("ROI Manager");
				run("Close");
				run("Close All");
			};
		};
					
		else{
			// DAPI staining.
			selectWindow("DAPI");
			run("Duplicate...", " ");
			rename("DAPI - "+name);
			run("Threshold...");
			waitForUser("Set the threshold values to select the DAPI staining.");
			getThreshold(lower_DAPI, upper_DAPI);
			setOption("BlackBackground", true);
			run("Convert to Mask");
			run("Fill Holes");
			run("Analyze Particles...", "size=80.00-900.00 circularity=0.20-1.00 show=Outlines exclude include summarize add in_situ");
			roiManager("Show All");
			Dialog.create("Quality Check");
			Dialog.addMessage("Proceed if all nuclei are selected. Otherwise repeat.");
			Dialog.show();
			roiManager("deselect");
			number = roiManager("count");
			roiManager("save", path+"roi_"+name+".zip");
			NumberofRows = Table.size("Summary");
			Table.deleteRows(0, NumberofRows-1, "Summary");
							
			// FOCI staining.
			selectWindow("DAPI - "+name);
			close();
			selectWindow("FOCI");
			run("Duplicate...", " ");
			selectWindow("FOCI");
			for (roi = 0; roi < number; roi++) {
				run("Duplicate...", " ");
				rename("nucleus_"+roi+1);
				roiManager("deselect");
				roiManager("select", roi);
				roiName = Roi.getName();
				rename(roiName+"_nucleus"+roi+1);
				run("Set Measurements...", "mean redirect=None decimal=2");
				run("Measure");
				close();
			};
			run("Threshold...");
			waitForUser("Set the threshold values to select the FOCI staining.");
			getThreshold(lower_FOCI, upper_FOCI);
			setOption("BlackBackground", true);
			run("Convert to Mask");
			selectWindow("Threshold");
			run("Close");
			selectWindow("FOCI-1");
			close();
			imageCalculator("AND create", "FOCI","DAPI");
			rename(name+" (pos)");
			run("Convert to Mask");
				
			for (roi = 0; roi < number; roi++) {
				run("Duplicate...", " ");
				rename("nucleus_"+roi+1);
				roiManager("deselect");
				roiManager("select", roi);
				roiName = Roi.getName();
				run("Watershed");
				rename(roiName+"_nucleus"+roi+1);
				run("Analyze Particles...", "size="+L+"-"+H+" circularity=0.50-1.00 show=Outlines include summarize in_situ");
				wait(500);
				close();				
			};
			waitForUser("Confirm your results.");
			selectWindow("Results");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] sheet="+name+" dataset_label=result_"+1+" file_mode = read_and_open");
			run("Close");
			Table.rename("Summary", "Results");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Results.xlsx] dataset_label=result_"+2+" file_mode = queue_write");
			run("Close");
			selectWindow("ROI Manager");
			run("Close");
			run("Close All");
		};
	};
};

//Developed by Clara Barreto =)