path = getDirectory("Choose file");
dir = getFileList(path);
folder = File.directory;
folder_name = File.getName(folder);

for (filenr=0; filenr < dir.length; filenr++) {
	if (endsWith(dir[filenr],".tif")) {
		file_name = File.getName(dir[filenr]);
		open(path+file_name);

		Dialog.create("Parameters");
		Dialog.addMessage("Determine staining method :");
		Dialog.addString("AFOG or TMkit :", "Y");
		Dialog.addString("TM :", "N");
		Dialog.addNumber("LT (threshold):", 250);
		Dialog.show();

		AFOG_TMkit = Dialog.getString();
		TM = Dialog.getString();
		LT = Dialog.getNumber();
		
		name = replace(file_name, "\\.tif", "");
		name = replace(name, " ", "_");
		rename("IMAGE");
		run("Colour Deconvolution", "vectors=[Masson Trichrome]");
		if ((TM == "Y") && (AFOG_TMkit == "N")) {
			selectWindow("Colour Deconvolution");
			close();
			selectWindow("IMAGE-(Colour_1)"); // blue
			close();
			selectWindow("IMAGE-(Colour_2)"); // red
			close();
			selectWindow("IMAGE-(Colour_3)"); // green
			rename(name);
		};
		if (((TM == "N") && (AFOG_TMkit == "Y"))) {
			selectWindow("Colour Deconvolution");
			close();
			selectWindow("IMAGE-(Colour_3)"); // green
			close();
			selectWindow("IMAGE-(Colour_2)"); // red
			close();
			selectWindow("IMAGE-(Colour_1)"); // blue
			rename(name); 
		};
		selectWindow(name);
		run("Local Thickness (complete process)", "threshold="+LT+" inverse");
		run("RGB Color");
		saveAs("Tiff", path+File.separator+name+" - LocalThickness.tif");
		close();
		
		selectWindow(name);		
		print(name);
		run("Directionality", "method=[Fourier components] nbins=90 histogram_start=-90 histogram_end=90 build display_color_wheel display_table");
		selectWindow("Orientation map for "+name);
		saveAs("Tiff", path+File.separator+name+" - Orientation map.tif");
		close();
		Table.rename("Directionality histograms for "+name+" (using Fourier components)", "Results");
		if (filenr == 0) {
			selectWindow("Color wheel");
			saveAs("Tiff", path+File.separator+" - Color wheel.tif");
			close();
			selectWindow("Results");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Directionality.xlsx] sheet="+name+" dataset_label=measurement file_mode = read_and_open");		
			run("Close");
			selectWindow(name);
			run("Get Main Direction MEGA");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Directionality.xlsx] sheet="+name+" dataset_label=stats file_mode = queue_write");
			run("Close");
		} else if (filenr == dir.length-1) {
			selectWindow("Color wheel");
			close();
			selectWindow("Results");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Directionality.xlsx] sheet="+name+" dataset_label=measurement file_mode =  queue_write");
			run("Close");
			selectWindow(name);
			run("Get Main Direction MEGA");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Directionality.xlsx] sheet="+name+" dataset_label=stats file_mode = write_and_close");
			run("Close");
		} else {
			selectWindow("Color wheel");
			close();
			selectWindow("Results");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Directionality.xlsx] sheet="+name+" dataset_label=measurement file_mode = queue_write");
			run("Close");
			selectWindow(name);
			run("Get Main Direction MEGA");
			run("Read and Write Excel", "no_count_column file=["+path+"/"+folder_name+"_Directionality.xlsx] sheet="+name+" dataset_label=stats file_mode = queue_write");
			run("Close");
		}
		
		waitForUser ("Save histogram if desired and close it. Close also the Results Table. Click 'OK'.");
		run("Close All");
	};
};


//Developed by Clara Barreto =)