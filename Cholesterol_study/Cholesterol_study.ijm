//Shading correction
waitForUser("Open images for shading determination");
run("Images to Stack", "name=Stack title=[]");
run("Z Project...", "projection=[Average Intensity]");
rename("SHADING");
selectWindow("Stack");
close()
run("Smooth");
run("Smooth");
run("Set Measurements...", "mean display redirect=None decimal=3");
run("Measure");
k1 = getResult("Mean");
selectWindow("Results");
run("Close");

waitForUser("Open image to quantify");
file_name = getTitle();
name = replace(file_name, ".JPG", "");
rename("IMAGE");

run("Calculator Plus", "i1=SHADING i2=IMAGE operation=[Divide: i2 = (i1/i2) x k1 + k2] k1="+k1+" k2=0 create");
rename("SHADING_IMAGE");
selectWindow("SHADING");
close();

//Background subtraction
selectWindow("IMAGE");
run("8-bit");
nBins = 256;
getHistogram(values, counts, 256);
cumHist = newArray(nBins);
cumHist[0] = values[counts[0]];
for (i = 1; i < nBins; i++){ cumHist[i] = counts[i] + cumHist[i-1]; }
normCumHist = newArray(nBins);
for (i = 0; i < nBins; i++){  normCumHist[i] = cumHist[i]/cumHist[nBins-1]; }
target = 0.05;
i = 0;
do {
	i = i + 1;
	percentile_5th = values[i];
	} 
while (normCumHist[i] < target);
run("Calculator Plus", "i1=IMAGE i2=SHADING_IMAGE operation=[Subtract: i2 = (i1-i2) x k1 + k2] k1="+percentile_5th+" k2=0 create");
rename("BACKREMOVED_SHADE_IMAGE");
selectWindow("SHADING_IMAGE");
close();
selectWindow("IMAGE");
close();

//Threshold
run("Duplicate...", " ");
run("8-bit");
run("Threshold...");
waitForUser("Set threshold value to outline cells, click 'Apply' and then press the 'OK' button in this window");
run("Set Measurements...", "area mean display redirect=None decimal=3");
run("Measure");
low_thresh_mean = getResult("Mean");
low_thresh_area = getResult("Area");
selectWindow("Results");
run("Close");
selectWindow("BACKREMOVED_SHADE_IMAGE-1");
close();
selectWindow("Threshold");
run("Close");
selectWindow("BACKREMOVED_SHADE_IMAGE");
run("Duplicate...", " ");
run("8-bit");
run("Threshold...");
waitForUser("Set threshold value to select specific staining, click 'Apply' and then press the 'OK' button in this window");
run("Set Measurements...", "integrated display redirect=None decimal=3");
run("Measure");
high_thresh_int = getResult("IntDen");
run("Clear Results");
selectWindow("BACKREMOVED_SHADE_IMAGE-1");
close();
selectWindow("Threshold");
run("Close");

//Final Results
Average_filipin_intensity = low_thresh_mean;
LSO_compartment_ratio = high_thresh_int / low_thresh_area;

setResult("Image_name", 0, name);
setResult("Average_filipin", 0, Average_filipin_intensity);
setResult("LSO_compartment", 0, LSO_compartment_ratio);