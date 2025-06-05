D=4.34
run("Set Scale...", "distance=D known=1 unit=micron");
run("16-bit");
original = getTitle();
Width = getWidth();
Height = getHeight();
run("Duplicate...", "title=reference");
S=10/D

makeRectangle(0, Height/4, Width, 1);
run("Duplicate...", "title=topthird intensity");
run("Duplicate...", "title=[topthird line]");
rename(original + "_topthird line")
setAutoThreshold("Default no-reset");
//run("Threshold...");
run("Convert to Mask");
run("Invert");
run("Set Measurements...", "area redirect=topthird intensity decimal=3");
run("Analyze Particles...", "size=S-Infinity show=[Overlay Masks] display");
IJ.renameResults("Results"); // otherwise below does not work...
for (row=0; row<nResults; row++) {
	sum = getResult("Area", row) * D;
    setResult("Airspace", row, sum);
}
updateResults();
run("Read and Write Excel");
Table.deleteRows(0, 1000);

selectImage("reference");
makeRectangle(0, Height/2, Width, 1);
run("Duplicate...", "title=midimage intensity");
run("Duplicate...", "title=[midimage line]");
rename(original + "_midimage line")
setAutoThreshold("Default no-reset");
//run("Threshold...");
run("Convert to Mask");
run("Invert");
run("Set Measurements...", "area redirect=midimage intensity decimal=3");
run("Analyze Particles...", "size=S-Infinity show=[Overlay Masks] display");
IJ.renameResults("Results"); // otherwise below does not work...
for (row=0; row<nResults; row++) {
	sum = getResult("Area", row) * D;
    setResult("Airspace", row, sum);
}
updateResults();
run("Read and Write Excel");
Table.deleteRows(0, 1000);

selectImage("reference");
makeRectangle(0, Height*3/4, Width, 1);
run("Duplicate...", "title=bottomthird intensity");
run("Duplicate...", "title=[bottomthird line]");
rename(original + "_bottomthird line")
setAutoThreshold("Default no-reset");
//run("Threshold...");
run("Convert to Mask");
run("Invert");
run("Set Measurements...", "area redirect=bottomthird intensity decimal=3");
run("Analyze Particles...", "size=S-Infinity show=[Overlay Masks] display");
IJ.renameResults("Results"); // otherwise below does not work...
for (row=0; row<nResults; row++) {
	sum = getResult("Area", row) * D;
    setResult("Airspace", row, sum);
}
updateResults();
run("Read and Write Excel");
Table.deleteRows(0, 1000);


run("Images to Stack", "use");
rename(original + "_quantified");