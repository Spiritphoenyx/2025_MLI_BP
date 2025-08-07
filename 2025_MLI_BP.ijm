D=4.34
run("Set Scale...", "distance=D known=1 unit=micron");
run("16-bit");
original = getTitle();
Width = getWidth();
Height = getHeight();
run("Duplicate...", "title=reference");
run("Duplicate...", "title=edges1")
run("Duplicate...", "title=edges2")
run("Duplicate...", "title=edges3")
S=10/D

makeRectangle(0, Height/4, Width, 1);
run("Duplicate...", "title=[topthird line]");
run("Select All");
run("Copy");
selectImage("edges1");
run("Select All");
run("Clear", "slice");
run("Paste");
rename(original + "_topthird line")
setAutoThreshold("Default no-reset");
//run("Threshold...");
run("Convert to Mask");
run("Invert");
run("Set Measurements...", "area decimal=3");
run("Analyze Particles...", "size=S-Infinity show=[Overlay Masks] display exclude");
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
run("Duplicate...", "title=[midimage line]");
run("Select All");
run("Copy");
selectImage("edges2");
run("Select All");
run("Clear", "slice");
run("Paste");
rename(original + "_midimage line")
setAutoThreshold("Default no-reset");
//run("Threshold...");
run("Convert to Mask");
run("Invert");
run("Set Measurements...", "area decimal=3");
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
run("Duplicate...", "title=[bottomthird line]");
run("Select All");
run("Copy");
selectImage("edges3");
run("Select All");
run("Clear", "slice");
run("Paste");
rename(original + "_bottomthird line")
setAutoThreshold("Default no-reset");
//run("Threshold...");
run("Convert to Mask");
run("Invert");
run("Set Measurements...", "area decimal=3");
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