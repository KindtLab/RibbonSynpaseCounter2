

// ********************************************************
// ** Neuromast_CompleteSynaspseCounter2D 2.0, Kindt Lab **
// ********************************************************


// Tested on ImageJ 1.53t, Fiji
// plugins required: AdaptiveThreshold (https://sites.imagej.net/adaptiveThreshold/)
// plugins required: Stack Normalizer 
// adopted from Candy Wong 2013 Jan.
// 2023 Mar. 24 by Zhengchang Lei

//--- update log 20230328
// save data of all samples in three master list: NMstats, rbbstats and magstats  
// enlarge rbb or mag rois for paired pre-sysnapse or post-synapse counting
// use HC mask to exclude signals outside of NM, 
//    if there's no mask file for a image file, the whole field of view is included. 

//--- update log 20230329
// contrast adjust before thresholding is important for accurate counting.
//*****************************************************************************


//## Parameter Settings #######################################################
  start_position=0;
  remove_from_end=0;
  suffix_for_rbb = "_rbb";
  suffix_for_mag = "_mag";
  suffix_for_HCmsk = "_hc_msk"; //if there's a mask, name it as ImageName + suffix_for_HCmsk
// channel asignment
ch_rbb = "C2";
ch_mag = "C3";
ch_hc = "C1";

//2D size threshold for "Analyze Particle"
size_mag_min=0.025;
size_rbb_min=0.025;

//adaptive thresholding parameters
blockSize_rbb=50;  //40
backGround_rbb = -40;
blockSize_mag=110;  
backGround_mag = -40; //-50
engPix = 1; //enlarge rois for overlapping calculation
rbbContrast = 0.20;
magContrast = 0.01;
zCorrectFactor = 0;
targetMean = 5000;

//That's all!
//############################################################################

// select input folder
inDir = getDirectory("--> INPUT: Choose Directory <--");
outDir = getDirectory("--> OUTPUT: Choose Directory for TIFF Output <--");

outXlsx = outDir + "\\xlsx\\";
if (!File.exists(outXlsx)) {
    File.makeDirectory(outXlsx);
}
outRoi = outDir + "\\Rois\\";
if (!File.exists(outRoi)) {
    File.makeDirectory(outRoi);
}
outJPG = outDir + "\\Jpgs\\";
if (!File.exists(outJPG)) {
    File.makeDirectory(outJPG);
}


dataSet = File.getName(inDir);
paraFile = outDir + dataSet + "_counting parameters.txt";
masterFileNM = outDir + dataSet + "_MasterFile_NMstats.csv";
masterFileRbb = outDir + dataSet + "_MasterFile_Rbb.csv";
masterFileMag = outDir + dataSet + "_MasterFile_Mag.csv";

      directory_tiff = outDir;
      directory_jpg = outDir;
      directory_spreadsheet = outDir;
      directory_roi = outDir;

inList = getFileList(inDir);
list = getFromFileList("czi", inList);  // select dirs only
fn = list.length;



// Checkpoint: get list of dirs
print("Below is a list of files to be processeded:");
printArray(list); // Implemented below
print("Result save to:");
print(outDir);
// save parameter to txt file
getDateAndTime(year, month, dayOfWeek, dayOfMonth, hour, minute, second, msec);
f = File.open(paraFile);
print(f, "Parameters used as follow: ( " + year +"-"+ month+"-" + dayOfMonth +"_"+ hour+"-"+ minute+"-"+ second + ")");
print(f, "size_mag_min = " + size_mag_min);
print(f, "size_rbb_min = " + size_rbb_min);
print(f, "blockSize_rbb = " + blockSize_rbb);
print(f, "backGround_rbb = " + backGround_rbb);
print(f, "blockSize_mag = " + blockSize_mag);
print(f, "backGround_mag = " + backGround_mag);
print(f, "Contrast_rbb = " + rbbContrast);
print(f, "Contrast_mag = " + magContrast);
print(f, "engPix = " + engPix);
print(f, "zCorrectFactor = " + zCorrectFactor);
File.close(f);
//###################################################################################
// Main Processing starts here

roiManager("show none"); // to avoid a weird error of the ROImanager reset function
setBatchMode(true);

SampleNames = newArray(fn);
NMstats = newArray();
curNM = newArray();
NMheaders = newArray("SampleID","Complete Syn","nRbb","upRbb","nMag","upMag");
Table.create("NMstats");
Table.update("NMstats");
 for (it = 0;it<6; it++){
 	Table.setColumn(NMheaders[it]);
 }
 Table.update(); 

Table.create("rbbstats");
Table.update("rbbstats");
Table.create("magstats");
Table.update("magstats");

for (i=0; i<fn; i++)
{
  spName = 	substring(list[i],0, lengthOf(list[i])-4);
  SampleNames[i] = spName;
  
  inFullname = inDir + list[i];
  inFullnameHCmsk = inDir + spName + suffix_for_HCmsk + ".tif";
  
    rbb_2Droi = spName + suffix_for_rbb + "2Droi.zip";
    rbb_count = spName + suffix_for_rbb +"Count.csv";
    rbb_area = spName + suffix_for_rbb +"Area.csv";

    mag_2Droi =  spName + suffix_for_mag + "2Droi.zip";
    mag_count = spName + suffix_for_mag +"Count.csv";
    mag_area = spName + suffix_for_mag + "Area.csv";

    Seg2D_tif =   spName  + "_Seg2D.tif";
    Seg2D_jpeg = spName  + "_Seg2D.jpeg";
    Seg2D_rbb =  spName  + "_SegRbb.jpeg";
    Seg2D_mag =  spName  + "_SegMag.jpeg";
    Seg2D_hc =  spName  + "_hc.jpeg";
    msk_jpeg =  spName  + "_msk.jpeg";

  outFullname_ZP = outDir + substring(list[i],0, lengthOf(list[i])-4) + ".png";
  // Checkpoint: Indicating progress
  print("Saving(",(i+1),"/",list.length,")...",list[i]); 


// Processing 
  open(inFullname);
  rename("Current");
  run("Split Channels");
  
// get HC mask ready 
 selectWindow(ch_hc + "-Current");
 run("Subtract Background...", "rolling=50 stack");
 rename("hc");
 run("Stack Normalizer", "minimum=0 maximum=65535");
 run("Z Project...", "projection=[Max Intensity]");
   if (File.exists(inFullnameHCmsk)){
  	open(inFullnameHCmsk);
  	rename("HCmsk");
   }else{
         selectWindow("MAX_hc");
         run("Duplicate...", "title=HCmsk");
         setMinAndMax(0, 1);
         run("8-bit");
         run("Apply LUT");
    }
  
// rbb counting //////////////////////////////////
 selectWindow(ch_rbb + "-Current");
 run("Subtract Background...", "rolling=50 stack");
 rename("rbb");
 run("Stack Normalizer", "minimum=0 maximum=65535");
 run("Z Project...", "projection=[Max Intensity]");
 run("Duplicate...", "title=MAX_rbb_m ");

 selectWindow("MAX_rbb_m");
 resetMinAndMax();
 run("Enhance Contrast", "saturated=rbbContrast");
 run("8-bit");
 run("adaptiveThr  Plugin", "using=[Weighted mean] from=blockSize_rbb then=backGround_rbb");
 setOption("BlackBackground", true);
 run("Convert to Mask");
 run("Watershed");
 imageCalculator("AND create", "MAX_rbb_m","HCmsk");
 selectWindow("Result of MAX_rbb_m");
 run("Analyze Particles...", "size=&size_rbb_min-Infinity show=Masks clear add");
 selectWindow("Mask of Result of MAX_rbb_m");
 run("Invert LUTs");
 rename("rbb_mask");

// maguk counting ///////////////////////////////
 selectWindow(ch_mag + "-Current");
 run("Subtract Background...", "rolling=50 stack");
 rename("mag");
 
 //z axis correction
 if (zCorrectFactor>0){
   for (ic = 1; ic <= nSlices; ic++) {
      setSlice(ic);
      factor = (1- zCorrectFactor*(nSlices-ic)/(nSlices-1);
      run("Multiply...", "value=factor slice");
      }
  }
 run("Stack Normalizer", "minimum=0 maximum=65535");
 run("Z Project...", "projection=[Max Intensity]");
 
 // mean normalization
 if(targetMean>0){
  getStatistics(area, mean, min, max);
  ("mean of the image is " + mean);
  factor = targetMean / mean;
  run("Multiply...", "value=" + factor);
 }
 
 run("Duplicate...", "title=MAX_mag_m ");
 selectWindow("MAX_mag_m");
 resetMinAndMax();
 run("Enhance Contrast", "saturated=magContrast");
 run("8-bit");
 run("adaptiveThr  Plugin", "using=[Weighted mean] from=blockSize_mag then=backGround_mag");
 setOption("BlackBackground", true);
 run("Convert to Mask");
 run("Analyze Particles...", "size=0.00-Infinity circularity=0.3-1.00 show=Masks clear add");
 selectWindow("Mask of MAX_mag_m");
 run("Invert LUTs"); 
// run("Watershed");
 imageCalculator("AND create", "Mask of MAX_mag_m","HCmsk");
 selectWindow("Result of Mask of MAX_mag_m");
 run("Analyze Particles...", "size=&size_mag_min-Infinity show=Masks exclude clear add");
 selectWindow("Mask of Result of Mask of MAX_mag_m");
 run("Invert LUTs");
 rename("mag_mask");

// data saving /////////////////////////////////
// build composite
 run("Merge Channels...", "c1=MAX_mag c2=MAX_rbb create keep");
 rename("Composite rbb-mag");
 run("8-bit");
 Stack.setChannel(2);
 run("Enhance Contrast", "saturated=0.35");
 run("Magenta");
 Stack.setChannel(1);
 run("Enhance Contrast", "saturated=0.35");
 run("Green");
// Save rbb
 roiManager("reset");
 selectWindow("rbb_mask");
 run("Analyze Particles...", "size=&size_rbb_min-Infinity clear add");
 selectWindow("MAX_rbb");
 roiManager("Show None");
 roiManager("Show All");
 roiManager("OR");
 roiManager("Measure");

 selectWindow("Results");
 curRbbSize = newArray(nResults,1);
 curRbbRawIntDen = newArray(nResults,1);
 curRbbn = nResults;
 for (ir=0; ir<nResults; ir++) {
    curRbbSize[ir] = getResult("Area",ir);
 }
 run("Clear Results");
 
// engPix = 1;
 // roi enlarge
 if (engPix > 0){
    roiManager("select","ROI Manager");
    for (j = 0; j < roiManager("count"); j++) {
       roiManager("select", j);
       run("Enlarge...", "enlarge=engPix pixel");
       roiManager("update");
    }    
  }
 roiManager("Deselect");
 roiManager("Save", outRoi+rbb_2Droi);
 
 selectWindow("mag_mask");
 roiManager("Show None");
 roiManager("Show All");
 roiManager("OR");
 roiManager("Measure");
 selectWindow("Results");
 curComSyn = 0;
 for (ir=0; ir<nResults; ir++) {
    setResult("Area",ir,curRbbSize[ir]);
    curRbbRawIntDen[ir] = getResult("RawIntDen",ir)/255;
    if (curRbbRawIntDen[ir]>0) {
    	curComSyn++;
    }
 }
 NMstats[0] = curComSyn;
 NMstats[1] = curRbbn;
 NMstats[2] = curRbbn-curComSyn;
 run("Input/Output...", "jpeg=100 gif=-1 file=.csv copy_column copy_row save_column save_row");
 saveAs("Results", outXlsx + rbb_count );
  
 selectWindow("Composite rbb-mag");
 run("Duplicate...", "title=rbb-save duplicate");
 roiManager("Show None");
 roiManager("Show All");
 roiManager("OR");
 setFont("Calibri", 22, "bold, antialised, white");
 drawString("rbb_"+spName, 10, 24);
 run("Flatten");
 run("Input/Output...", "jpeg=100");
 saveAs("Jpeg", outJPG+Seg2D_rbb);
 roiManager("Reset");
 run("Clear Results");
 // update rbb-stats table
 selectWindow("rbbstats");
 Table.setColumn(spName + "-Area",curRbbSize);
 Table.setColumn(spName + "-RawIntDen",curRbbRawIntDen);
 Table.update(); 
// Save mag
 selectWindow("mag_mask");
 run("Analyze Particles...", "size=&size_mag_min-Infinity clear add");
  
 selectWindow("MAX_mag");
 roiManager("Show None");
 roiManager("Show All");
 roiManager("OR");
 roiManager("Measure");
 
  selectWindow("Results");
 curMagSize = newArray(nResults,1);
 curMagRawIntDen = newArray(nResults,1);
 curMagn = nResults;
 for (ir=0; ir<nResults; ir++) {
    curMagSize[ir] = getResult("Area",ir);;
 }
 run("Clear Results"); 
 
  if (engPix > 0){
    for (jj = 0; jj < roiManager("count"); jj++) {
       roiManager("select", jj);
       run("Enlarge...", "enlarge=engPix pixel");
       roiManager("update");
    }
   }
 roiManager("Deselect");
 roiManager("Save", outRoi + mag_2Droi);
 selectWindow("rbb_mask");
 roiManager("Show None");
 roiManager("Show All");
 roiManager("OR");
 roiManager("Measure");
 selectWindow("Results");
 curUpMag = 0;
 for (ir=0; ir<nResults; ir++) {
    setResult("Area",ir,curMagSize[ir]);
    curMagRawIntDen[ir] = getResult("RawIntDen",ir)/255;
    if (curMagRawIntDen[ir]==0){
    	curUpMag++;
    }
 }
 NMstats[3] = curMagn;
 NMstats[4] = curUpMag;

 run("Input/Output...", "jpeg=100 gif=-1 file=.csv copy_column copy_row save_column save_row");
 saveAs("Results", outXlsx + mag_count);
 
 selectWindow("Composite rbb-mag");
 run("Duplicate...", "title=mag-save duplicate");
 roiManager("Show None");
 roiManager("Show All");
 roiManager("OR");
  setFont("Calibri", 22, "bold, antialised, white");
 drawString("mag_"+spName, 10, 24);
 run("Flatten");
 run("Input/Output...", "jpeg=100");
 saveAs("Jpeg", outJPG + Seg2D_mag);
 roiManager("Reset");
 run("Clear Results");
 // update mag-stats table
 selectWindow("magstats");
 Table.setColumn(spName + "-Area",curMagSize);
 Table.setColumn(spName + "-RawIntDen",curMagRawIntDen);
 Table.update(); 
  // update NM-stats table
 selectWindow("NMstats");
 Table.set("SampleID", i, spName);
 Table.set("Complete Syn",i, NMstats[0]);
 Table.set("nRbb",i, NMstats[1]);
 Table.set("upRbb",i, NMstats[2]);
 Table.set("nMag",i, NMstats[3]);
 Table.set("upMag",i, NMstats[4]);
 Table.update(); 
 
// save the compostite jpg
 selectWindow("Composite rbb-mag"); 
 run("Duplicate...", "title=CompositeJPG duplicate");
 setFont("Calibri", 22, "bold, antialised, white");
 drawString(spName, 10, 24);
 run("Flatten");
 run("Input/Output...", "jpeg=100");
 saveAs("Jpeg",  outJPG + Seg2D_jpeg);
 
 // save composite tif
 selectWindow("Composite rbb-mag");
 save(outRoi + Seg2D_tif);
 
 // save masks
 run("Merge Channels...", "c1=mag_mask c2=rbb_mask create keep");
 rename("CompositeMSK");
 Stack.setChannel(2);
 run("Magenta");
 Stack.setChannel(1);
 run("Green");
 setFont("Calibri", 22, "bold, antialised, white");
 drawString(spName, 10, 24);
 run("Flatten");
 run("Input/Output...", "jpeg=100");
 saveAs("Jpeg",  outJPG + msk_jpeg);
 
 // save hc
 selectWindow("MAX_hc");
 resetMinAndMax();
 run("Enhance Contrast", "saturated=0.35");
 run("8-bit");
 run("Flatten");
 run("Input/Output...", "jpeg=100");
 saveAs("Jpeg", outJPG +Seg2D_hc);
 
 run("Close All");
 if (isOpen("Results")) {
         selectWindow("Results"); 
         run("Close" );
  }
}

// Save stats tables
selectWindow("NMstats");

run("Input/Output...", "jpeg=100 gif=-1 file=.csv copy_column copy_row save_column save_row");
saveAs("Results", masterFileNM );

selectWindow("magstats");
run("Input/Output...", "jpeg=100 gif=-1 file=.csv copy_column copy_row save_column save_row");
 saveAs("Results",  masterFileMag );
 
selectWindow("rbbstats");
run("Input/Output...", "jpeg=100 gif=-1 file=.csv copy_column copy_row save_column save_row");
 saveAs("Results",  masterFileRbb);


setBatchMode("exit and display");
print("--- All Done ---");

// --- Main procedure end ---
//###############################################################################

function getFromFileList(ext, fileList)
{
  selectedFileList = newArray(fileList.length);
  selectedDirList = newArray(fileList.length);
  ext = toLowerCase(ext);
  j = 0;
  iDir = 0;
  for (i=0; i<fileList.length; i++)
    {
      extHere = toLowerCase(getExtension(fileList[i]));
      if (endsWith(fileList[i], "/"))
        {
      	  selectedDirList[iDir] = fileList[i];
      	  iDir++;
        }
      else if (extHere == ext)
        {
          selectedFileList[j] = fileList[i];
          j++;
        }
    }

  selectedFileList = Array.trim(selectedFileList, j);
  selectedDirList = Array.trim(selectedDirList, iDir);
  if (ext == "")
    {
    	return selectedDirList;
    }
  else
    {
    	return selectedFileList;
    }
}

function printArray(array)
{
  // Print array elements recursively
  for (i=0; i<array.length; i++)
    print(array[i]);
}

function getExtension(filename)
{
  ext = substring( filename, lastIndexOf(filename, ".") + 1 );
  return ext;
}
