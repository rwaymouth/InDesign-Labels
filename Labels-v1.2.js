///// Variables ////////
var oldInteractionPrefs = app.scriptPreferences.userInteractionLevel;
 app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAlerts; 


var layerName = "Labels-Layer";

var redLabelColor = 'Labels-Red';
var redColor = {
	name: redLabelColor,
	space: ColorSpace.CMYK,
	model: ColorModel.PROCESS,
	colorValue:[0, 100, 100, 0]
	};	

var yellowLabelColor = 'Labels-Yellow';
var yellowColor = {
	name: yellowLabelColor,
	model: ColorModel.PROCESS,
	space: ColorSpace.CMYK,
	colorValue:[0,  0, 100, 0]
	};

var paragraphStyleName = "Labels-Text"
var labelParagraphStyle = {
                name: paragraphStyleName,
                basedOn: "None",
                nextStyle: "None",
                capitalization: Capitalization.ALL_CAPS,
                enableFill: true,
                pointSize: 16,
                justification: Justification.CENTER_ALIGN,
                fillColor: "Black"
            };

var labelName = 'Base-Label';
var baseLabelStyle = {
	  name: "Test Style",
                enableFill: true,
                enableStroke: false,
                enableParagraphStyle: true,
                appliedParagraphStyle: "None",
                enableTextWrapAndOthers: true,
                enableStrokeAndCornerOptions: false,
                enableTextFrameBaselineOptions: true,
                enableTextFrameGeneralOptions: true,
	   textFramePreferences: {
	   	verticalJustification: VerticalJustification.CENTER_ALIGN,
	   	ignoreWrap: true
	   },
                fillTransparencySettings: {
                	blendingSettings: {
                		opacity: 65
                	}
                }
};

var redLabelStyle = clone(baseLabelStyle);
redLabelStyle.fillColor = redLabelColor;
redLabelStyle.name = "Red-Label";

var yellowLabelStyle = clone(baseLabelStyle);
yellowLabelStyle.fillColor = yellowLabelColor;
yellowLabelStyle.name = "Yellow-Label";


///// Helper Functions //////

function clone(obj) {
    if (null == obj || "object" != typeof obj) return obj;
    var copy = obj.constructor();
    for (var attr in obj) {
        if (obj.hasOwnProperty(attr)) copy[attr] = obj[attr];
    }
    return copy;
}

function smashAlerts(){
               app.dialogs.everyItem().destroy();
}

function addStyle(theDoc, theAttr, theStyle, theStyleName){
	smashAlerts();  
	try {
		addedStyle = theDoc[theAttr].item(theStyleName);
		styleTest = addedStyle.name;
	} catch(theError){
		added = theDoc[theAttr].add(theStyle);
	}
}


	
function prepDoc(theDoc) {
	//app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

	if (app.dialogs.length > 0){
		smashAlerts();
		alert('smash!');
	};
	addStyle(theDoc, 'colors', redColor, redLabelColor);
	addStyle(theDoc, 'colors', yellowColor, yellowLabelColor);
	addStyle(theDoc, 'objectStyles',  redLabelStyle, redLabelStyle.name);
	addStyle(theDoc, 'objectStyles',  yellowLabelStyle,  yellowLabelStyle.name);
	addStyle(theDoc, 'paragraphStyles', labelParagraphStyle, labelParagraphStyle.name);
	var layers = theDoc.layers;
    	var theLayer = layers.itemByName(layerName);
        		if (theLayer === null) {
            			theLayer = theDoc.layers.add();
            			theLayer.name = layerName;
        		}
        	theLayer.move(LocationOptions.BEFORE, theDoc.layers[0]);

        	theDoc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.PICAS;
    	theDoc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.PICAS;
    	theDoc.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;
   	theDoc.zeroPoint = [0,0];

}

function pickStyle(theImage){
	var path = theImage.location;
	if (path === "Downloads") {
		return yellowLabelStyle;
	} else {
		return redLabelStyle;
	}
}

// Checks to see if the image is on the Spread
var onPage = null;

function onSpread(theLink) {
    if (theLink.parent instanceof Spread) {
    	if ( theLink.parentPage) {
    		onPage = true;
    	} else {
    		onPage = false;
    	}
    } else if (theLink.parent instanceof Document) {
    	onPage = false;
    } else {
    	onSpread(theLink.parent);
    }
};

// Checks to see if the label overlaps other images

function checkLabels(theDoc) {
	var labels = theDoc.layers.itemByName(layerName).textFrames;
	checkForDuplicate(labels);
		for (var i = 0; i < labels.length; i++){
   			var theLabel = labels[i];
   		             checkLabelOverlap(theLabel);
            			checkBounds(theLabel);
			checkLabelOverlap(theLabel);
   		};
}

function checkForDuplicate(labels) {
	for (var i = 0; i < labels.length -1; i++){ 
		var theLabel = labels[i]
		var labelContent = theLabel.contents;
		var otherLabels = labels[i + 1];
		if (labelContent === otherLabels.contents) {
			otherLabels.remove();
		}
	};
}

function checkLabelOverlap(theLabel){
	var otherLabels = app.activeDocument.layers.itemByName(layerName).textFrames;
	var theDoc = app.activeDocument;
	for (var i = 0; i < otherLabels.length; i++) {
		    var otherLabel = otherLabels[i];
		    var labelBounds = theLabel.visibleBounds;
		    var otherBounds = otherLabel.visibleBounds;
		    var height = labelBounds[2] - labelBounds[0];
		    if (theDoc.pages.length > 1) {
		        var pageBounds = [theDoc.pages[0].bounds[0], theDoc.pages[0].bounds[1], theDoc.pages[1].bounds[2], theDoc.pages[1].bounds[3] ];}
		    else {
		        var pageBounds = theDoc.pages[0].bounds;
		    }
		    if (labelBounds[0] > pageBounds[2]/2) {
		    	height = height*-1;
		    }
		    if (    !(
		                (labelBounds[0]>otherBounds[2]) || 
		                (labelBounds[2]<otherBounds[0]) ||
		                (labelBounds[1]>otherBounds[3]) || 
		                (labelBounds[3]<otherBounds[1]) ) && 
		                (labelBounds[0] !== otherBounds[0])
		        ){
		            labelBounds[0] = labelBounds[0] + height;
		            labelBounds[2] = labelBounds[2] + height;
		            theLabel.visibleBounds = labelBounds;
		        }; 
	    }
};

// Checks to see if the label is on the spread, adjusts the bounds to fit the page if it isn't.
function checkBounds(theLabel) {
	    var fixedBounds = theLabel.geometricBounds;
	    var theDoc = app.documents[0];
	    //Maybe check to see if parentPage + 1 has the same bounds? Checks both the length and the value at once.
	    if (theDoc.pages.length > 1) {
		var pageBounds = [theDoc.pages[0].bounds[0], theDoc.pages[0].bounds[1], theDoc.pages[1].bounds[2], theDoc.pages[1].bounds[3] ];}
	    else {
	        	var pageBounds = theDoc.pages[0].bounds;
	    }
	    if (fixedBounds[0] < pageBounds[0]){
	        	fixedBounds[2] = fixedBounds[2] + Math.abs((pageBounds[0] - fixedBounds[0]));
	        	fixedBounds[0] = pageBounds[0];
	    };
	    if (fixedBounds[2] > pageBounds[2]){
	        	fixedBounds[0] = fixedBounds[0] - (pageBounds[2] - fixedBounds[2]);
	        	fixedBounds[2] = pageBounds[2];
	    };
	    if (fixedBounds[1] < pageBounds[1]) {
	        	fixedBounds[3] = fixedBounds[3] + Math.abs((fixedBounds[1] - pageBounds[1]));
	        	fixedBounds[1] = pageBounds[1];
	    };
	    if (fixedBounds[3] > pageBounds[3]) {
	        	fixedBounds[1] = fixedBounds[1] - (fixedBounds[3] - pageBounds[3]);
	        	fixedBounds[3] = pageBounds[3];
	    };
	    	theLabel.geometricBounds = fixedBounds;
}

// Calculates the size of the image on the page
function calcSize(theLink) {
	// Calculate the pages area
	var theDoc = app.documents[0];
	var pageBounds = theDoc.pages[0].bounds;
	var pageWidth = pageBounds[3] - pageBounds[1];
	var pageHeight = pageBounds[2] - pageBounds[0];
	var pageArea = pageWidth * pageHeight;
	//Calculate the images area
	var bounds = theLink.parent.visibleBounds;
	var width = (bounds[3] - bounds[1]);
	var height = (bounds[2] - bounds[0]);
	var imageArea = width * height;
	var size = imageArea / pageArea;
	if (size < 0.25) {
		return '1/4 page size';
	}  
	if (size > 0.25 && size < 0.5) {
		return '1/2 page size';
	}  
	if (size > 0.5 && size < 0.75) {
		return '3/4 page size';
	}  
	if ( size > 0.75 && size < 1.05) {
		return 'Full page size';
	} 
	if ( size > 1.05 ) {
		return 'Spread size';
	}
}

// Link Location 

function linkLocation(theLink) {

	var linkFolder = theLink.filePath;
	if (theLink.filePath.indexOf('Downloads')  > 0 ||
            		theLink.filePath.indexOf('FPOs')  > 0
    	){ 
		return 'Downloads';
     	}
	if (linkFolder.indexOf('Client') !== -1) {
		return 'Client';
	}
	if (linkFolder.indexOf('Ads') !== -1) {
		return 'Ads';
	}
	if (linkFolder.indexOf('CD') !== -1){
		return 'CD';
	}
	if (linkFolder.indexOf('Standing Art') !== -1){
		return 'Standing Art';
	}

}

function getVendor(theLink) {
	var linkInfo = theLink.name.split("-");
	if ((linkInfo.length - 1) >= 1) {
		return linkInfo[0];
	} else {
		return "Vendor Not Listed";
	}
}

function getRights(theLink) {
	var linkInfo = theLink.name.split("-");
	if ((linkInfo.length - 1) >= 1 && linkInfo[1] == "RM" || (linkInfo.length - 1) >= 1 && linkInfo[1] === "RF" ) {
		return linkInfo[1];
	} else {
		return "Rights Not Listed";
	}
}

// Image Data Structure

function MyImage() {
	var imageInfo = ' ';

	var __construct = function() {
		imageInfo = 'Information not set yet';
	}()

	this.getInfo = function() {
		return imageInfo;
	}
	this.setImageInfo = function(theLink) {
		 imageInfo = {
		 	fileName: theLink.name,
		 	page: theLink.parent.parentPage,
		 	location: linkLocation(theLink),
		 	bounds: theLink.parent.geometricBounds,
		 	size: calcSize(theLink),
		 	vendor: getVendor(theLink),
		 	rights: getRights(theLink)
		 };
	}
}

// Building functions

function processImages (theDoc) {
	prepDoc(theDoc);
	var theLinks = theDoc.links;
	var selectedImages = [];
	for (var i = 0; i < theLinks.length; i++) {
		onSpread(theLinks[i]);
		if (onPage === true){
			var theImage = new MyImage();
			theImage.setImageInfo(theLinks[i]);
			var theImage = theImage.getInfo();
			labelImage(theImage);
			selectedImages.push(theImage);
		}
	};
}

function labelImage(theImage) {
	//Labeling part of the script goes here
	var theDoc = app.activeDocument;
	var thePage = theImage.page;
	var imageBounds = theImage.bounds;
	var vCenter = imageBounds[0] + (imageBounds[2] - imageBounds[0])/2;
	var hCenter =  imageBounds[1]+ (imageBounds[3] - imageBounds[1])/2;
	var newBounds = [vCenter  - 2.5,  hCenter -15, vCenter + 2.5, hCenter + 15];
	
	smashAlerts();

	var theLabel = thePage.textFrames.add();


	var labelColor = pickStyle(theImage);
	var style = theDoc.objectStyles.item(labelColor.name);
    	theLabel.applyObjectStyle(style, true);
	theLabel.visibleBounds = newBounds;
             checkLabelOverlap(theLabel);
	if (theImage.location ===  'CD'){
		theLabel.contents = 'CD - ' + theLabel.contents;
	}
	if (theImage.location === 'Client') {
		theLabel.contents = 'CLIENT - ' + theLabel.contents;
	}
	if (theImage.location === 'Ads') {
		theLabel.contents = 'ADS - ' + theLabel.contents;
	}
	if (theImage.location === 'Standing Art') {
		theLabel.contents = 'SA - ' + theLabel.contents;
	}
	theLabel.contents = theLabel.contents + theImage.fileName;

	var theText = theLabel.paragraphs[0];
	var paraStyle = theDoc.paragraphStyles.item(labelParagraphStyle.name);
	var charStyle = theDoc.characterStyles[0];

	theText.applyCharacterStyle(charStyle, true);
	theText.applyParagraphStyle(paraStyle, true);
    	theLabel.fit(FitOptions.frameToContent);
        	
        	var oldBounds = theLabel.visibleBounds;
        	var newBounds = [oldBounds[0] - 1, oldBounds[1] - 1, oldBounds[2] + 1, oldBounds[3] + 1];
            theLabel.visibleBounds = newBounds;
}


// Create PDFs in selected Folder.
function createPDF(theDoc) {
    var myPath = folderPath;
    var theDocname = theDoc.name+"-Labeled";
    var prefs = app.pdfExportPreferences;
    var pdfPath = myPath+"/"+theDocname+".pdf";
    var preset = app.pdfExportPresets.itemByName("[Smallest File Size]");
    preset.exportReaderSpreads = true;
    app.activeDocument.asynchronousExportFile(ExportFormat.pdfType, File(new File(pdfPath)),  false, preset).waitForTask();
    theDoc.close();
}

var docFolder = Folder.selectDialog("Select Indesign Folder");
var folderPath = Folder.selectDialog("Select a folder for the created PDFs");

var activeDocs = {
       labelDocs: function (){
        	var docIndsnFiles = docFolder.getFiles("*.indd");
        	for (var i = 0; i < docIndsnFiles.length; i++) {
        		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAlerts; 
		if (app.dialogs.length > 0){
			smashAlerts();
		};
            		app.open(docIndsnFiles[i]);
		processImages(app.activeDocument);
		checkLabels(app.activeDocument);
             	createPDF(app.activeDocument);
             	smashAlerts();
        		};
    	},
        labelDoc: function() {
    	            	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    	            	if (app.dialogs.length > 0){
			smashAlerts();
		};	
        		processImages(app.activeDocument);
		checkLabels(app.activeDocument);
             	createPDF(app.activeDocument);
        }

}
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAlerts; 

activeDocs.labelDocs();
