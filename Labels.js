// Label the created elements.
var layerName = "Label-Layer";
var styleName = "Label-Styles";
var paraStyleName = "Para-label-Style";
var folderName = "Labeled PDFs"

// Save old settings to restore later
var oldInteractionPrefs = app.scriptPreferences.userInteractionLevel;

// Turn of modal interuptions
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

// Create global Rectangle variable for storing the visible window co-ordinates.
var theRect;

var activeDocs = {
    allDocs: app.documents,
    labelOpen: function () {
        app.activeDocument.zeroPoint = [0,0];
        prepDoc(app.activeDocument);
        labelLinks(app.activeDocument);
    },
    labelDocs: function (){
        var docIndsnFiles = docFolder.getFiles("*.indd");
        for (var i = 0; i < docIndsnFiles.length; i++) {
            app.open(docIndsnFiles[i]);
            app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
            if (app.dialogs.length > 0){
                app.dialogs.everyItem().destroy();
            }
            prepDoc(app.activeDocument);
            labelLinks(app.activeDocument);
            createPDF(app.activeDocument);
            app.activeDocument.close();
        };
    }
}

// Create PDFs in selected Folder.
function createPDF(theDoc) {
    var myPath = folderPath;
    var theDocname = theDoc.name+"-Labeled";
    var prefs = theDoc.parent.pdfExportPreferences;
    var pdfPath = myPath+"/"+theDocname+".pdf";
    prefs.exportReaderSpreads = true;
    theDoc.exportFile(ExportFormat.PDF_TYPE, File(new File(pdfPath)));
}


// Add any necessary styles/layers to the documents.
function prepDoc(theDoc) {
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    theDoc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.PICAS;
    theDoc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.PICAS;
    theDoc.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;
            if (app.dialogs.length > 0){
                app.dialogs.everyItem().destroy();
            }

// Check for label layer
    var layers = theDoc.layers;
    var theLayer = layers.itemByName(layerName);
        if (theLayer === null) {
            theLayer = theDoc.layers.add();
            theLayer.name = layerName;
        }
        theLayer.move(LocationOptions.BEFORE, theDoc.layers[0]);

// Check for label styles
    var color = theDoc.colors.itemByName("Labels-Yellow");
        if (color === null){
            color = theDoc.colors.add();
            color.name = "Labels-Yellow";
            color.space = ColorSpace.CMYK;
            color.colorValue = [0, 0, 100, 0];
        };

// Add Text Area Style
    var labelObjectStyle = theDoc.objectStyles.itemByName(styleName);
        if (labelObjectStyle === null) {
            labelObjectStyle = theDoc.objectStyles.add();
            with(labelObjectStyle){
                name = styleName;
                enableFill = true;
                enableStroke = false;
                enableTextWrapAndOthers = true;
                textFramePreferences.ignoreWrap = true;
                enableStrokeAndCornerOptions = false;
                enableTextFrameBaselineOptions = true;
                enableTextFrameGeneralOptions = true;
                fillColor = color;
                textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;
                fillTransparencySettings.blendingSettings.opacity = 65;
                };
            };

// Add paragraph Style
    var labelParagraphStyle = theDoc.paragraphStyles.itemByName(paraStyleName);
        if  (labelParagraphStyle === null) {
            labelParagrapheStyle = theDoc.paragraphStyles.add();
            with(labelParagrapheStyle){
                name = paraStyleName;
                capitalization = Capitalization.ALL_CAPS;
                try {
                    appliedFont = "Arial";
                    fontStyle = "Bold";
                } catch(e){};
                pointSize = 16;
                justification = Justification.CENTER_ALIGN;
                fillColor = "Black";
            }
        };
    }

// Find the Rectangle containing the link, this fixes issues with Groups of images.
function findRect(theLink) {
    if (theLink.parent instanceof Spread) {
        theRect = theLink;
    } else if (theLink.parent instanceof Document){
        return null;
    }
    else {
        findRect(theLink.parent);
    }
};

function labelLinks(theDoc) {
// Get all links
    var allLinks = theDoc.links;
    var theLinks = [];
// Remove links that aren't in the spread
    for (var i = 0; i <  allLinks.length; i++) {
         if (allLinks[i].parent.parent.parent instanceof Spread || allLinks[i].parent.parent.parent.parent instanceof Spread){
            if (allLinks[i] !== null) {
                if(allLinks[i].parent.parentPage !== null){
                theLinks.push(allLinks[i]);
            }
            };
         };
    };
// Label appropriate links
    for (var i = 0; i < theLinks.length; i++) {
        var theLink = theLinks[i];
        findRect(theLink);
        var thePage = theLink.parent.parentPage;
        var visibleWindow = theRect.visibleBounds;
        var newBounds = [visibleWindow[2] - 2.5 , visibleWindow[1] - 15, (visibleWindow[2] + 2.5), visibleWindow[3] + 15]
        var theLabel = thePage.textFrames.add();
        var style = theDoc.objectStyles.itemByName(styleName);
            theLabel.applyObjectStyle(style, true);
            theLabel.geometricBounds = newBounds;
            theLabel.contents = theLink.name;

// Fit and Style the labels
        var text = theLabel.paragraphs[0];
        var paraStyle = theDoc.paragraphStyles.itemByName(paraStyleName);
        var charStyle = theDoc.characterStyles[0];
            text.applyCharacterStyle(charStyle, true);
            text.applyParagraphStyle(paraStyle, true);
            theLabel.fit(FitOptions.frameToContent);


// Add padding to labels
        var oldBounds = theLabel.visibleBounds;
        var newBounds = [oldBounds[0] - 1, oldBounds[1] - 1, oldBounds[2] + 1, oldBounds[3] + 1];
            theLabel.visibleBounds = newBounds;

// Make sure label is on the printable page
        if (theDoc.pages.length > 1) {
                var leftBounds = theDoc.pages[1].bounds[3];
            } else {
                var rightBounds = theDoc.pages[0].bounds[3];}

       
            checkBounds(theLabel);
       
    };
}

// Checks to see if the label is on the spread, adjusts the bounds to fit the page if it isn't.
function checkBounds(theLabel) {
    var fixedBounds = theLabel.geometricBounds;
    var theDoc = app.documents[0];
    if (theDoc.pages.length > 1) {
        var pageBounds = [theDoc.pages[0].bounds[0], theDoc.pages[0].bounds[1], theDoc.pages[1].bounds[2], theDoc.pages[1].bounds[3] ];}
    else {
        var pageBounds = theDoc.pages[0].bounds;
    }
    if (fixedBounds[0] < pageBounds[0]){
        fixedBounds[2] = fixedBounds[2] - (fixedBounds[2] - pageBounds[2]);
        fixedBounds[0] = pageBounds[0];
    };
    if (fixedBounds[2] > pageBounds[2]){
        fixedBounds[0] = fixedBounds[0] + (pageBounds[2] - fixedBounds[2]);
        fixedBounds[2] = pageBounds[2];
    };
    if (fixedBounds[1] < pageBounds[1]) {
        fixedBounds[3] = fixedBounds[3] + (fixedBounds[1] - pageBounds[1]);
        fixedBounds[1] = pageBounds[1];
    };
    if (fixedBounds[3] > pageBounds[3]) {
        fixedBounds[1] = fixedBounds[1] - (fixedBounds[3] - pageBounds[3]);
        fixedBounds[3] = pageBounds[3];
    };
    theLabel.geometricBounds = fixedBounds;
}

app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
var docFolder = Folder.selectDialog("Select Indesign Folder");
var folderPath = Folder.selectDialog("Select a folder for the created PDFs");
activeDocs.labelDocs();
//activeDocs.labelOpen();

// Restore previous settings
app.scriptPreferences.userInteractionLevel = oldInteractionPrefs;

// Maybe add an additional "Remove all added styles", though probably not necessary as the docs are all closed without saving afterwards.
