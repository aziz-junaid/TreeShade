//Tree Shade - Javascript for Adobe InDesign
//dulajun@gmail.com
//Tel.: +966508018243
var copyNumber = "9.3.7";
var tsClientID = "";
var userID = "null";
var userName = "unknown";
var userEmail = "no email";
var rootFolderPath = "";
var linksIdleTask;
var savedDocIdleTask;
var initiateIdleTask;
var workshopPath = "";
var pagesPath = "";
var dataPath = "";
var versionsPath = "";
var finalPDFPath = "";
var previewPath = "";
var trashPath = "";
var isSettingObtained = false;
var startPDFExportingAction;
var startImagesExportingAction;
var setImagesExportingAction;
var pausePDFAutoExportingAction;
var startAutoCheckInAction;
var autoImportOutOfTreeLinksAction;
var autoDynamicLinksAndLeafFolderAction;
var autoUpdateLinksRecordsAction;
var openAndSavePSDAction;
var solveOverflowsAction;
var leaveEmptyPagesAction;
var noTextCoordinatesAction;
var isStartPDFExporting;
var isStartImagesExporting;
var isStartAutoCheckIn;
var beforeUpdateList = new Array;
var saveWithoutAskList = new Array;
var CSExtensionList = new Array;
var nameAtPathsList = new Array;
var openAndSavePSDList = new Array;
var isAlertedAboutBridge = false;
var treeShadeNamespace = "http://ns.dulajun.com/treeshade/";
var treeShadePrefix = "ts:";
var autoUpdateLinksRecordsInterval = 1;
var autoUpdateLinksRecordsStep = 0;
var snippetFilePath_ID_List = new Array;
var bridgeScanningQueue = new Array;
var tsPlaceSelectionAndFileAndDoc = null;
var bridgeScanningStep = 0;
var tsDynamicSnippetTag = "::";
var tsEditText = null;
var stepsAmount = 2.0;
var minColWidth = 5.0;
var isToAdjust = true;
var isResetWidth = true;
var isExpandColumnsHaveExtraLines = true;
var isNeedToSwitch = false;

try {
    if (ExternalObject.AdobeXMPScript == undefined)
          ExternalObject.AdobeXMPScript = new ExternalObject('lib:AdobeXMPScript');
    XMPMeta.registerNamespace(treeShadeNamespace, treeShadePrefix);
}
catch (error) {
    /*Error*/$.writeln ("Error Tree Shade No. " + $.line);
}

var xLib = null;
try {  
    xLib = new ExternalObject("lib:\PlugPlugExternalObject");  
} catch (e) { alert(e) }
var eventObj = null;

//live snippets
var autoSubmitAllLiveSnippetsAction;
var autoUpdateAllLiveSnippetsAction;
var autoCheckDocumentAction;
var currentLiveSnippetStartChar = null;

var number_relinkDialog = "";

var message_warningAndDesDialog = "";
var isUpdated = false;

function addToBridgeScanningQueue (fullPath) {
    if ((fullPath + "/").indexOf (workshopPath + "/") == 0) {
        fullPath = fullPath.replace (workshopPath, "");
        for (var bsq = bridgeScanningQueue.length - 1; bsq >= 0; bsq--) {
            if (bridgeScanningQueue[bsq] == fullPath) {
                return false;
            }
        }
        bridgeScanningQueue.push (fullPath);
        return true;
    }
    return false;
}

function installStartupScripts (extensionPath, userDataPath, userDocumentsPath, hostApplicationPath, senderLOADorBUTTON)
{ 
    //determining os
    var bridgeStartupFileName = "Tree Shade Bridge.jsx";
    if ($.os.indexOf ("Macintosh") != 0) {
        hostApplicationPath = hostApplicationPath.slice (0, hostApplicationPath.lastIndexOf ("/"));
    }
    var isUpdated = false;
    //Defining files to be copied and set Thumbnail objects
    var bridgeStartupFileThumb = new Thumbnail(new File(extensionPath + "/ExtendScript Startup/" + bridgeStartupFileName));
    var resourcesFolderThumb = new Thumbnail(Folder(extensionPath + "/ExtendScript Startup/Tree Shade Resources"));
    //target all Bridge Versions
    var theHostEnd = hostApplicationPath.slice (hostApplicationPath.lastIndexOf ("/"));
    theHostEnd = theHostEnd.slice (7);
    var theDesFolder = new Folder (userDataPath + "/Adobe/" + theHostEnd + "/Startup Scripts");
    
    var desBridgeStartupFile = new File (theDesFolder.fsName.replace(/\\/g, "/") + "/" + bridgeStartupFileName);
    if (!theDesFolder.exists) {
        theDesFolder.create ();
    }
    if (desBridgeStartupFile.exists) {
        if (desBridgeStartupFile.modified.getTime () >= bridgeStartupFileThumb.spec.modified.getTime ()) {
            if (xLib) {  
                var eventObj = new CSXSEvent();   
                eventObj.type = "com.treeshade.copystartupresult"; 
                eventObj.data = "Uptodate";  
                eventObj.dispatch();  
                
            }
            return true;
        }
        else {
            isUpdated = true;
        }
    }
        
    //removing old version
    var oldFilesAndFolders = theDesFolder.getFiles ("Tree Shade *");
    for (var aff = 0; aff < oldFilesAndFolders.length; aff++) {
        if (File.decode (oldFilesAndFolders[aff].name).indexOf ("Tree Shade Path for") == -1) {
            forceRemove (oldFilesAndFolders[aff], true);
        }
    }

    //copying bridge startup file
    bridgeStartupFileThumb.copyTo (theDesFolder);
    //copying resources
    resourcesFolderThumb.copyTo (theDesFolder);
    //write extension path
    var extensionPathRecordFile = new File (theDesFolder.fsName.replace(/\\/g, "/") + "/Tree Shade Extension Path");
    writeEncodedFile (extensionPathRecordFile, extensionPath);

    if (xLib) {  
        var eventObj = new CSXSEvent();   
        eventObj.type = "com.treeshade.copystartupresult"; 
        eventObj.data = (isUpdated? "Updated" : "Welcome");  
        eventObj.dispatch();
    }
    return true;
 }

function forceRemove (fileOrFolder, isToRemoveRoot) {
    if (fileOrFolder instanceof File) {
        fileOrFolder.remove ();
        return true;
    }
    else {
        var allFilesAndFolders = fileOrFolder.getFiles ("*");
        for (var aff = 0; aff < allFilesAndFolders.length; aff++) {
            forceRemove (allFilesAndFolders[aff], true);
        }
        if (isToRemoveRoot) {
            fileOrFolder.remove();
        }
        return true;
    }
    return false;
}

//InDesign Section

function mainInDesign() {
    /*function_start*///alert ("mainInDesign-start");
    if (CSExtensionList.length == 0) {
        tsClientID = Folder.desktop.parent.fsName.replace(/\\/g, "/").slice (Folder.desktop.parent.fsName.replace(/\\/g, "/").lastIndexOf ("/")+1);
        if (tsClientID == "OneDrive") {
            tsClientID = Folder.desktop.parent.parent.fsName.replace(/\\/g, "/").slice (Folder.desktop.parent.parent.fsName.replace(/\\/g, "/").lastIndexOf ("/")+1);
        }

        if (xLib) {  
            eventObj = new CSXSEvent();   
            eventObj.type = "com.treeshade.rootset";   
        }

        app.eventListeners.add("beforeOpen", beforeOpenHandler, false);
        app.eventListeners.add("afterActivate", traceChanges, false);
        app.eventListeners.add("beforeDeactivate", beforeDeactivateHandler, false);
        app.eventListeners.add("afterQuit", afterQuitHandler, false);
        app.eventListeners.add("afterSelectionChanged", activityTock, false);
        app.insertLabel ("activityTock", "0");
        createMenu ();
        initiateIdleTask = app.idleTasks.add ();
        initiateIdleTask.eventListeners.add("onIdle", initiateSettingHandler);
        initiateIdleTask.sleep = 3000;

        CSExtensionList.push (".psd");
        CSExtensionList.push (".png");
        CSExtensionList.push (".jpg");
        CSExtensionList.push (".jpeg");
        CSExtensionList.push (".tif");
        CSExtensionList.push (".tiff");
        CSExtensionList.push (".ai");
        CSExtensionList.push (".eps");
        CSExtensionList.push (".pdf");
        CSExtensionList.push (".indd");
        CSExtensionList.push (".indt");
        CSExtensionList.push (".ait");
        CSExtensionList.push (".swf");
        CSExtensionList.push (".psb");
        CSExtensionList.push (".mov");
        CSExtensionList.push (".mp4");
        CSExtensionList.push (".aep");
        CSExtensionList.push (".svg");
    }
}

function createMenu () {
    /*function_start*///alert ("createMenu-start");
    var Str;
	Str = "Report External Links";
	var reportOutAction = app.scriptMenuActions.add(Str);
    reportOutAction.eventListeners.add("onInvoke", function () {reportOutHandler ();}, false);
	Str = "Reveal the Document in Bridge";
	var revealDocumentAction = app.scriptMenuActions.add(Str);
	revealDocumentAction.eventListeners.add("onInvoke", function () {revealDocumentHandler ();}, false);
	Str = "PDF Exporting After Checking in";
	startPDFExportingAction = app.scriptMenuActions.add(Str);
	startPDFExportingAction.eventListeners.add("onInvoke", function () {startPDFExportingHandler ();}, false);
	Str = "Images Exporting After Saving";
	startImagesExportingAction = app.scriptMenuActions.add(Str);
	startImagesExportingAction.eventListeners.add("onInvoke", function () {startImagesExportingHandler ();}, false);
	Str = "Set Pages Exporting Resolution";
	setImagesExportingAction = app.scriptMenuActions.add(Str);
	setImagesExportingAction.eventListeners.add("onInvoke", function () {setPagesExportingHandler ();}, false);
	Str = "Set PDF Preset";
	var setPDFPresetAction = app.scriptMenuActions.add(Str);
	setPDFPresetAction.eventListeners.add("onInvoke", function () {setPDFPresetHandler ();}, false);
	Str = "Auto Check In is Active";
	pausePDFAutoExportingAction = app.scriptMenuActions.add(Str);
	pausePDFAutoExportingAction.eventListeners.add("onInvoke", function () {pausePDFAutoExportingHandler ();}, false);
    pausePDFAutoExportingAction.eventListeners.add("beforeDisplay", function () {pausePDFAutoExportingDisplayHandler ();}, false);
	Str = "Checking in After Saving";
	startAutoCheckInAction = app.scriptMenuActions.add(Str);
	startAutoCheckInAction.eventListeners.add("onInvoke", function () {startAutoCheckInHandler ();}, false);
	Str = "Auto Import Links";
	autoImportOutOfTreeLinksAction = app.scriptMenuActions.add(Str);
	autoImportOutOfTreeLinksAction.eventListeners.add("onInvoke", function () {autoImportOutOfTreeLinksHandler ();}, false);
    //Auto Move Linked Files in Hierarchy
	Str = "Dynamic Links and Leaf Folder On";
	autoDynamicLinksAndLeafFolderAction = app.scriptMenuActions.add(Str);
	autoDynamicLinksAndLeafFolderAction.eventListeners.add("onInvoke", function () {autoDynamicLinksAndLeafFolderHandler ();}, false);
    autoDynamicLinksAndLeafFolderAction.eventListeners.add("beforeDisplay", function () {autoDynamicLinksAndLeafFolderDisplayHandler ();}, false);
    //Update Links Records after Place or Delete
	Str = "Auto Update Links Records";
	autoUpdateLinksRecordsAction = app.scriptMenuActions.add(Str);
	autoUpdateLinksRecordsAction.eventListeners.add("onInvoke", function () {autoUpdateLinksRecordsHandler ();}, false);
    //Open and Save PSD files before check in
	Str = "Update PSDs Before Checking in";
	openAndSavePSDAction = app.scriptMenuActions.add(Str);
	openAndSavePSDAction.eventListeners.add("onInvoke", function () {openAndSavePSDHandler ();}, false);
    openAndSavePSDAction.eventListeners.add("beforeDisplay", function () {openAndSavePSDBeforeDisplayHandler ();}, false);
    //ignore overflows
	Str = "Solve Overset Before Checking in";
	solveOverflowsAction = app.scriptMenuActions.add(Str);
	solveOverflowsAction.eventListeners.add("onInvoke", function () {solveOverflowsHandler ();}, false);
    solveOverflowsAction.eventListeners.add("beforeDisplay", function () {solveOverflowsBeforeDisplayHandler ();}, false);
    //Delete Empty Pages
	Str = "Don't Export Empty Pages";
	leaveEmptyPagesAction = app.scriptMenuActions.add(Str);
	leaveEmptyPagesAction.eventListeners.add("onInvoke", function () {leaveEmptyPagesHandler ();}, false);
    leaveEmptyPagesAction.eventListeners.add("beforeDisplay", function () {leaveEmptyPagesBeforeDisplayHandler ();}, false);
    //Export Text Coordinates
	Str = "Don't Export Text Coordinates";
	noTextCoordinatesAction = app.scriptMenuActions.add(Str);
	noTextCoordinatesAction.eventListeners.add("onInvoke", function () {noTextCoordinatesHandler ();}, false);
    noTextCoordinatesAction.eventListeners.add("beforeDisplay", function () {noTextCoordinatesDisplayHandler ();}, false);
	Str = "Set Dynamic Link on Selected";
	var dynamicLinkAction = app.scriptMenuActions.add(Str);
	dynamicLinkAction.eventListeners.add("onInvoke", function () {dynamicLinkHandler ();}, false);
	Str = "Load Selected Link File";
	var loadFileAction = app.scriptMenuActions.add(Str );
	loadFileAction.eventListeners.add("onInvoke", function () {loadFileHandler ();}, false);
	Str = "Update Links Records";
	var updateLinksRecordsAction = app.scriptMenuActions.add(Str );
	updateLinksRecordsAction.eventListeners.add("onInvoke", function () {updateLinksRecordsHandler ();}, false);
	Str = "Export Fonts";
	var exportFontsAction = app.scriptMenuActions.add(Str );
	exportFontsAction.eventListeners.add("onInvoke", function () {exportFontsHandler ();}, false);
	Str = "About Tree Shade";
	var aboutAction = app.scriptMenuActions.add(Str );
	aboutAction.eventListeners.add("onInvoke", function () {aboutHandler ();}, false);
    
    try{
        app.menus.item("$ID/Main").submenus.item("Tree Shade").remove();
    }catch(myError){}
	var treeShadeMenu = app.menus.item("$ID/Main").submenus.add("Tree Shade");
    treeShadeMenu.menuItems.add(reportOutAction);
    treeShadeMenu.menuItems.add(revealDocumentAction);
    treeShadeMenu.menuSeparators.add();
    treeShadeMenu.menuItems.add(dynamicLinkAction);
    treeShadeMenu.menuItems.add(loadFileAction);
    treeShadeMenu.menuItems.add(updateLinksRecordsAction);
    treeShadeMenu.menuItems.add(exportFontsAction);
    treeShadeMenu.menuSeparators.add();
    treeShadeMenu.menuItems.add(setImagesExportingAction);
    treeShadeMenu.menuItems.add(setPDFPresetAction);
    treeShadeMenu.menuSeparators.add();
    treeShadeMenu.menuItems.add(autoDynamicLinksAndLeafFolderAction);
    treeShadeMenu.menuItems.add(openAndSavePSDAction);
    treeShadeMenu.menuItems.add(solveOverflowsAction);
    treeShadeMenu.menuItems.add(leaveEmptyPagesAction);
    treeShadeMenu.menuItems.add(noTextCoordinatesAction);
    treeShadeMenu.menuSeparators.add();
    treeShadeMenu.menuItems.add(startAutoCheckInAction);
    treeShadeMenu.menuItems.add(startPDFExportingAction);
    treeShadeMenu.menuItems.add(pausePDFAutoExportingAction);
    treeShadeMenu.menuItems.add(startImagesExportingAction);
    treeShadeMenu.menuItems.add(autoImportOutOfTreeLinksAction);
    treeShadeMenu.menuItems.add(autoUpdateLinksRecordsAction);
    treeShadeMenu.menuSeparators.add();
    treeShadeMenu.menuItems.add(aboutAction);
    
    //live snippets
	Str = "New Live Snippet";
	var newLiveSnippetAction = app.scriptMenuActions.add("New Live Snippet");
	newLiveSnippetAction.eventListeners.add("onInvoke", function () {newLiveSnippetHandler();}, false);
	Str = "Dynamic Identifier";
	var dynamicPathAndNameAction = app.scriptMenuActions.add(Str);
    dynamicPathAndNameAction.eventListeners.add("onInvoke", function () {dynamicPathAndNameHandler ();}, false);
	Str = "Dynamic Content";
	var dynamicLiveSnippetAction = app.scriptMenuActions.add(Str);
    dynamicLiveSnippetAction.eventListeners.add("onInvoke", function () {dynamicLiveSnippetHandler ();}, false);
	Str = "Dynamic Style";
	var dynamicStyleAction = app.scriptMenuActions.add(Str);
    dynamicStyleAction.eventListeners.add("onInvoke", function () {dynamicStyleHandler ();}, false);
	Str = "Dynamic Replace";
	var dynamicReplaceAction = app.scriptMenuActions.add(Str);
    dynamicReplaceAction.eventListeners.add("onInvoke", function () {dynamicReplaceHandler ();}, false);
    Str = "Place Live Snippet";
	var placeAction = app.scriptMenuActions.add(Str);
	placeAction.eventListeners.add("onInvoke", function () {placeLiveSnippetHandler ();}, false);
	Str = "Unlink Instance";
	var unlinkInstanceAction = app.scriptMenuActions.add(Str);
	unlinkInstanceAction.eventListeners.add("onInvoke", function () {unlinkInstanceHandler ();}, false);
	Str = "Show Info.";
	var showInfoAction = app.scriptMenuActions.add(Str);
    showInfoAction.eventListeners.add("onInvoke", function () {showInfoHandler ();}, false);
	Str = "Reveal in Bridge";
	var revealInBridgeAction = app.scriptMenuActions.add(Str);
	revealInBridgeAction.eventListeners.add("onInvoke", function () {revealInBridgeHandler ();}, false);
	Str = "Submit Selected";
	var submitSelectedAction = app.scriptMenuActions.add("Selected");
	submitSelectedAction.eventListeners.add("onInvoke", function () {submitSelectedHandler ();}, false);
	Str = "Submit All";
	var submitAllAction = app.scriptMenuActions.add("All");
	submitAllAction.eventListeners.add("onInvoke", function () {submitAllHandler ();}, false);
	Str = "Auto Submit";
	autoSubmitAllLiveSnippetsAction = app.scriptMenuActions.add("Auto Submit");
	autoSubmitAllLiveSnippetsAction.eventListeners.add("onInvoke", function () {autoSubmitAllLiveSnippetsHandler ();}, false);
	Str = "Update Selected";
	var updateSelectedAction = app.scriptMenuActions.add("Selected");
	updateSelectedAction.eventListeners.add("onInvoke", function () {updateSelectedHandler ();}, false);
	Str = "Update All";
	var updateAllAction = app.scriptMenuActions.add("All");
	updateAllAction.eventListeners.add("onInvoke", function () {updateAllHandler ();}, false);
	Str = "Auto Update";
	autoUpdateAllLiveSnippetsAction = app.scriptMenuActions.add("Auto Update");
	autoUpdateAllLiveSnippetsAction.eventListeners.add("onInvoke", function () {autoUpdateAllLiveSnippetsHandler ();}, false);
	Str = "Check Document";
	var checkDocumentAction = app.scriptMenuActions.add("Check Document");
	checkDocumentAction.eventListeners.add("onInvoke", function () {checkDocumentHandler ();}, false);
    Str = "Auto Check Document";
    autoCheckDocumentAction = app.scriptMenuActions.add("Auto Check Document");
    autoCheckDocumentAction.eventListeners.add("onInvoke", function () {autoCheckDocumentHandler ();}, false);
	Str = "Select Content";
	var selectContentAction = app.scriptMenuActions.add("Content");
	selectContentAction.eventListeners.add("onInvoke", function () {selectContentHandler ();}, false);
	Str = "Select End Marker";
	var selectEndMarkerAction = app.scriptMenuActions.add("End Marker");
	selectEndMarkerAction.eventListeners.add("onInvoke", function () {selectEndMarkerHandler ();}, false);
	Str = "Select First Child";
	var selectFirstChildAction = app.scriptMenuActions.add("First Child");
	selectFirstChildAction.eventListeners.add("onInvoke", function () {selectFirstChildHandler ();}, false);
	Str = "Select Next Brother";
	var selectNextBrotherAction = app.scriptMenuActions.add("Next Brother");
	selectNextBrotherAction.eventListeners.add("onInvoke", function () {selectNextBrotherHandler ();}, false);
	Str = "Select Previous Brother";
	var selectPreviousBrotherAction = app.scriptMenuActions.add("Previous Brother");
	selectPreviousBrotherAction.eventListeners.add("onInvoke", function () {selectPreviousBrotherHandler ();}, false);
	Str = "Select Start Marker";
	var selectStartMarkerAction = app.scriptMenuActions.add("Start Marker");
	selectStartMarkerAction.eventListeners.add("onInvoke", function () {selectStartMarkerHandler ();}, false);
	Str = "Select to End Border";
	var selectToEndBorderAction = app.scriptMenuActions.add("To End Border");
	selectToEndBorderAction.eventListeners.add("onInvoke", function () {selectToEndBorderHandler ();}, false);
	Str = "Select to Start Border";
	var selectToStartBorderAction = app.scriptMenuActions.add("To Start Border");
	selectToStartBorderAction.eventListeners.add("onInvoke", function () {selectToStartBorderHandler ();}, false);
	Str = "Select Whole Snippet";
	var selectWholeSnippetAction = app.scriptMenuActions.add("Whole Snippet");
	selectWholeSnippetAction.eventListeners.add("onInvoke", function () {selectWholeSnippetHandler ();}, false);
	Str = "Reflect Marker Horizontally";
	var reflectMarkerHorizontallyAction = app.scriptMenuActions.add(Str );
	reflectMarkerHorizontallyAction.eventListeners.add("onInvoke", function () {reflectMarkerHorizontallyHandler ();}, false);
	Str = "Reflect Marker Vertically";
	var reflectMarkerVerticallyAction = app.scriptMenuActions.add(Str);
	reflectMarkerVerticallyAction.eventListeners.add("onInvoke", function () {reflectMarkerVerticallyHandler ();}, false);
    
    try{
        app.menus.item("$ID/Main").submenus.item("Live Snippets").remove();
    }catch(myError){}
	var liveSnippetsMenu = app.menus.item("$ID/Main").submenus.add("Live Snippets");
	liveSnippetsMenu.menuItems.add(newLiveSnippetAction);
    var dynamicMenu = liveSnippetsMenu.submenus.add("Dynamic");
    dynamicMenu.menuItems.add(dynamicPathAndNameAction);
    dynamicMenu.menuItems.add(dynamicLiveSnippetAction);
    dynamicMenu.menuItems.add(dynamicStyleAction);
    dynamicMenu.menuItems.add(dynamicReplaceAction); 
    liveSnippetsMenu.menuSeparators.add();
	liveSnippetsMenu.menuItems.add(placeAction);
    var submitMenu = liveSnippetsMenu.submenus.add("Submit");
	submitMenu.menuItems.add(submitSelectedAction);
    submitMenu.menuItems.add(submitAllAction);
    submitMenu.menuItems.add(autoSubmitAllLiveSnippetsAction);
	var updateMenu = liveSnippetsMenu.submenus.add("Update");
	updateMenu.menuItems.add(updateSelectedAction);
	updateMenu.menuItems.add(updateAllAction);
    updateMenu.menuItems.add(autoUpdateAllLiveSnippetsAction);
    liveSnippetsMenu.menuItems.add(checkDocumentAction);
    liveSnippetsMenu.menuItems.add(autoCheckDocumentAction);
	liveSnippetsMenu.menuItems.add(unlinkInstanceAction);
    liveSnippetsMenu.menuItems.add(showInfoAction);
    liveSnippetsMenu.menuItems.add(revealInBridgeAction);
    liveSnippetsMenu.menuSeparators.add();
	var selectMenu = liveSnippetsMenu.submenus.add("Select");
	selectMenu.menuItems.add(selectContentAction);
	selectMenu.menuItems.add(selectToStartBorderAction);
	selectMenu.menuItems.add(selectToEndBorderAction);
	selectMenu.menuItems.add(selectWholeSnippetAction);
	selectMenu.menuItems.add(selectFirstChildAction);
	selectMenu.menuItems.add(selectNextBrotherAction);
	selectMenu.menuItems.add(selectPreviousBrotherAction);
	selectMenu.menuItems.add(selectStartMarkerAction);
	selectMenu.menuItems.add(selectEndMarkerAction);
	liveSnippetsMenu.menuItems.add(reflectMarkerHorizontallyAction);
	liveSnippetsMenu.menuItems.add(reflectMarkerVerticallyAction);
    
    //auto actions checked settings for live snippets
    if (app.extractLabel ("LS_AUTO_SUBMIT") == "")
        app.insertLabel ("LS_AUTO_SUBMIT", "no");
    if (app.extractLabel ("LS_AUTO_UPDATE") == "")
        app.insertLabel ("LS_AUTO_UPDATE", "yes");
    if (app.extractLabel ("LS_AUTO_CHECK") == "")
        app.insertLabel ("LS_AUTO_CHECK", "yes");
    autoSubmitAllLiveSnippetsAction.checked = (app.extractLabel ("LS_AUTO_SUBMIT") == "yes");
    autoUpdateAllLiveSnippetsAction.checked = (app.extractLabel ("LS_AUTO_UPDATE") == "yes");
    autoCheckDocumentAction.checked = (app.extractLabel ("LS_AUTO_CHECK") == "yes");
}

function aboutHandler () {
    alert ("Developed by Abdulaziz A. Junaid\ndulajun@gmail.com\nTree Shade Root Folder:\n" + Folder (workshopPath).parent.fsName.replace(/\\/g, "/") + "\n\nVersion No. " + copyNumber);
}

function getUserInfo (targetFile) {  
    /**///$.writeln ($.line);
    var infoArray = new Array;
    infoArray.push (""); //User ID
    infoArray.push (""); //User User Name
    infoArray.push (""); //User Email
    infoArray.push (""); //User Key
    try {
        if (ExternalObject.AdobeXMPScript == undefined)
              ExternalObject.AdobeXMPScript = new ExternalObject('lib:AdobeXMPScript');
        XMPMeta.registerNamespace(treeShadeNamespace, treeShadePrefix);
        var xmpFile = new XMPFile(targetFile.fsName.replace(/\\/g, "/"), XMPConst.UNKNOWN, XMPConst.OPEN_FOR_READ);
        var xmp = xmpFile.getXMP();
        var TS_USER_ID = xmp.getProperty(treeShadeNamespace, "TS_USER_ID");
        TS_USER_ID = TS_USER_ID + "";
        infoArray[0] = TS_USER_ID;
        var TS_USER_NAME = xmp.getProperty(treeShadeNamespace, "TS_USER_NAME");
        TS_USER_NAME = TS_USER_NAME + "";
        infoArray[1] = TS_USER_NAME;
        var TS_USER_EMAIL = xmp.getProperty(treeShadeNamespace, "TS_USER_EMAIL");
        TS_USER_EMAIL = TS_USER_EMAIL + "";
        infoArray[2] = TS_USER_EMAIL;
        var TS_USER_KEY = xmp.getProperty(treeShadeNamespace, "TS_USER_KEY");
        TS_USER_KEY = TS_USER_KEY + "";
        infoArray[3] = TS_USER_KEY;
        xmpFile.closeFile(XMPConst.CLOSE_UPDATE_SAFELY);
        return infoArray;
    }
    catch (e) {
        return null;
    }
}

function obtainRootSettings () {
    /*function_start*///alert ("obtainRootSettings-start");
    var rootTest = app.extractLabel ("Tree Shade Path for " + tsClientID);
    var metadataUser = "";
    if (!rootTest) {
        if (xLib) {  
            var eventObj = new CSXSEvent();   
            eventObj.type = "com.treeshade.rootset"; 
            eventObj.data = "Please open Adobe Bridge and select 'Tree Shade' from Window -> Extensions.<br><br>No Root Path.";  
            eventObj.dispatch();
        }
        return false;
    }
    if (rootFolderPath == rootTest) {
        return true;
    }
    
    workshopPath = rootTest + "/Workshop";
    pagesPath = rootTest + "/Pages"; 
    dataPath = rootTest + "/.Data";
    versionsPath = rootTest + "/Versions";
    finalPDFPath = rootTest + "/Final PDF";
    previewPath = rootTest + "/Versions PDFs";
    trashPath = rootTest + "/Trash";

    if (app.extractLabel ("TS_VERSIONS_OUTPUT_TO_WORKSHOP") == "TRUE") {
        previewPath = workshopPath;
    }

    var userImage = Folder (rootTest).getFiles ("I'm *.jpg");
    if (userImage.length > 0) {
        var isMoreThanOne = (userImage.length > 1);
        userImage = userImage[0];
        var userInfo = getUserInfo (userImage);
        if (!userInfo) {
            metadataUser = "<br><br>Couldn't read the Metadata User Info from his Image:<br>" + userImage.fsName.replace(/\\/g, "/");
        }
        else {
            userID = userInfo[0];
            userName = userInfo[1];
            userEmail = userInfo[2];
        }
        if (isMoreThanOne) {
            metadataUser = "<br><br>There is more than one user Images in local data folder.<br>";
        }
    }
    else {
        metadataUser = "<br><br>The user image is missing!<br>";
    }
    if (xLib) {  
        var eventObj = new CSXSEvent();   
        eventObj.type = "com.treeshade.rootset"; 
        eventObj.data = "Root Path:<br>" + rootTest + metadataUser;  
        eventObj.dispatch();
    }
    rootFolderPath = rootTest;
    return true;
}

function isFolder (targetFile) {
    if (targetFile instanceof Folder) {
        if (targetFile.name[0] != ".")
            return true;
    }
    return false;
}

function tsFillZeros (number, digitsCount) {
	var digits = number.toString();
	while (digits.length < digitsCount)
		digits = "0" + digits;
	return digits;
}

function startPDFExportingHandler () {
    /*function_start*///alert ("startPDFExportingHandler-start");
    isStartPDFExporting = !isStartPDFExporting;
    if (isStartPDFExporting) {
        startPDFExportingAction.checked = true;
        app.insertLabel ("startPDFExporting", "start");
        var PDFExportingFile = new File (dataPath + "/No PDF Exporting");
        PDFExportingFile.remove ();
    }
    else {
        startPDFExportingAction.checked = false;
        app.insertLabel ("startPDFExporting", "quit");
        var PDFExportingFile = new File (dataPath + "/No PDF Exporting");
        if (!PDFExportingFile.parent.exists)
            PDFExportingFile.parent.create ();
        writeFile (PDFExportingFile, "TS");
    }
}

function startImagesExportingHandler () {
    /*function_start*///alert ("startImagesExportingHandler-start");
    isStartImagesExporting = !isStartImagesExporting;
    if (isStartImagesExporting) {
        startImagesExportingAction.checked = true;
        app.insertLabel ("startImagesExporting", "start");
    }
    else {
        startImagesExportingAction.checked = false;
        app.insertLabel ("startImagesExporting", "quit");
    }
}

function setPagesExportingHandler () {
    /*function_start*///alert ("setPagesExportingHandler-start");
    try {
        app.activeDocument;
    }
    catch (e) {
        return false;
    }
    var docPageRes = app.activeDocument.extractLabel ("TS_PAGES_RES");
    if (docPageRes == "") {
        docPageRes = "72";
    }
    var docPageQua = app.activeDocument.extractLabel ("TS_PAGES_QUALITY");
    if (docPageQua == "") {
        docPageQua = "MEDIUM";
    }
    var docPageExt = app.activeDocument.extractLabel ("TS_PAGES_EXTENSION");
    if (docPageExt == "") {
        docPageExt = "jpg";
    }
    var newRes = prompt ("Document Pages Exporting Settings:\nResolution: 1 - 2400\nQuality: LOW, MEDIUM, HIGH or MAXIMUM\nExtension: jpg, png", docPageRes + ":" + docPageQua + ":" + docPageExt);
    if (newRes) {
        newRes = newRes.split (":");
        var resValue = parseInt (newRes[0], 10);
        if (resValue > 0 && resValue < 2401) { //1 - 2400
            app.activeDocument.insertLabel ("TS_PAGES_RES", resValue.toString ());
        }
        if (newRes.length > 1) {
            if (newRes[1] == "LOW" || newRes[1] == "MEDIUM" || newRes[1] == "HIGH" || newRes[1] == "MAXIMUM") {
                app.activeDocument.insertLabel ("TS_PAGES_QUALITY", newRes[1]);
            }
        }
        if (newRes.length > 2) {
            if (newRes[2] == "png" || newRes[2] == "PNG") {
                app.activeDocument.insertLabel ("TS_PAGES_EXTENSION", "png");
            }
            else {
                app.activeDocument.insertLabel ("TS_PAGES_EXTENSION", "jpg");
            }
        }
    }
}

function setPDFPresetHandler () {
    /*function_start*///alert ("setPDFPresetHandler-start");
    try {
        app.activeDocument;
    }
    catch (e) {
        return false;
    }
    var docPDFPreset = app.activeDocument.extractLabel ("TS_PDF_PRESET");
    var newDocPDFPreset = prompt ("Auto Exporting PDF Preset:", docPDFPreset);
    if (newDocPDFPreset) {
        app.activeDocument.insertLabel ("TS_PDF_PRESET", newDocPDFPreset);
    }
}

function pausePDFAutoExportingHandler () {
    app.insertLabel ("TS_PDF_AUTO_EXPORTING_PAUSED", (app.extractLabel ("TS_PDF_AUTO_EXPORTING_PAUSED") == "Yes")? "No" : "Yes"); 
    if (app.extractLabel ("TS_PDF_AUTO_EXPORTING_PAUSED") == "Yes") {
        pausePDFAutoExportingAction.checked = false;
    }
    else {
        pausePDFAutoExportingAction.checked = true;
    }
}

function pausePDFAutoExportingDisplayHandler () {
    if (app.extractLabel ("TS_PDF_AUTO_EXPORTING_PAUSED") == "Yes") {
        pausePDFAutoExportingAction.checked = false;
    }
    else {
        pausePDFAutoExportingAction.checked = true;
    }
}

function startAutoCheckInHandler () {
    /*function_start*///alert ("startAutoCheckInHandler-start");
    isStartAutoCheckIn = !isStartAutoCheckIn;
    if (isStartAutoCheckIn) {
        startAutoCheckInAction.checked = true;
        app.insertLabel ("startAutoCheckIn", "start");
        var result = confirm ("Do you want to create new version after every save?");
        if (result) {
            startAutoCheckInAction.title = "Create New Version After Save";
            app.insertLabel ("isToCreateNewVersion", "Yes");
        }
        else {
            startAutoCheckInAction.title = "Checking in After Save";
            app.insertLabel ("isToCreateNewVersion", "No");
        }
        result = confirm ("Do you want to be asked after every save?");
        if (result) {
            startAutoCheckInAction.title = "Ask for " + startAutoCheckInAction.title;
            app.insertLabel ("isAskingForCheckingIn", "Yes");
        }
        else {
            app.insertLabel ("isAskingForCheckingIn", "No");
        }
    }
    else {
        startAutoCheckInAction.checked = false;
        app.insertLabel ("startAutoCheckIn", "quit");
    }
}

function autoImportOutOfTreeLinksHandler () {
    app.insertLabel ("autoImportOutOfTreeLinks", (app.extractLabel("autoImportOutOfTreeLinks") == "Yes")? "No" : "Yes");
    autoImportOutOfTreeLinksAction.checked = (app.extractLabel("autoImportOutOfTreeLinks") == "Yes");
    if ((app.extractLabel("autoImportOutOfTreeLinks") == "Yes")) {
        var result = confirm ("Auto import could be to 'Links' or 'Leaf' folder beside the document file.\nClick Yes to import to 'Leaf', or No to import to 'Links'?\n");
        if (result)
            app.insertLabel ("autoImportOutOfTreeLinksFolder", "Leaf");
        else
            app.insertLabel ("autoImportOutOfTreeLinksFolder", "Links");
    }
}

function autoDynamicLinksAndLeafFolderHandler () {
    var isThereDoc = true;
    try {
        app.activeDocument;
    }
    catch (myError) {
        isThereDoc = false;
    }
    if (isThereDoc) {
        var theDoc = app.activeDocument;
        theDoc.insertLabel ("TS_DYNAMIC_AND_LEAF", (theDoc.extractLabel("TS_DYNAMIC_AND_LEAF") != "No")? "No" : "Yes");
    }
    else {
        
    }
}

function autoDynamicLinksAndLeafFolderDisplayHandler () {
    var isThereDoc = true;
    try {
        app.activeDocument;
    }
    catch (myError) {
        isThereDoc = false;
    }
    if (isThereDoc) {
        autoDynamicLinksAndLeafFolderAction.enabled = true;
        var theDoc = app.activeDocument;
        autoDynamicLinksAndLeafFolderAction.checked = !(theDoc.extractLabel("TS_DYNAMIC_AND_LEAF") == "No");
    }
    else {
        autoDynamicLinksAndLeafFolderAction.enabled = false;
        autoDynamicLinksAndLeafFolderAction.checked = false;
    }
}

function autoUpdateLinksRecordsHandler () {
    app.insertLabel ("autoUpdateLinksRecords", (app.extractLabel("autoUpdateLinksRecords") == "Yes")? "No" : "Yes");
    autoUpdateLinksRecordsAction.checked = (app.extractLabel("autoUpdateLinksRecords") == "Yes");
}

function openAndSavePSDHandler () {
    var isThereDoc = true;
    try {
        app.activeDocument;
    }
    catch (myError) {
        isThereDoc = false;
    }
    if (isThereDoc) {
        var theDoc = app.activeDocument;
        theDoc.insertLabel ("TS_OPEN_AND_SAVE_PSD", (theDoc.extractLabel("TS_OPEN_AND_SAVE_PSD") != "No")? "No" : "Yes");
        if (theDoc.extractLabel("TS_OPEN_AND_SAVE_PSD") != "No") {
            alert ("This targets files with 'ts_lyr' flag in their name. The file will be opened, and its layers updated. If the flag includes an output format like 'ts_lyr_jpg', 'ts_lyr_png', or 'ts_lyr_gif', the PSD file will be duplicated to the page thumbnails directory, exported to the specified format, and the duplicate will be deleted.");
        }
    }
    else {
        
    }
}

function openAndSavePSDBeforeDisplayHandler () {
    var isThereDoc = true;
    try {
        app.activeDocument;
    }
    catch (myError) {
        isThereDoc = false;
    }
    if (isThereDoc) {
        openAndSavePSDAction.enabled = true;
        var theDoc = app.activeDocument;
        openAndSavePSDAction.checked = (theDoc.extractLabel("TS_OPEN_AND_SAVE_PSD") != "No");
    }
    else {
        openAndSavePSDAction.enabled = false;
        openAndSavePSDAction.checked = false;
    }
}

function solveOverflowsHandler () {
    var isThereDoc = true;
    try {
        app.activeDocument;
    }
    catch (myError) {
        isThereDoc = false;
    }
    if (isThereDoc) {
        var theDoc = app.activeDocument;
        theDoc.insertLabel ("TS_SOLVE_OVERFLOWS", (theDoc.extractLabel("TS_SOLVE_OVERFLOWS") == "No")? "Yes" : "No");
        if (theDoc.extractLabel("TS_SOLVE_OVERFLOWS") != "No") {
            var isToAutoSolveOverflows = confirm ("If text gets overset, should Tree Shade fix it by shrinking the text? Click No to leave documents open for you to fix manually.");
            theDoc.insertLabel ("TS_AUTO_SOLVE_OVERFLOWS", isToAutoSolveOverflows? "Yes" : "No");
        }
    }
    else {
        
    }
}

function solveOverflowsBeforeDisplayHandler () {
    var isThereDoc = true;
    try {
        app.activeDocument;
    }
    catch (myError) {
        isThereDoc = false;
    }
    if (isThereDoc) {
        solveOverflowsAction.enabled = true;
        var theDoc = app.activeDocument;
        solveOverflowsAction.checked = !(theDoc.extractLabel("TS_SOLVE_OVERFLOWS") == "No");
    }
    else {
        solveOverflowsAction.enabled = false;
        solveOverflowsAction.checked = false;
    }
}

function leaveEmptyPagesHandler () {
    var isThereDoc = true;
    try {
        app.activeDocument;
    }
    catch (myError) {
        isThereDoc = false;
    }
    if (isThereDoc) {
        var theDoc = app.activeDocument;
        theDoc.insertLabel ("TS_LEAVE_EMPTY_PAGES", (theDoc.extractLabel("TS_LEAVE_EMPTY_PAGES") == "No")? "Yes" : "No");
    }
    else {
        
    }
}

function noTextCoordinatesHandler () {
    var isThereDoc = true;
    try {
        app.activeDocument;
    }
    catch (myError) {
        isThereDoc = false;
    }
    if (isThereDoc) {
        var theDoc = app.activeDocument;
        theDoc.insertLabel ("TS_NO_TEXT_COORDINATES", (theDoc.extractLabel("TS_NO_TEXT_COORDINATES") == "No")? "Yes" : "No");
    }
    else {
        
    }
}

function leaveEmptyPagesBeforeDisplayHandler () {
    var isThereDoc = true;
    try {
        app.activeDocument;
    }
    catch (myError) {
        isThereDoc = false;
    }
    if (isThereDoc) {
        leaveEmptyPagesAction.enabled = true;
        var theDoc = app.activeDocument;
        leaveEmptyPagesAction.checked = !(theDoc.extractLabel("TS_LEAVE_EMPTY_PAGES") == "No");
    }
    else {
        leaveEmptyPagesAction.enabled = false;
        leaveEmptyPagesAction.checked = false;
    }
}

function noTextCoordinatesDisplayHandler () {
    var isThereDoc = true;
    try {
        app.activeDocument;
    }
    catch (myError) {
        isThereDoc = false;
    }
    if (isThereDoc) {
        noTextCoordinatesAction.enabled = true;
        var theDoc = app.activeDocument;
        noTextCoordinatesAction.checked = !(theDoc.extractLabel("TS_NO_TEXT_COORDINATES") == "No");
    }
    else {
        noTextCoordinatesAction.enabled = false;
        noTextCoordinatesAction.checked = false;
    }
}

function dynamicLinkHandler () {
    var targetLink;
    if (app.selection.length == 1) {
        if (app.selection[0] instanceof Rectangle) {
            if (app.selection[0].allGraphics.length > 0) {
                if (app.selection[0].allGraphics[0].itemLink) {
                    targetLink = app.selection[0].allGraphics[0].itemLink;
                }
            }
        }
        else if (app.selection[0] instanceof Image) {
            if (app.selection[0].itemLink) {
                targetLink = app.selection[0].itemLink;
            }
        }
    }
    if (targetLink) {
        var dynamicLinkPhrase = targetLink.extractLabel ("TS_DYNAMIC_LINK");
        if (!dynamicLinkPhrase) {
            dynamicLinkPhrase = ".$/<SP>/<LN>.*";
        }
        dynamicLinkPhrase = tsGetText (["Dynamic Link Phrase:",
                                        "//: Root of Tree Shade path.",
                                        "./: Link file parent folder path. add dots for parent of parent...",
                                        "<.>: Link file parent folder name. add dots for parent of parent...",
                                        "To reference to the document instead of the link file write like this '.$/' or '<.$>'...",
                                        "Also use these tags with path:",
                                        "<Link_Name>, <Sub_Path>, <Doc_Name>, <Doc_Order>, <Page_Num>, <Applied_Section_Prefix>,",
                                        "<Applied_Section_Marker>, <Master_Prefix>, <Master_Name>, <Master_Full_Name>,",
                                        "<Para_First>, <Para_Before>, <Para_This>, <Char_Order>, <Anchored_Order>, <Para_Order>,",
                                        "<Para_After>, <Para_Style>, <Char_Style>, <Name_At>, <Shift1>.",
                                        "",
                                        "The following accepts abbreviation and used for date, time and random code:",
                                        "<Day_Seconds>, <Day_Minutes>, <Day_Hours>, <Weekday_Order>, <Weekday_Name>, <Weekday_Abbreviation>,",
                                        "<Month_Date>, <Month_Order>, <Month_Name>, <Month_Abbreviation>, <Year_Full>,",
                                        "<Year_Abbreviation>, <Time_Decimal>, <Time_Hexadecimal>, <Time_Shortest>,",
                                        "<Random_Style_Long> Where 'Style' a combination of symbols 'a' for small letters,",
                                        "'A' capital letters, 'h' small hexadecimal, 'H' large hexadecimal and '1' for numbers.",
                                        "For example <Random_1A_8> means random 8 characters consists of capital letters and numbers",
                                        ""
                                    ], dynamicLinkPhrase, false);
        if (dynamicLinkPhrase) {
            targetLink.insertLabel ("TS_DYNAMIC_LINK", dynamicLinkPhrase);
        }
    }
}

function loadFileHandler () {
    var targetLink;
    if (app.selection.length == 1) {
        if (app.selection[0] instanceof Rectangle) {
            if (app.selection[0].allGraphics.length > 0) {
                if (app.selection[0].allGraphics[0].itemLink) {
                    targetLink = app.selection[0].allGraphics[0].itemLink;
                }
            }
        }
        else if (app.selection[0] instanceof Image) {
            if (app.selection[0].itemLink) {
                targetLink = app.selection[0].itemLink;
            }
        }
    }
    if (targetLink) {
        var fileID = targetLink.extractLabel ("TSID");
        if (fileID) {
            var loadMessageFolder = new Folder (dataPath + "/Messages/To Load");
            if (!loadMessageFolder.exists) {
                loadMessageFolder.create ();
            }
            var messageFileName = fileID.replace (/\//g,"\.");
            messageFileName = messageFileName.slice (1);
            writeFile (File (loadMessageFolder.fsName.replace(/\\/g, "/") + "/" + messageFileName), "TS");
            alert ("The loading under processing.");
        }
    }
}

function updateLinksRecordsHandler () {
    try {
        app.activeDocument.filePath;
    }
    catch (e) {
        return false;
    }
    var targetDocument = app.activeDocument;
    var docPath = preparePath (targetDocument.filePath.fsName.replace(/\\/g, "/")) + "/";
    if (docPath.indexOf (workshopPath + "/") == 0 && workshopPath != "") {
        var docName = docPath + "/" + targetDocument.name;
        updateLinksRecords (targetDocument, docName);
    }
}

function exportFontsHandler () {
    try {
        app.activeDocument.filePath;
    }
    catch (e) {
        return false;
    }
    var targetDocument = app.activeDocument;
    if (workshopPath == "")
        return false;
    if (!Folder (workshopPath).exists)
        return false;
    var docPath = preparePath (targetDocument.filePath.fsName.replace(/\\/g, "/"));
    var docName = docPath + "/" + targetDocument.name;
    
    var fontsFolder = new Folder (docPath + "/Document fonts");
    var isFontsExist = fontsFolder.exists;
    if (!isFontsExist)
        fontsFolder.create ();
    for (var f = 0; f < targetDocument.fonts.length; f++) {
        var fontLocation;
        try {
            fontLocation = targetDocument.fonts[f].location;
        }
        catch (e) {continue;}
        var fontFullPath = preparePath (fontLocation);
        var fontName = fontFullPath.slice (fontFullPath.lastIndexOf ("/") + 1);
        targetFontFile = File (fontsFolder.fsName.replace(/\\/g, "/") + "/" + fontName);
        if (!targetFontFile.exists) {
            File (fontFullPath).copy (targetFontFile);
        }
    }
}

function reportOutHandler () {
    /*function_start*///alert ("reportOutHandler-start");
    try {
        app.activeDocument.filePath;
    }
    catch (e) {
        return false;
    }
    var docPath = preparePath (app.activeDocument.filePath.fsName.replace(/\\/g, "/")) + "/";
    if (docPath.indexOf (workshopPath + "/") == 0 && workshopPath != "") {
        var outList = getOutLinks (app.activeDocument.links);
        if (outList.length > 0) {
            var MessageContent = "Links out of " + workshopPath.slice (workshopPath.lastIndexOf ("/") + 1) + " folder.";
            for (var n = 0; n < outList.length; n++) {
                MessageContent += "\n" + File.decode (outList[n].name);
            }
            alert (MessageContent);
        }
        else {
            alert ("It's OK.\nNo links out of " + workshopPath.slice (workshopPath.lastIndexOf ("/") + 1) + " folder."); 
        }
    }
    else {
        alert ("This document out of 'Workshop' folder!");
    }
}

function revealDocumentHandler () {
    try {
        app.activeDocument.filePath;
    }
    catch (e) {
        return false;
    }
    var targetDocument = app.activeDocument;
    try {
        targetDocument.filePath;
    }
    catch (e) {
        return false;
    }

    var docPath = preparePath (targetDocument.filePath.fsName.replace(/\\/g, "/")) + "/";
    var docFullPath = docPath + "/" + targetDocument.name;
    var documentFile = new File (docFullPath);
    if (documentFile.exists) {
        revealFileInBridge (documentFile);
    }
}

function getOutLinks (allLinks) {
    /*function_start*///alert ("getOutLinks-start");
    var outList = new Array;
    for (var c = 0; c < allLinks.length; c++) {
        var linkFile = new File (preparePath (allLinks[c].filePath));
        if ((linkFile.parent.fsName.replace(/\\/g, "/") + "/").indexOf (workshopPath + "/") != 0)
            outList.push (linkFile);
    }
    return outList;
}

function initiateSettingHandler () {
    /*function_start*///alert ("initiateSettingHandler-start");
    if (!obtainRootSettings ()) {
        return false;
    }
    if (isSettingObtained) {
        return true;
    }
    var cooperateInExportingLabel = app.extractLabel ("startPDFExporting");
    var cooperateInLinksUpdatingLabel = app.extractLabel ("startImagesExporting");
    var cooperateInAutoCheckInLabel = app.extractLabel ("startAutoCheckIn");
    if (cooperateInExportingLabel == "quit") {
        startPDFExportingAction.checked = false;
        isStartPDFExporting = false;
    }
    else {
        startPDFExportingAction.checked = true;
        isStartPDFExporting = true;
        app.insertLabel ("startPDFExporting", "start");
    }
    if (cooperateInLinksUpdatingLabel == "quit") {
        startImagesExportingAction.checked = false;
        isStartImagesExporting = false;
    }
    else {
        startImagesExportingAction.checked = true;
        isStartImagesExporting = true;
        app.insertLabel ("startImagesExporting", "start");
    }
    if (cooperateInAutoCheckInLabel == "quit") {
        startAutoCheckInAction.checked = false;
        isStartAutoCheckIn = false;
    }
    else {
        startAutoCheckInAction.checked = true;
        isStartAutoCheckIn = true;
        app.insertLabel ("startAutoCheckIn", "start");
    }
    autoImportOutOfTreeLinksAction.checked = (app.extractLabel("autoImportOutOfTreeLinks") == "Yes");
    autoUpdateLinksRecordsAction.checked = (app.extractLabel("autoUpdateLinksRecords") == "Yes");
    if (app.extractLabel("autoImportOutOfTreeLinksFolder") == "")
        app.insertLabel ("autoImportOutOfTreeLinksFolder", "Leaf");
    linksIdleTask = app.idleTasks.add ();
    linksIdleTask.eventListeners.add("onIdle", linksIdleTaskHandler);
    linksIdleTask.sleep = 3000;
    savedDocIdleTask = app.idleTasks.add ();
    savedDocIdleTask.eventListeners.add("onIdle", savedDocIdleTaskHandler); 
    savedDocIdleTask.sleep = 2000;

    var tsPlaceSelectionAndFileAndDocTask = app.idleTasks.add ();
    tsPlaceSelectionAndFileAndDocTask.eventListeners.add("onIdle", tsPlaceSelectionAndFileAndDocHandler); 
    tsPlaceSelectionAndFileAndDocTask.sleep = 500;

    //set linking Preferences
    app.linkingPreferences.checkLinksAtOpen = false;
    app.linkingPreferences.findMissingLinksAtOpen = false;
    if (previewPath != "") {
        var exportPDFPreset = app.pdfExportPresets.itemByName ("Tree Shade");
        try {
            exportPDFPreset.name;
        }
        catch (e) {
            exportPDFPreset = app.pdfExportPresets.add ({name: "Tree Shade"});
            with (exportPDFPreset) {
                acrobatCompatibility = AcrobatCompatibility.ACROBAT_8;
                bleedTop = "5mm";
                bleedBottom =  "5mm";
                bleedInside = 0;
                bleedOutside =  "5mm";
                colorBars = false;
                bleedMarks = false;
                colorBitmapSamplingDPI = 300;
                cropImagesToFrames = true;
                cropMarks = true;
                exportLayers = true;
                exportNonprintingObjects = false;
                exportReaderSpreads = true;
                exportWhichLayers = ExportLayerOptions.EXPORT_VISIBLE_PRINTABLE_LAYERS;
                generateThumbnails = true;
                grayscaleBitmapSamplingDPI = 300;
                includeBookmarks = false;
                includeHyperlinks = false;
                includeICCProfiles = false;
                includeSlugWithPDF = false;
                includeStructure = true;
                interactiveElementsOption = InteractiveElementsOptions.DO_NOT_INCLUDE;
                optimizePDF = true;
                pageInformationMarks = true;
                pageMarksOffset = "2.5mm";
                registrationMarks = false;
            }
        }
    }
    isSettingObtained = true;
}

function activityTock (myEvent){
    var current = app.extractLabel ("activityTock");
    current = current ? parseInt (current, 10) : 0;
    current++;
    app.insertLabel ("activityTock", current.toString ());
    
    //Live Snippets
    //Determine current live snippet
	currentLiveSnippetStartChar = null;
    if (app.extractLabel ("LS_AUTO_CHECK") == "yes") {
        var levelObject = app.selection[0];
        currentLiveSnippetStartChar = getStartChar (levelObject);
        var aStartChar = currentLiveSnippetStartChar;
        var wasException = false;
        while (aStartChar) {
            currentLiveSnippetStartChar = aStartChar;
            if (aStartChar.textFrames.item(0).extractLabel ("LS_ID").indexOf("exception") != -1) {
                if (wasException) {
                    currentLiveSnippetStartChar = null;
                    break
                }
                wasException = true;
            }
            else {
                wasException = false;
            }
            if (aStartChar .index > 0) {
                aStartChar = getStartChar (aStartChar.parent.characters.item(aStartChar.index - 1));
            }
            else {
                if (aStartChar.parent.constructor.name == "Cell")
                    aStartChar = getStartChar (aStartChar.parent);
                else
                    aStartChar = getStartChar (aStartChar.parentTextFrames[0]);
            }
        }
    }
}

function afterQuitHandler (myEvent){
    var openedDocFolder = new Folder (dataPath + "/Opened Documents");
    var allFiles = openedDocFolder.getFiles ("*");
    if (allFiles) {
        for (var odf = 0; odf < allFiles.length; odf++) {
            allFiles[odf].remove();
        }
    }
    openedDocFolder.remove ();
}

function beforeOpenHandler (myEvent){
    /*function_start*///alert ("beforeOpenHandler-start");
    var docToBeOpenedFile = File (myEvent.fullName);
    if (docToBeOpenedFile.fsName.replace(/\\/g, "/").indexOf (versionsPath + "/") == 0) {
        myEvent.preventDefault ();
        var untitledDoc = app.open (docToBeOpenedFile, true, OpenOptions.OPEN_COPY);
        untitledDoc.addEventListener ("beforeSaveAs", beforeSaveAsHandler, false);
    }
}

function traceChanges (myEvent){
    /*function_start*///alert ("traceChanges-start");
    if (!app.modalState)
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    var targetObject;
    var targetDocument;
    var allLinks;
    try {
        targetObject = myEvent.parent;
        if (!(targetObject instanceof Application || targetObject instanceof Document)) {
            return;
        }
        if (targetObject instanceof Application) {
            if (targetObject.documents.length == 0) {
                return;
            }
            targetDocument = targetObject.activeDocument;
        }
        else {
            targetDocument = targetObject;
        }
    }
    catch (e) {
        return;
    }
    try {
        targetDocument.filePath;
    }
    catch (e) {
        return false;
    }
    if (workshopPath == "")
        return false;
    return setDocument (targetDocument);
}

function beforeDeactivateHandler (myEvent) {
    /*function_start*///alert ("traceChanges-start");
    if (!app.modalState)
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
    var targetObject;
    var targetDocument;
    try {
        targetObject = myEvent.parent;
        if (!(targetObject instanceof Application || targetObject instanceof Document)) {
            return;
        }
        if (targetObject instanceof Application) {
            if (targetObject.documents.length == 0) {
                return;
            }
            targetDocument = targetObject.activeDocument;
        }
        else {
            targetDocument = targetObject;
        }
    }
    catch (e) {
        return;
    }

    //auto Update Links Records
    if (app.extractLabel("autoUpdateLinksRecords") == "Yes") {
        try {
            targetDocument.filePath;
        }
        catch (e) {
            return false;
        }
        var docPath = preparePath (targetDocument.filePath.fsName.replace(/\\/g, "/")) + "/";
        if (docPath.indexOf (workshopPath + "/") == 0 && workshopPath != "") {
            var docName = docPath + "/" + targetDocument.name;
            updateLinksRecords (targetDocument, docName);
        }
    }
}
function setDocument (targetDocument) {
    var docPath = preparePath (targetDocument.filePath.fsName.replace(/\\/g, "/"));
    var docPathWithSlash = docPath + "/";
    if (docPathWithSlash.indexOf (workshopPath + "/") != 0) {
        return false;
    }
    var isListenerExist = false;
    for (var i = 0; i < targetDocument.eventListeners.length; i++) {
        if (targetDocument.eventListeners[i].handler.name == "linkPlaceHandler") {
            isListenerExist = true;
            break;
        }
    }
    if (!isListenerExist) {
        addToBridgeScanningQueue (docPath);
        targetDocument.addEventListener ("afterImport", linkPlaceHandler, false);
        targetDocument.addEventListener ("beforeSave", beforeSaveHandler, false);
        targetDocument.addEventListener ("afterSave", afterSaveHandler, false);
        targetDocument.addEventListener ("afterOpen", afterOpenHandler, false);
        targetDocument.addEventListener ("beforeClose", beforeCloseHandler, false);
        //targetDocument.addEventListener ("afterSaveAs", afterSaveAsHandler, false);
        targetDocument.addEventListener ("beforeSaveAs", beforeSaveAsHandler, false);
        //targetDocument.addEventListener ("beforeSaveACopy", beforeSaveACopyHandler, false);
        targetDocument.links.everyItem ().addEventListener ("afterUpdate", linkAfterUpdateHandler, false);
        targetDocument.links.everyItem ().addEventListener ("beforeUpdate", linkBeforeUpdateHandler, false);
        targetDocument.links.everyItem ().addEventListener ("beforeDelete", beforeDeleteLink, false);
        targetDocument.links.everyItem ().addEventListener ("afterAttributeChanged", afterAttributeChangedHandler, false);
        var docPathCurrent = targetDocument.extractLabel ("TS_PATH_CURRENT");
        if (docPathCurrent != docPath) {
            if (docPathCurrent) {
                targetDocument.insertLabel ("TS_PATH_PREVIOUS", docPathCurrent);
            }
            else {
                targetDocument.insertLabel ("TS_PATH_PREVIOUS", docPath);
            }
            targetDocument.insertLabel ("TS_PATH_CURRENT", docPath);
        }
    }
    else {
        allLinks = targetDocument.links;
        var speedChain = new Array;
        var bridgeTalkLinesArr = new Array;
        for (var c = 0; c < allLinks.length; c++) {
            var isLinkListenerExist = false;
            for (var i = 0; i < allLinks[c].eventListeners.length; i++) {
                if (allLinks[c].eventListeners[i].handler.name == "afterAttributeChangedHandler") {
                    isLinkListenerExist = true;
                    break;
                }
            }
            if (!isLinkListenerExist) {
                allLinks[c].addEventListener ("afterUpdate", linkAfterUpdateHandler, false);
                allLinks[c].addEventListener ("beforeUpdate", linkBeforeUpdateHandler, false);
                allLinks[c].addEventListener ("beforeDelete", beforeDeleteLink, false);
                allLinks[c].addEventListener ("afterAttributeChanged", afterAttributeChangedHandler, false);
            }
            testLink (allLinks[c], speedChain, false, bridgeTalkLinesArr);
        } 
        if (bridgeTalkLinesArr.length > 0) {
            sendImportListToInDesign (bridgeTalkLinesArr);
        } 
    }
    return true;
}

function preparePath (osPath) {
    /*function_start*///alert ("preparePath-start");
    if ($.os.indexOf ("Macintosh OS") == 0) {
        var resultPath = osPath.replace (/:/g, "/");
        if (resultPath.indexOf ("/Users/") > 0) {
            resultPath = resultPath.slice (resultPath.indexOf  ("/Users/"));
            return resultPath;
        }
        if (resultPath.indexOf ("~") == 0) {
            resultPath = resultPath.replace ("~", "/Users/" + tsClientID);
        }
        if (resultPath[0] != '/') {
            resultPath = "/Volumes/" + resultPath;
        }
        return resultPath;
    }
    else {
        return osPath;
    }
}

function getActualModified (modified, fileID) {
    /**///$.writeln ($.line);
    var fakeDateFile = new File (dataPath + "/IDs" + fileID + "/Workshop/FakeDate");
    if (fakeDateFile.exists) {
        var fakeDate = readFile (fakeDateFile);
        if (modified.toString () == fakeDate) {
            var beforeFakeFile = new File (dataPath + "/IDs" + fileID + "/Workshop/BeforeFake");
            var beforeFakeDate = readFile (beforeFakeFile);
            if (beforeFakeDate) {
                return parseInt (beforeFakeDate, 10);
            }
            else {
                return modified;
            }
        }
        else {
            fakeDateFile.remove ();
            var beforeFakeFile = new File (dataPath + "/IDs" + fileID + "/Workshop/BeforeFake");
            beforeFakeFile.remove ();
            return modified;
        }
    }
    else
        return modified;
}

function getAnchoredChar (currentObject) {
	var rootChar = null;
    do {
        if (currentObject.constructor.name == "InsertionPoint" && currentObject.parentTextFrames.length) {
            currentObject = currentObject.parentTextFrames[0];
        }
        else if (currentObject.constructor.name == "Table") {
            currentObject = currentObject.storyOffset.parent.characters.item(currentObject.storyOffset.index);
        }
        else {
            try {
                currentObject = currentObject.parent;
            }
            catch (e) {
                break;
            }
        }
        if (!currentObject) {
            break;
        }
        else if (currentObject.constructor.name == "Character") {
            rootChar = currentObject;
        }
    } while (currentObject.constructor.name != "Spread" && currentObject.constructor.name != "MasterSpread" && currentObject.constructor.name != "Page" && currentObject.constructor.name != "Document");
	return rootChar;
}

function testLink (targetLink, PathIDPathArray, isAll, bridgeTalkLinesArr) { //soso
    /*function_start*///alert ("testLink-start");
    if (!isAll) {
        if (targetLink.status != LinkStatus.LINK_MISSING && targetLink.status != LinkStatus.LINK_OUT_OF_DATE) {
            if (targetLink.extractLabel ("TSID") != "")
                return false;
        }
    }
    var linkFile = new File (preparePath (targetLink.filePath));
    for (var p = 0; p < PathIDPathArray.length; p++) {
        if (PathIDPathArray[p][0] == linkFile.fsName.replace(/\\/g, "/")) {
            if (PathIDPathArray[p][1]) {
                if (PathIDPathArray[p][1] == "Reset")
                    targetLink.insertLabel ("TSID", "");
                else
                    targetLink.insertLabel ("TSID", PathIDPathArray[p][1]);
            }
            if (PathIDPathArray[p][2]) {
                if (PathIDPathArray[p][2] == "out")
                    return "out";
                try {
                    targetLink.relink (File (PathIDPathArray[p][2]));
                }
                catch (e) {return false;}   
                return true;
            }
            if (PathIDPathArray[p].length > 3) {
                targetLink.update ();
                targetLink.insertLabel ("Updated", (new Date ().getTime ()).toString ());
            }
            return false;
        }
    }
    var fileID = targetLink.extractLabel ("TSID");
    /*fixing links*///$.writeln ("fileID " + fileID);
    var speedElement = new Array;
    speedElement.push (linkFile.fsName.replace(/\\/g, "/"));
    /*fixing links*///$.writeln ("linkFile.fsName.replace(/\\/g, "/") " + linkFile.fsName.replace(/\\/g, "/"));
    if ((linkFile.fsName.replace(/\\/g, "/")).indexOf (workshopPath + "/") != 0) { //The Link file out of Tree Shade
        if (linkFile.exists) {
            if (app.extractLabel("autoImportOutOfTreeLinks") == "Yes") {
                var targetBridge = BridgeTalk.getSpecifier ("bridge");
                if (targetBridge && targetBridge.appStatus != "not installed") {
                    var docPath = null;
                    var parentDoc = targetLink.parent;
                    while (!(parentDoc instanceof Document)) {
                        parentDoc = parentDoc.parent;
                    }
                    docPath = preparePath (parentDoc.filePath.fsName.replace(/\\/g, "/"));
                    if (bridgeTalkLinesArr instanceof Array) {
                        bridgeTalkLinesArr.push ("theTotalQuery.push (tsImportFile (new File (\"" + linkFile.fsName.replace(/\\/g, "/") + "\"), new Folder (\"" + docPath + "/" + app.extractLabel("autoImportOutOfTreeLinksFolder") + "\"), false, \"" + targetLink.id + "@" + parentDoc.id + "\", true, 6).join (\",\"));");
                    }
                    else {
                        var talkBridge = new BridgeTalk;
                        talkBridge.target = targetBridge;
                        talkBridge.onResult = function (returnBtObj) { setBridgeScriptInTestLink (returnBtObj.body); }
                        talkBridge.body = "tsImportFile (new File (\"" + linkFile.fsName.replace(/\\/g, "/") + "\"), new Folder (\"" + docPath + "/" + app.extractLabel("autoImportOutOfTreeLinksFolder") + "\"), false, \"" + targetLink.id + "@" + parentDoc.id + "\", true, 6);";
                        talkBridge.send ();
                        function setBridgeScriptInTestLink (theBody) {
                            if (theBody) {
                                var importingData = theBody.split (",");
                                if (importingData.length == 3) {
                                    var importedPath = importingData[0];
                                    var importedTSID = importingData[1];
                                    var nativeLinkAndDocIDs = importingData[2].split ("@");
                                    if (nativeLinkAndDocIDs.length == 2) {
                                        var nativeLinkID = nativeLinkAndDocIDs[0];
                                        var targetDocNativeID = nativeLinkAndDocIDs[1];
                                        var importedLinkTargetDoc = null;
                                        for (var ad = 0; ad < app.documents.length; ad++) {
                                            if (targetDocNativeID == app.documents[ad].id) {
                                                importedLinkTargetDoc = app.documents[ad];
                                                break;
                                            }
                                        }
                                        if (importedLinkTargetDoc) {
                                            var targetLink = importedLinkTargetDoc.links.itemByID (parseInt (nativeLinkID, 10));
                                            if (targetLink) {
                                                targetLink.relink (new File (workshopPath + importedPath));
                                                targetLink.insertLabel ("TSID", importedTSID);
                                                targetLink.insertLabel ("Updated", (new Date ().getTime ()).toString ());
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    return false;
                }
            }
        }
        else {
            var outTraceFile = new File (dataPath + "/Traces/Outside" + linkFile.fsName.replace(/\\/g, "/"));
            var outTraceContent = readFile (outTraceFile);
            /*fixing links*///$.writeln ("line " + $.line);
            if (outTraceContent) {
                var outTraceSplitted = outTraceContent.split (":");
                if (outTraceSplitted.length > 1) {
                    if (outTraceSplitted[0] == "TS_ID") {
                        targetLink.insertLabel ("TSID", outTraceSplitted[1]);
                        fileID = outTraceSplitted[1];
                    }
                }
            }
            if (!fileID) {
                return false;
            }
        }
    }
    /*fixing links*///$.writeln ("line " + $.line);
    if (targetLink.status == LinkStatus.LINK_MISSING && fileID == "") {
        /*fixing links*///$.writeln ("line " + $.line);
        var traceFile = new File (dataPath + "/Traces/Inside" + linkFile.fsName.replace(/\\/g, "/").replace (workshopPath, ""));
        var traceContent = readFile (traceFile);
        if (traceContent) {
            var traceSplitted = traceContent.split (":");
            if (traceSplitted.length > 1) {
                if (traceSplitted[0] == "TS_ID") {
                    targetLink.insertLabel ("TSID", traceSplitted[1]);
                    fileID = traceSplitted[1];
                }
            }
        }
    }
    /*fixing links*///$.writeln ("line " + $.line);
    if (fileID == "") {
        /*fixing links*///$.writeln ("line " + $.line);
        if (targetLink.status != LinkStatus.LINK_MISSING) {
            var gainedID = prepareLink (targetLink);
            if (gainedID) {
                /*fixing links*///$.writeln ("line " + $.line);
                speedElement.push (gainedID);
                speedElement.push (null);
                if (targetLink.status == LinkStatus.LINK_OUT_OF_DATE) {
                    /*fixing links*///$.writeln ("line " + $.line);
                    try {
                        targetLink.update ();
                    }
                    catch (r) {}
                    /*fixing links*///$.writeln ("line " + $.line);
                    speedElement.push ("GetNewID");
                    /*fixing links*///$.writeln ("line " + $.line);
                    targetLink.insertLabel ("Updated", (new Date ().getTime ()).toString ());
                    /*fixing links*///$.writeln ("line " + $.line);
                }
                PathIDPathArray.push (speedElement);
                /*fixing links*///$.writeln ("line " + $.line);
            }
            else {
                /*fixing links*///$.writeln ("line " + $.line);
                speedElement.push (null);
                speedElement.push (null);
                PathIDPathArray.push (speedElement);
                return false;
            }
        }
        else {

        }
    }
    else {
        var actualFileID;
        /*fixing links*///$.writeln ("line " + $.line);
        if (targetLink.status != LinkStatus.LINK_MISSING) {
             actualFileID = getTSID (linkFile);
             if (!actualFileID) {
                addToBridgeScanningQueue (linkFile.parent.fsName.replace(/\\/g, "/"));
             }
             var isToUpdate = false;
            if (targetLink.status == LinkStatus.LINK_OUT_OF_DATE) {
                var ggg = getActualModified (linkFile.modified.getTime (), actualFileID);
                var oldDate = targetLink.extractLabel ("Updated");
                if (oldDate == "") {
                    isToUpdate = true;
                }
                else {
                    if (parseInt (oldDate, 10) >= ggg) {
                        isToUpdate = true;
                    }
                }
            }
            if (actualFileID == fileID) {
                speedElement.push (null);
                speedElement.push (null);
                if (isToUpdate) {
                    targetLink.update ();
                    speedElement.push ("FakeModifing");
                    targetLink.insertLabel ("Updated", (new Date ().getTime ()).toString ());
                }
                PathIDPathArray.push (speedElement);
                return false;
            }
            else if (actualFileID) { //change ID
                targetLink.insertLabel ("TSID", actualFileID);
                speedElement.push (actualFileID);
                speedElement.push (null);
                if (isToUpdate) {
                    targetLink.update ();
                    speedElement.push ("FakeModifing");
                    targetLink.insertLabel ("Updated", (new Date ().getTime ()).toString ());
                }
                PathIDPathArray.push (speedElement);
                return true;
            }
        }
        /*fixing links*///$.writeln ("line " + $.line);
        var fileID_Path_newID = [fileID];
        getPath (fileID_Path_newID);
        /*fixing links*///$.writeln ("fileID_Path_newID[1] " + fileID_Path_newID[1]);
        if (!fileID_Path_newID[1]) {
            speedElement.push (null);
            speedElement.push (null);
            PathIDPathArray.push (speedElement);
            return false;
        }
        var newLinkFile = new File (workshopPath + fileID_Path_newID[1]);
        /*fixing links*///$.writeln ("newLinkFile " + (workshopPath + fileID_Path_newID[1]));
        if (!newLinkFile.exists) {
            if (fileID_Path_newID[2])
                speedElement.push (fileID_Path_newID[2]);
            else
                speedElement.push (null);
            speedElement.push (null);
            PathIDPathArray.push (speedElement);
            return false;
        }
        var testFileID = getTSID (newLinkFile);
        if (!testFileID) {
            addToBridgeScanningQueue (newLinkFile.parent.fsName.replace(/\\/g, "/"));
        }
        /*fixing links*///$.writeln ("testFileID " + testFileID);
        /*fixing links*///$.writeln ("fileID " + fileID);
        if (fileID_Path_newID[2])
            fileID = fileID_Path_newID[2];
        /*fixing links*///$.writeln ("fileID_Path_newID[2] " + fileID_Path_newID[2]);
        if (testFileID != fileID) {
            speedElement.push ("Reset");
            speedElement.push (null);
            PathIDPathArray.push (speedElement);
            return false;
        }
        try {
            /*fixing links*///$.writeln ("\nrelinked to " + newLinkFile.fsName.replace(/\\/g, "/"));
            if (!fileID_Path_newID[2]) {
                speedElement.push (null);
            }
            else {
                targetLink.insertLabel ("TSID", fileID);
                speedElement.push (fileID_Path_newID[2]);
            }
            if (isShade (File (workshopPath + fileID_Path_newID[1]))) {
                speedElement.push (null);
            }
            else {
                speedElement.push (workshopPath + fileID_Path_newID[1]);
                targetLink.relink (newLinkFile);
            }
            PathIDPathArray.push (speedElement);
            return true;
        }
        catch (e) {
            speedElement.push (null);
            speedElement.push (null);
            PathIDPathArray.push (speedElement);
            return false;
        }
    }
    /*fixing links*///$.writeln ("line " + $.line);
    return true;
}

function linksIdleTaskHandler (myEvent) {
    /*function_start*///alert ("linksIdleTaskHandler-start");
    if (workshopPath != app.extractLabel ("Tree Shade Path for " + tsClientID) + "/Workshop") {
        obtainRootSettings ();
    }
    var bridgeTalkLinesArr = new Array;
    try {
        app.activeDocument.filePath;
    }
    catch (e) {
        return false;
    }
    var docPath = preparePath (app.activeDocument.filePath.fsName.replace(/\\/g, "/")) + "/";
    if (docPath.indexOf (workshopPath + "/") == 0 && workshopPath != "") {
        var allLinks = app.activeDocument.links;
        var speedChain = new Array;
        for (var c = 0; c < allLinks.length; c++) {
            testLink (allLinks[c], speedChain, false, bridgeTalkLinesArr);
        }
    }

    //live snippets auto check
    if (app.extractLabel ("LS_AUTO_CHECK") == "yes") {
        checkAllLiveSnippets (app.activeDocument, (app.extractLabel ("LS_AUTO_SUBMIT") == "yes"), (app.extractLabel ("LS_AUTO_UPDATE") == "yes"));
    }

    //bridgeScanning
    var isToGetReturn = (bridgeTalkLinesArr.length > 0);
    if (bridgeScanningStep++ > 3) {
        bridgeScanningStep = 0;
        if (bridgeScanningQueue.length > 0) {
            bridgeTalkLinesArr.push ("tsAddToScanningQueue ([\"" + bridgeScanningQueue.join ("\", \"") + "\"]);");
            bridgeScanningQueue = new Array;
        }        
    }
    if (bridgeTalkLinesArr.length > 0) {
        if (isToGetReturn) {
            sendImportListToInDesign (bridgeTalkLinesArr);
        }
        else {
            var targetBridge = BridgeTalk.getSpecifier ("bridge");
            if (targetBridge && targetBridge.appStatus != "not installed") {
                var talkBridge = new BridgeTalk;
                talkBridge.target = targetBridge;
                talkBridge.body = bridgeTalkLinesArr.join ("\n");
                talkBridge.send ();     
            }
        }
    }
}

function sendImportListToInDesign (bridgeTalkLinesArr) {
    bridgeTalkLinesArr.unshift ("var theTotalQuery = new Array;");
    bridgeTalkLinesArr.push ("theTotalQuery.join (\":\");");
    var targetBridge = BridgeTalk.getSpecifier ("bridge");
    if (targetBridge && targetBridge.appStatus != "not installed") {
        var talkBridge = new BridgeTalk;
        talkBridge.target = targetBridge;
        talkBridge.onResult = function (returnBtObj) { setBridgeScriptList (returnBtObj.body); }
        talkBridge.body = bridgeTalkLinesArr.join ("\n");
        talkBridge.send ();
        function setBridgeScriptList (theBody) {
            if (theBody) {
                var bodyList = theBody.split (":");
                for (var bl = 0; bl < bodyList.length; bl++) {
                    theBody = bodyList[bl];
                    var importingData = theBody.split (",");
                    if (importingData.length == 3) {
                        var importedPath = importingData[0];
                        var importedTSID = importingData[1];
                        var nativeLinkAndDocIDs = importingData[2].split ("@");
                        if (nativeLinkAndDocIDs.length == 2) {
                            var nativeLinkID = nativeLinkAndDocIDs[0];
                            var targetDocNativeID = nativeLinkAndDocIDs[1];
                            var importedLinkTargetDoc = null;
                            for (var ad = 0; ad < app.documents.length; ad++) {
                                if (targetDocNativeID == app.documents[ad].id) {
                                    importedLinkTargetDoc = app.documents[ad];
                                    break;
                                }
                            }
                            if (importedLinkTargetDoc) {
                                var targetLink = importedLinkTargetDoc.links.itemByID (parseInt (nativeLinkID, 10));
                                if (targetLink) {
                                    targetLink.relink (new File (workshopPath + importedPath));
                                    targetLink.insertLabel ("TSID", importedTSID);
                                    targetLink.insertLabel ("Updated", (new Date ().getTime ()).toString ());
                                }
                            }
                        }
                    }
                }
            }
        }     
    }
}

function getPath (fileID_Path_newID) {
    fileID_Path_newID.push (null);
    fileID_Path_newID.push (null);
    var pathRecordFolder;
    var isDirected = false;
    var isFailed = false;
    var newID = fileID_Path_newID[0];
    var pathTimeFile;
    var pathTime;
    do {
        pathTimeFile = new File (dataPath + "/IDs" + newID + "/Path/Change Time");
        pathTime = readFile (pathTimeFile);
        if (!pathTime) {
            isFailed = true;
            fileID_Path_newID[1] = null;
            if (fileID_Path_newID[0] != newID)
                fileID_Path_newID[2] = newID;
        }
        else {
            pathRecordFile = new File (dataPath + "/IDs" + newID + "/Path/V/" + pathTime + "/Value");
            if (pathRecordFile.exists) {
                fileID_Path_newID[1] = readEncodedFile (pathRecordFile);
            }
            if (fileID_Path_newID[1]) {
                if (fileID_Path_newID[1][0] ==':') {
                    isDirected = true;
                    newID = fileID_Path_newID[1].slice (1);
                }
                else {
                    isDirected = false;
                    if (fileID_Path_newID[0] != newID)
                        fileID_Path_newID[2] = newID;
                }
            }
            else {
                isFailed = true;
                fileID_Path_newID[1] = null;
                if (fileID_Path_newID[0] != newID)
                    fileID_Path_newID[2] = newID;
            }
        }
    } while (isDirected && !isFailed);
    return pathTime;
}

function getPDFPreviewLevel (fileName) {
    var match = fileName.match(/\.nested(\d+)$/);
    return match ? parseInt(match[1], 10) : 0;
}

function filterPDFPreviewFiles(files) {
    // Function to extract level from the file extension (must be at the end)
    // Find the maximum level
    var maxLevel = 0;
    for (var i = 0; i < files.length; i++) {
        var level = getPDFPreviewLevel (files[i].name);
        if (level > maxLevel) {
            maxLevel = level;
        }
    }
    // Filter files to keep only those with the max level
    var filteredFiles = [];
    for (var j = 0; j < files.length; j++) {
        if (getPDFPreviewLevel (files[j].name) === maxLevel) {
            filteredFiles.push(files[j]);
        }
    }
    return filteredFiles;
}


function checkNestedMissingLinks(doc, currentLevel) {
    var links = doc.links;
    var isAllUpdated = true;
    var newLevel = currentLevel + 1;
    for (var i = 0; i < links.length; i++) {
        if (links[i].links.length > 0) {
            if (links[i].linkType == "InDesign Format Name") {
                for (var ii = 0; ii < links[i].links.length; ii++) {
                    if (links[i].links[ii].status == LinkStatus.LINK_MISSING || links[i].links[ii].status == LinkStatus.LINK_OUT_OF_DATE) {
                        var linkFile = new File (preparePath (links[i].filePath));
                        if (linkFile.exists) {
                            fileID = getTSID (linkFile);
                            if (fileID) {
                                var digitsDocIDTriple = getWorkshopVersionInfo (linkFile, fileID);
                                var currentVersion = digitsDocIDTriple[0];
                                var fileDotID = fileID.replace (/\//g,"\.");
                                fileDotID = fileDotID.slice (1);
                                alertFile = new File (dataPath + "/Messages/PDFPreview/" + "P." + fileDotID + "." + currentVersion + ".nested" + newLevel);
                                if (!alertFile.parent.exists)
                                    alertFile.parent.create ();
                                writeFile (alertFile, new Date().getTime());
                                isAllUpdated = false;
                            }
                        }
                        break;
                    }
                }
            }
        }
    }
    return isAllUpdated;
}

function getWorkshopVersionInfo (targetFile, fileID) {
    /**///$.writeln ($.line);
    var versionDigits;
    var versionFile;
    if (!fileID)
        fileID = getTSID (targetFile);
    if (fileID) {
        var currentVersionFile = new File (dataPath + "/IDs" + fileID + "/CV");
        var recordContent = readFile (currentVersionFile);
        if (recordContent) {
            versionDigits = recordContent.split (":")[0];
            versionFile = new File (targetFile.parent.fsName.replace (workshopPath, versionsPath) + "/ver" + versionDigits + " " + File.decode (targetFile.name));
            var digitsFileIDTriple = new Array;
            digitsFileIDTriple.push (versionDigits);
            digitsFileIDTriple.push (versionFile);
            digitsFileIDTriple.push (fileID);
            return digitsFileIDTriple;
        }
        else {
            versionDigits = "01";
            versionFile = new File (targetFile.parent.fsName.replace (workshopPath, versionsPath) + "/ver" + versionDigits + " " + File.decode (targetFile.name));
            var digitsFileIDTriple = new Array;
            digitsFileIDTriple.push ("new");
            digitsFileIDTriple.push (versionFile);
            digitsFileIDTriple.push (fileID);
            return digitsFileIDTriple;
        }
    }
    return null;
}

function tsPlaceSelectionAndFileAndDocHandler (myEvent) {
    if (tsPlaceSelectionAndFileAndDoc) {
        if (tsPlaceSelectionAndFileAndDoc[0] && tsPlaceSelectionAndFileAndDoc[1] && tsPlaceSelectionAndFileAndDoc[2]) {
            var thisPathAndName = getPathAndName (tsPlaceSelectionAndFileAndDoc[2], tsPlaceSelectionAndFileAndDoc[1]);
            tsPlaceLiveSnippet (tsPlaceSelectionAndFileAndDoc[0], null, thisPathAndName, tsPlaceSelectionAndFileAndDoc[2], tsPlaceSelectionAndFileAndDoc[1], null, false, true, [], [], -1, -1, []);
        }
        tsPlaceSelectionAndFileAndDoc = null;
    }
}

function savedDocIdleTaskHandler (myEvent) {
    /*function_start*///alert ("savedDocIdleTaskHandler-start");
    var dailySheetsFiles = Folder (dataPath + "/Messages/DailySheets").getFiles (isUnhiddenFile);
    if (dailySheetsFiles.length > 0) {
        var infoContent = readEncodedFile (dailySheetsFiles[0]);
        var scopeID = File.decode (dailySheetsFiles[0].name).slice (0, File.decode (dailySheetsFiles[0].name).indexOf ("."));
        var splitted = infoContent.split ("\n");
        var senderName = splitted[0];
        var producingPath = splitted[1];
        splitted.splice (0, 2);
        if (splitted[0] == "Work Time") {
            var fromTime = splitted[1];
            var toTime = splitted[2];
            if (splitted.length > 3) {
                var InDesignList = new Array;
                var timeList = new Array;
                for (var wt = 3; wt < splitted.length; wt++) {
                    var span = splitted[wt].split (":");
                    var spanItem = new Array;
                    spanItem.push (span[0]);
                    for (var ind = 1; ind < span.length; ind++) {
                        var spanSplitted = span[ind].split (",");
                        var InDesignID = spanSplitted[0];
                        var InDesignMod = spanSplitted[1];
                        var activityTock = spanSplitted[2];
                        var existIndex = -1;
                        for (var st = 0; st < InDesignList.length; st++) {
                            if (InDesignList[st][0] == InDesignID) {
                                existIndex = st;
                                if (InDesignID != "No Document") {
                                    InDesignList[st][1] += 5;
                                    if (parseInt (activityTock, 10) >= 5)
                                        InDesignList[st][2] += 5;
                                }
                                break;
                            }
                        }
                        var cellItem = new Array;
                        if (existIndex != -1) {
                            cellItem.push (existIndex);
                            cellItem.push (InDesignMod);
                            cellItem.push (spanSplitted[2]);
                            spanItem.push (cellItem);
                        }
                        else {
                            var InDesignItem = new Array;
                            InDesignItem.push (InDesignID);
                            if (InDesignID != "No Document") {
                                InDesignItem.push (1*5);
                                if (parseInt (activityTock, 10) >= 5)
                                    InDesignItem.push (5);
                                else
                                    InDesignItem.push (0);
                            }
                            else {
                                InDesignItem.push (0);
                                InDesignItem.push (0);
                            }
                            InDesignList.push (InDesignItem);
                            cellItem.push (InDesignList.length - 1);
                            cellItem.push (InDesignMod);
                            cellItem.push (activityTock);
                            spanItem.push (cellItem);
                        }
                        
                    }
                    timeList.push (spanItem);
                }
                var bodyMatrix = new Array;
                
                for (var bm = 0; bm < timeList.length + 3; bm++) {
                    var bodyRow = new Array;
                    for (var vv = 0; vv < InDesignList.length + 1; vv++) {
                        bodyRow.push ("");
                    }
                    bodyMatrix.push (bodyRow);
                }
                var total = 0;
                var pure = 0;
                var pathsArray = new Array;
                for (var bo = 0; bo < InDesignList.length; bo++) {
                    pathsArray.push (""); 
                    var InDesignName = InDesignList[bo][0];
                    if (InDesignName != "No Document") {
                        var fileID_Path_newID = [InDesignName];
                        getPath (fileID_Path_newID);
                        if (fileID_Path_newID[1]) {
                            pathsArray[pathsArray.length - 1] = fileID_Path_newID[1];
                            InDesignName = fileID_Path_newID[1].slice (fileID_Path_newID[1].lastIndexOf ("/") + 1);
                        }
                        else {
                            InDesignList[bo][0] = "Unknown";
                        }
                    }
                    bodyMatrix[0][bo + 1] = InDesignName;
                    total += InDesignList[bo][1];
                    pure += InDesignList[bo][2];
                    var totalHours = Math.floor (InDesignList[bo][1] / 60);
                    var totalMinutes = Math.round (InDesignList[bo][1] % 60);
                    var pureHours = Math.floor (InDesignList[bo][2] / 60);
                    var pureMinutes = Math.round (InDesignList[bo][2] % 60);      
                    if (totalHours < 10)
                        totalHours = "0" + totalHours;
                    if (pureHours < 10)
                        pureHours = "0" + pureHours;
                    bodyMatrix[1][bo + 1] = totalHours + ":" + totalMinutes;
                    bodyMatrix[2][bo + 1] = pureHours + ":" + pureMinutes;
                }
                var totalH = Math.floor (total / 60);
                var totalM = Math.round (total % 60);
                var pureH = Math.floor (pure / 60);
                var pureM = Math.round (pure % 60);
                
                totalM = (totalM > 9)? totalM.toString () : "0" + totalM; 
                bodyMatrix[1][0] = "Total Hours: " + totalH + ":" + totalM;
                bodyMatrix[2][0] = "Pure Hours: " + pureH + ":" + pureM;
                for (var tl = 0; tl < timeList.length; tl++) {
                    bodyMatrix[tl + 3][0] = timeList[tl][0];
                    for (var qt = 1; qt < timeList[tl].length; qt++) {
                        bodyMatrix[tl + 3][timeList[tl][qt][0] + 1] = timeList[tl].length - 1 + "-" + timeList[tl][qt][1] + "-" + timeList[tl][qt][2];
                    }
                }

                var is2From = true;
                var timeBeforeChange = "";
                if (bodyMatrix.length > 2)
                    timeBeforeChange = bodyMatrix[3][0];
                for (var ft = 3; ft < bodyMatrix.length - 1; ft++) {
                    textSplitted1 = timeBeforeChange.split ("/");
                    timeBeforeChange = bodyMatrix[ft + 1][0];
                    textSplitted2 = bodyMatrix[ft + 1][0].split ("/");
                    var isPreviousFrom = is2From;
                    var date1 = new Date (parseInt (textSplitted1[0], 10), 
                                                        parseInt (textSplitted1[1], 10), 
                                                        parseInt (textSplitted1[2], 10), 
                                                        parseInt (textSplitted1[3], 10), 
                                                        parseInt (textSplitted1[4], 10), 0, 0);
                    var date2 = new Date (parseInt (textSplitted2[0], 10), 
                                                        parseInt (textSplitted2[1], 10), 
                                                        parseInt (textSplitted2[2], 10), 
                                                        parseInt (textSplitted2[3], 10), 
                                                        parseInt (textSplitted2[4], 10), 0, 0);
                    if (date1.getTime () + 300000 == date2.getTime ()) { //1: Not At, Not To, May From, May In    //2: Not At, Not From, May To, May In
                        if (is2From) { //1: From 
                            bodyMatrix[ft][0] = "From " + textSplitted1[3] + ":" + textSplitted1[4];
                        }
                        else { //1: In 
                            bodyMatrix[ft][0] = textSplitted1[3] + ":" + textSplitted1[4];
                        }
                        if (ft == bodyMatrix.length - 2) { //2: To
                            bodyMatrix[ft + 1][0] = "To " + textSplitted2[3] + ":" + textSplitted2[4];
                        }
                        is2From = false;
                    }
                    else { //1: May At, May To    //2: May At, May From, Not To, Not In
                        if (is2From) { //1: At
                            bodyMatrix[ft][0] = "At " + textSplitted1[3] + ":" + textSplitted1[4];
                        }
                        else {  //1: To
                            bodyMatrix[ft][0] = "To " + textSplitted1[3] + ":" + textSplitted1[4];
                        }
                        if (ft == bodyMatrix.length - 2) { //2: To
                            bodyMatrix[ft + 1][0] = "At " + textSplitted2[3] + ":" + textSplitted2[4];
                        }
                        is2From = true;
                    }
                }
                for (var fg = 0; fg < bodyMatrix.length; fg++) {
                    bodyMatrix[fg] = bodyMatrix[fg].join ("\t");
                }
                
                bodyMatrix.splice (2, 1);
                bodyMatrix = bodyMatrix.join ("*");
                var headLines = "Work Time Sheet for: " + senderName + "*";
                headLines += "From: " + fromTime.slice (0, 10) + " - " + fromTime.slice (11).replace ("/", ":") + "*";
                headLines += "To: " + toTime.slice (0, 10) + " - " + toTime.slice (11).replace ("/", ":") + "$";
                bodyMatrix = headLines + bodyMatrix;
                var sheetDoc = app.documents.add (false);
                sheetDoc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
                sheetDoc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;
                sheetDoc.documentPreferences.pageHeight = 300;
                sheetDoc.documentPreferences.pageWidth = 100;
                sheetDoc.documentPreferences.facingPages = false;
                var page = sheetDoc.pages.item(0);
                var frame = page.textFrames.add();
                var firstFrame = frame;
                var width = sheetDoc.documentPreferences.pageWidth;
                var height = sheetDoc.documentPreferences.pageHeight;
                var x1 = (page.side == PageSideOptions.leftHand)? page.marginPreferences.right : page.marginPreferences.left;
                var x2 = (page.side == PageSideOptions.leftHand)? page.marginPreferences.left : page.marginPreferences.right;
                x2 = width - x2;
                frame.geometricBounds = [page.marginPreferences.top, x1, height - page.marginPreferences.bottom, x2];
                frame.contents = bodyMatrix;
                
                //frame.contents = frame.contents.replace (/*/g, String.fromCharCode(13));
                var myStory = sheetDoc.stories.item(0);
                myStory.digitsType = DigitsTypeOptions.ARABIC_DIGITS;
                myStory.characterDirection = CharacterDirectionOptions.LEFT_TO_RIGHT_DIRECTION;
                var myStartCharacter1 = myStory.paragraphs.item(0).characters.item(0);
                var myEndCharacter1 = myStory.paragraphs.item(-1).characters.item(frame.contents.indexOf ("$"));
                var myText1 = myStory.texts.itemByRange(myStartCharacter1, myEndCharacter1);
                var myStartCharacter = myStory.paragraphs.item(0).characters.item(frame.contents.indexOf ("$") + 1);
                var myEndCharacter = myStory.paragraphs.item(-1).characters.item(-1);
                var myText = myStory.texts.itemByRange(myStartCharacter, myEndCharacter);
                myText.convertToTable("\t","*");
                myText1.convertToTable("\t","*");
                var myTable1 = frame.tables.item (0);
                myTable1.rows.item (2).cells.item (0).paragraphs.item (0).characters.item (-1).remove ();
                myTable1.bottomBorderStrokeWeight = 0;
                myTable1.leftBorderStrokeWeight = 0;
                myTable1.rightBorderStrokeWeight = 0;
                myTable1.topBorderStrokeWeight = 0;
                for (var e = 0; e < myTable1.rows.length; e++) {
                    myTable1.rows.item (e).topEdgeStrokeWeight = "0.25pt";
                    for (var rc = 0; rc < myTable1.rows.item (e).cells.length; rc++) {
                        for (var rcp = 0; rcp < myTable1.rows.item (e).cells.item (rc).paragraphs.length; rcp++) {
                            myTable1.rows.item (e).cells.item (rc).paragraphs.item (rcp).pointSize = "14pt";
                        }
                    }
                }
                var myTable = frame.tables.item (1);
                for (var ac = myTable.columns.length - 1; ac >= 0; ac--) {
                    myTable.columns.item (ac).width = 100;
                }
                myTable.rows.item (0).rotationAngle = 90;
                myTable.rows.item (0).rowType = RowTypes.HEADER_ROW;
                myTable.rows.item (0).height = 200;
                myTable.rows.item (0).verticalJustification = VerticalJustification.CENTER_ALIGN;
                var maxHeight = 0;
                for (var cc = 1; cc < myTable.rows.item (0).cells.length; cc++) {
                    if (myTable.rows.item (0).cells[cc].paragraphs.length > 0) {
                        myTable.rows.item (0).cells[cc].paragraphs[0].pointSize = "14pt";
                        var thisObject = myTable.rows.item (0).cells[cc].paragraphs[0].createOutlines (false);
                        var thisHeight = thisObject[0].geometricBounds[2] - thisObject[0].geometricBounds[0];
                        if (thisHeight > maxHeight) {
                            maxHeight = thisHeight;
                        }
                        thisObject[0].remove ();
                    }
                }
                myTable.rows.item (0).height = maxHeight + myTable.rows.item (0).cells.item (1).bottomInset + myTable.rows.item (0).cells.item (1).topInset + 1;
                while (frame.overflows) {
                    var oldFrame = frame;
                    page = sheetDoc.pages.add ();
                    frame = page.textFrames.add();
                    width = sheetDoc.documentPreferences.pageWidth;
                    height = sheetDoc.documentPreferences.pageHeight;
                    x1 = (page.side == PageSideOptions.leftHand)? page.marginPreferences.right : page.marginPreferences.left;
                    x2 = (page.side == PageSideOptions.leftHand)? page.marginPreferences.left : page.marginPreferences.right;
                    x2 = width - x2;
                    frame.geometricBounds = [page.marginPreferences.top, x1, height - page.marginPreferences.bottom, x2];
                    oldFrame.nextTextFrame = frame;
                }
                var mergeIndex = 2;
                var toTimeDigits = myTable.rows.item (myTable.rows.length-1).cells.item (0).contents.slice (-5);
                myTable.rows.item (myTable.rows.length-1).remove ();
                for (var t = 2; t < myTable.rows.length; t++) {
                    if (myTable.rows.item (t).cells.item (0).contents.length < 6 || myTable.rows.item (t).cells.item (0).contents[0] == 'T') {
                        myTable.rows.item (t).topEdgeStrokeWeight = "0.25pt";
                    }
                    for (var c = 1; c < myTable.rows.item (t).cells.length; c++) {
                        myTable.rows.item (t).cells.item (c).topInset = "0.2mm";
                        myTable.rows.item (t).cells.item (c).bottomInset = "0.5mm";
                        if (myTable.rows.item (0).cells.item (c).contents == "No Document")
                            myTable.rows.item (t).cells.item (c).contents = "";
                        var cellValue = myTable.rows.item (t).cells.item (c).contents;
                        if (cellValue != "") {
                            var splitted = cellValue.split ("-");
                            myTable.rows.item (t).cells.item (c).contents = "";
                            myTable.rows.item (t).cells.item (c).fillColor = sheetDoc.colors.item("Black");
                            //"myTable.rows.item (t).cells.item (c).fillTint = ((1/parseInt (splitted[0], 10)) * 100)/2;
                            myTable.rows.item (t).cells.item (c).fillTint = 40;
                            if (splitted.length > 1) {
                                if (splitted[1] == "undefined") {
                                    splitted[1] = "";
                                }
                                if (splitted.length > 2) {
                                    if (parseInt (splitted[2], 10) < 5) {
                                        myTable.rows.item (t).cells.item (c).fillTint = 15;
                                    }
                                }
                                myTable.rows.item (t).cells.item (c).contents = splitted[1];
                            }
                        }
                        if (myTable.rows.item (t).cells.item (c).paragraphs.length > 0) {
                            myTable.rows.item (t).cells.item (c).paragraphs.item (0).pointSize = "6pt";
                            myTable.rows.item (t).cells.item (c).paragraphs.item (0).justification = Justification.CENTER_ALIGN;
                        }
                        else 
                            myTable.rows.item (t).cells.item (c).insertionPoints.item (0).pointSize = "6pt";
                    }
                    var isToMerge = false;
                    if (mergeIndex != t) {
                        if (myTable.rows.item (t).cells.item (0).paragraphs.item (0).contents.slice (-2) == "00") {
                            mergeIndex = t;
                        }
                        else {
                            isToMerge = true;
                        }
                    }
                    if (isToMerge) {
                        if (t == myTable.rows.length - 1) {
                            if (toTimeDigits.slice (-2) != "00")
                                myTable.rows.item (mergeIndex).cells.item (0).contents = "To " + toTimeDigits;
                            else 
                                myTable.rows.item (mergeIndex).cells.item (0).contents = myTable.rows.item (mergeIndex).cells.item (0).contents.slice (0, 2) + ":00";
                        }
                        myTable.rows.item (t).cells.item (0).contents = "";
                        myTable.rows.item (mergeIndex).cells.item (0).merge (myTable.rows.item (t).cells.item (0));
                    }
                }
                
                var totalWidth = 0;
                var isThereWork = (myTable.columns.length > 2? true : false);
                for (var fc = 0; fc < myTable.columns.length; fc++) {
                    if (isThereWork && myTable.columns.item (fc).cells.item (0).contents == "No Document") {
                        myTable.columns.item (fc).remove ();
                        fc--;
                        continue;
                    }
                    var maxWidth = 0;
                    for (var ar = 1; ar < myTable.columns.item (fc).cells.length; ar ++) {
                        if (myTable.columns.item (fc).cells.item (ar).paragraphs.length > 0) {
                            var theObject = myTable.columns.item (fc).cells.item (ar).paragraphs[0].createOutlines (false);
                            var thisWidth = theObject[0].geometricBounds[3] - theObject[0].geometricBounds[1];
                            if (maxWidth < thisWidth) {
                                maxWidth = thisWidth;
                            }
                            theObject[0].remove ();
                        }
                    }
                    myTable.columns.item (fc).width = maxWidth + myTable.columns.item (fc).cells.item (1).leftInset + myTable.columns.item (fc).cells.item (1).rightInset + 1;
                    totalWidth += myTable.columns.item (fc).width;
                }
                //set buttons
                for (var cc = 1; cc < myTable.rows.item (0).cells.length; cc++) {
                    if (myTable.rows.item (0).cells[cc].paragraphs.item (0).contents == "No Document")
                        continue;
                    if (myTable.rows.item (0).cells[cc].paragraphs.length > 0) {
                        var refButton = myTable.rows.item (0).cells[cc].paragraphs.item (0).buttons.add (sheetDoc.layers.firstItem (), LocationOptions.AT_BEGINNING);
                        var docDesc = "";
                        if (pathsArray[cc - 1] != "") {
                            var fullPath = workshopPath + pathsArray[cc - 1];
                            var descArray = new Array;
                            do {
                                fullPath = fullPath.slice (0, fullPath.lastIndexOf ("/"));
                                var labelFile = Folder (fullPath).getFiles ("- *");
                                descArray.push ("*");
                                if (labelFile.length > 0) {
                                    descArray [descArray.length - 1] = File.decode (labelFile[0].name).slice (2, File.decode (labelFile[0].name).lastIndexOf ("."));
                                }
                            } while (fullPath != workshopPath);
                            docDesc = descArray.join ("\n");
                        }
                        with (refButton) {
                            with (anchoredObjectSettings) {
                                anchoredPosition = AnchorPosition.ANCHORED;
                                horizontalReferencePoint = AnchoredRelativeTo.COLUMN_EDGE;
                                verticalReferencePoint = VerticallyRelativeTo.COLUMN_EDGE;
                                anchorPoint = AnchorPoint.TOP_LEFT_ANCHOR;
                                horizontalAlignment = HorizontalAlignment.LEFT_ALIGN;
                                verticalAlignment = VerticalAlignment.TOP_ALIGN;
                                anchorXoffset = myTable.rows.item (0).cells[cc].topInset;
                                anchorYoffset = - myTable.rows.item (0).cells[cc].rightInset;
                            }
                            geometricBounds = [0, 0, myTable.rows.item (0).cells[cc].height, myTable.rows.item (0).cells[cc].width];
                            description = docDesc;
                        }
                    }
                }
                
                var requiredWidth = totalWidth + page.marginPreferences.left + page.marginPreferences.right;
                if (requiredWidth < 210) {
                    sheetDoc.documentPreferences.pageWidth = 210;
                    sheetDoc.documentPreferences.pageHeight = 150;
                } else if (requiredWidth < 297) {
                    sheetDoc.documentPreferences.pageWidth = 297;
                    sheetDoc.documentPreferences.pageHeight = 150;
                } else if (requiredWidth < 420) {
                    sheetDoc.documentPreferences.pageWidth = 420;
                    sheetDoc.documentPreferences.pageHeight = 150;
                } else {
                    sheetDoc.documentPreferences.pageWidth = requiredWidth;
                    sheetDoc.documentPreferences.pageHeight = 150;
                }
                for (var tf = 0; tf < sheetDoc.textFrames.length; tf++) {
                    var thisPage = sheetDoc.textFrames.item (tf).parentPage;
                    fx1 = (thisPage.side == PageSideOptions.leftHand)? thisPage.marginPreferences.right : thisPage.marginPreferences.left;
                    fx2 = (thisPage.side == PageSideOptions.leftHand)? thisPage.marginPreferences.left : thisPage.marginPreferences.right;
                    fx2 = sheetDoc.documentPreferences.pageWidth - fx2;
                    sheetDoc.textFrames.item (tf).geometricBounds = [thisPage.marginPreferences.top, fx1, sheetDoc.documentPreferences.pageHeight - page.marginPreferences.bottom, fx2];
                }
                
                if (firstFrame.nextTextFrame) {
                    while (firstFrame.nextTextFrame.texts.length != 0) {
                        sheetDoc.documentPreferences.pageHeight += 20;
                        for (var tf = 0; tf < sheetDoc.textFrames.length; tf++) {
                            var thisPage = sheetDoc.textFrames.item (tf).parentPage;
                            fx1 = (thisPage.side == PageSideOptions.leftHand)? thisPage.marginPreferences.right : thisPage.marginPreferences.left;
                            fx2 = (thisPage.side == PageSideOptions.leftHand)? thisPage.marginPreferences.left : thisPage.marginPreferences.right;
                            fx2 = sheetDoc.documentPreferences.pageWidth - fx2;
                            sheetDoc.textFrames.item (tf).geometricBounds = [thisPage.marginPreferences.top, fx1, sheetDoc.documentPreferences.pageHeight - page.marginPreferences.bottom, fx2];
                        }
                    }
                }
                
                myTable1.columns.item (0).width = frame.geometricBounds[3] - frame.geometricBounds[1];
                var isExtra = false;
                for (var rf = sheetDoc.textFrames.length - 1; rf >= 0; rf--) {
                    var thisPage = sheetDoc.textFrames.item (rf).parentPage;
                    if (sheetDoc.textFrames.item (rf).texts.length == 0) {
                        sheetDoc.textFrames.item (rf).remove ();
                        thisPage.remove ();
                        isExtra = true;
                    } else {
                        break;
                    }
                }
                if (!isExtra) {
                    while (frame.overflows) {
                        var oldFrame = frame;
                        page = sheetDoc.pages.add ();
                        frame = page.textFrames.add();
                        width = sheetDoc.documentPreferences.pageWidth;
                        height = sheetDoc.documentPreferences.pageHeight;
                        x1 = (page.side == PageSideOptions.leftHand)? page.marginPreferences.right : page.marginPreferences.left;
                        x2 = (page.side == PageSideOptions.leftHand)? page.marginPreferences.left : page.marginPreferences.right;
                        x2 = width - x2;
                        frame.geometricBounds = [page.marginPreferences.top, x1, height - page.marginPreferences.bottom, x2];
                        oldFrame.nextTextFrame = frame;
                    }
                }
                //producing pdfs
                fromTime = fromTime.split ("/");
                toTime = toTime.split ("/");
                var monthName;
                switch (fromTime[1]) {
                    case "01":
                        monthName = "January";
                        break;
                    case "02":
                        monthName = "Febuary";
                        break;
                    case "03":
                        monthName = "March";
                        break;
                    case "04":
                        monthName = "April";
                        break;
                    case "05":
                        monthName = "May";
                        break;
                    case "06":
                        monthName = "June";
                        break;
                    case "07":
                        monthName = "July";
                        break;
                    case "08":
                        monthName = "August";
                        break;
                    case "09":
                        monthName = "September";
                        break;
                    case "10":
                        monthName = "October";
                        break;
                    case "11":
                        monthName = "November";
                        break;
                    case "12":
                        monthName = "December";
                        break;
                }
                var sheetDate = new Date ();
                sheetDate.setYear (parseInt (fromTime[0], 10));
                sheetDate.setMonth (parseInt (fromTime[1], 10)-1);
                sheetDate.setDate (parseInt (fromTime[2], 10));
                var dayName;
                switch (sheetDate.getDay ()) {
                    case 0:
                        dayName = "Sunday";
                        break;
                    case 1:
                        dayName = "Monday";
                        break;
                    case 2:
                        dayName = "Tuesday";
                        break;
                    case 3:
                        dayName = "Wednesday";
                        break;
                    case 4:
                        dayName = "Thursday";
                        break;
                    case 5:
                        dayName = "Friday";
                        break;
                    case 6:
                        dayName = "Saturday";
                        break;
                }
                
                var sheetPDFFile;
                if (producingPath == "MANUALLY") {
                    sheetPDFFile = new File (Folder.desktop.fsName.replace(/\\/g, "/") + "/" + senderName + "-" + fromTime.join (".") + "-" + toTime.join (".") + ".pdf");
                }
                else {
                    if (fromTime[2] == toTime[2]) {
                        sheetPDFFile = new File (producingPath + "/" + senderName + "/" + "Daily" + "/" + fromTime[0] + "/" + fromTime[1] + " " + monthName + "/" + fromTime[2] + " " + dayName + ".pdf");
                        
                        //produce total day for monthly reports
                        var dayTotalFile = new File (dataPath + "/Sheets Summary/" + scopeID + "/" + fromTime[0] + "." + fromTime[1] + "." + fromTime[2]);
                        var dayTotalLines = new Array;
                        for (var inl = 0; inl < InDesignList.length; inl++) {
                            dayTotalLines.push (InDesignList[inl].join (":"));
                        }
                        if (!dayTotalFile.parent.exists)
                            dayTotalFile.parent.create ();
                        writeFile (dayTotalFile, dayTotalLines.join ("\n"));
                    }
                    else {
                        sheetPDFFile = new File (producingPath + "/" + senderName + "/" + "Monthly" + "/" + fromTime[0] + "/" + fromTime[1] + " " + monthName + ".pdf");
                    }
                }
                if (!sheetPDFFile.parent.exists)
                    sheetPDFFile.parent.create ();
                var sheetPreset = app.pdfExportPresets.itemByName ("TS Auto Working Sheets");
                try {
                    sheetPreset.name;
                }
                catch (e) {
                    sheetPreset = app.pdfExportPresets.add ({name: "TS Auto Working Sheets"});
                    with (sheetPreset) {
                        acrobatCompatibility = AcrobatCompatibility.ACROBAT_8;
                        bleedTop = 0;
                        bleedBottom =  0;
                        bleedInside = 0;
                        bleedOutside =  0;
                        colorBars = false;
                        bleedMarks = false;
                        colorBitmapSamplingDPI = 300;
                        cropImagesToFrames = true;
                        cropMarks = false;
                        exportLayers = false;
                        exportNonprintingObjects = false;
                        exportReaderSpreads = true;
                        exportWhichLayers = ExportLayerOptions.EXPORT_VISIBLE_PRINTABLE_LAYERS;
                        generateThumbnails = true;
                        grayscaleBitmapSamplingDPI = 300;
                        includeBookmarks = false;
                        includeHyperlinks = false;
                        includeICCProfiles = false;
                        includeSlugWithPDF = false;
                        includeStructure = true;
                        interactiveElementsOption = InteractiveElementsOptions.DO_NOT_INCLUDE;
                        optimizePDF = true;
                        pageInformationMarks = true;
                        pageMarksOffset = 0;
                        registrationMarks = false;
                    }
                }
                
                app.interactivePDFExportPreferences.pageRange = PageRange.ALL_PAGES;
                app.interactivePDFExportPreferences.viewPDF = false;
                try {
                    sheetDoc.exportFile (ExportFormat.INTERACTIVE_PDF, sheetPDFFile, false, sheetPreset);
                }
                catch (err) { }
                sheetDoc.close (SaveOptions.NO);
            }
        }
        dailySheetsFiles[0].remove ();
    }
    if (previewPath != "" && app.extractLabel ("TS_PDF_AUTO_EXPORTING_PAUSED") != "Yes") {
        for (var swal = saveWithoutAskList.length -1; swal >= 0; swal--) { 
            if (saveWithoutAskList[swal][1] == -1 || saveWithoutAskList[swal][1] == 1) {
                if (saveWithoutAskList[swal][1] == -1) {
                    saveWithoutAskList[swal][1] = -2;
                }
                else {
                    saveWithoutAskList[swal][1] = 2;
                }
                var docIndex = -1;
                for (var ad = 0; ad < app.documents.length; ad++) {
                    if (saveWithoutAskList[swal][0] == app.documents[ad].metadataPreferences.getProperty(treeShadeNamespace, "TS_ID")) {
                        docIndex = ad;
                        break;
                    }
                }
                if (docIndex != -1) {
                    if (saveWithoutAskList[swal][1] == -2) {
                        app.documents[docIndex].save ();
                    }
                    else {
                        app.documents[docIndex].close (SaveOptions.YES);
                    }
                }
                else {
                    saveWithoutAskList.splice (swal, 1);
                    if (app.extractLabel ("TS_SWITCH_WITH_INDESIGN") == "TRUE" && saveWithoutAskList.length == 0) {
                        switchToBridge ();
                    }
                }
            }
        }
        var PDFPreviewFolder = Folder (dataPath + "/Messages/PDFPreview");
        if (PDFPreviewFolder.exists) {
            var allFiles = PDFPreviewFolder.getFiles (isUnhiddenFile);
            allFiles = filterPDFPreviewFiles (allFiles);
            
            if (allFiles.length > 20) {
                allFiles = allFiles.slice (0, 20);
            }
            else if (allFiles.length == 0) {
                if (app.extractLabel ("TS_SWITCH_WITH_INDESIGN") == "TRUE" && isNeedToSwitch) {
                    switchToBridge ();
                }
            }
            outerloop:
            for (var afis = 0; afis < allFiles.length; afis++) {
                isNeedToSwitch = true;
                var pdfPreviewFileName = allFiles[afis].name;
                var currentLevel = getPDFPreviewLevel (pdfPreviewFileName);
                var pdfPreviewNameSplitted = pdfPreviewFileName.split (".");
                var fileName = pdfPreviewNameSplitted[1] + "." + pdfPreviewNameSplitted[2] + "." + pdfPreviewNameSplitted[3] + "." + pdfPreviewNameSplitted[4];
                var versionNumber = pdfPreviewNameSplitted[5];
                var isFinal = pdfPreviewNameSplitted[0] == "F" ? true : false;
                var fileID_Path_newID = ["/" + fileName.replace (/\./g, "\/")];
                getPath (fileID_Path_newID);
                var docFilePath = fileID_Path_newID[1];
                var oslIndex = -1;
                if (docFilePath) {
                    var docParentRelativePath = docFilePath.slice (0, docFilePath.lastIndexOf ("/"));
                    var docName = docFilePath.slice (docFilePath.lastIndexOf ("/") + 1);
                    var docWithoutExtName = docName.slice (docName.length - 5) == ".indd" ? docName.slice (0, docName.length - 5) : docName;
                    var docFile = new File (workshopPath + docFilePath);
                    if (docFile.exists) {
                        if (docFile.length <= 128) {
                            allFiles[afis].remove ();
                            continue outerloop;
                        }
                        var isDocAlreadyOpened = true;
                        var doc = null;
                        for (var a = 0; a < app.documents.length; a++) {
                            if (app.documents[a].layoutWindows.length > 0) {
                                if (fileID_Path_newID[0] == app.documents[a].metadataPreferences.getProperty(treeShadeNamespace, "TS_ID")) {
                                    doc = app.documents[a];
                                    if (doc.extractLabel ("TS_WAS_CLOSED") == "YES") {
                                        isDocAlreadyOpened = false;
                                    }
                                    break;
                                }
                            }
                        }
                        if (!app.modalState)
                            app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
                        if (!doc) {
                            isDocAlreadyOpened = false;
                            doc = app.open (docFile, true, OpenOptions.OPEN_ORIGINAL);
                            doc.insertLabel ("TS_WAS_CLOSED", "YES");
                            continue outerloop;
                        }
                        //pdf preset
                        var exportPDFPreset = null;
                        var docPDFPresetName = doc.extractLabel ("TS_PDF_PRESET");
                        if (docPDFPresetName) {
                            var docPDFPreset = app.pdfExportPresets.itemByName (docPDFPresetName);
                            var isDocPreset = true;
                            try {
                                docPDFPreset.name;
                            }
                            catch (e) {
                                isDocPreset = false;
                            }
                            if (isDocPreset) {
                                exportPDFPreset = docPDFPreset;
                            }
                        }
                        if (exportPDFPreset == null) {
                            exportPDFPreset = app.pdfExportPresets.itemByName ("Tree Shade");
                            try {
                                exportPDFPreset.name;
                            }
                            catch (e) {
                                exportPDFPreset = app.pdfExportPresets.add ({name: "Tree Shade"});
                                with (exportPDFPreset) {
                                    acrobatCompatibility = AcrobatCompatibility.ACROBAT_8;
                                    bleedTop = "5mm";
                                    bleedBottom =  "5mm";
                                    bleedInside = 0;
                                    bleedOutside =  "5mm";
                                    colorBars = false;
                                    bleedMarks = false;
                                    colorBitmapSamplingDPI = 300;
                                    cropImagesToFrames = true;
                                    cropMarks = true;
                                    exportLayers = true;
                                    exportNonprintingObjects = false;
                                    exportReaderSpreads = true;
                                    exportWhichLayers = ExportLayerOptions.EXPORT_VISIBLE_PRINTABLE_LAYERS;
                                    generateThumbnails = true;
                                    grayscaleBitmapSamplingDPI = 300;
                                    includeBookmarks = false;
                                    includeHyperlinks = false;
                                    includeICCProfiles = false;
                                    includeSlugWithPDF = false;
                                    includeStructure = true;
                                    interactiveElementsOption = InteractiveElementsOptions.DO_NOT_INCLUDE;
                                    optimizePDF = true;
                                    pageInformationMarks = true;
                                    pageMarksOffset = "2.5mm";
                                    registrationMarks = false;
                                }
                            }
                        }

                        //open and save psd
                        if (doc.extractLabel("TS_OPEN_AND_SAVE_PSD") != "No") {
                            //openAndSavePSDList
                            //openAndSavePSDList[osl][0] Document ID
                            //openAndSavePSDList[osl][1] Proccess Stage... 0 added, 1 
                            //openAndSavePSDList[osl][2] PSD list paths
                            //openAndSavePSDList[osl][3] ... Initial, Open, Opened, Update, Close, Closed, Done
                            for (var osl = openAndSavePSDList.length - 1; osl >= 0; osl--) {
                                if (openAndSavePSDList[osl][0] == fileID_Path_newID[0]) {
                                    oslIndex = osl;
                                    break;
                                }
                            }
                            if (oslIndex == -1) {
                                oslIndex = openAndSavePSDList.length;
                                var openAndSavePSDItem = new Array;
                                openAndSavePSDItem.push (fileID_Path_newID[0]);
                                openAndSavePSDItem.push (0);

                                openAndSavePSDItem.push ([]);
                                openAndSavePSDItem.push ("");
                                openAndSavePSDList.push (openAndSavePSDItem);
                            }
                        }
                        //updating all modified links
                        var isToUpdateAndExportNow = true;
                        if (oslIndex != -1) {
                            if (openAndSavePSDList[oslIndex][3] != "") {
                                isToUpdateAndExportNow = false;
                            }
                        }
                        if (isToUpdateAndExportNow) {
                            var allLinks = doc.links;
                            for (var c = 0; c < allLinks.length; c++) {
                                if (allLinks[c].status == LinkStatus.LINK_OUT_OF_DATE) {
                                    allLinks[c].update ();
                                    allLinks[c].insertLabel ("Updated", (new Date ().getTime ()).toString ());
                                }
                            }
                            //update live snippets
                            checkAllLiveSnippets (doc, false, true);

                            //Check Nested Links
                            var isAllUpdated = checkNestedMissingLinks (doc, currentLevel);
                            if (!isAllUpdated) {
                                continue outerloop;
                            }

                            var delayCheckIn = doc.extractLabel ("TS_DELAY_CHECKIN");
                            var isSolveOverflows = (doc.extractLabel("TS_SOLVE_OVERFLOWS") != "No");
                            var isToLeaveEmptyPages = (doc.extractLabel("TS_LEAVE_EMPTY_PAGES") != "No");
                            if (!delayCheckIn) {
                                if (isSolveOverflows) {
                                    if (doc.extractLabel("TS_AUTO_SOLVE_OVERFLOWS") != "No") {
                                        var isOverflowsSolved = true;
                                        try {
                                            isOverflowsSolved = tsSolveDocumentOverflows (doc, false);                                                                        
                                        }
                                        catch (err) {
                                            isOverflowsSolved = false;
                                        }
                                        if (!isOverflowsSolved) {
                                            allFiles[afis].remove ();
                                            continue;
                                        }
                                    }
                                    else {
                                        if (!tsSolveDocumentOverflows (doc, true)) {
                                            allFiles[afis].remove ();
                                            continue;
                                        }
                                    }
                                }
                            }
                            if (delayCheckIn) {
                                doc.insertLabel ("TS_DELAY_CHECKIN", "");
                                var lastDelayTime = doc.extractLabel ("TS_DELAY_TIME");
                                if (lastDelayTime) {
                                    lastDelayTime = parseInt (lastDelayTime, 10);
                                    if (new Date ().getTime () - lastDelayTime > 120000) {
                                        doc.insertLabel ("TS_DELAY_TIME", "");
                                        allFiles[afis].remove ();
                                        continue outerloop;
                                    }
                                }
                                else {
                                    doc.insertLabel ("TS_DELAY_TIME", (new Date ().getTime ()).toString ());
                                }
                                continue outerloop;
                            }
                            else {
                                doc.insertLabel ("TS_DELAY_TIME", "");
                            }
                            updateLinksRecords (doc, workshopPath + docFilePath);

                            var thePageRange = PageRange.ALL_PAGES;
                            if (isToLeaveEmptyPages) {
                                var testedPageRange = tsGetUsedPagesRange (doc);
                                if (testedPageRange[2] != false) {
                                    thePageRange = testedPageRange[0] + "-" + testedPageRange[1];
                                }
                            }
                            if (app.extractLabel ("startPDFExporting") == "start") {
                                var manualPDFFile = new File (workshopPath + docParentRelativePath + "/pdf/" + docWithoutExtName + ".pdf");
                                if (!manualPDFFile.exists) {
                                    var verPDFFile = new File (previewPath + docParentRelativePath + "/ver" + (versionNumber=="new"? "01" : versionNumber) + " " + docWithoutExtName + ".pdf");
                                    if (currentLevel == 0) {
                                        if (!verPDFFile.parent.exists) {
                                            verPDFFile.parent.create ();
                                        }
                                        
                                        if (exportPDFPreset.name.search (/Interactive/i) != -1) {
                                            app.interactivePDFExportPreferences.pageRange = thePageRange;
                                            app.interactivePDFExportPreferences.viewPDF = false;
                                            try {
                                                doc.exportFile (ExportFormat.INTERACTIVE_PDF, verPDFFile, false, exportPDFPreset);
                                            }
                                            catch (err) { }
                                        }
                                        else {
                                            app.pdfExportPreferences.pageRange = thePageRange;
                                            app.pdfExportPreferences.viewPDF = false;
                                            try {
                                                doc.exportFile (ExportFormat.PDF_TYPE, verPDFFile, false, exportPDFPreset);
                                            }
                                            catch (err) { }
                                        }
                                    }
                                    else {
                                        verPDFFile.remove ();
                                    }
                                }
                            }
                        }
                        if (oslIndex != -1) {
                            if (openAndSavePSDList[oslIndex][3] == "") {
                                var allLinks = doc.links;
                                for (var aln = 0; aln < allLinks.length; aln++) {
                                    if (allLinks[aln].linkType == "Photoshop" && allLinks[aln].name.search (/ts_lyr/) != -1) {
                                        if (allLinks[aln].status != LinkStatus.LINK_MISSING) {
                                            var linkFile = new File (preparePath (allLinks[aln].filePath));
                                            if (linkFile.exists) {
                                                if (linkFile.name.search (/ts_lyr_(gif|png|jpg)/) != -1) {
                                                    var psdDesFile = new File (pagesPath + "/Workshop" + docFilePath + "/" + File.decode (linkFile.name));
                                                    openAndSavePSDList[oslIndex][2].push (psdDesFile.fsName); //replace(/\\/g, "/")
                                                    linkFile.copy (psdDesFile);
                                                }
                                                else {
                                                    openAndSavePSDList[oslIndex][2].push (linkFile.fsName); //replace(/\\/g, "/")
                                                }
                                            }
                                        }
                                    }
                                }
                                openAndSavePSDList[oslIndex][3] = "Initial";
                            }
                            if (openAndSavePSDList[oslIndex][3] != "Done" && openAndSavePSDList[oslIndex][2].length > 0 && openAndSavePSDList[oslIndex][1] < 20) {
                                if (openAndSavePSDList[oslIndex][3] == "Closed") {
                                    openAndSavePSDList[oslIndex][3] = "Done";
                                    continue outerloop;
                                }
                                else {
                                    if (openAndSavePSDList[oslIndex][3] == "Initial") {
                                        excuteInPhotoshop (oslIndex);
                                    }
                                    openAndSavePSDList[oslIndex][1]++;
                                    continue outerloop;
                                }
                            }
                            else if (openAndSavePSDList[oslIndex][3] == "Done") {
                                var allPSDFiles = Folder (pagesPath + "/Workshop" + docFilePath).getFiles ("*.psd");
                                for (var apf = 0; apf < allPSDFiles.length; apf++) {
                                    allPSDFiles[apf].remove ();
                                }
                            }
                        }
                        doc.insertLabel ("TS_WAS_CLOSED", "");
                        if (doc.modified) {
                            var alreadyIndex = -1;
                            for (var swa = 0; swa < saveWithoutAskList.length; swa++) { 
                                if (fileID_Path_newID[0] == saveWithoutAskList[swa][0]) {
                                    alreadyIndex = swa;
                                    break;
                                }
                            }
                            if (alreadyIndex == -1) {
                                alreadyIndex = saveWithoutAskList.length;
                                saveWithoutAskList.push ([fileID_Path_newID[0]]);
                            }
                            saveWithoutAskList[alreadyIndex].push (isDocAlreadyOpened? -1 : 1);
                            saveWithoutAskList[alreadyIndex].push (versionNumber);
                        }
                        else {
                            var checkInMessageFolder = new Folder (dataPath + "/Messages/To CheckIn");
                            if (!checkInMessageFolder.exists) {
                                checkInMessageFolder.create ();
                            }
                            var pdfPreview = "noPDF";
                            var messageContent = versionNumber + ":" + pdfPreview;
                            var docDotID = fileID_Path_newID[0].replace (/\//g,"\.");
                            docDotID = docDotID.slice (1);
                            var targetFile = new File (checkInMessageFolder.fsName.replace(/\\/g, "/") + "/" + docDotID);
                            if (!targetFile.exists)
                                writeFile (targetFile, messageContent);
                            if (!isDocAlreadyOpened) {
                                doc.close (SaveOptions.YES);
                            }
                        }
                    }
                }
                allFiles[afis].remove ();
                if (oslIndex != -1) {
                    openAndSavePSDList.splice (oslIndex, 1);
                }
            }
        }
    }
}

function tsSolveDocumentOverflows (doc, isOnlyCheck) {
    var minFontSize = 1;
    var maxReductions = 9;
    var reductionFactor = 0.9;
    var isOverflowsSolved = true;
    for (var ads = 0; ads < doc.stories.length; ads++) {
        var story = doc.stories[ads];
        //scan the textFrames
        for (var stfs = 0; stfs < story.textFrames.length; stfs++)  {
            var textFrame = story.textFrames[stfs];
            var isSolved = tsSolveOverflow (textFrame, minFontSize, maxReductions, reductionFactor, isOnlyCheck);
            if (!isSolved) {
                if (isOnlyCheck) {
                    return false;
                }
                isOverflowsSolved = false;
            }
        }
        //Check rest of paragraphs in the story
        var isSolved = tsSolveOverflow (story, minFontSize, maxReductions, reductionFactor, isOnlyCheck);
        if (!isSolved) {
            if (isOnlyCheck) {
                return false;
            }
            isOverflowsSolved = false;

        }
        //scan the tables cells
        for (var sts = 0; sts < story.tables.length; sts++)  {
            var table = story.tables[sts];
            for (var tcs = 0; tcs < table.cells.length; tcs++) {
                var cell = table.cells[tcs];
                var isSolved = tsSolveOverflow (cell, minFontSize, maxReductions, reductionFactor, isOnlyCheck);
                if (!isSolved) {
                    if (isOnlyCheck) {
                        return false;
                    }
                    isOverflowsSolved = false;
                }
            }
        }
    }
    return isOverflowsSolved;
}

function tsSolveOverflow(container, minFontSize, maxReductions, reductionFactor, isOnlyCheck) {
    if (!container.overflows) {
        return true; // If there is no overflow then no need to process anything
    }
    else if (isOnlyCheck) {
        return false;
    }
    var stylesToReduce = [];

    // Step 1: Collect all unique paragraph styles used in the container
    var paragraphs = container.paragraphs;
    collectingloop:
    for (var i = 0; i < paragraphs.length; i++) {
        var paragraphStyle = paragraphs[i].appliedParagraphStyle;
        if (paragraphStyle != null) {
            for (var sstr = stylesToReduce.length - 1; sstr >= 0; sstr--) {
                if (stylesToReduce[sstr] === paragraphStyle) {
                    continue collectingloop;
                }
            }
            stylesToReduce.push(paragraphStyle);
        }
    }

    // Step 2: Loop through each style and reduce size
    for (var reductionCount = 0; reductionCount < maxReductions; reductionCount++) {
        // Iterate through each collected style
        for (var j = 0; j < stylesToReduce.length; j++) {
            var style = stylesToReduce[j];

            // Reduce the font size and space if the current font size is greater than the minimum
            if (style.pointSize > minFontSize) {
                style.pointSize = Math.max(style.pointSize * reductionFactor, minFontSize);
                style.spaceAfter *= reductionFactor;
                style.spaceBefore *= reductionFactor;
            }
        }

        // Step 3: Recompose container to apply changes
        container.recompose();

        // Step 4: Check if overflow is resolved after updating styles
        if (!container.overflows) {
            break; // Exit loop if overflow is resolved
        }
    }

    return !container.overflows;
}

function tsGetUsedPagesRange (targetDoc) {
    var firstPageName = targetDoc.pages.item (0).name;
    var lastPageName = targetDoc.pages.item (0).name;
    var isAllEmpty = null;
    for (var pgs = targetDoc.pages.length - 1; pgs >= 0; pgs--) {
        for (var aps = 0; aps < targetDoc.pages[pgs].allPageItems.length; aps++) {
            if (targetDoc.pages[pgs].allPageItems[aps].constructor.name == "TextFrame") {
                if (targetDoc.pages[pgs].allPageItems[aps].characters.length != 0 
                    || targetDoc.pages[pgs].allPageItems[aps].overflows) {
                    if (pgs == targetDoc.pages.length - 1) {
                        isAllEmpty = false;
                    }
                    lastPageName = targetDoc.pages.item (pgs).name;
                    return [firstPageName, lastPageName, isAllEmpty];
                }
            }
            else {
                lastPageName = targetDoc.pages.item (pgs).name;
                if (pgs == targetDoc.pages.length - 1) {
                    isAllEmpty = false;
                }
                return [firstPageName, lastPageName, isAllEmpty];
            }
        }
    }
    return [firstPageName, lastPageName, true];
}

function excuteInPhotoshop (oslIndex) {
    //openAndSavePSDList
    //openAndSavePSDList[osl][0] Document ID
    //openAndSavePSDList[osl][1] Proccess Stage... 0 added, 1 
    //openAndSavePSDList[osl][2] PSD list paths
    //openAndSavePSDList[osl][3] ... Initial, Open, Opened, Update, Close, Closed, Done
    var bt = new BridgeTalk();  
    bt.target = "photoshop";  
    //a string representaion of variables and the function to run. 
    //a string var. Note string single, double quote syntax '"+s+"'
    bt.body = "var docID = '" + openAndSavePSDList[oslIndex][0] + "';";
    //an array. Note use toSource()
    bt.body += "var psdPaths = [\"" + openAndSavePSDList[oslIndex][2].join ("\", \"") + "\"];";
    //the function and the call
    if (openAndSavePSDList[oslIndex][3] == "Initial") {
        openAndSavePSDList[oslIndex][3] = "Open";
        bt.body += psOpenScript.toString() + "psOpenScript();";
        bt.onResult = function(resObj) { 
            psdOpenReply (resObj.body);
        }
    }
    else if (openAndSavePSDList[oslIndex][3] == "Open") {
        bt.body += psCheckScript.toString() + "psCheckScript();";
        bt.onResult = function(resObj) { 
            psdOpenedReply (resObj.body);
        }
    }
    else if (openAndSavePSDList[oslIndex][3] == "Opened") {
        bt.body += psUpdateScript.toString() + "psUpdateScript();";
        bt.onResult = function(resObj) { 
            psdUpdateReply (resObj.body);
        }
    }
    else if (openAndSavePSDList[oslIndex][3] == "Update") {
        openAndSavePSDList[oslIndex][3] = "Close";
        bt.body += psCloseScript.toString() + "psCloseScript();";
        bt.onResult = function(resObj) { 
            psdCloseReply (resObj.body);
        }
    }
    else if (openAndSavePSDList[oslIndex][3] == "Close") {
        bt.body += psCheckScript.toString() + "psCheckScript();";
        bt.onResult = function(resObj) { 
            psdClosedReply (resObj.body);
        }
    }
    bt.onError = function(e) { alert(e.body); };  
    bt.send(8);
    function psOpenScript() {
        for (var pp = 0; pp < psdPaths.length; pp++) {
            app.open (File (psdPaths[pp]));
        }
        return docID;
    }
    function psCheckScript() {
        var foundList = new Array;
        var openedDocs = [];
        for (var dl = 0; dl < app.documents.length; dl++) {
            openedDocs.push (app.documents[dl]);
        }
        for (dl = 0; dl < openedDocs.length; dl++) {
            var isFound = false;
            var psdOpenedFile = null;
            try {
                psdOpenedFile = File (openedDocs[dl].fullName);
            }
            catch (e) {
                continue;
            }
            if (psdOpenedFile && psdOpenedFile.exists) {
                var psdFilePath = psdOpenedFile.fsName; //replace(/\\/g, "/")
                for (var pp = 0; pp < psdPaths.length; pp++) {
                    if (psdFilePath == psdPaths[pp]) {
                        isFound = true;
                        break;
                    }
                }
                if (isFound) {
                    foundList.push (psdFilePath);
                }
            }
        }
        var returned = docID;
        if (foundList.length > 0) {
            returned += ":TS:" + foundList.join (":TS:");
        }
        return returned;
    }
    function psUpdateScript() {
        var foundList = new Array;
        var openedDocs = [];
        for (var dl = 0; dl < app.documents.length; dl++) {
            openedDocs.push (app.documents[dl]);
        }
        for (dl = 0; dl < openedDocs.length; dl++) {
            var isFound = false;
            var psdOpenedFile = null;
            try {
                psdOpenedFile = File (openedDocs[dl].fullName);
            }
            catch (e) {
                continue;
            }
            if (psdOpenedFile && psdOpenedFile.exists) {
                var psdFilePath = psdOpenedFile.fsName; //replace(/\\/g, "/")
                for (var pp = 0; pp < psdPaths.length; pp++) {
                    if (psdFilePath == psdPaths[pp]) {
                        isFound = true;
                        app.activeDocument = openedDocs[dl];
                        try {
                            app.runMenuItem(stringIDToTypeID("placedLayerUpdateAllModified"));
                        }
                        catch (e) {}
                        break;
                    }
                }
                if (isFound) {
                    foundList.push (psdFilePath);
                }
            }
        }
        var returned = docID;
        if (foundList.length > 0) {
            returned += ":TS:" + foundList.join (":TS:");
        }
        return returned;
    }
    function psCloseScript() {
        var foundList = new Array;
        var openedDocs = [];
        for (var dl = 0; dl < app.documents.length; dl++) {
            openedDocs.push (app.documents[dl]);
        }
        for (dl = openedDocs.length - 1; dl >= 0; dl--) {
            var isFound = false;
            var psdOpenedFile = null;
            try {
                psdOpenedFile = File (openedDocs[dl].fullName);
            }
            catch (e) {
                continue;
            }
            if (psdOpenedFile && psdOpenedFile.exists) {
                var psdFilePath = psdOpenedFile.fsName; //replace(/\\/g, "/")
                for (var pp = 0; pp < psdPaths.length; pp++) {
                    if (psdFilePath == psdPaths[pp]) {
                        isFound = true;
                        var so = null;
                        if (psdOpenedFile.name.search (/ts_lyr_gif/) != -1) {
                            so = new GIFSaveOptions();
                            /*so.palette = Palette.LOCALSELECTIVE;
                            so.colors = 256;
                            so.dither = Dither.DIFFUSION;
                            so.ditherAmout = 100;
                            so.forced = ForcedColors.WEB;
                            so.interlaced = true;
                            so.matte = MatteType.WHITE;
                            so.preserveExactColors = true;
                            so.transparency = true;*/
                            var gifSaveAsFile = File (psdFilePath.slice (0, psdFilePath.lastIndexOf ('.')) + ".gif");
                            openedDocs[dl].saveAs (gifSaveAsFile, so, true);
                            openedDocs[dl].close (SaveOptions.DONOTSAVECHANGES);
                        }
                        else if (psdOpenedFile.name.search (/ts_lyr_png/) != -1) {
                            so = new PNGSaveOptions();
                            so.compression = 0
                            so.interlaced = false;
                            var pngSaveAsFile = File (psdFilePath.slice (0, psdFilePath.lastIndexOf ('.')) + ".png");
                            openedDocs[dl].saveAs (pngSaveAsFile, so, true, Extension.LOWERCASE);
                            openedDocs[dl].close (SaveOptions.DONOTSAVECHANGES);
                        }
                        else if (psdOpenedFile.name.search (/ts_lyr_jpg/) != -1) {
                            so = new JPEGSaveOptions();
                            so.formatOptions = FormatOptions.STANDARDBASELINE;
                            so.quality = 12;
                            var jpgSaveAsFile = File (psdFilePath.slice (0, psdFilePath.lastIndexOf ('.')) + ".jpg");
                            openedDocs[dl].saveAs (jpgSaveAsFile, so, true);
                            openedDocs[dl].close (SaveOptions.DONOTSAVECHANGES);
                        }
                        else {
                            openedDocs[dl].close (SaveOptions.SAVECHANGES);
                        }
                        break;
                    }
                }
                if (isFound) {
                    foundList.push (psdFilePath);
                }
            }
        }
        var returned = docID;
        if (foundList.length > 0) {
            returned += ":TS:" + foundList.join (":TS:");
        }
        return returned;
    }
};

function psdOpenReply (psdOpenAnswer) {
    var oslIndex = -1;
    for (var osl = openAndSavePSDList.length - 1; osl >= 0; osl--) {
        if (openAndSavePSDList[osl][0] == psdOpenAnswer) {
            oslIndex = osl;
            break;
        }
    }
    if (oslIndex != -1) {
        excuteInPhotoshop (oslIndex);
    }
}

function psdOpenedReply (psdCheckAnswer) {
    psdCheckAnswer = psdCheckAnswer.split (":TS:");
    var oslIndex = -1;
    for (var osl = openAndSavePSDList.length - 1; osl >= 0; osl--) {
        if (openAndSavePSDList[osl][0] == psdCheckAnswer[0]) {
            oslIndex = osl;
            break;
        }
    }
    if (oslIndex != -1) {
        if (psdCheckAnswer.length > openAndSavePSDList[oslIndex][2].length) {
            openAndSavePSDList[oslIndex][3] = "Opened";
        }
        excuteInPhotoshop (oslIndex);
    }
    else {
        alert ("Wrong!");
    }
}

function psdUpdateReply (psdCheckAnswer) {
    psdCheckAnswer = psdCheckAnswer.split (":TS:");
    var oslIndex = -1;
    for (var osl = openAndSavePSDList.length - 1; osl >= 0; osl--) {
        if (openAndSavePSDList[osl][0] == psdCheckAnswer[0]) {
            oslIndex = osl;
            break;
        }
    }
    if (oslIndex != -1) {
        openAndSavePSDList[oslIndex][3] = "Update";
        excuteInPhotoshop (oslIndex);
    }
    else {
        alert ("Wrong!");
    }
}

function psdCloseReply (psdCheckAnswer) {
    psdCheckAnswer = psdCheckAnswer.split (":TS:");
    var oslIndex = -1;
    for (var osl = openAndSavePSDList.length - 1; osl >= 0; osl--) {
        if (openAndSavePSDList[osl][0] == psdCheckAnswer[0]) {
            oslIndex = osl;
            break;
        }
    }
    if (oslIndex != -1) {
        excuteInPhotoshop (oslIndex);
    }
    else {
        alert ("Wrong!");
    }
}

function psdClosedReply (psdCheckAnswer) {
    psdCheckAnswer = psdCheckAnswer.split (":TS:");
    var oslIndex = -1;
    for (var osl = openAndSavePSDList.length - 1; osl >= 0; osl--) {
        if (openAndSavePSDList[osl][0] == psdCheckAnswer[0]) {
            oslIndex = osl;
            break;
        }
    }
    if (oslIndex != -1) {
        if (psdCheckAnswer.length == 1) {
            openAndSavePSDList[oslIndex][3] = "Closed";
        }
        else {
            excuteInPhotoshop (oslIndex);
        }
    }
    else {
        alert ("Wrong!");
    }
}

function switchToBridge () {
    isNeedToSwitch = false;
    var targetBridge = BridgeTalk.getSpecifier ("bridge");
    if (targetBridge && targetBridge.appStatus != "not installed") {
        var talkBridge = new BridgeTalk;
        talkBridge.target = targetBridge;
        talkBridge.body = "app.bringToFront ();";
        talkBridge.send ();
        return true;
    }
    return false;
}

function beforeDeleteLink (myEvent) {
    /*function_start*///alert ("beforeDeleteLink-start");
    var targetLink = myEvent.parent;
    if (!targetLink instanceof Link)
        return;
    //myEvent.stopPropagation();
    //myEvent.preventDefault();
}

function afterAttributeChangedHandler (myEvent) {
    /*function_start*///alert ("afterAttributeChangedHandler-start");
    var targetLink = myEvent.parent;
    if (!targetLink instanceof Link)
        return;
    var speedChain = new Array;
    testLink (targetLink, speedChain, false, null);
}

function linkPlaceHandler (myEvent) {
    /*function_start*///alert ("linkPlaceHandler-start");
    for (var c = myEvent.parent.placeGuns.pageItems.length - 1; c >= 0; c--) {
        if (myEvent.parent.placeGuns.pageItems[c].parentStory) {
            if (myEvent.parent.placeGuns.pageItems[c].parentStory.itemLink) {
                if (app.selection.length == 1) {
                    tsPlaceSelectionAndFileAndDoc = new Array;
                    tsPlaceSelectionAndFileAndDoc.push (app.selection[0]);
                    var snippetFullPath = preparePath (myEvent.parent.placeGuns.pageItems[c].parentStory.itemLink.filePath);
                    myEvent.parent.placeGuns.pageItems[c].parentStory.remove ();
                    var toPlaceSnippetFile = new File (snippetFullPath);
                    tsPlaceSelectionAndFileAndDoc.push (toPlaceSnippetFile);
                    tsPlaceSelectionAndFileAndDoc.push (myEvent.parent);
                    break;
                }
            }
        }
        if (!myEvent.parent.placeGuns.pageItems[c].allGraphics[0].itemLink)
            continue;
        var docPath = null;
        var newLink = myEvent.parent.placeGuns.pageItems[c].allGraphics[0].itemLink;
        var linkFullPath = preparePath (newLink.filePath);
        var linkFile = new File (linkFullPath);
        var folderParent = linkFile.parent;
        addToBridgeScanningQueue (folderParent.fsName.replace(/\\/g, "/"));
        if ((folderParent.fsName.replace(/\\/g, "/") + "/").indexOf (versionsPath + "/") == 0) {
            var versionFile = new File (preparePath (newLink.filePath));
            var tratedName = File.decode (versionFile.name);
            if (tratedName.search(/ver\d\d /i) == 0)
                tratedName = tratedName.slice (6);
            var workshopFile = new File (versionFile.parent.fsName.replace(/\\/g, "/").replace (versionsPath, workshopPath) + "/" + tratedName);
            versionFile.readonly = false;
            versionFile.copy (workshopFile);
            newLink.relink (workshopFile);
            linkFile = workshopFile;        
        }
        else if (app.extractLabel("autoImportOutOfTreeLinks") == "Yes") {
            if ((folderParent.fsName.replace(/\\/g, "/") + "/").indexOf (workshopPath + "/") != 0) {
                var targetBridge = BridgeTalk.getSpecifier ("bridge");
                if (targetBridge && targetBridge.appStatus != "not installed") {
                    var parentDoc = newLink.parent;
                    while (!(parentDoc instanceof Document)) {
                        parentDoc = parentDoc.parent;
                    }
                    docPath = preparePath (parentDoc.filePath.fsName.replace(/\\/g, "/"));
                    var talkBridge = new BridgeTalk;
                    talkBridge.target = targetBridge;
                    talkBridge.onResult = function (returnBtObj) { setBridgeScriptInPlace (returnBtObj.body); }
                    talkBridge.body = "tsImportFile (new File (\"" + linkFile.fsName.replace(/\\/g, "/") + "\"), new Folder (\"" + docPath + "/" + app.extractLabel("autoImportOutOfTreeLinksFolder") + "\"), false, \"" + newLink.id + "@" + parentDoc.id + "\", true, 6);";
                    talkBridge.send ();
                    function setBridgeScriptInPlace (theBody) {
                        if (theBody) {
                            var importingData = theBody.split (",");
                            if (importingData.length == 3) {
                                var importedPath = importingData[0];
                                var importedTSID = importingData[1];
                                var nativeLinkAndDocIDs = importingData[2].split ("@");
                                if (nativeLinkAndDocIDs.length == 2) {
                                    var nativeLinkID = nativeLinkAndDocIDs[0];
                                    var targetDocNativeID = nativeLinkAndDocIDs[1];
                                    var importedLinkTargetDoc = null;
                                    for (var ad = 0; ad < app.documents.length; ad++) {
                                        if (targetDocNativeID == app.documents[ad].id) {
                                            importedLinkTargetDoc = app.documents[ad];
                                            break;
                                        }
                                    }
                                    if (importedLinkTargetDoc) {
                                        var targetLink = importedLinkTargetDoc.links.itemByID (parseInt (nativeLinkID, 10));
                                        if (targetLink) {
                                            var importedLinkFile = new File (workshopPath + importedPath);
                                            targetLink.relink (importedLinkFile);
                                            targetLink.insertLabel ("TSID", importedTSID);
                                            targetLink.insertLabel ("Updated", (new Date ().getTime ()).toString ());

                                            var importedLinkParentPath = importedLinkFile.parent.fsName.replace(/\\/g, "/") + "/";
                                            var importedLinkTargetDoc = targetLink.parent;
                                            while (!(importedLinkTargetDoc instanceof Document)) {
                                                importedLinkTargetDoc = importedLinkTargetDoc.parent;
                                            }
                                            var targetDocPath = preparePath (importedLinkTargetDoc.filePath.fsName.replace(/\\/g, "/"));
                                            if (importedLinkParentPath.indexOf (targetDocPath + "/Links/") == 0 && !targetLink.extractLabel ("TS_DYNAMIC_LINK")) {
                                                targetLink.insertLabel ("TS_DYNAMIC_LINK", ".$/<SP>/<LN>.*");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        if (docPath == null) {
            prepareLink (newLink);
            var linkParentPath = linkFile.parent.fsName.replace(/\\/g, "/") + "/";
            var parentDoc = newLink.parent;
            while (!(parentDoc instanceof Document)) {
                parentDoc = parentDoc.parent;
            }
            var targetDocPath = preparePath (parentDoc.filePath.fsName.replace(/\\/g, "/"));
            if (linkParentPath.indexOf (targetDocPath + "/Links/") == 0 && !newLink.extractLabel ("TS_DYNAMIC_LINK")) {
                newLink.insertLabel ("TS_DYNAMIC_LINK", ".$/<SP>/<LN>.*");
            }
        }
        newLink.addEventListener ("afterUpdate", linkAfterUpdateHandler, false);
        newLink.addEventListener ("beforeUpdate", linkBeforeUpdateHandler, false);
        newLink.addEventListener ("beforeDelete", beforeDeleteLink, false);
        newLink.addEventListener ("afterAttributeChanged", afterAttributeChangedHandler, false);
    }
}

function isUnhiddenFile (targetFile) {
    /*function_start*///alert ("isUnhiddenFile-start");
    if (targetFile instanceof Folder)
        return false;
    if (targetFile.hidden)
        return false;
    return true;
}

function isUnhiddenFolder (targetFolder) {
    /*function_start*///alert ("isUnhiddenFolder-start");
    if (targetFolder instanceof File)
        return false;
    if (targetFolder.hidden)
        return false;
    return true;
}

function afterOpenHandler (myEvent) {
    /*function_start*///alert ("afterOpenHandler-start");
    var targetDocument = myEvent.parent.parent;
    if (!targetDocument instanceof Document) {
        return false;
    }
    if (workshopPath == "")
        return false;
    if (!Folder (workshopPath).exists)
        return false;
    targetDocument.preflightOptions.preflightOff = true;
    var docID = targetDocument.metadataPreferences.getProperty(treeShadeNamespace, "TS_ID");
    if (docID) {
        var docDotID = docID.replace (/\//g,"\.");
        docDotID = docDotID.slice (1);
        var openedMarkFile = new File (dataPath + "/Opened Documents/" + docDotID);
        if (!openedMarkFile.parent.exists)
            openedMarkFile.parent.create ();
        writeFile (openedMarkFile, new Date().getTime());
    }
}

function beforeCloseHandler (myEvent) {
    /*function_start*///alert ("beforeCloseHandler-start");
    var targetDocument = myEvent.parent;
    if (targetDocument instanceof LayoutWindow)
        return false;
    if (workshopPath == "")
        return false;
    if (!Folder (workshopPath).exists)
        return false;
    var docID = targetDocument.metadataPreferences.getProperty(treeShadeNamespace, "TS_ID");
    if (docID) {
        var docDotID = docID.replace (/\//g,"\.");
        docDotID = docDotID.slice (1);
        var openedMarkFile = new File (dataPath + "/Opened Documents/" + docDotID);
        if (openedMarkFile.exists)
            openedMarkFile.remove ();
    }
}

function beforeSaveHandler (myEvent) {
    /*function_start*///alert ("beforeSaveHandler-start");
    var targetDocument = myEvent.parent;
    if (!targetDocument instanceof Document) {
        return;
    }
    if (workshopPath == "")
        return false;
    if (!Folder (workshopPath).exists)
        return false;
    var docPath = preparePath (targetDocument.filePath.fsName.replace(/\\/g, "/"));
    var docName = docPath + "/" + targetDocument.name;
    var allLinks = targetDocument.links;
    var speedChain = new Array;
    var bridgeTalkLinesArr = new Array;
    for (var c = 0; c < allLinks.length; c++) {
        testLink (allLinks[c], speedChain, true, bridgeTalkLinesArr);
    }
    if (bridgeTalkLinesArr.length > 0) {
        sendImportListToInDesign (bridgeTalkLinesArr);
    }
    var docID = targetDocument.metadataPreferences.getProperty(treeShadeNamespace, "TS_ID");
    if (docID) {
        for (var swal = saveWithoutAskList.length -1; swal >= 0; swal--) {
            if (saveWithoutAskList[swal][0] == docID) {
                return true;
            }
        }
    }
    updateLinksRecords (targetDocument, docName);
}

function afterSaveHandler (myEvent) {
    /*function_start*///alert ("beforeSaveHandler-start");
    var targetDocument = myEvent.parent;
    if (!targetDocument instanceof Document) {
        return;
    }
    if (workshopPath == "")
        return false;
    if (!Folder (workshopPath).exists)
        return false;
        
    var alreadyIndex = -1;
    for (var swa = 0; swa < saveWithoutAskList.length; swa++) {
        if (targetDocument.metadataPreferences.getProperty(treeShadeNamespace, "TS_ID") == saveWithoutAskList[swa][0]) {
            alreadyIndex = swa;
            break;
        }
    }
    var docID = targetDocument.metadataPreferences.getProperty(treeShadeNamespace, "TS_ID");
    if (!docID) {
        return false;
    }
    var docDotID = docID.replace (/\//g,"\.");
    docDotID = docDotID.slice (1);
    var isOKAfterAsking = true;
    if (isStartAutoCheckIn && app.extractLabel ("isAskingForCheckingIn") == "Yes" && alreadyIndex == -1) {
        var messageBody = (app.extractLabel ("isToCreateNewVersion") == "Yes") ? "create new version" : "check in";
        isOKAfterAsking = confirm ("Do you want to " + messageBody + "?");
    }

    if ((isStartAutoCheckIn && isOKAfterAsking) || alreadyIndex != -1) {
        var checkInMessageFolder = new Folder (dataPath + "/Messages/To CheckIn");
        if (!checkInMessageFolder.exists) {
            checkInMessageFolder.create ();
        }
        var isToCreateNew = (app.extractLabel ("isToCreateNewVersion") == "Yes") ? true : false;
        var pdfPreview = (alreadyIndex != -1)? "noPDF" : "yesPDF";
        var newOrOverwriteOrDigits = isToCreateNew? "NewVersion" : "OverwriteCurrent";
        if (alreadyIndex != -1) {
            newOrOverwriteOrDigits = saveWithoutAskList[alreadyIndex][2];
        }
        var messageContent = newOrOverwriteOrDigits + ":" + pdfPreview;
        var targetFile = new File (checkInMessageFolder.fsName.replace(/\\/g, "/") + "/" + docDotID);
        if (!targetFile.exists)
            writeFile (targetFile, messageContent);
        if (alreadyIndex != -1) {
            saveWithoutAskList.splice (alreadyIndex, 1);
            if (app.extractLabel ("TS_SWITCH_WITH_INDESIGN") == "TRUE" && saveWithoutAskList.length == 0) {
                switchToBridge ();
            }
        }
    }
    else {
        var modifiedMessageFolder = new Folder (dataPath + "/Messages/Modified");
        if (!modifiedMessageFolder.exists) {
            modifiedMessageFolder.create ();
        }
        var targetFile = new File (modifiedMessageFolder.fsName.replace(/\\/g, "/") + "/" + docDotID);
        if (!targetFile) 
            writeFile (File (modifiedMessageFolder.fsName.replace(/\\/g, "/") + "/" + docDotID), "TS");
    }
}

function afterSaveAsHandler (myEvent) {
    /*function_start*///alert ("afterSaveAsHandler-start");
    var targetDocument = myEvent.parent;
}

function beforeSaveAsHandler (myEvent) {
    /*function_start*///alert ("beforeSaveAsHandler-start");
    var toBeSavedAsFile = new File (myEvent.fullName);
    if (toBeSavedAsFile.fsName.replace(/\\/g, "/").indexOf (versionsPath + "/") == 0) {
        myEvent.preventDefault ();
    }
}

function beforeSaveACopy (myEvent) {
    /*function_start*///alert ("beforeSaveACopy-start");
    var targetDocument = myEvent.parent;
}

function updateLinksRecords (targetDocument, docName) {
    /*function_start*///alert ("updateLinksRecords-start");
    var docID = targetDocument.metadataPreferences.getProperty(treeShadeNamespace, "TS_ID");
    if (!docID) {
        return;
    }
    var docDotID = docID.replace (/\//g,"\.");
    docDotID = docDotID.slice (1);
    if (!Folder (dataPath).exists)
        return;
    //Test all links
    var allLinks = targetDocument.links;
    var speedChain = new Array;
    var bridgeTalkLinesArr = new Array;
    for (var c = 0; c < allLinks.length; c++) {
        testLink (allLinks[c], speedChain, true, bridgeTalkLinesArr);
    }
    if (bridgeTalkLinesArr.length > 0) {
        sendImportListToInDesign (bridgeTalkLinesArr);
    }
    
    var linksReportFile = new File (dataPath + "/IDs" + docID + "/Workshop/LinksReport");
    var linksMissedFile = new File (dataPath + "/IDs" + docID + "/Workshop/MissedReport");
    var linksOutFile = new File (dataPath + "/IDs" + docID + "/Workshop/OutReport");
    var oldLinksLines = new Array;
    var content = readFile (linksReportFile);
    if (content) {
        oldLinksLines = content.split ("\n");
    }
    var docPath = null;
    docPath = preparePath (targetDocument.filePath.fsName.replace(/\\/g, "/"));
    docPath = docPath.replace (workshopPath, "");
    if (targetDocument.extractLabel("TS_DYNAMIC_AND_LEAF") != "No") {
        var bridgeMessage = "";
        var messageNewLine = "";
        for (var g = 0; g < targetDocument.links.length; g++) {
            var linkFile = new File (preparePath (targetDocument.links[g].filePath));
            var desFile = null;
            var cancelOrOverwriteOrSolve = "1";
            var isToCopy = true;
            var targetLink = targetDocument.links[g];
            if (File.decode (linkFile.parent.name) == "Leaf") {
                isToCopy = false;
                var linkID = targetLink.extractLabel ("TSID");
                if (linkID != "") {
                    var linkTSIDLabel = linkID;
                    var placesFolder = new Folder (dataPath + "/IDs" + linkTSIDLabel + "/Workshop/Places");
                    if (placesFolder.exists) {
                        var placesFiles = placesFolder.getFiles (isUnhiddenFile);
                        var docAbsolutePathsList = new Array;
                        for (var e = 0; e < placesFiles.length; e++) {
                            var pagesFolder;
                            var fileDotID = File.decode (placesFiles[e].name);
                            var fileID = "/" + fileDotID.replace (/\./g, "\/");
                            var fileID_Path_newID = [fileID];
                            getPath (fileID_Path_newID);
                            if (fileID_Path_newID[1]) {
                                docAbsolutePathsList.push (fileID_Path_newID[1]);
                            }
                        }
                        if (docAbsolutePathsList.length > 0) {
                            var homePath = "";
                            var ii = 0;
                            while (true) {
                                var nextChar;
                                if (docAbsolutePathsList[0].length > ii) {
                                    nextChar = docAbsolutePathsList[0][ii];
                                }
                                else {
                                    break;
                                }
                                var isToAdd = true;
                                for (var sh = 1; sh < docAbsolutePathsList.length; sh++) {
                                    if (docAbsolutePathsList[sh].length > ii) {
                                        if (nextChar != docAbsolutePathsList[sh][ii]) {
                                            isToAdd = false;
                                            break;
                                        }
                                    }
                                    else {
                                        isToAdd = false;
                                        break;
                                    }
                                }
                                if (isToAdd) {
                                    homePath += nextChar;
                                }
                                else {
                                    break;
                                }
                                ii++;
                            }
                            homePath = homePath.slice (0, homePath.lastIndexOf ("/"));
                            homePath = workshopPath + homePath + "/Leaf";
                            if (homePath != linkFile.parent.fsName.replace(/\\/g, "/")) {
                                desFile = new File (homePath + "/"+ File.decode (linkFile.name));
                            }
                        }
                    }
                }
            }
            else {
                var dynamicLinkPhrase = targetLink.extractLabel ("TS_DYNAMIC_LINK");
                var newSolvedPath = "";
                if (dynamicLinkPhrase) {
                    //<LN>  <Link_Name>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<LN>/ig, "<Link_Name>");
                    //<SP>  <Sub_Path>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<SP>/ig, "<Sub_Path>");
                    //<DN>  <Doc_Name>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<DN>/ig, "<Doc_Name>");
                    //<DO>  <Doc_Order>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<DO>/ig, "<Doc_Order>");
                    //<PN>  <Page_Num>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<PN>/ig, "<Page_Num>");
                    //<ASP>  <Applied_Section_Prefix>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<ASP>/ig, "<Applied_Section_Prefix>");
                    //<ASM>  <Applied_Section_Marker>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<ASM>/ig, "<Applied_Section_Marker>");
                    //<MP>  <Master_Prefix>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<MP>/ig, "<Master_Prefix>");
                    //<MN>  <Master_Name>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<MN>/ig, "<Master_Name>");
                    //<MFN> <Master_Full_Name>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<MFN>/ig, "<Master_Full_Name>");
                    //<PF>  <Para_First>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<PF>/ig, "<Para_First>");
                    //<PB>  <Para_Before>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<PB>/ig, "<Para_Before>");
                    //<PT>  <Para_This>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<PT>/ig, "<Para_This>");
                    //<CO>  <Char_Order>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<CO>/ig, "<Char_Order>");
                    //<AO>  <Anchored_Order>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<AO>/ig, "<Anchored_Order>");
                    //<PO>  <Para_Order>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<PO>/ig, "<Para_Order>");
                    //<PA>  <Para_After>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<PA>/ig, "<Para_After>");
                    //<PS>  <Para_Style>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<PS>/ig, "<Para_Style>");
                    //<CS>  <CS>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<CS>/ig, "<Char_Style>");
                    //<NA>  <Name_At>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace(/<NA>/ig, "<Name_At>");
                    //<S1>  <Shift1>
                    dynamicLinkPhrase = dynamicLinkPhrase.replace (/<S(\-?\d+>)/ig, "<Shift$1");

                    newSolvedPath = dynamicLinkPhrase;
                    if (newSolvedPath[0] == ':') {
                        isToCopy = false;
                        newSolvedPath = newSolvedPath.slice (1);
                    }
                    if (newSolvedPath.indexOf ("//") == 0) {
                        newSolvedPath = newSolvedPath.replace("//", Folder (workshopPath).parent.fsName.replace(/\\/g, "/") + "/");
                    }
                    var prefixDotsIndex = newSolvedPath.search (/\/?\.+\$?\//);
                    while (prefixDotsIndex != -1) {
                        var targetParent = null;
                        var deepCount = 2;
                        var isRelative = false;
                        var dotsFirstMatch = newSolvedPath.match (/\/?\.+\$?\//)[0];
                        if (dotsFirstMatch[dotsFirstMatch.length - 2] == '$') {
                            targetParent = Folder (workshopPath + docPath);
                            deepCount = 3;
                        }
                        else {
                            targetParent = linkFile.parent;
                        }
                        if (dotsFirstMatch[0] == "/") {
                            isRelative = true;
                            deepCount++;
                        }
                        for (var pml = 0; pml < dotsFirstMatch.length - deepCount; pml++) {
                            if (!targetParent.parent) {
                                break;
                            }
                            targetParent = targetParent.parent;
                        }
                        var parentPath = targetParent.fsName.replace(/\\/g, "/") + "/";
                        if (isRelative) {
                            parentPath = parentPath.replace (workshopPath, "");
                        }
                        newSolvedPath = newSolvedPath.replace (dotsFirstMatch, parentPath);
                        prefixDotsIndex = newSolvedPath.search (/\/?\.+\$?\//);
                    }
                    var nameDotsIndex = newSolvedPath.search (/<\.+\$?>/);
                    while (nameDotsIndex != -1) {
                        var deepCount = 3;
                        var dotsFirstMatch = newSolvedPath.match (/<\.+\$?>/)[0];
                        var targetParent = null;
                        if (dotsFirstMatch[dotsFirstMatch.length - 2] == '$') {
                            targetParent = Folder (workshopPath + docPath);
                            deepCount = 4;
                        }
                        else {
                            targetParent = linkFile.parent;
                        }
                        for (var pml = 0; pml < dotsFirstMatch.length - deepCount; pml++) {
                            if (!targetParent.parent) {
                                break;
                            }
                            targetParent = targetParent.parent;
                        }
                        var parentFolderName = "";
                        if (targetParent && targetParent.name) {
                            parentFolderName = File.decode (targetParent.name);
                        }
                        newSolvedPath = newSolvedPath.replace (dotsFirstMatch, parentFolderName);
                        nameDotsIndex = newSolvedPath.search (/<\.+\$?>/);
                    }

                    if (newSolvedPath.search (/<Link_Name>/i) != -1) {
                        var theName = File.decode (linkFile.name);
                        if (theName.lastIndexOf (".") != -1) {
                            theName = theName.slice (0, theName.lastIndexOf ("."));
                        }
                        newSolvedPath = newSolvedPath.replace(/<Link_Name>/gi, theName);
                    }
                    if (newSolvedPath.search (/<Sub_Path>/i) != -1) {
                        var subPath = "";
                        var linkParentPath = linkFile.parent.fsName.replace(/\\/g, "/") + "/";
                        var previousDocPath = targetDocument.extractLabel ("TS_PATH_PREVIOUS");
                        if (previousDocPath && linkParentPath.indexOf (previousDocPath + "/") == 0) {
                            subPath = linkParentPath.replace (previousDocPath + "/", "");
                        }
                        if (subPath == "") {
                            newSolvedPath = newSolvedPath.replace(/<Sub_Path>\/?/gi, subPath);
                        }
                        else {
                            newSolvedPath = newSolvedPath.replace(/<Sub_Path>/gi, subPath.slice (0, -1));
                        }
                    }
                    if (newSolvedPath.search (/<Doc_Name>/i) != -1) {
                        var pureName = getPureName (targetDocument.name, false);
                        newSolvedPath = newSolvedPath.replace(/<Doc_Name>/gi, pureName);
                    }
                    if (newSolvedPath.search (/<Doc_Order>/i) != -1) {
                        var docFileOrder = 1;
                        var docParentFolder = new Folder (workshopPath + docPath);
                        var allDocs = docParentFolder.getFiles ("*.indd");
                        sortFilesList (docParentFolder, allDocs, false);
                        for (var ad = 0; ad < allDocs.length; ad++) {
                            if (File.decode (allDocs[ad].name) == File.decode (targetDocument.name)) {
                                docFileOrder = ad + 1;
                                break;
                            }
                        }
                        newSolvedPath = newSolvedPath.replace(/<Doc_Order>/gi, tsFillZerosIfAllDigits (docFileOrder.toString (), 4));
                    }
                    if (newSolvedPath.search (/<Page_Num>/i) != -1) {
                        if (!(targetLink.parent instanceof Story)) {
                            newSolvedPath = newSolvedPath.replace(/<Page_Num>/gi, tsFillZerosIfAllDigits (targetLink.parent.parentPage.name, 4));
                        }
                    }
                    if (newSolvedPath.search (/<Applied_Section_Prefix>/i) != -1) {
                        if (!(targetLink.parent instanceof Story)) {
                            if (targetLink.parent.parentPage.appliedSection) {
                                newSolvedPath = newSolvedPath.replace(/<Applied_Section_Prefix>/gi, targetLink.parent.parentPage.appliedSection.sectionPrefix);
                            }
                        }
                    }
                    if (newSolvedPath.search (/<Applied_Section_Marker>/i) != -1) {
                        if (!(targetLink.parent instanceof Story)) {
                            if (targetLink.parent.parentPage.appliedSection) {
                                newSolvedPath = newSolvedPath.replace(/<Applied_Section_Marker>/gi, targetLink.parent.parentPage.appliedSection.marker);
                            }
                        }
                    }
                    if (newSolvedPath.search (/<Master_Prefix>/) != -1) {
                        if (!(targetLink.parent instanceof Story)) {
                            newSolvedPath = newSolvedPath.replace(/<Master_Prefix>/gi, targetLink.parent.parentPage.appliedMaster.namePrefix);
                        }
                    }
                    if (newSolvedPath.search (/<Master_Name>/i) != -1) {
                        if (!(targetLink.parent instanceof Story)) {
                            newSolvedPath = newSolvedPath.replace(/<Master_Name>/gi, targetLink.parent.parentPage.appliedMaster.baseName);
                        }
                    }
                    if (newSolvedPath.search (/<Master_Full_Name>/i) != -1) {
                        if (!(targetLink.parent instanceof Story)) {
                            newSolvedPath = newSolvedPath.replace(/<Master_Full_Name>/gi, targetLink.parent.parentPage.appliedMaster.name);
                        }
                    }
                    if (newSolvedPath.search (/<Para_First>/i) != -1) {
                        var anchorChar = getAnchoredChar (targetLink);
                        if (anchorChar) {
                            var firstPara = anchorChar.parentStory.paragraphs.firstItem ();
                            if (firstPara) {
                                var paraContents = getParaPlainContents (firstPara);
                                if (paraContents) {
                                    newSolvedPath = newSolvedPath.replace(/<Para_First>/gi, paraContents);
                                }
                            }
                        }
                    }
                    if (newSolvedPath.search (/<Para_Before>/i) != -1) {
                        var anchorChar = getAnchoredChar (targetLink);
                        if (anchorChar) {
                            var beforePara = anchorChar.parent.paragraphs.previousItem (anchorChar.paragraphs.item (0));
                            if (beforePara) {
                                var paraContents = getParaPlainContents (beforePara);
                                if (paraContents) {
                                    newSolvedPath = newSolvedPath.replace(/<Para_Before>/gi, paraContents);
                                }
                            }
                        }
                    }
                    if (newSolvedPath.search (/<Para_This>/i) != -1) {
                        var anchorChar = getAnchoredChar (targetLink);
                        if (anchorChar) {
                            var thisPara = anchorChar.paragraphs.item (0);
                            if (thisPara) {
                                var paraContents = getParaPlainContents (thisPara);
                                if (paraContents) {
                                    newSolvedPath = newSolvedPath.replace(/<Para_This>/gi, paraContents);
                                }
                            }
                        }
                    }
                    if (newSolvedPath.search (/<Char_Order>/i) != -1) {
                        var anchorChar = getAnchoredChar (targetLink);
                        if (anchorChar) {
                            var thisPara = anchorChar.paragraphs.item(0);
                            if (thisPara) {
                                var charOrder = thisPara.characters.itemByRange(thisPara.characters.firstItem(), anchorChar).characters.length - 1;
                                if (charOrder) {
                                    newSolvedPath = newSolvedPath.replace(/<Char_Order>/gi, tsFillZeros(charOrder, 4));
                                }
                            }
                        }
                    }
                    if (newSolvedPath.search (/<Anchored_Order>/i) != -1) {
                        var anchorChar = getAnchoredChar (targetLink);
                        if (anchorChar) {
                            var thisPara = anchorChar.paragraphs.item(0);
                            if (thisPara) {
                                var anchoredChars = thisPara.characters.itemByRange(thisPara.characters.firstItem(), anchorChar).contents.toString ();
                                var charOrder = -1;
                                for (var i = 0; i < anchoredChars.length; i++) {
                                    if (anchoredChars.charCodeAt(i) == 65532) {
                                        charOrder++;
                                    }
                                }
                                newSolvedPath = newSolvedPath.replace(/<Anchored_Order>/gi, tsFillZeros(charOrder, 4));
                            }
                        }
                    }
                    if (newSolvedPath.search (/<Para_Order>/i) != -1) {
                        var anchorChar = getAnchoredChar (targetLink);
                        if (anchorChar) {
                            var paraOrderParent = anchorChar.parentStory;
                            if (anchorChar.parent.constructor.name == "Cell") {
                                paraOrderParent = anchorChar.parent;
                            }
                            var firstPara = paraOrderParent.paragraphs.firstItem ();
                            if (firstPara) {
                                var thisPara = anchorChar.paragraphs.item (0);
                                if (thisPara) {
                                    var paraOrder = paraOrderParent.paragraphs.itemByRange (firstPara, thisPara).paragraphs.length;
                                    if (paraOrder) {
                                        newSolvedPath = newSolvedPath.replace(/<Para_Order>/gi, tsFillZeros (paraOrder, 4));
                                    }
                                }
                            }
                        }
                    }
                    if (newSolvedPath.search (/<Para_After>/i) != -1) {
                        var anchorChar = getAnchoredChar (targetLink);
                        if (anchorChar) {
                            var beforePara = anchorChar.parent.paragraphs.nextItem (anchorChar.paragraphs.item (0));
                            if (beforePara) {
                                var paraContents = getParaPlainContents (beforePara);
                                if (paraContents) {
                                    newSolvedPath = newSolvedPath.replace(/<Para_After>/gi, paraContents);
                                }
                            }
                        }
                    }
                    if (newSolvedPath.search (/<Para_Style>/i) != -1) {
                        var anchorChar = getAnchoredChar (targetLink);
                        if (anchorChar) {
                            var paraStyleName = anchorChar.appliedParagraphStyle.name;
                            if (paraStyleName) {
                                newSolvedPath = newSolvedPath.replace(/<Para_Style>/gi, paraStyleName);
                            }
                        }
                    }
                    if (newSolvedPath.search (/<Char_Style>/i) != -1) {
                        var anchorChar = getAnchoredChar (targetLink);
                        if (anchorChar) {
                            var charStyleName = anchorChar.appliedCharacterStyle.name;
                            if (charStyleName) {
                                newSolvedPath = newSolvedPath.replace(/<Char_Style>/gi, charStyleName);
                            }
                        }
                    }
                    newSolvedPath = solveShift (newSolvedPath);
                    newSolvedPath = solveNameAt (newSolvedPath);
                    if (newSolvedPath.lastIndexOf ("/") != -1) {
                        var lastPortion = newSolvedPath.slice (newSolvedPath.lastIndexOf ("/") + 1);
                        var newPortion = null;
                        if (lastPortion.indexOf ("*") != -1 || lastPortion.indexOf ("?") != -1) {
                            var toBeSearched = new Folder (newSolvedPath.slice (0, newSolvedPath.lastIndexOf ("/")));
                            if (toBeSearched.exists) {
                                var expectedFiles = toBeSearched.getFiles (lastPortion);
                                if (expectedFiles.length > 0) {
                                    var theChosen = expectedFiles[0];
                                    for (var exf = 1; exf < expectedFiles.length; exf++) {
                                        if (theChosen.modified.getTime () < expectedFiles[exf].modified.getTime ()) {
                                            theChosen = expectedFiles[exf];
                                        }
                                    }
                                    newPortion = File.decode (theChosen.name);
                                }
                            }
                            if (newPortion == null) { 
                                newPortion = false;
                            }
                        }
                        if (newPortion == false) { 
                            newSolvedPath = false;
                        }
                        else if (newPortion) {
                            newSolvedPath = newSolvedPath.slice (0, newSolvedPath.lastIndexOf ("/") + 1) + newPortion;
                        }  
                    }
                    if (newSolvedPath) {
                        if (newSolvedPath != linkFile.fsName.replace(/\\/g, "/")) {
                            desFile = new File (newSolvedPath);
                        }
                        else if (targetLink.status == LinkStatus.LINK_OUT_OF_DATE) {
                            targetLink.update ();
                        }
                    }
                }
            }
            if (desFile) {
                if (!desFile.exists && linkFile.exists) {
                    desFile.parent.create ();
                    bridgeMessage += messageNewLine + "tsChangeFile (new File (\"" + linkFile.fsName.replace(/\\/g, "/") + "\"), new File (\"" + desFile.fsName.replace(/\\/g, "/") + "\"), " + (isToCopy? "true" : "false") + ", " + cancelOrOverwriteOrSolve + ", false);";
                    messageNewLine = "\n";
                    if (isToCopy) {
                        targetLink.insertLabel ("TSID", "");
                    }
                    continue;
                }
                else {
                    if (isShade (desFile)) { //the destination file is exist destination file is shade
                        var targetShadeID = getTSID (desFile);
                        if (targetShadeID) {
                            var dynamicLoadTime = targetLink.extractLabel ("TS_DYNAMIC_LOAD");
                            if (dynamicLoadTime) {
                                continue;
                            }
                            var loadMessageFolder = new Folder (dataPath + "/Messages/To Load");
                            if (!loadMessageFolder.exists) {
                                loadMessageFolder.create ();
                            }
                            var messageFileName = targetShadeID.replace (/\//g,"\.");
                            messageFileName = messageFileName.slice (1);
                            var loadRequestFile = new File (loadMessageFolder.fsName.replace(/\\/g, "/") + "/" + messageFileName);
                            writeFile (loadRequestFile, "TS");
                            targetLink.insertLabel ("TS_DYNAMIC_LOAD", new Date ().getTime ());
                            continue;
                        }
                        else {
                            addToBridgeScanningQueue (desFile.parent.fsName.replace(/\\/g, "/"));
                        }
                    }
                    else { //the destination file is exist and is actual
                        try {
                            targetLink.relink (desFile);
                        }
                        catch (e) {}
                        targetLink.insertLabel ("TS_DYNAMIC_LOAD", "");
                        var newLinkFile = new File (preparePath (targetLink.filePath));
                        //if (newLinkFile.fsName.replace(/\\/g, "/") == linkFile.fsName.replace(/\\/g, "/")) {
                            targetLink.insertLabel ("TSID", "");
                        //}      
                    }
                }
            }
        }
        if (bridgeMessage != "") {
            var targetBridge = BridgeTalk.getSpecifier ("bridge");
            if (targetBridge && targetBridge.appStatus != "not installed") {
                var talkBridge = new BridgeTalk;
                talkBridge.target = targetBridge;
                talkBridge.body = bridgeMessage;
                talkBridge.send ();
            }
        }
    }
    var newLinksReports = new Array;
    var newLinksLines = new Array;
    var missedLinks = new Array;
    var outLinks = new Array;
    var soleyPlacedList = new Array;
    for (var p = 0; p < targetDocument.pages.length; p++) {
        var parentName = targetDocument.pages[p].name;
        var allPageLinks = new Array;
        var allPageSnippets = new Array;
        for (var g = 0; g < targetDocument.pages[p].allGraphics.length; g++) {
            if (!targetDocument.pages[p].allGraphics[g].itemLink)
                continue;
            allPageLinks.push (targetDocument.pages[p].allGraphics[g].itemLink);
        }
        for (var mtf = 0; mtf < targetDocument.pages[p].allPageItems.length; mtf++) {
            if (targetDocument.pages[p].allPageItems[mtf].extractLabel ("LS_TSID")) {
                allPageSnippets.push (targetDocument.pages[p].allPageItems[mtf]);
            }
        }
        if ((allPageLinks.length + allPageSnippets.length) == 1) {
            if (allPageLinks.length == 1) {
                soleyPlacedList.push ([allPageLinks[0], targetDocument.pages[p], true]); //true means it's link, false means live snippet
            }
            else {
                soleyPlacedList.push ([allPageSnippets[0], targetDocument.pages[p], false]);
            }
        }
        for (var s = 0; s < targetDocument.pages[p].masterPageItems.length; s++) {
            try {
                targetDocument.pages[p].masterPageItems[s].graphics;
            }
            catch (e) {
                continue;
            }
            for (var e = 0; e < targetDocument.pages[p].masterPageItems[s].graphics.length; e++) {
                if (!targetDocument.pages[p].masterPageItems[s].graphics[e].itemLink)
                    continue;
                allPageLinks.push (targetDocument.pages[p].masterPageItems[s].graphics[e].itemLink);
            }
        }
        var myPageWidth = targetDocument.documentPreferences.pageWidth;
        var myPageHeight = targetDocument.documentPreferences.pageHeight;
        myPageWidth = parseFloat (myPageWidth.toString ());
        myPageHeight = parseFloat (myPageHeight.toString ());
        for (var t = 0; t < allPageLinks.length; t++) {
            checkMissedOrOut (allPageLinks[t], missedLinks, outLinks);
            var linkTSIDLabel = allPageLinks[t].extractLabel ("TSID");
            if (!linkTSIDLabel)
                continue;
            var linkIndex = -1;
            for (var n = 0; n < newLinksReports.length; n++) {
                if (linkTSIDLabel == newLinksReports[n][0]) {
                    linkIndex = n;
                    break;
                }
            }
            var pageWithCoordinates = [parentName];
            if (allPageLinks[t].parent.constructor.name != "Story") {
                pageWithCoordinates.push (allPageLinks[t].parent.geometricBounds);
                for (var pwc = 0; pwc < pageWithCoordinates[1].length; pwc++) {
                    pageWithCoordinates[1][pwc] = parseFloat (pageWithCoordinates[1][pwc].toString ());
                    if (pwc == 0 || pwc == 2) {
                        pageWithCoordinates[1][pwc] = (pageWithCoordinates[1][pwc]/myPageHeight) * 100;
                    }
                    else {
                        pageWithCoordinates[1][pwc] = (pageWithCoordinates[1][pwc]/myPageWidth) * 100;
                    }
                    pageWithCoordinates[1][pwc] = pageWithCoordinates[1][pwc].toFixed(2);
                    if (pageWithCoordinates[1][pwc].slice (-3) == ".00") {
                        pageWithCoordinates[1][pwc] = pageWithCoordinates[1][pwc].slice (0, -3);
                    }
                }
            }
            if (linkIndex == -1) {
                var isAlready = false;
                for (var d = 1; d < oldLinksLines.length; d++) {
                    if (linkTSIDLabel == oldLinksLines[d]) {
                        oldLinksLines.splice (d, 1);
                        isAlready = true;
                        break;
                    }
                }
                newLinksLines.push (linkTSIDLabel);
                var linkReport = new Array;
                linkReport.push (linkTSIDLabel);
                linkReport.push (new Array);
                linkReport.push (isAlready? null : allPageLinks[t]);
                newLinksReports.push (linkReport);
                linkIndex = newLinksReports.length - 1;
            }
            var linkPageIndex = -1;
            for (var nlr = 0; nlr < newLinksReports[linkIndex][1].length; nlr++) {
                if (pageWithCoordinates[0] == newLinksReports[linkIndex][1][nlr][0]) {
                    linkPageIndex = nlr;
                    break;
                }
            }
            if (linkPageIndex == -1) {
                newLinksReports[linkIndex][1].push ([pageWithCoordinates[0]]);
                linkPageIndex = newLinksReports[linkIndex][1].length - 1;
            }
            if (pageWithCoordinates.length > 1) {
                newLinksReports[linkIndex][1][linkPageIndex].push (pageWithCoordinates[1].join (","));
            }
        }
        for (var aps = 0; aps < allPageSnippets.length; aps++) {
            var snippetTSIDLabel = allPageSnippets[aps].extractLabel("LS_TSID");
            if (!snippetTSIDLabel)
                continue;
            var snippetIndex = -1;
            for (var n = 0; n < newLinksReports.length; n++) {
                if (snippetTSIDLabel == newLinksReports[n][0]) {
                    snippetIndex = n;
                    break;
                }
            }
            var textCoordinates = null;
            var isNoTextCoordinates = (targetDocument.extractLabel("TS_NO_TEXT_COORDINATES") != "No");
            if (!isNoTextCoordinates) {
                var lastCharIndex = new Array;
                var nestedMarkers = new Array;
                if (getContainerSnippet (allPageSnippets[aps], lastCharIndex, nestedMarkers, "first marker")) {
                    if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") != "lonely" && nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") != "spoiled") {
                        textCoordinates = getTextCoordinates (nestedMarkers[0][0].parent.characters.itemByRange(nestedMarkers[0][0], nestedMarkers[nestedMarkers.length-1][0]), null, myPageWidth, myPageHeight);
                    }    
                }
            }
            if (snippetIndex == -1) {
                var isAlready = false;
                for (var d = 1; d < oldLinksLines.length; d++) {
                    if (snippetTSIDLabel == oldLinksLines[d]) {
                        oldLinksLines.splice (d, 1);
                        isAlready = true;
                        break;
                    }
                }
                newLinksLines.push (snippetTSIDLabel);
                var linkReport = new Array;
                linkReport.push (snippetTSIDLabel);
                linkReport.push (new Array);
                linkReport.push (isAlready? null : allPageSnippets[aps]);
                newLinksReports.push (linkReport);
                snippetIndex = newLinksReports.length - 1;
            }
            if (!textCoordinates) {
                textCoordinates = [[parentName]];
            }
            for (var tc = 0; tc < textCoordinates.length; tc++) {
                var snippetPageIndex = -1;
                for (var nlr = 0; nlr < newLinksReports[snippetIndex][1].length; nlr++) {
                    if (textCoordinates[tc][0] == newLinksReports[snippetIndex][1][nlr][0]) {
                        snippetPageIndex = nlr;
                        break;
                    }
                }
                if (snippetPageIndex == -1) {
                    newLinksReports[snippetIndex][1].push ([textCoordinates[tc][0]]);
                    snippetPageIndex = newLinksReports[snippetIndex][1].length - 1;
                }
                if (textCoordinates[tc].length > 1) {
                    newLinksReports[snippetIndex][1][snippetPageIndex].push (textCoordinates[tc][1].join (","));
                }
            }
        }
    }
    for (var k = 0; k < targetDocument.links.length; k++) {
        if (targetDocument.links[k].linkType == "XML") {
            checkMissedOrOut (targetDocument.links[k], missedLinks, outLinks);
            var linkTSIDLabel = targetDocument.links[k].extractLabel ("TSID");
            if (!linkTSIDLabel)
                continue;
            var isAlready = false;
            for (var d = 1; d < oldLinksLines.length; d++) {
                if (linkTSIDLabel == oldLinksLines[d]) {
                    oldLinksLines.splice (d, 1);
                    isAlready = true;
                    break;
                }
            }
            newLinksLines.push (linkTSIDLabel);
            var linkReport = new Array;
            linkReport.push (linkTSIDLabel);
            linkReport.push (new Array);
            linkReport[1].push (["1"]);
            linkReport.push (isAlready? null : targetDocument.links[k]);
            newLinksReports.push (linkReport);
        }
    }
    for (var n = 0; n < newLinksReports.length; n++) {
        for (var q = 0; q < newLinksReports[n][1].length; q++) {
            newLinksReports[n][1][q] = newLinksReports[n][1][q].join ("|");
        }
        var docReportFile = new File (dataPath + "/IDs" + newLinksReports[n][0] + "/Workshop/Places/" + docDotID);
        if (!docReportFile.parent.exists)
            docReportFile.parent.create ();
        if (!writeFile (docReportFile, newLinksReports[n][1].join ("\n")))
            return false;
    }
    for (var e = 1; e < oldLinksLines.length; e++) {
        var docReportFile = new File (dataPath + "/IDs" + oldLinksLines[e] + "/Workshop/Places/" + docDotID);
        docReportFile.remove ();
    }
    var newContent = newLinksLines.length == 0? userID : (userID + "\n" + newLinksLines.join ("\n"));
    if (!linksReportFile.parent.exists)
        linksReportFile.parent.create ();
    if (!writeFile (linksReportFile, newContent))
        return false;
    if (missedLinks.length > 0) {
        writeEncodedFile (linksMissedFile, missedLinks.join ("\n"));
    }
    else {
        linksMissedFile.remove ();
    }
    if (outLinks.length > 0) {
        writeEncodedFile (linksOutFile, outLinks.join ("\n"));
    }
    else {
        linksOutFile.remove ();
    }
    //auto update links records
    if (app.extractLabel("autoUpdateLinksRecords") == "Yes") {
        var toBeLabeled = new Array;
        for (var auto = 0; auto < newLinksReports.length; auto++) {
            if (newLinksReports[auto][2]) {
                if (newLinksReports[auto][2].constructor.name == "TextFrame") {
                    var toLabelSnippetFile = findTargetFile (targetDocument, newLinksReports[auto][2].parent);
                    if (toLabelSnippetFile) {
                        if (toLabelSnippetFile instanceof File) {
                            toBeLabeled.push (toLabelSnippetFile);
                        }
                    }
                }
                else {
                    toBeLabeled.push (new File (preparePath (newLinksReports[auto][2].filePath)));
                }
            }
        }
        for (var ol = 1; ol < oldLinksLines.length; ol++) {
            var fileID_Path_newID = [oldLinksLines[ol]];
            getPath (fileID_Path_newID);
            if (fileID_Path_newID[1]) {
                toBeLabeled.push (new File (workshopPath + fileID_Path_newID[1]));
            }
        }
        var bridgeTalkBody = "";
        var newLine = "";
        for (var tbl = 0; tbl < toBeLabeled.length; tbl++) {
            if (toBeLabeled[tbl].exists) {
                bridgeTalkBody += newLine + "labelWorkshopFile (new File (\"" + toBeLabeled[tbl].fsName.replace(/\\/g, "/") + "\"), [null], true);";
                newLine = "\n";
            }
        }
        var targetBridge = BridgeTalk.getSpecifier ("bridge");
        if (targetBridge && targetBridge.appStatus != "not installed") {
            var talkBridge = new BridgeTalk;
            talkBridge.target = targetBridge;
            talkBridge.body = bridgeTalkBody;
            talkBridge.send ();
        }

    }
    //exportPagesThumbnail

    var pagesFolder = Folder (docName.replace (workshopPath, pagesPath + "/Workshop"));

    if (!isStartImagesExporting) {
        if (pagesFolder.exists) {
            var allFiles = pagesFolder.getFiles (isUnhiddenFile);
            for (var f = 0; f < allFiles.length; f++) {
                allFiles[f].remove ();
            }
            pagesFolder.remove();
        }
    }
    else {
        if (pagesFolder.exists) {
            var allFiles = pagesFolder.getFiles (isUnhiddenFile);
            for (var f = 0; f < allFiles.length; f++) {
                allFiles[f].remove ();
            }
        }
        else {
            pagesFolder.create ();
        }
        var pageFileExtension = ".jpg";
        var exportFormatEnum = ExportFormat.JPG;
        if (targetDocument.extractLabel ("TS_PAGES_EXTENSION") == "png") {
            pageFileExtension = ".png";
            exportFormatEnum = ExportFormat.PNG_FORMAT;
            var docPngRes = targetDocument.extractLabel ("TS_PAGES_RES");
            if (docPngRes != "") {
                docPngRes = parseInt (docPngRes, 10);
            }
            else {
                docPngRes = 120;
            }
            var docPngQua = targetDocument.extractLabel ("TS_PAGES_QUALITY");
            if (docPngQua != "") {
                switch (docPngQua) {
                    case "LOW":
                        docPngQua = PNGQualityEnum.LOW;
                        break;
                    case "MEDIUM":
                        docPngQua = PNGQualityEnum.MEDIUM;
                        break;
                    case "HIGH":
                        docPngQua = PNGQualityEnum.HIGH;
                        break;
                    case "MAXIMUM":
                        docPngQua = PNGQualityEnum.MAXIMUM;
                        break;
                }
            }
            else {
                docPngQua = PNGQualityEnum.LOW;
            }
            with (app.pngExportPreferences)
            {
                antiAlias = true;
                transparentBackground = true;
                useDocumentBleeds = false;
                pngColorSpace = PNGColorSpaceEnum.RGB;
                exportingSpread = false;
                pngExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
                pngQuality = docPngQua; //LOW, MEDIUM, HIGH, MAXIMUM
                exportResolution = docPngRes; //1 - 2400
            }
        }
        else {
            var docJpegRes = targetDocument.extractLabel ("TS_PAGES_RES");
            if (docJpegRes != "") {
                docJpegRes = parseInt (docJpegRes, 10);
            }
            else {
                docJpegRes = 120;
            }
            var docJpegQua = targetDocument.extractLabel ("TS_PAGES_QUALITY");
            if (docJpegQua != "") {
                switch (docJpegQua) {
                    case "LOW":
                        docJpegQua = JPEGOptionsQuality.LOW;
                        break;
                    case "MEDIUM":
                        docJpegQua = JPEGOptionsQuality.MEDIUM;
                        break;
                    case "HIGH":
                        docJpegQua = JPEGOptionsQuality.HIGH;
                        break;
                    case "MAXIMUM":
                        docJpegQua = JPEGOptionsQuality.MAXIMUM;
                        break;
                }
            }
            else {
                docJpegQua = JPEGOptionsQuality.LOW;
            }
            with (app.jpegExportPreferences)
            {
                exportingSpread = false;
                jpegExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
                jpegQuality = docJpegQua; //LOW, MEDIUM, HIGH, MAXIMUM
                jpegRenderingStyle = JPEGOptionsFormat.BASELINE_ENCODING; //PROGRASSIVE_ENCODING
                //pageString = "1-2";
                exportResolution = docJpegRes; //1 - 2400
            }
        }
        var thePageRange = null;
        var isToLeaveEmptyPages = (targetDocument.extractLabel("TS_LEAVE_EMPTY_PAGES") != "No");
        if (isToLeaveEmptyPages) {
            thePageRange = tsGetUsedPagesRange (targetDocument);
        }
        var lastPageReached = false;
        //saveFilesList[x][0] page thumbnail file object
        //saveFilesList[x][1] page id
        var saveFilesList = new Array;
        for (var p = 0; p < targetDocument.pages.length; p++) {
            if (lastPageReached) {
                break;
            }
            var currentPage = targetDocument.pages.item(p);
            if (thePageRange != null) {
                if (currentPage.name == thePageRange[1]) {
                    lastPageReached = true;
                }
            }
            if (pageFileExtension == ".png") {
                app.pngExportPreferences.pageString = currentPage.name;
            }
            else {
                app.jpegExportPreferences.pageString = currentPage.name;
            }
            var pageFileName = tsFillZerosIfAllDigits (currentPage.name, 4) + pageFileExtension;
            var saveFile = new File (pagesFolder.fsName.replace(/\\/g, "/") + "/" + pageFileName);
            targetDocument.exportFile (exportFormatEnum, saveFile, false);
            saveFilesList.push ([saveFile, currentPage.id]);
        }
        if (pageFileExtension != ".jpg") {
            pageFileExtension = ".jpg";
            exportFormatEnum = ExportFormat.JPG;
            var docJpegRes = targetDocument.extractLabel ("TS_PAGES_RES");
            if (docJpegRes != "") {
                docJpegRes = parseInt (docJpegRes, 10);
            }
            else {
                docJpegRes = 120;
            }
            var docJpegQua = targetDocument.extractLabel ("TS_PAGES_QUALITY");
            if (docJpegQua != "") {
                switch (docJpegQua) {
                    case "LOW":
                        docJpegQua = JPEGOptionsQuality.LOW;
                        break;
                    case "MEDIUM":
                        docJpegQua = JPEGOptionsQuality.MEDIUM;
                        break;
                    case "HIGH":
                        docJpegQua = JPEGOptionsQuality.HIGH;
                        break;
                    case "MAXIMUM":
                        docJpegQua = JPEGOptionsQuality.MAXIMUM;
                        break;
                }
            }
            else {
                docJpegQua = JPEGOptionsQuality.LOW;
            }
            with (app.jpegExportPreferences)
            {
                exportingSpread = false;
                jpegExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
                jpegQuality = docJpegQua; //LOW, MEDIUM, HIGH, MAXIMUM
                jpegRenderingStyle = JPEGOptionsFormat.BASELINE_ENCODING; //PROGRASSIVE_ENCODING
                exportResolution = docJpegRes; //1 - 2400
            }
        }
        for (var spl = 0; spl < soleyPlacedList.length; spl++) {
            var linkFile = null;
            if (soleyPlacedList[spl][2]) {
                linkFile = new File (preparePath (soleyPlacedList[spl][0].filePath));
                if (linkFile.name.slice (-5) == ".indd") {
                    continue;
                }
            }
            else {
                var snippetID = soleyPlacedList[spl][0].extractLabel ("LS_TSID");
                if (snippetID) {
                    var fileID_Path_newID = [snippetID];
                    getPath (fileID_Path_newID);
                    if (fileID_Path_newID[1]) {
                        linkFile = new File (workshopPath + fileID_Path_newID[1]);
                    }
                }
            }
            if (linkFile != null && linkFile.exists) {  
                app.jpegExportPreferences.pageString = soleyPlacedList[spl][1].name;
                var thumbFile = new File (linkFile.parent.fsName.replace(/\\/g, "/") + "/.TREESHADE/.THUMB/" + File.decode (linkFile.name) + ".jpg");
                if (!thumbFile.parent.exists) {
                    thumbFile.parent.create ();
                }
                targetDocument.exportFile (exportFormatEnum, thumbFile, false);
                tsSetMetadata (thumbFile, "TS_THUMB_FABRICATED", "ASSIGNED");
            }
        }
        for (p = 0; p < saveFilesList.length; p++) {
            setPgDocID (saveFilesList[p][0], docID, saveFilesList[p][1]);
        }
        var dataFolder = new Folder (workshopPath + docPath + "/Data");
        if (dataFolder.exists) {
            var targetBridge = BridgeTalk.getSpecifier ("bridge");
            if (targetBridge && targetBridge.appStatus != "not installed") {
                var talkBridge = new BridgeTalk;
                talkBridge.target = targetBridge;
                talkBridge.body = "tsThumbFromPageFolder (new Folder (\"" + workshopPath + docPath + "/Data" + "\"));";
                talkBridge.send ();
            }
        }
    }
}

function tsSetMetadata (targetImage, key, value) {
    /**///$.writeln ($.line);
    try {
        if (ExternalObject.AdobeXMPScript == undefined)
              ExternalObject.AdobeXMPScript = new ExternalObject('lib:AdobeXMPScript');
        XMPMeta.registerNamespace(treeShadeNamespace, treeShadePrefix);
        var xmpUpdateFile = new XMPFile(targetImage.fsName, XMPConst.UNKNOWN, XMPConst.OPEN_FOR_UPDATE);
        var xmpUpdate = xmpUpdateFile.getXMP();
        xmpUpdate.setProperty( treeShadeNamespace, key, value);
        if (xmpUpdateFile.canPutXMP(xmpUpdate)) {
            xmpUpdateFile.putXMP(xmpUpdate);
        }
        xmpUpdateFile.closeFile (XMPConst.CLOSE_UPDATE_SAFELY);
        return true;
    }
    catch (error) {
        returned = false;
    }
}

function tsGetText (theLabelArr, theOldText, isMultiLine) {
    /*
    Code for Import https://scriptui.joonas.me — (Triple click to select): 
    {"activeId":0,"items":{"item-0":{"id":0,"type":"Dialog","parentId":false,"style":{"text":"Import Multiple PDF pages","preferredSize":[0,0],"margins":16,"orientation":"column","spacing":10,"alignChildren":["left","top"],"varName":null,"windowType":"Dialog","creationProps":{"su1PanelCoordinates":false,"maximizeButton":false,"minimizeButton":false,"independent":false,"closeButton":true,"borderless":false,"resizeable":false},"enabled":true}},"item-1":{"id":1,"type":"Panel","parentId":20,"style":{"text":"Page Selection","preferredSize":[0,99],"margins":10,"orientation":"column","spacing":10,"alignChildren":["left","top"],"alignment":null,"varName":null,"creationProps":{"borderStyle":"etched","su1PanelCoordinates":false},"enabled":true}},"item-3":{"id":3,"type":"EditText","parentId":6,"style":{"text":"1","preferredSize":[600,0],"alignment":null,"varName":null,"helpTip":null,"softWrap":false,"creationProps":{"noecho":false,"readonly":false,"multiline":false,"scrollable":false,"borderless":false,"enterKeySignalsOnChange":false},"enabled":true,"justify":"left"}},"item-4":{"id":4,"type":"StaticText","parentId":1,"style":{"text":"To edit long text use\n<a href = 'http://docs.google.com/?action=newdoc'>Google New Document</a>","justify":"left","preferredSize":[0,0],"alignment":null,"varName":null,"helpTip":null,"softWrap":true,"creationProps":{"truncate":"none","multiline":true,"scrolling":false},"enabled":true}},"item-6":{"id":6,"type":"Group","parentId":1,"style":{"preferredSize":[0,0],"margins":0,"orientation":"row","spacing":10,"alignChildren":["left","center"],"alignment":null,"varName":null,"enabled":true}},"item-20":{"id":20,"type":"Group","parentId":0,"style":{"preferredSize":[0,0],"margins":0,"orientation":"column","spacing":10,"alignChildren":["fill","top"],"alignment":null,"varName":null,"enabled":true}},"item-36":{"id":36,"type":"Group","parentId":0,"style":{"preferredSize":[0,0],"margins":0,"orientation":"row","spacing":10,"alignChildren":["left","top"],"alignment":null,"varName":null,"enabled":true}},"item-37":{"id":37,"type":"Button","parentId":36,"style":{"text":"OK","justify":"center","preferredSize":[0,0],"alignment":null,"varName":"ok","helpTip":null,"enabled":true}},"item-38":{"id":38,"type":"Button","parentId":36,"style":{"text":"Cancel","justify":"center","preferredSize":[0,0],"alignment":null,"varName":"cancel","helpTip":null,"enabled":true}}},"order":[0,20,1,6,3,4,36,37,38],"settings":{"importJSON":true,"indentSize":false,"cepExport":false,"includeCSSJS":true,"functionWrapper":false,"compactCode":false,"showDialog":true,"itemReferenceList":"None"}}
    */ 

    // DIALOG
    // ======
    var dialog = new Window("dialog"); 
        dialog.text = "Tree Shade"; 
        dialog.orientation = "column"; 
        dialog.alignChildren = ["left","top"]; 
        dialog.spacing = 10; 
        dialog.margins = 16; 

    // GROUP1
    // ======
    var group1 = dialog.add("group", undefined, {name: "group1"}); 
        group1.orientation = "column"; 
        group1.alignChildren = ["fill","top"]; 
        group1.spacing = 10; 
        group1.margins = 0; 

    var staticText = group1.add('edittext {properties: {name: "edittext1", readonly: true, multiline: true, scrollable: true}}');
    staticText.preferredSize.width = 600; 
    staticText.preferredSize.height = (15 * theLabelArr.length) + 10;
    if (staticText.preferredSize.height > 200) {
        staticText.preferredSize.height = 200;
    }
    for (var tla = 0; tla < theLabelArr.length; tla++) {
        staticText.text += theLabelArr[tla];
        if (tla < theLabelArr.length - 1) {
            staticText.text += "\n";
        }
    }

    if (isMultiLine) {
        tsEditText = group1.add('edittext {properties: {name: "tsEditText", multiline: true, scrollable: true}}');
        tsEditText.preferredSize.width = 600;
        tsEditText.preferredSize.height = 100;
    }
    else {
        tsEditText = group1.add('edittext {properties: {name: "tsEditText"}}');
        tsEditText.preferredSize.width = 600;
    }
    tsEditText.text = theOldText; 
    tsEditText.addEventListener ('focus', edtActivate, false);
    function edtActivate () {

    }

    // GROUP2
    // ======
    var group2 = dialog.add("group", undefined, {name: "group2"}); 
    group2.orientation = "row"; 
    group2.alignChildren = ["fill","top"]; 
    group2.spacing = 10; 
    group2.margins = 0; 

    var group3 = group2.add("group", undefined, {name: "group3"}); 
    group3.orientation = "column"; 
    group3.alignChildren = ["left","top"]; 
    group3.spacing = 10; 
    group3.margins = 0; 
    if (isMultiLine) {
        var button0 = group3.add("button", undefined, undefined, {name: "button0"}); 
        button0.text = "Text Editor"; 
        button0.onClick = function() {
            var textEditorFile = new File (dataPath + "/Text Editor.htm");
            if (textEditorFile && textEditorFile.exists) {
                var targetBridge = BridgeTalk.getSpecifier ("bridge");
                if (targetBridge && targetBridge.appStatus != "not installed") {
                    var talkBridge = new BridgeTalk;
                    talkBridge.target = targetBridge;
                    talkBridge.body = "var toOpenThumbnail = new Thumbnail (new File (\"" + textEditorFile.fsName.replace(/\\/g, "/") + "\"));"
                                    + "\n" + "toOpenThumbnail.open ();";
                    talkBridge.send ();
                }
            }
        };
    }
    dialog.addEventListener ('focus', dlgActivate, false);
    function dlgActivate () {
        tsEditText.active = true;
        if (tsEditText.preferredSize.height == 100) {
            tsEditText.textselection = "";
        }
    }

    var group4 = group2.add("group", undefined, {name: "group4"}); 
    group4.orientation = "column"; 
    group4.alignChildren = ["fill","top"]; 
    group4.spacing = 10;
    group4.margins = 0;

    var group5 = group4.add("group", undefined, {name: "group5"}); 
    group5.orientation = "row"; 
    group5.alignChildren = ["right","top"]; 
    group5.spacing = 10;
    group5.margins = 0;

    var statictext2 = group5.add("statictext", undefined, "", {name: "statictext2"});
    statictext2.preferredSize.width = 230;
    var button1 = group5.add("button", undefined, undefined, {name: "button1"}); 
    button1.text = "Cancel"; 
    button1.onClick = function() {
        this.parent.parent.close(null);
    };

    var button2 = group5.add("button", undefined, undefined, {name: "button2"}); 
    if (isMultiLine) {
        button2.text = "OKAY";
    }
    else {
        button2.text = "OK";
    }
    button2.preferredSize.width = 100;
    button2.onClick = function() {
        this.parent.parent.close(1);
    }; 

    var showResult = dialog.show();
    if (showResult == 1) {
        return tsEditText.text;
    }
    else {
        return null;
    }
}

function solveShift (phrase) {
    while (phrase.search (/<Shift\-?\d+(\.\d+)?>/i) != -1) {
        var shiftPhrase = phrase.match (/<Shift\-?\d+(\.\d+)?>/i)[0];
        var theNum = phrase.match (/\-?\d+(\.\d+)?<Shift/i);
        if (theNum) {
            theNum = theNum[0];
            var isFactorNegative = false;
            var shiftFactor = shiftPhrase.match (/\-?\d+(\.\d+)?/)[0];
            if (shiftFactor[0] == '-') {
                isFactorNegative = true;
                shiftFactor = shiftFactor.slice (1);
            }
            var decimalCount = 0;
            var isDecimalForce = false;
            if (shiftFactor.indexOf (".") != -1) {
                var decimalPart = shiftFactor.slice (shiftFactor.indexOf (".") + 1);
                if (decimalPart.length > 0) {
                    if (decimalPart[0] == '1') {
                        isDecimalForce = true;
                    }
                    decimalCount = decimalPart.length;
                }
                shiftFactor = shiftFactor.slice (0, shiftFactor.indexOf ("."));
            }
            var digitsCount = shiftFactor.length;
            var isToFill = false;
            if (shiftFactor) {
                if (shiftFactor[0] == '0' && shiftFactor.length > 1) {
                    isToFill = true;
                }
                if (isFactorNegative) {
                    shiftFactor = "-" + shiftFactor;
                }
                shiftFactor = parseInt (shiftFactor, 10);
            }
            else {
                shiftFactor = 0;
            }
            theNum = theNum.slice (0, -6);
            theNum = parseFloat (theNum);
            if (shiftFactor != 0) {
                theNum += shiftFactor;
            }
            var numNegative = "";
            theNum = theNum.toString();
            if (theNum) {
                if (theNum[0] == '-') {
                    numNegative = "-";
                    theNum = theNum.slice (1);
                }
                var leftPart = theNum;
                var rightPart = "";
                if (theNum.indexOf (".") != -1) {
                    leftPart = theNum.slice (0, theNum.indexOf ("."));
                    rightPart = theNum.slice (theNum.indexOf (".") + 1);
                }
                if (rightPart != "") {
                    if (rightPart.length > decimalCount) {
                        if (rightPart[decimalCount] == '5' || rightPart[decimalCount] == '6' || rightPart[decimalCount] == '7' || rightPart[decimalCount] == '8' || rightPart[decimalCount] == '9') {
                            rightPart = rightPart.slice (0, decimalCount);
                            var beforeLen = rightPart.length;
                            rightPart = parseInt (rightPart, 10) + 1;
                            rightPart = rightPart.toString ();
                            if (rightPart.length > beforeLen) {
                                rightPart = "";
                                var zeroCount = 0;
                                while (zeroCount < decimalCount) {
                                    rightPart += "0";
                                    zeroCount++;
                                }
                                leftPart = parseInt(leftPart, 10) + 1;
                                leftPart = leftPart.toString ();
                            }
                        }
                        else {
                            rightPart = rightPart.slice (0, decimalCount);
                        }
                    }
                }
                if (isToFill) {
                    leftPart = tsFillZeros (parseInt(leftPart, 10), digitsCount);
                }
                if (isDecimalForce) {
                    while (rightPart.length < decimalCount) {
                        rightPart += "0";
                    }
                }
                if (rightPart != "") {
                    rightPart = "." + rightPart;
                }
                theNum = numNegative + leftPart + rightPart;
                phrase = phrase.replace (/\-?\d+(\.\d+)?<Shift/i, theNum + "<Shift");
            }
        }
        phrase = phrase.replace (/<Shift\-?\d+(\.\d+)?>/i, "");
    }
    return phrase;
}

function solveNameAt (phrase) {
    var i = 0;
    while (phrase.search (/<Name_At>\d+/i) != -1) {
        i++;
        if (i > 10) {
            return phrase;
        }
        if (phrase.search (/<Name_At>\d+/i) != phrase.search (/<Name_At>/i)) {
            return phrase;
        }
        var nameAtPhrase = phrase.match (/<Name_At>\d+/i)[0];
        var orderDigits = nameAtPhrase.slice (9);
        var orderNum = parseInt (orderDigits, 10);
        if (orderNum == 0) {
            return phrase;
        }
        var pathPrefix = phrase.slice (0, phrase.search (/<Name_At>/i));
        var isForFolder = (phrase.slice (phrase.search (/<Name_At>/i)).indexOf ("/") != -1);
        if (pathPrefix.indexOf ("/") == -1) {
            return phrase;
        }
        var pathIndex = -1;
        var now = new Date ();
        if (nameAtPathsList.length > 50) {
            nameAtPathsList.splice (40, nameAtPathsList.length);
        }
        for (var napl = 0; napl < nameAtPathsList.length; napl++) {
            if (nameAtPathsList[napl][0] == pathPrefix) {
                if (now - nameAtPathsList[napl][1] > 30000) {
                    nameAtPathsList.splice (napl, 1);
                    break;
                }
                else {
                    pathIndex = napl;
                    break;
                }
            }
        }
        var nameAtPrefix = pathPrefix.slice (pathPrefix.lastIndexOf ("/")+1);
        if (pathIndex == -1) {
            var nameAtFolder = Folder (pathPrefix.slice (0, pathPrefix.lastIndexOf ("/")));
            var namesList = new Array;
            if (nameAtFolder.exists) {
                if (isForFolder) {
                    var allFolders = nameAtFolder.getFiles (isCSFolder);
                    sortFilesList (nameAtFolder, allFolders, true);
                    for (var afs = 0; afs < allFolders.length; afs++) {
                        var testName = File.decode (allFolders[afs].name);
                        if (testName.indexOf (nameAtPrefix) == 0) {
                            namesList.push (testName);
                        }
    
                    }
                }
                else {
                    var allFiles = nameAtFolder.getFiles (isUnhiddenFile);
                    sortFilesList (nameAtFolder, allFiles, false);
                    for (var afl = 0; afl < allFiles.length; afl++) {
                        var testName = File.decode (allFiles[afl].name);
                        if (testName.indexOf (".") != -1) {
                            testName = testName.slice (0, testName.lastIndexOf ("."));
                        }
                        if (testName.indexOf (nameAtPrefix) == 0) {
                            namesList.push (testName);
                        }
    
                    }
                }
                //nameAtPathsList
                //nameAtPathsList[x][0] path with prefix
                //nameAtPathsList[x][1] adding time
                //nameAtPathsList[x][2] names list
                nameAtPathsList.unshift ([pathPrefix, now, namesList]);
                pathIndex = 0;
            }
            else {
                return phrase;
            }
        }
        if (orderNum <= nameAtPathsList[pathIndex][2].length) {
            phrase = phrase.replace (new RegExp(escapeRegExp (nameAtPrefix) + "<[Nn][Aa][Mm][Ee]_[Aa][Tt]>" + escapeRegExp(orderDigits), 'g'), nameAtPathsList[pathIndex][2][orderNum-1]);
        }
        else {
            return phrase;
        }
    }
    return phrase;
}

function escapeRegExp(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function writeFile (targetFile, fileNewContent) {
    try {
        targetFile.open ("write");
        targetFile.write (fileNewContent);
        targetFile.close ();
    }
    catch (e) {return false;}
    return true;
}

function writeEncodedFile (targetFile, fileNewContent) {
    /*function_start*///alert ("writeEncodedFile-start");
    try {
        targetFile.open ("write");
        targetFile.write (File.encode (fileNewContent));
        targetFile.close ();
    }
    catch (e) {return false;}
    return true;
}

function checkMissedOrOut (targetLink, missedLinks, outLinks) {
    /*function_start*///alert ("checkMissedOrOut-start");
    var linkFullPath = preparePath (targetLink.filePath);
    var isMissed = (targetLink.status == LinkStatus.LINK_MISSING)? true : false;
    if (isMissed) {
        var isRepeated = false;
        for (var m = 0; m < missedLinks.length; m++) {
            if (missedLinks[m] == linkFullPath) {
                isRepeated = true;
                break;
            }
        }
        if (!isRepeated) {
            missedLinks.push (linkFullPath);
        }
    }
    var isOut = (linkFullPath.indexOf (workshopPath + "/") == 0)? false : true;
    if (isOut) {
        var isRepeated = false;
        for (var o = 0; o < outLinks.length; o++) {
            if (outLinks[o] == linkFullPath) {
                isRepeated = true;
                break;
            }
        }
        if (!isRepeated) {
            outLinks.push (linkFullPath);
        }
    }
    var result = 1;
    result += isMissed? 2 : 0;
    result += isOut? 3 : 0;
    return result;
}

function tsFillZerosIfAllDigits (digits, digitsCount) {
    /*function_start*///alert ("tsFillZerosIfAllDigits-start");
    for (var d = 0; d < digits.length; d++) {
        if (digits.charCodeAt (d) < 48 || digits.charCodeAt (d) > 57)
            return digits;
    }
	while (digits.length < digitsCount)
		digits = "0" + digits;
	return digits;
}

function linkAfterUpdateHandler (myEvent) {
    /*function_start*///alert ("linkAfterUpdateHandler-start");
    var targetLink = myEvent.parent;
    if (!targetLink instanceof Link)
        return;
    /*var targetLinkID = targetLink.extractLabel ("TSID");
    var targetLinkFullPath = preparePath (targetLink.filePath);
    if (targetLinkID) {
        for (var c = 0; c < beforeUpdateList.length; c++) {
            if (targetLinkID == beforeUpdateList[c][0]) {
                if (targetLinkFullPath != beforeUpdateList[c][1]) { //The updating is a relinking, so we must clear old ID
                    targetLink.insertLabel ("TSID", "");
                    prepareLink (targetLink);
                    beforeUpdateList.splice (c, 1);
                    return true;
                }
                beforeUpdateList.splice (c, 1);
                return false;
            }
        }        
    }*/
    var speedChain = new Array;
    testLink (targetLink, speedChain, true, null);
}

function linkBeforeUpdateHandler (myEvent) {
    /*function_start*///alert ("linkBeforeUpdateHandler-start");
    var targetLink = myEvent.parent;
    if (!targetLink instanceof Link)
        return;
    /*var targetLinkID = targetLink.extractLabel ("TSID");
    var targetLinkFullPath = preparePath (targetLink.filePath);
    if (targetLinkID) {
        var linkIDTwin = new Array;
        linkIDTwin.push (targetLinkID);
        linkIDTwin.push (targetLinkFullPath);
        beforeUpdateList.push (linkIDTwin);
    }*/
    if (isShade (File (targetLink.filePath))) {
        myEvent.stopPropagation();
        myEvent.preventDefault();
    }
}

function prepareLink (targetLink) {
    /*function_start*///alert ("prepareLink-start");
    var linkFile = new File (preparePath (targetLink.filePath));
    var fileID = getTSID (linkFile);
    if (fileID) {
        targetLink.insertLabel ("TSID", fileID);
            return fileID;
    }
    else {
        addToBridgeScanningQueue (linkFile.parent.fsName.replace(/\\/g, "/"));
    }
    return false;
}

function isShade (targetFile) {
    /**///$.writeln ("tsIsShade " + targetFile.fsName.replace(/\\/g, "/") + " " + $.line);
    if (targetFile.length > 128) {
        return false;
    }
    else {
        var fileID = readFile (targetFile);
        if (fileID) {
            fileID = fileID.split (":");
            if (fileID.length > 1) {
                if (fileID[0] == "TS_ID") {
                    return fileID[1];
                }
                else {
                    return false;
                }
            }
            else {
                return false;
            }
        }
        else {
            return null;
        }
    }
}

function getTSID (targetFile) {  
    /**///$.writeln ($.line);
    var shadeFileID = isShade (targetFile);
    if (shadeFileID) {
        return shadeFileID;
    }
    isCS = isCSFile (targetFile);
    if (isCS) {
        try {
            if (ExternalObject.AdobeXMPScript == undefined)
                  ExternalObject.AdobeXMPScript = new ExternalObject('lib:AdobeXMPScript');
            var xmpFile = new XMPFile(targetFile.fsName.replace(/\\/g, "/"), XMPConst.UNKNOWN, XMPConst.OPEN_FOR_READ);
            var xmp = xmpFile.getXMP();
            var fileID = xmp.getProperty(treeShadeNamespace, "TS_ID");
            if (!fileID)
                fileID = "";
            fileID = fileID + "";
            xmpFile.closeFile(XMPConst.CLOSE_UPDATE_SAFELY);
            return fileID;
        }
        catch (e) {
            return null;
        }
    }
    else {
        var idFile = new File (targetFile.parent.fsName.replace(/\\/g, "/") + "/.TREESHADE/" + File.decode (targetFile.name));
        if (idFile.exists) {
            var fileID = readFile (idFile);
            if (fileID) {
                fileID = fileID.split (":");
                if (fileID.length > 1) {
                    return fileID[1];
                }
            }
        }
        return "";
    }
}

function setPgDocID (targetFile, targetID, pageID) {
    /**///$.writeln ($.line);
    if (!targetFile.exists) {
        return false;
    }
    else {
        try {
            if (ExternalObject.AdobeXMPScript == undefined)
                  ExternalObject.AdobeXMPScript = new ExternalObject('lib:AdobeXMPScript');
             XMPMeta.registerNamespace(treeShadeNamespace, treeShadePrefix);
            var xmpUpdateFile = new XMPFile(targetFile.fsName.replace(/\\/g, "/"), XMPConst.UNKNOWN, XMPConst.OPEN_FOR_UPDATE);
            var xmpUpdate = xmpUpdateFile.getXMP();
            xmpUpdate.setProperty( treeShadeNamespace, "TS_PG_DOCID", targetID);
            xmpUpdate.setProperty( treeShadeNamespace, "TS_PG_ID", pageID);
            if (xmpUpdateFile.canPutXMP(xmpUpdate)) {
                xmpUpdateFile.putXMP(xmpUpdate);
                xmpUpdateFile.closeFile (XMPConst.CLOSE_UPDATE_SAFELY);
            }
            else {
                xmpUpdateFile.closeFile (XMPConst.CLOSE_UPDATE_SAFELY);
                return false;
            }
        }
        catch (error) {
            return false;
        }
    }
    return true;
}

function isCSFile (targetFile) {
    if (targetFile instanceof Folder) {
        return false;
    }
    if (targetFile.hidden)
        return false;
    var dotIndex = File.decode (targetFile.name).lastIndexOf (".");
    if (dotIndex == -1)
        return false;
    else {
        var fileExtension = File.decode (targetFile.name).slice (dotIndex);
        fileExtension = fileExtension.toLowerCase();
        for (var c = 0; c < CSExtensionList.length; c++) {
            if (fileExtension == CSExtensionList[c]) {
                return true;
            }
        }
    }
    return false;
}

function isCSFolder (targetFile) {
    if (targetFile instanceof Folder) {
        if (File.decode (targetFile.name) == "pdf" || File.decode (targetFile.name).indexOf ("untitled folder") == 0 || File.decode (targetFile.name)[0] == '.') 
            return false;
        return true;
    }
    return false;
}

function readEncodedFile (targetFile) {
    /*function_start*///alert ("readEncodedFile-start");
    var content = "";
    if (targetFile.exists) {
        try {
            targetFile.open ("read");
            var content = File.decode (targetFile.read());
            targetFile.close ();
        }
        catch (e) {
            return false;
        }
        return content;
    }
    else {
        return false;
    }
}

function readFile (targetFile) {
    /*function_start*///alert ("readFile-start");
    var content = "";
    if (targetFile.exists) {
        try {
            targetFile.open ("read");
            content = targetFile.read();
            targetFile.close ();
        }
        catch (e) {return false}
        content = content.replace (/\r\n/g, "\n");
        return content;
    }
    else {
        return false;
    }
}

function readLog (oldLogPath) {
    /*function_start*///alert ("readLog-start");
    var logFile = new File (oldLogPath);
    var logContent = "";
    if (!logFile.exists) return "BROKEN_LOG_LINK";
    try {
        logFile.open ("read");
        logContent = File.decode (logFile.read ());
        logFile.close ();
    }
    catch (e) {
        return "OUT_OF_REACH";
    }
    return logContent;
}

function newLiveSnippetHandler () {
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            //[targetDocument, selectedObject, isWithFormat] 
            var dataArray = new Array;
            dataArray.push(app.activeDocument);
            dataArray.push(app.selection[0]);
            app.doScript(newLiveSnippet, ScriptLanguage.JAVASCRIPT, dataArray, UndoModes.ENTIRE_SCRIPT, "New Live Snippet");
            //newLiveSnippet(dataArray);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
		case "not_saved":
            alert ("Please save the document.");
            exit();
            break;
	}
}

function dynamicPathAndNameHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedObject = app.selection[0];
            var isForcibly = true;
            var selectedObjectAndisForcibly = new Array;
            selectedObjectAndisForcibly.push(selectedObject);
            selectedObjectAndisForcibly.push(isForcibly);
            //app.doScript(dynamicPathAndNameSelected, ScriptLanguage.JAVASCRIPT, selectedObjectAndisForcibly, UndoModes.ENTIRE_SCRIPT, "Update Live Snippet");
            dynamicPathAndNameSelected(selectedObjectAndisForcibly);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
		case "not_saved":
            alert ("Please save the document.");
            exit();
            break;
	}
}

function dynamicLiveSnippetHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedObject = app.selection[0];
            var isForcibly = true;
            var selectedObjectAndisForcibly = new Array;
            selectedObjectAndisForcibly.push(selectedObject);
            selectedObjectAndisForcibly.push(isForcibly);
            dynamicLiveSnippetSelected(selectedObjectAndisForcibly);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
		case "not_saved":
            alert ("Please save the document.");
            exit();
            break;
	}
}

function dynamicStyleHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedObject = app.selection[0];
            var isForcibly = true;
            var selectedObjectAndisForcibly = new Array;
            selectedObjectAndisForcibly.push(selectedObject);
            selectedObjectAndisForcibly.push(isForcibly);
            dynamicStyleSelected(selectedObjectAndisForcibly);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
		case "not_saved":
            alert ("Please save the document.");
            exit();
            break;
	}
}

function dynamicReplaceHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedObject = app.selection[0];
            var isForcibly = true;
            var selectedObjectAndisForcibly = new Array;
            selectedObjectAndisForcibly.push(selectedObject);
            selectedObjectAndisForcibly.push(isForcibly);
            dynamicReplaceSelected(selectedObjectAndisForcibly);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
		case "not_saved":
            alert ("Please save the document.");
            exit();
            break;
	}
}

function placeLiveSnippetHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedObject = app.selection[0];
            app.doScript(tsPlaceLiveSnippetFromUser, ScriptLanguage.JAVASCRIPT, selectedObject, UndoModes.ENTIRE_SCRIPT, "Place Live Snippet");
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
		case "not_saved":
            alert ("Please save the document.");
            exit();
            break;
	}
}

function reflectMarkerHorizontallyHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
		case "not_saved":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedObject = app.selection[0];
            app.doScript(reflectHorizontally, ScriptLanguage.JAVASCRIPT, selectedObject, UndoModes.ENTIRE_SCRIPT, "Reflect Marker Horizontally");
            //reflectHorizontally(selectedObject);
            break;
		case "no_document":
            alert ("Please open or create a document.");
            exit();
            break;
	}
}

function reflectMarkerVerticallyHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
		case "not_saved":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedObject = app.selection[0];
            app.doScript(reflectVertically, ScriptLanguage.JAVASCRIPT, selectedObject, UndoModes.ENTIRE_SCRIPT, "Reflect Marker Vertically");
            //reflectVertically(selectedObject);
            break;
		case "no_document":
            alert ("Please open or create a document.");
            exit();
            break;
	}
}

function unlinkInstanceHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            //[targetDocument, selectedObject, targetNumber] 
            var dataArray = new Array;
            dataArray.push(app.activeDocument);
            dataArray.push(app.selection[0]);
            app.doScript(unlinkInstance, ScriptLanguage.JAVASCRIPT, dataArray, UndoModes.ENTIRE_SCRIPT, "Unlink Instance");
            //unlinkInstance(dataArray);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
		case "not_saved":
            alert ("Please save the document.");
            exit();
            break;
	}
}

function revealInBridgeHandler () {
    var targetDocument = app.activeDocument;
    var selectedPortion = app.selection[0];
	var lastCharIndex = new Array;
    var nestedMarkers = new Array;
	if (!getContainerSnippet (selectedPortion, lastCharIndex, nestedMarkers, "first marker")) { //"first marker"  or "base marker"
        alert ("Please, first determine Live Snippet by selecting it's marker or inserting the text cursor in it's content."); 
	}
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(startChar.textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
	else if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
		reflectSpoiled (nestedMarkers);
		alert ("The instance is spoiled");
		return false;
    }
    var liveSnippetID = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_ID");
    var dynamicSnippetCode = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_PREFIX_NUMBER");
    var isToChange = false;
    if (dynamicSnippetCode) {
        var pathAndName = solveSnippetCode (targetDocument, nestedMarkers[0][0].parent.insertionPoints.item (nestedMarkers[0][0].index), dynamicSnippetCode, null);
        if (pathAndName[1] == "") {
            pathAndName[1] = "0001";
        }
        if (pathAndName[0] || pathAndName[1]) {
            liveSnippetIDSplitted = new Array;
            liveSnippetIDSplitted.push (liveSnippetID.slice (0, liveSnippetID.lastIndexOf ("/")));
            liveSnippetIDSplitted.push (liveSnippetID.slice (liveSnippetID.lastIndexOf ("/") + 1));
            if (pathAndName[0]) {
                if (liveSnippetIDSplitted[0] != pathAndName[0]) {
                    isToChange = true;
                }
            }
            if (pathAndName[1]) {
                if (liveSnippetIDSplitted[1] != pathAndName[1]) {
                    isToChange = true;
                }
            }
            if (isToChange) {
                alert ("The instance has a Dynamic Identifier (Branch and Name) and requires an update.");
                return false;
            }
        }
    }
    var snippetFile = findTargetFile (targetDocument, nestedMarkers[0][0]);
    if (snippetFile instanceof File) {
        revealFileInBridge (snippetFile);
    }
    else if (snippetFile == false) {
        alert ("It's an Inline Snippet and has no file!");
    }
    else {
        alert ("The file is missing!");
    }
}

function revealFileInBridge (targetFile) {
    var targetBridge = BridgeTalk.getSpecifier ("bridge");
    if (targetBridge && targetBridge.appStatus != "not installed") {
        var talkBridge = new BridgeTalk;
        talkBridge.target = targetBridge;
        talkBridge.body = "tsToBeSelectedThumbnail = new Thumbnail (new File (\"" + targetFile.fsName.replace(/\\/g, "/") + "\"));"
                            + "\n" + "tsLoadStage = 1;"
                            + "\n" + "app.document.thumbnail = tsToBeSelectedThumbnail.parent;"
                            + "\n" + "app.bringToFront ();";
        talkBridge.send ();
    }
}

function showInfoHandler () {
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            //[targetDocument, selectedObject, targetNumber] 
            var dataArray = new Array;
            dataArray.push(app.activeDocument);
            dataArray.push(app.selection[0]);
            app.doScript(showInfo, ScriptLanguage.JAVASCRIPT, dataArray, UndoModes.ENTIRE_SCRIPT, "Show Info.");
            //showInfo(dataArray);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
		case "not_saved":
            alert ("Please save the document.");
            exit();
            break;
	}
}

function selectContentHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
		case "not_saved":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedPortion = app.selection[0];
            var selectedPortionAndTargetWay = new Array;
            selectedPortionAndTargetWay.push(selectedPortion);
            selectedPortionAndTargetWay.push("content");
            app.doScript(selectInstance, ScriptLanguage.JAVASCRIPT, selectedPortionAndTargetWay, UndoModes.ENTIRE_SCRIPT, "Select to End border");
            //selectInstance(selectedPortionAndTargetWay);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
	}
}

function selectEndMarkerHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
		case "not_saved":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedPortion = app.selection[0];
            var selectedPortionAndTargetWay = new Array;
            selectedPortionAndTargetWay.push(selectedPortion);
            selectedPortionAndTargetWay.push("end_marker");
            app.doScript(selectInstance, ScriptLanguage.JAVASCRIPT, selectedPortionAndTargetWay, UndoModes.ENTIRE_SCRIPT, "Select End Marker");
            //selectInstance(selectedPortionAndTargetWay);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
	}
}

function selectFirstChildHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
		case "not_saved":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedPortion = app.selection[0];
            app.doScript(selectFirstChild, ScriptLanguage.JAVASCRIPT, selectedPortion, UndoModes.ENTIRE_SCRIPT, "Select First Child");
            //selectFirstChild(selectedPortion);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
	}
}

function selectNextBrotherHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
		case "not_saved":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedPortion = app.selection[0];
            app.doScript(selectNextBrother, ScriptLanguage.JAVASCRIPT, selectedPortion, UndoModes.ENTIRE_SCRIPT, "Select Next Brother");
            //selectNextBrother(selectedPortion);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
	}
}

function selectPreviousBrotherHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
		case "not_saved":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedPortion = app.selection[0];
            app.doScript(selectPreviousBrother, ScriptLanguage.JAVASCRIPT, selectedPortion, UndoModes.ENTIRE_SCRIPT, "Select Previous Brother");
            //selectPreviousBrother(selectedPortion);
            break;
		case "no_document":
            alert ("Please open or create a document.");
            exit();
            break;
	}
}

function selectStartMarkerHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
		case "not_saved":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedPortion = app.selection[0];
            var selectedPortionAndTargetWay = new Array;
            selectedPortionAndTargetWay.push(selectedPortion);
            selectedPortionAndTargetWay.push("start_marker");
            app.doScript(selectInstance, ScriptLanguage.JAVASCRIPT, selectedPortionAndTargetWay, UndoModes.ENTIRE_SCRIPT, "Select Start Marker");
            //selectInstance(selectedPortionAndTargetWay);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
	}
}

function selectToEndBorderHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
		case "not_saved":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedPortion = app.selection[0];
            var selectedPortionAndTargetWay = new Array;
            selectedPortionAndTargetWay.push(selectedPortion);
            selectedPortionAndTargetWay.push("to_end_border");
            app.doScript(selectInstance, ScriptLanguage.JAVASCRIPT, selectedPortionAndTargetWay, UndoModes.ENTIRE_SCRIPT, "Select to End border");
            //selectInstance(selectedPortionAndTargetWay);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
	}
}

function selectToStartBorderHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
		case "not_saved":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedPortion = app.selection[0];
            var selectedPortionAndTargetWay = new Array;
            selectedPortionAndTargetWay.push(selectedPortion);
            selectedPortionAndTargetWay.push("to_start_border");
            app.doScript(selectInstance, ScriptLanguage.JAVASCRIPT, selectedPortionAndTargetWay, UndoModes.ENTIRE_SCRIPT, "Select To Start Border");
            //selectInstance(selectedPortionAndTargetWay);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
	}
}

function selectWholeSnippetHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
		case "not_saved":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedPortion = app.selection[0];
            var selectedPortionAndTargetWay = new Array;
            selectedPortionAndTargetWay.push(selectedPortion);
            selectedPortionAndTargetWay.push("whole_instance");
            app.doScript(selectInstance, ScriptLanguage.JAVASCRIPT, selectedPortionAndTargetWay, UndoModes.ENTIRE_SCRIPT, "Update Live Snippet");
            //selectInstance(selectedPortionAndTargetWay);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
	}
}

function submitSelectedHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var dataArray = new Array;
            dataArray.push (app.activeDocument);
            dataArray.push (app.selection[0]);
            //app.doScript(submit, ScriptLanguage.JAVASCRIPT, dataArray, UndoModes.ENTIRE_SCRIPT, "Submit Live Snippet");
            submitSelected (dataArray);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
		case "not_saved":
            alert ("Please save the document.");
            exit();
            break;
	}
}

function submitAllHandler (){
    // There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
            var targetDocument = app.activeDocument;
            app.doScript(submitAllLiveSnippets, ScriptLanguage.JAVASCRIPT, targetDocument, UndoModes.ENTIRE_SCRIPT, "Submit Document");
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
		case "not_saved":
            alert ("Please save the document.");
            exit();
            break;
	}
}

function autoSubmitAllLiveSnippetsHandler(){
    if (app.extractLabel ("LS_AUTO_SUBMIT") == "yes")
        app.insertLabel ("LS_AUTO_SUBMIT", "no");
    else
        app.insertLabel ("LS_AUTO_SUBMIT", "yes");
    autoSubmitAllLiveSnippetsAction.checked = (app.extractLabel ("LS_AUTO_SUBMIT") == "yes");
}

function submitAllLiveSnippets (targetDocument) {
    checkAllLiveSnippets (targetDocument, true, false);
}

function checkDocumentHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
        	var currentDocument = app.activeDocument;
            app.doScript(checkDocument, ScriptLanguage.JAVASCRIPT, currentDocument, UndoModes.ENTIRE_SCRIPT, "Check Document");
            //checkDocument(currentDocument);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
		case "not_saved":
            alert ("Please save the document.");
            exit();
            break;
	}
}

function autoCheckDocumentHandler(){
    if (app.extractLabel ("LS_AUTO_CHECK") == "yes")
        app.insertLabel ("LS_AUTO_CHECK", "no");
    else
        app.insertLabel ("LS_AUTO_CHECK", "yes");
    autoCheckDocumentAction.checked = (app.extractLabel ("LS_AUTO_CHECK") == "yes");
}

function updateSelectedHandler (){
	// There must be an active and already saved document so we can create a folder inside its parent
	switch (isActiveDocReady ()){
		case "ready":
            // Ensuring that the user has selected a single text or a single object.
            if (app.selection.length != 1)
                exit();
            var selectedObject = app.selection[0];
            var isForcibly = true;
            var selectedObjectAndisForcibly = new Array;
            selectedObjectAndisForcibly.push(selectedObject);
            selectedObjectAndisForcibly.push(isForcibly);
            //app.doScript(updateSelected, ScriptLanguage.JAVASCRIPT, selectedObjectAndisForcibly, UndoModes.ENTIRE_SCRIPT, "Update Live Snippet");
            updateSelected(selectedObjectAndisForcibly);
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
		case "not_saved":
            alert ("Please save the document.");
            exit();
            break;
	}
}

function updateAllHandler (){
	switch (isActiveDocReady ()){
		case "ready":
            var targetDocument = app.activeDocument;
            //app.doScript(updateDocument, ScriptLanguage.JAVASCRIPT, targetDocument, UndoModes.ENTIRE_SCRIPT, "Update Document");
            app.doScript(updateAllLiveSnippets, ScriptLanguage.JAVASCRIPT, targetDocument, UndoModes.ENTIRE_SCRIPT, "Update Document");
            break;
		case "no_document":
            alert ("Please open a document.");
            exit();
            break;
		case "not_saved":
            alert ("Please save the document.");
            exit();
            break;
	}
}

function updateAllLiveSnippets(targetDocument){
    checkAllLiveSnippets (targetDocument, false, true);
}

function autoUpdateAllLiveSnippetsHandler(){
    if (app.extractLabel ("LS_AUTO_UPDATE") == "yes")
        app.insertLabel ("LS_AUTO_UPDATE", "no");
    else
        app.insertLabel ("LS_AUTO_UPDATE", "yes");
    autoUpdateAllLiveSnippetsAction.checked = (app.extractLabel ("LS_AUTO_UPDATE") == "yes");
}

function createLiveSnippetFile (targetDocument, liveSnippetsFolder, nestedMarkers, dynamicLiveSnippetPhrase) {
	var startChar = nestedMarkers[0][0];
	var endChar = nestedMarkers[nestedMarkers.length-1][0];
	var startFrame = startChar.textFrames.item(0);
    var extractLabelLS_ID = startFrame.extractLabel("LS_ID");
    var liveSnippetIDSplitted = new Array;
    liveSnippetIDSplitted.push (extractLabelLS_ID.slice (0, extractLabelLS_ID.lastIndexOf ("/")));
    liveSnippetIDSplitted.push (extractLabelLS_ID.slice (extractLabelLS_ID.lastIndexOf ("/") + 1));
    var newLiveSnippetFile = null;
    if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_SUM") == "DYNAMIC") {
        var expectedFile = liveSnippetsFolder.getFiles (liveSnippetIDSplitted[1] + "*.txt");
        if (expectedFile.length > 0) {
            newLiveSnippetFile = expectedFile[0];
        }
        else {
            newLiveSnippetFile =  new File (liveSnippetsFolder.fsName.replace(/\\/g, "/") + "/" + liveSnippetIDSplitted[1] + ".txt");
        }
        newLiveSnippetFile.encoding = "UTF8";
        var dynamicContent = readFile (newLiveSnippetFile);
        if (dynamicContent != false && dynamicContent.indexOf (tsDynamicSnippetTag) == 0 && dynamicContent.indexOf (":\n") != -1) {
            dynamicLiveSnippetPhrase = dynamicLiveSnippetPhrase + dynamicContent.slice (dynamicContent.indexOf (":\n"));
        }
        writeFile (newLiveSnippetFile, tsDynamicSnippetTag + dynamicLiveSnippetPhrase);
    }
    else if (nestedMarkers.length == 2 && !nestedMarkers[0][2]) { //
        var expectedFile = liveSnippetsFolder.getFiles (liveSnippetIDSplitted[1] + "*.txt");
        if (expectedFile.length > 0) {
            newLiveSnippetFile = expectedFile[0];
        }
        else {
            newLiveSnippetFile =  new File (liveSnippetsFolder.fsName.replace(/\\/g, "/") + "/" + liveSnippetIDSplitted[1] + ".txt");
        }
        newLiveSnippetFile.encoding = "UTF8";
        var liveSnippetContentAndSUM = getLiveSnippetContentAndSUM (nestedMarkers, 0);
        var toBeExported = startChar.parent.characters.itemByRange(startChar.parent.characters.item(startChar.index+1), startChar.parent.characters.item(endChar.index-1));
        var toBeExportedContents = toBeExported.contents;
        if (toBeExportedContents == "") {
            toBeExportedContents = "no_text";
        }
        writeFile (newLiveSnippetFile, toBeExportedContents);
        startFrame.insertLabel("LS_SUM", liveSnippetContentAndSUM[1]);
    }
    else {
        var expectedFile = liveSnippetsFolder.getFiles (liveSnippetIDSplitted[1] + "*.IDMS");
        if (expectedFile.length > 0) {
            newLiveSnippetFile = expectedFile[0];
            newLiveSnippetFile.remove();
        }
        else {
            newLiveSnippetFile =  new File (liveSnippetsFolder.fsName.replace(/\\/g, "/") + "/" + liveSnippetIDSplitted[1] + ".IDMS");
        }
        var liveSnippetContentAndSUM = getLiveSnippetContentAndSUM (nestedMarkers, 0);
        //Adding a temporary text frame to paste the text inside it.
        var myOldXUnits;
        var myOldYUnits;
        with (targetDocument.viewPreferences) {
            myOldXUnits = horizontalMeasurementUnits;
            myOldYUnits = verticalMeasurementUnits;
            horizontalMeasurementUnits = MeasurementUnits.points;
            verticalMeasurementUnits = MeasurementUnits.points;
        }
        var originalBounds = null;
        var targetPage = null;
        if (startChar.parentTextFrames.length > 0) {
            originalBounds = startChar.parentTextFrames[0].geometricBounds;
            targetPage = startChar.parentTextFrames[0].parentPage;
        }
        else {
            originalBounds = startChar.parentStory.textFrames.lastItem ().geometricBounds;
            targetPage = startChar.parentStory.textFrames.lastItem ().parentPage;
        }
        var scoopFrame = targetPage.textFrames.add({geometricBounds: [originalBounds[0], originalBounds[1], originalBounds[0]+8000, originalBounds[3]], 
                                                                                            appliedObjectStyle: targetDocument.objectStyles.itemByName("[Basic Text Frame]")});
        with (targetDocument.viewPreferences) {
            try{
                horizontalMeasurementUnits = myOldXUnits;
                verticalMeasurementUnits = myOldYUnits;
            }
            catch(myError) {
                alert("Could not reset custom measurement units.");
            }
        }
        startChar.parent.characters.itemByRange(startChar, endChar).duplicate(LocationOptions.atBeginning, scoopFrame.insertionPoints.item(0));
        scoopFrame.fit(FitOptions.FRAME_TO_CONTENT);
        scoopFrame.select ();
        try {
            scoopFrame.exportFile(ExportFormat.INDESIGN_SNIPPET, newLiveSnippetFile);
        }
        catch (myError) {
            alert(myError);
            return false;
        }
        scoopFrame.remove();
        startFrame.insertLabel("LS_SUM", liveSnippetContentAndSUM[1]);
    }
	//Mark the snippet instrance with the modification date of the snippet file.
    var newModification = (newLiveSnippetFile.modified.getTime ()).toString ();
    startFrame.insertLabel("LS_DATE", newModification);
    //Order Bridge to checkin the file
    var checkInMessageFolder = new Folder (dataPath + "/Messages/To CheckIn Live Snippet");
    if (!checkInMessageFolder.exists) {
        checkInMessageFolder.create ();
    }
    var c = 1;
    while (File (checkInMessageFolder.fsName.replace(/\\/g, "/") + "/" + newModification + c).exists) {
        c++;
    }
    writeEncodedFile (File (checkInMessageFolder.fsName.replace(/\\/g, "/") + "/" + newModification + c), newLiveSnippetFile.fsName.replace(/\\/g, "/"));
	return newLiveSnippetFile;
}

function reflectMarkerVertically(textFrameObject) {
	if (textFrameObject.extractLabel("LS_TAG") == "snippet start") {
		with (textFrameObject.anchoredObjectSettings) {
			if (anchorPoint == AnchorPoint.TOP_LEFT_ANCHOR)
				anchorPoint = AnchorPoint.BOTTOM_LEFT_ANCHOR;
			else if (anchorPoint == AnchorPoint.TOP_RIGHT_ANCHOR)
				anchorPoint = AnchorPoint.BOTTOM_RIGHT_ANCHOR;
			else if (anchorPoint == AnchorPoint.BOTTOM_LEFT_ANCHOR)
				anchorPoint = AnchorPoint.TOP_LEFT_ANCHOR;
			else if (anchorPoint == AnchorPoint.BOTTOM_RIGHT_ANCHOR)
				anchorPoint = AnchorPoint.TOP_RIGHT_ANCHOR;
		}
	}
	return true;
}

function reflectMarkerHorizontally(textFrameObject){
	if (textFrameObject.extractLabel("LS_TAG") == "snippet start") {
		with (textFrameObject.anchoredObjectSettings) {
			if (anchorPoint == AnchorPoint.TOP_LEFT_ANCHOR)
				anchorPoint = AnchorPoint.TOP_RIGHT_ANCHOR;
			else if (anchorPoint == AnchorPoint.TOP_RIGHT_ANCHOR)
				anchorPoint = AnchorPoint.TOP_LEFT_ANCHOR;
			else if (anchorPoint == AnchorPoint.BOTTOM_LEFT_ANCHOR)
				anchorPoint = AnchorPoint.BOTTOM_RIGHT_ANCHOR;
			else if (anchorPoint == AnchorPoint.BOTTOM_RIGHT_ANCHOR)
				anchorPoint = AnchorPoint.BOTTOM_LEFT_ANCHOR;
		}
	}
	return true;
}

function createMarkersStyles (targetDocument) {
	//Adding colors for the text that mark the instances of the snippets.
	if (targetDocument.colors.itemByName("Snippet Marker Background") == null) {
		targetDocument.colors.add({name: "Snippet Marker Background", //The yellow background of the small markers the the start and end of the instances.
								colorValue: [94, 136, 194], 
								space: ColorSpace.RGB});
	}

	//Adding the paragraph styles for the starting and ending texts
	if (targetDocument.paragraphStyles.itemByName("Snippet Start Marker") == null) {
		targetDocument.paragraphStyles.add({name: "Snippet Start Marker",
											paragraphDirection: ParagraphDirectionOptions.LEFT_TO_RIGHT_DIRECTION,
											appliedFont: "Arial", 
											fontStyle: "Regular", 
											pointSize: 5,
											digitsType: DigitsTypeOptions.ARABIC_DIGITS,
											fillColor: targetDocument.colors.itemByName("Snippet Marker Background") }); // Normal color
	}

	//Adding style for the anchored text frame that mark the instances starting and ending positions and mark the number of the snippet
	if (targetDocument.objectStyles.itemByName("Snippet Start Marker") == null) {
		targetDocument.objectStyles.add({name: "Snippet Start Marker", 
                                                enableFill: true,
                                                enableStroke: true,
                                                strokeWeight: 0,
                                                fillColor: targetDocument.colors.itemByName("Snippet Marker Background"),
                                                enableParagraphStyle: true,
                                                appliedParagraphStyle: targetDocument.paragraphStyles.itemByName("Snippet Start Marker"),
                                                enableTextFrameAutoSizingOptions: true,
                                                enableAnchoredObjectOptions: true,
                                                enableTextWrapAndOthers: true,
                                                enableTextFrameGeneralOptions: true,
                                                nonprinting: true}); //The user can change this value before printing so the instances tags will be printed for reviewing.

		with (targetDocument.objectStyles.itemByName("Snippet Start Marker")) {
			textFramePreferences.insetSpacing = ["0pt", "0pt", "0pt", "0pt"];
			with (anchoredObjectSettings) {
				anchoredPosition = AnchorPosition.ANCHORED;
				anchorPoint = AnchorPoint.TOP_LEFT_ANCHOR;
				verticalAlignment = VerticalAlignment.BOTTOM_ALIGN;
				horizontalReferencePoint = AnchoredRelativeTo.ANCHOR_LOCATION;
				verticalReferencePoint = VerticallyRelativeTo.TOP_OF_LEADING;
				//anchorYoffset = "2pt";
				pinPosition = false;
			}
			fillTransparencySettings.blendingSettings.opacity = 20;
            with (textFramePreferences) {
                autoSizingType = AutoSizingTypeEnum.HEIGHT_AND_WIDTH;
                useNoLineBreaksForAutoSizing = true;
            }
		}
	}

	if (targetDocument.objectStyles.itemByName("Snippet Temporarily Anchored") == null) {
		targetDocument.objectStyles.add({name: "Snippet Temporarily Anchored", 
                                                enableAnchoredObjectOptions: true});

		with (targetDocument.objectStyles.itemByName("Snippet Temporarily Anchored")) {
			with (anchoredObjectSettings) {
				anchoredPosition = AnchorPosition.ANCHORED;
				anchorPoint = AnchorPoint.TOP_LEFT_ANCHOR;
			}
		}
	}

	if (targetDocument.colors.itemByName("Live Snippet Modified") == null) {
		targetDocument.colors.add({name: "Live Snippet Modified", //Live Snippet Modifieds, Orange color
										colorValue: [255, 190, 0], 
										space: ColorSpace.RGB});
	}
	if (targetDocument.colors.itemByName("Live Snippet Missing") == null) {
		targetDocument.colors.add({name: "Live Snippet Missing", //Live Snippet Missings, Red color
										colorValue: [250, 40, 40], 
                                        space: ColorSpace.RGB});
    }
	if (targetDocument.colors.itemByName("Live Snippet Need to Submit") == null) {
		targetDocument.colors.add({name: "Live Snippet Need to Submit", //Live Snippet Need to Submit, Blue color
										colorValue: [30, 50, 160], 
										space: ColorSpace.RGB});
	}
	if (targetDocument.characterStyles.itemByName("Live Snippet Modified") == null) {
		targetDocument.characterStyles.add({name: "Live Snippet Modified", //Live Snippet Modifieds, Blue color
										fillColor: targetDocument.colors.itemByName("Live Snippet Modified"), 
										fontStyle: "Modified Snippet"});
	}
	if (targetDocument.characterStyles.itemByName("Live Snippet Missing") == null) {
		targetDocument.characterStyles.add({name: "Live Snippet Missing", //Live Snippet Missings, Red color
										fillColor: targetDocument.colors.itemByName("Live Snippet Missing"), 
										fontStyle: "Missing Snippet"});
	}
	if (targetDocument.characterStyles.itemByName("Live Snippet Lonely") == null) {
		targetDocument.characterStyles.add({name: "Live Snippet Lonely", //Live Snippet Missings, Red color
										strikeThru: true,
										strikeThroughColor: targetDocument.colors.itemByName("Live Snippet Missing"), 
										fontStyle: "Lonely Snippet"});
	}
	if (targetDocument.characterStyles.itemByName("Live Snippet Need to Submit") == null) {
		targetDocument.characterStyles.add({name: "Live Snippet Need to Submit", //Live Snippet Missings, Red color
										underline: true,
										underlineColor: targetDocument.colors.itemByName("Live Snippet Need to Submit"), 
										fontStyle: "Live Snippet Need to Submit"});
	}
	if (targetDocument.characterStyles.itemByName("Live Snippet Spoiled") == null) {
		targetDocument.characterStyles.add({name: "Live Snippet Spoiled", //Live Snippet Missings, Red color
										fillColor: targetDocument.colors.itemByName("Black"), 
										fontStyle: "Spoiled Snippet"});
	}
	return true;
}

function changeMarkers (nestedMarkers, liveSnippetPrefix, liveSnippetDigits) {
    var startMarkerFrame = nestedMarkers[0][0].textFrames.item(0);
    
    //Inserting the information in the start anchored text frame as labels
    startMarkerFrame.insertLabel("LS_ID", liveSnippetPrefix + "/" + liveSnippetDigits);
    startMarkerFrame.insertLabel("LS_DATE", "0");
    startMarkerFrame.insertLabel("LS_SUM", "");
    startMarkerFrame.paragraphs.item(0).contents = liveSnippetPrefix + "/" + liveSnippetDigits; //This only for preview for the user and modifying it doesn't affect anything
}

function insertMarkers (targetDocument, targetText, liveSnippetPrefix, liveSnippetDigits, trimEnd, isIndividuals) {
    targetText.insertionPoints.item(trimEnd).contents = SpecialCharacters.END_NESTED_STYLE;
    var startMarkerFrame = targetText.insertionPoints.item(0).textFrames.add ({appliedObjectStyle: targetDocument.objectStyles.itemByName("Snippet Start Marker")});
    try {
        startMarkerFrame.geometricBounds = ["-30pt", "0pt", "-20pt", "20pt"];
    }
    catch (e) { }
    //Inserting the information in the start anchored text frame as labels
    startMarkerFrame.insertLabel("LS_TAG", "snippet start");
    startMarkerFrame.insertLabel("LS_ID", liveSnippetPrefix + "/" + liveSnippetDigits);
    startMarkerFrame.insertLabel("LS_DATE", "0");
    startMarkerFrame.insertLabel("LS_SUM", "");
    if (isIndividuals == false) {
        startMarkerFrame.insertLabel("TS_LS_REPLACE", "\n:\r");
    }
    startMarkerFrame.insertionPoints.item(-1).contents = liveSnippetPrefix + "/" + liveSnippetDigits; //This only for preview for the user and modifying it doesn't affect anything
    startMarkerFrame.paragraphs.item(0).appliedCharacterStyle = targetDocument.characterStyles.itemByName("[None]");
    startMarkerFrame.paragraphs.item(0).clearOverrides(OverrideType.ALL);
    return startMarkerFrame;
}

function removeZeros (str) {
    for (var s = 0; s < str.length - 1; s++) {
        if (str[s] == '0') {
            str = removeByIndex (str, s);
            s--;
        }
        else if (str.charCodeAt (s) >= 48 && str.charCodeAt (s) <= 57) {
            break;
        }
    }
    return str;
}

function removeByIndex (str,index) {
    if (index == 0) {
        return  str.slice (1)
    } else {
        return str.slice (0, index) + str.slice (index + 1);
    } 
}

function removeMarkers (nestedMarkers) {
    nestedMarkers[nestedMarkers.length-1][0].remove();
    nestedMarkers[0][0].remove(); 
}

function dynamicPathAndNameSelected(selectedObjectAndisForcibly) {
    var selectedObject = selectedObjectAndisForcibly[0];
    var isForcibly = selectedObjectAndisForcibly[1];
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
	var targetDocument = app.activeDocument;
	var lastCharIndex = new Array;
	var nestedMarkers = new Array;
	if (!getContainerSnippet (selectedObject, lastCharIndex, nestedMarkers, "first marker")) { //"first marker"  or "base marker"
		return false; 
	}
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(startChar.textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
	else if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
		reflectSpoiled (nestedMarkers);
		alert ("The instance is spoiled");
		return false;
	}
    var dynamicSnippetCode = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_PREFIX_NUMBER");
    dynamicSnippetCode = tsGetText (["Dynamic Identifier (Branch and Name):",
                                         "Use this <.> for the snippet file parent folder name, add more dots for parent of parent...",
                                         "To reference to the document instead of the snippet file add dollar sign like this '<.$>'.",
                                         "Use also: <Increment>, <Container_File_Base>, <Container_File_Name>,",
                                         "<Doc_Name>, <Doc_Order>, <Page_Num>, <Applied_Section_Prefix>, <Applied_Section_Marker>,",
                                         "<Master_Prefix>, <Master_Name>, <Master_Full_Name>, <Char_Order>,",
                                         "<Anchored_Order>, <Para_Order>, <Para_First>, <Para_Before>, <Para_This>, <Para_After>, <Para_Style>,",
                                         "<Col_Head>, <Col_Num>, <Row_Head>, <Row_Num>, <Shift1>.",
                                         "",
                                         "The following accepts abbreviation and used for date, time and random code:",
                                         "<Day_Seconds>, <Day_Minutes>, <Day_Hours>, <Weekday_Order>, <Weekday_Name>, <Weekday_Abbreviation>,",
                                         "<Month_Date>, <Month_Order>, <Month_Name>, <Month_Abbreviation>, <Year_Full>,",
                                         "<Year_Abbreviation>, <Time_Decimal>, <Time_Hexadecimal>, <Time_Shortest>,",
                                         "<Random_Style_Long> Where 'Style' a combination of symbols 'a' for small letters,",
                                         "'A' capital letters, 'h' small hexadecimal, 'H' large hexadecimal and '1' for numbers.",
                                         "For example <Random_1A_8> means random 8 characters consists of capital letters and numbers",
                                         ""
                                        ], dynamicSnippetCode, false);
    if (dynamicSnippetCode) {
        nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_PREFIX_NUMBER", dynamicSnippetCode);
        var startStory = nestedMarkers[0][0].parent;
        var startIndex = nestedMarkers[0][0].index;
        var snippetsCacheList = new Array;
        updateInstance (targetDocument, null, nestedMarkers, isForcibly, null, null, null, false, snippetsCacheList, true, [], [], -1, -1, []);
        
        var toSelect = new Array;
        surveyAndCheck (startStory.characters.item(startIndex), lastCharIndex, toSelect, false);
        var toSelectRange = toSelect[0][0].parent.characters.itemByRange(toSelect[0][0], toSelect[toSelect.length-1][0]);
        toSelectRange.select();
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
    }
}

function dynamicLiveSnippetSelected(selectedObjectAndisForcibly) {
    var selectedObject = selectedObjectAndisForcibly[0];
    var isForcibly = selectedObjectAndisForcibly[1];
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
	var targetDocument = app.activeDocument;
	var lastCharIndex = new Array;
	var nestedMarkers = new Array;
	if (!getContainerSnippet (selectedObject, lastCharIndex, nestedMarkers, "first marker")) { //"first marker"  or "base marker"
		return false; 
	}
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(startChar.textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
	else if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
		reflectSpoiled (nestedMarkers);
		alert ("The instance is spoiled");
		return false;
	}
    var dynamicLiveSnippetPhrase = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_DYNAMIC");
    newDynamicLiveSnippet (targetDocument, nestedMarkers, dynamicLiveSnippetPhrase, isForcibly, true);
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
}

function newDynamicLiveSnippet (targetDocument, nestedMarkers, dynamicLiveSnippetPhrase, isForcibly, isToAsk) {
    if (dynamicLiveSnippetPhrase == "" || isToAsk)
        dynamicLiveSnippetPhrase = tsGetText (["Path:Method:Separator",
                                               "",
                                               "Path (or table formula):",
                                               "//: Root of Tree Shade path.",
                                               "./: Snippet parent folder path. add dots for parent of parent...",
                                               "<.>: Snippet parent folder name. add dots for parent of parent...",
                                               "To reference to the document instead of the snippet file write like this '.$/' or '<.$>'...",
                                               "Also use these tags with path:",
                                               "<File_Branch>, <File_Name>, <File_Display>, <Original_File_Name>, <Previous_File_Name>,",
                                               "<Container_File_Branch>, <Container_File_Name>, <File_Level>, <Doc_Name>, <Doc_Order>, <Page_Num>,",
                                               "<Applied_Section_Prefix>, <Applied_Section_Marker>, <Master_Prefix>, <Master_Name>,",
                                               "<Master_Full_Name>, <Char_Order>, <Anchored_Order>, <Para_Order>, <Para_First>, <Para_Before>,",
                                               "<Para_This>, <Para_After>, <Para_Style>, <Char_Style>, <Col_Head>, <Col_Num>, <Row_Head>,",
                                               "<Row_Num>, <Name_At>, <Shift1>. The target number comes after <Name_At> but before <Shift1>.",
                                               "Use <File_Name_Separator_NewSeparatorOrIndex> like <FN_ _1> will take first word.",
                                               "Dynamic Table Formula:",
                                               "'Sum', 'Subtract', 'Divide', 'Multiply', 'Average', 'Count', or 'Join'.",
                                               "For an example <Sum (<Col_Num>, 1)(<Col_Num>, <Row_Num> - 1)>",
                                               "This will get the sum of all cells above.",
                                               "You could abbreviate any one like <FB> instead of <File_Branch>.",
                                               "",
                                               "The following also accepts abbreviation and used for date, time and random code:",
                                               "<Day_Seconds>, <Day_Minutes>, <Day_Hours>, <Weekday_Order>, <Weekday_Name>, <Weekday_Abbreviation>,",
                                               "<Month_Date>, <Month_Order>, <Month_Name>, <Month_Abbreviation>, <Year_Full>,",
                                               "<Year_Abbreviation>, <Time_Decimal>, <Time_Hexadecimal>, <Time_Shortest>,",
                                               "<Random_Style_Long> Where 'Style' a combination of symbols 'a' for small letters,",
                                               "'A' capital letters, 'h' small hexadecimal, 'H' large hexadecimal and '1' for numbers.",
                                               "For example <Random_1A_8> means random 8 characters consists of capital letters and numbers",
                                               "",
                                               "Method:",
                                               "'Snippets', 'Place', 'Names', 'Repeat', 'Key' or 'Data'.",
                                               "",
                                               "Separator:",
                                               "'Frames', 'Table', 'Rows', 'Columns', 'Colon' or any text letters.",
                                               "",
                                               "Example: '.$/<PN>/*.png:Place:|'",
                                               "This will place all the png files in the folder beside the document that",
                                               "its name as the page number. These PNGs will separated by '|' character."], dynamicLiveSnippetPhrase, false);
    if (dynamicLiveSnippetPhrase) {
        nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_SUM", "DYNAMIC");
        nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_DYNAMIC", dynamicLiveSnippetPhrase);
        var inlineDynamicMark = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_DYNAMIC_INLINE");
        if (inlineDynamicMark == "") {
            var liveSnippetID = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_ID");
            var liveSnippetIDSplitted = new Array;
            liveSnippetIDSplitted.push (liveSnippetID.slice (0, liveSnippetID.lastIndexOf ("/")));
            liveSnippetIDSplitted.push (liveSnippetID.slice (liveSnippetID.lastIndexOf ("/") + 1));
            if (liveSnippetIDSplitted[0] == "#") {
                nestedMarkers[0][0].textFrames.item(0).insertLabel ("LS_DYNAMIC_INLINE", "YES");
            }
            else {
                var liveSnippetsFolder = getLiveSnippetFolder (targetDocument, liveSnippetIDSplitted[0], null, null);
                createLiveSnippetFile (targetDocument, liveSnippetsFolder, nestedMarkers, dynamicLiveSnippetPhrase);
            }
        }
        var startStory = nestedMarkers[0][0].parent;
        var startIndex = nestedMarkers[0][0].index;
        var snippetsCacheList = new Array;
        updateInstance (targetDocument, null, nestedMarkers, isForcibly, dynamicLiveSnippetPhrase, null, null, false, snippetsCacheList, true, [], [], -1, -1, []);
        
        var toSelect = new Array;
        var lastCharIndex = new Array;
        surveyAndCheck (startStory.characters.item(startIndex), lastCharIndex, toSelect, false);
        var toSelectRange = toSelect[0][0].parent.characters.itemByRange(toSelect[0][0], toSelect[toSelect.length-1][0]);
        toSelectRange.select();
    }
}

function dynamicStyleSelected(selectedObjectAndisForcibly) {
    var selectedObject = selectedObjectAndisForcibly[0];
    var isForcibly = selectedObjectAndisForcibly[1];
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
	var targetDocument = app.activeDocument;
	var lastCharIndex = new Array;
	var nestedMarkers = new Array;
	if (!getContainerSnippet (selectedObject, lastCharIndex, nestedMarkers, "first marker")) { //"first marker"  or "base marker"
		return false; 
	}
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(startChar.textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
	else if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
		reflectSpoiled (nestedMarkers);
		alert ("The instance is spoiled");
		return false;
	}
    var dynamicStylePhrase = nestedMarkers[0][0].textFrames.item(0).extractLabel("TS_LS_STYLE");
    dynamicStylePhrase = tsGetText (["PathOrStyle1:Method1:Separator1::PathOrStyle2:Method2:Separator2::...",
                                     "",
                                     "PathOrStyle:",
                                     "//: Root of Tree Shade path.",
                                     "./: Snippet parent folder path. add dots for parent of parent...",
                                     "<.>: Snippet parent folder name. add dots for parent of parent...",
                                     "To reference to the document instead of the snippet file write like this '.$/' or '<.$>'...",
                                     "Also use these tags with path:",
                                     "<File_Branch>, <File_Name>, <File_Display>, <Original_File_Name>, <Previous_File_Name>,",
                                     "<Container_File_Branch>, <Container_File_Name>, <File_Level>, <Doc_Name>, <Doc_Order>,",
                                     "<Page_Num>, <Applied_Section_Prefix>, <Applied_Section_Marker>, <Master_Prefix>, <Master_Name>,",
                                     "<Master_Full_Name>, <Char_Order>, <Anchored_Order>, <Para_Order>, <Para_First>,",
                                     "<Para_Before>, <Para_This>, <Para_After>, <Para_Style>, <Char_Style>, <Col_Head>,",
                                     "<Col_Num>, <Row_Head>, <Row_Num>, <Name_At>, <Shift1>, <Remove_If_Empty>.",
                                     "After the style name write <Based_On> to add the style if not exist, then the base style if there is.",
                                     "You could abbreviate like <FB> instead of <File_Branch>.",
                                     "Use <File_Name_Separator_NewSeparatorOrIndex> like <FN_ _1> will take first word.",
                                     "",
                                     "The following accepts abbreviation and used for date, time and random code:",
                                     "<Day_Seconds>, <Day_Minutes>, <Day_Hours>, <Weekday_Order>, <Weekday_Name>, <Weekday_Abbreviation>,",
                                     "<Month_Date>, <Month_Order>, <Month_Name>, <Month_Abbreviation>, <Year_Full>,",
                                     "<Year_Abbreviation>, <Time_Decimal>, <Time_Hexadecimal>, <Time_Shortest>,",
                                     "<Random_Style_Long> Where 'Style' a combination of symbols 'a' for small letters,",
                                     "'A' capital letters, 'h' small hexadecimal, 'H' large hexadecimal and '1' for numbers.",
                                     "For example <Random_1A_8> means random 8 characters consists of capital letters and numbers",
                                     "",
                                     "Method:",
                                     "The target style or action: 'Font', 'Style', 'Size', 'Color', 'Justification',",
                                     "'characterDirection', 'paragraphDirection', 'Para', 'Char', 'Obj', 'Height', 'Width', 'Area',",
                                     "'Rotate', 'Shear', 'Short', 'Long', 'Merge', 'Cell', 'Row', 'Column' or 'Table'.",
                                     "",
                                     "Separator:",
                                     "'Colon' or any jointer between files' texts (space is the default)."], dynamicStylePhrase, false);
    if (dynamicStylePhrase == null)
        return false;
    nestedMarkers[0][0].textFrames.item(0).insertLabel("TS_LS_STYLE", dynamicStylePhrase);
    var snippetsCacheList = new Array;
    updateInstance (targetDocument, null, nestedMarkers, isForcibly, null, null, null, false, snippetsCacheList, true, [], [], -1, -1, []);
    
    var toSelect = new Array;
    surveyAndCheck (startStory.characters.item(startIndex), lastCharIndex, toSelect, false);
    var toSelectRange = toSelect[0][0].parent.characters.itemByRange(toSelect[0][0], toSelect[toSelect.length-1][0]);
    toSelectRange.select();
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
}

function dynamicReplaceSelected(selectedObjectAndisForcibly) {
    var selectedObject = selectedObjectAndisForcibly[0];
    var isForcibly = selectedObjectAndisForcibly[1];
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
	var targetDocument = app.activeDocument;
	var lastCharIndex = new Array;
	var nestedMarkers = new Array;
	if (!getContainerSnippet (selectedObject, lastCharIndex, nestedMarkers, "first marker")) { //"first marker"  or "base marker"
		return false; 
	}
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(startChar.textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
	else if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
		reflectSpoiled (nestedMarkers);
		alert ("The instance is spoiled");
		return false;
	}
    var dynamicReplacePhrase = nestedMarkers[0][0].textFrames.item(0).extractLabel("TS_LS_REPLACE");
    dynamicReplacePhrase = tsGetText (["Find1:Replace1:Find2:Replace2...",
                                     "",
                                     "//: Root of Tree Shade path.",
                                     "./: Snippet parent folder path. add dots for parent of parent...",
                                     "<.>: Snippet parent folder name. add dots for parent of parent...",
                                     "To reference to the document instead of the snippet file write like this '.$/' or '<.$>'...",
                                     "Also use these tags with path:",
                                     "<File_Branch>, <File_Name>, <File_Display>, <Original_File_Name>, <Previous_File_Name>,",
                                     "<Container_File_Branch>, <Container_File_Name>, <File_Level>, <Doc_Name>, <Doc_Order>,",
                                     "<Page_Num>, <Applied_Section_Prefix>, <Applied_Section_Marker>, <Master_Prefix>, <Master_Name>,",
                                     "<Master_Full_Name>, <Char_Order>, <Anchored_Order>, <Para_Order>, <Para_First>,",
                                     "<Para_Before>, <Para_This>, <Para_After>, <Para_Style>, <Char_Style>, <Col_Head>,",
                                     "<Col_Num>, <Row_Head>, <Row_Num>, <Name_At>, <Shift1>, <Remove_If_Empty>.",
                                     "Use <File_Name_Separator_NewSeparatorOrIndex> like <FN_ _1> will take first word.",
                                     "You could abbreviate like <FB> instead of <File_Branch>.",
                                     "",
                                     "The following also accepts abbreviation and used for date, time and random code:",
                                     "<Day_Seconds>, <Day_Minutes>, <Day_Hours>, <Weekday_Order>, <Weekday_Name>, <Weekday_Abbreviation>,",
                                     "<Month_Date>, <Month_Order>, <Month_Name>, <Month_Abbreviation>, <Year_Full>,",
                                     "<Year_Abbreviation>, <Time_Decimal>, <Time_Hexadecimal>, <Time_Shortest>,",
                                     "<Random_Style_Long> Where 'Style' a combination of symbols 'a' for small letters,",
                                     "'A' capital letters, 'h' small hexadecimal, 'H' large hexadecimal and '1' for numbers.",
                                     "For example <Random_1A_8> means random 8 characters consists of capital letters and numbers",
                                     ""
                                    ], dynamicReplacePhrase, false);
    if (dynamicReplacePhrase == null)
        return false;
    nestedMarkers[0][0].textFrames.item(0).insertLabel("TS_LS_REPLACE", dynamicReplacePhrase);
    var snippetsCacheList = new Array;
    updateInstance (targetDocument, null, nestedMarkers, isForcibly, null, null, null, false, snippetsCacheList, true, [], [], -1, -1, []);
    
    var toSelect = new Array;
    surveyAndCheck (startStory.characters.item(startIndex), lastCharIndex, toSelect, false);
    var toSelectRange = toSelect[0][0].parent.characters.itemByRange(toSelect[0][0], toSelect[toSelect.length-1][0]);
    toSelectRange.select();
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
}

function updateSelected(selectedObjectAndisForcibly) {
    var selectedObject = selectedObjectAndisForcibly[0];
    var isForcibly = selectedObjectAndisForcibly[1];
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
	var targetDocument = app.activeDocument;
	var lastCharIndex = new Array;
	var nestedMarkers = new Array;
	if (!getContainerSnippet (selectedObject, lastCharIndex, nestedMarkers, "first marker")) { //"first marker"  or "base marker"
		return false; 
	}
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(startChar.textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
	else if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
		reflectSpoiled (nestedMarkers);
		alert ("The instance is spoiled");
		return false;
	}

	var startStory = nestedMarkers[0][0].parent;
    var startIndex = nestedMarkers[0][0].index;
    var snippetsCacheList = new Array;
	updateInstance (targetDocument, null, nestedMarkers, isForcibly, null, null, null, false, snippetsCacheList, true, [], [], -1, -1, []);
	var toSelect = new Array;
	surveyAndCheck (startStory.characters.item(startIndex), lastCharIndex, toSelect, false);
	var toSelectRange = toSelect[0][0].parent.characters.itemByRange(toSelect[0][0], toSelect[toSelect.length-1][0]);
	toSelectRange.select();
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
}

function getContainerSnippet (selectedObject, lastCharIndex, nestedMarkers, optionStr) { 
	var startChar = null; //optionStr = "first marker" or "base marker"
	var levelObject = selectedObject;
	startChar = getStartChar (levelObject);
	switch (optionStr) {
		case "first marker":
			break;
		case "base marker": //If it is a long story the process will take much time
			var aStartChar = startChar;
			while (aStartChar) {
				startChar = aStartChar;
				if (aStartChar.index > 0) {
					aStartChar = getStartChar (aStartChar.parent.characters.item(aStartChar.index - 1));
				}
				else {
					if (aStartChar.parent.constructor.name == "Cell")
						aStartChar = getStartChar (aStartChar.parent);
					else
						aStartChar = getStartChar (aStartChar.parentTextFrames[0]);
				}
			}
			break;
		default:
			alert ("Wrong in optionStr parameter in getContainerSnippet (selectedObject, optionStr) function.");
	}
	if (startChar) {
		lastCharIndex.push(0);
		surveyAndCheck(startChar, lastCharIndex, nestedMarkers, false);
		return true;
	}
	else
		return false;
}

function searchAllIndexes(str, regex) {
    var indexes = [];
    var offset = 0;
    var matchLength;
    var result;

    // Remove 'g' if present — not used by search()
    regex = new RegExp(regex.source, regex.flags.replace('g', ''));

    while ((result = str.search(regex)) !== -1) {
        indexes.push(offset + result);
        matchLength = str.match(regex)[0].length;
        offset += result + matchLength;
        str = str.slice(result + matchLength);
    }

    return indexes;
}

function surveyAndCheck(startChar, lastCharIndex, nestedMarkers, isWithFixing) {
    nestedMarkers.push([startChar, "", false]);
    if (isWithFixing)
        startChar.textFrames.item(0).insertLabel("checked", "yes");

    var currentStory = startChar.parent;
    var currentIndex = startChar.index + 1;
    var currentChar;
    var isMarkerChar;
    var isSpoiled = false;
    var storyContents = currentStory.contents;

    var indexes = [];
    var offset = 0;
    var matchLength = 1;
    var result;
    var str = storyContents;
    var isSameLength = (str.length == currentStory.characters.length);
    if (isSameLength) {
        while ((result = str.search(/[\u0003\uFFFC]/)) !== -1) {
            indexes.push(offset + result);
            offset += result + matchLength;
            str = str.slice(result + matchLength);
        }
    }
    while (currentIndex < currentStory.characters.length) {
        if (isSameLength) {
            var foundNext = false;
            var isAlreadyIndex = false;
            for (var j = 0; j < indexes.length; j++) {
                if (indexes[j] >= currentIndex) {
                    if (indexes[j] > currentIndex) {
                        currentIndex = indexes[j];
                        foundNext = true;
                    }
                    else {
                        isAlreadyIndex = true;
                    }
                    break;
                }
            }
            if (!foundNext && !isAlreadyIndex) {
                currentIndex = currentStory.characters.length - 1;
            }
        }
        currentChar = currentStory.characters.item(currentIndex);
        isMarkerChar = false;

        if (currentChar.contents == SpecialCharacters.END_NESTED_STYLE) {
            isMarkerChar = true;
            nestedMarkers.push([currentChar, "", false]);
            lastCharIndex[0] = currentIndex;

            if (!isSpoiled) {
                startChar.textFrames.item(0).insertLabel("LS_STATE", "normal");
                return true;
            } else {
                startChar.textFrames.item(0).insertLabel("LS_STATE", "spoiled");
                return false;
            }
        } 
        else if (currentChar.textFrames.length > 0 &&
            currentChar.textFrames.item(0).extractLabel("LS_TAG") == "snippet start") {
            isMarkerChar = true;

            if (isWithFixing && currentChar.textFrames.item(0).extractLabel("checked") == "yes") {
                fixNesting(currentChar, nestedMarkers);
                currentIndex = nestedMarkers[nestedMarkers.length - 1][0].index;
                var state = "";
                if (nestedMarkers[nestedMarkers.length - 1][0].textFrames.length > 0)
                    state = nestedMarkers[nestedMarkers.length - 1][0].textFrames.item(0).extractLabel("LS_STATE");

                if (state == "spoiled" || state == "lonely") {
                    isSpoiled = true;
                }
            } else {
                if (!surveyAndCheck(currentChar, lastCharIndex, nestedMarkers, isWithFixing)) {
                    isSpoiled = true;
                }
                currentIndex = lastCharIndex[0];
            }
        }

        if (!isMarkerChar) {
            if (!surveyAndCheckSubStories(currentChar, nestedMarkers)) {
                isSpoiled = true;
            }
            nestedMarkers[nestedMarkers.length - 1][1] += currentChar.contents;

            if (nestedMarkers[nestedMarkers.length - 1][2] === false &&
                (currentChar.allPageItems.length > 0 || currentChar.tables.length > 0)) {
                nestedMarkers[nestedMarkers.length - 1][2] = true;
            }
        }

        currentIndex++;
    }

    startChar.textFrames.item(0).insertLabel("LS_STATE", "lonely");
    lastCharIndex[0] = currentStory.characters.length - 1;
    return false;
}


function surveyAndCheckSubStories(rootChar, nestedMarkers, isWithFixing) {
	var subStoriesArray = new Array;
	if (rootChar.tables.length > 0) {
		solelyTable = rootChar.tables.item(0);
		for (var c = 0; c < solelyTable.cells.length; c++) {
			if (solelyTable.cells.item(c).characters.length > 0)
				subStoriesArray.push (solelyTable.cells.item(c).characters.item(0).parent);
		}
	}
	else {
		var c;
		var subFramesArray = new Array;
		for (c = 0; c < rootChar.allPageItems.length; c++) {
			if (rootChar.allPageItems[c] instanceof TextFrame) {
				if ((rootChar.allPageItems[c].parent instanceof Character)) {
					if (rootChar.allPageItems[c].parent == rootChar)
						subFramesArray.push(rootChar.allPageItems[c]);
				}
				else {
					subFramesArray.push(rootChar.allPageItems[c]);
				}
			}
		}
		for (c = 0; c < subFramesArray.length; c++) {
			subStoriesArray.push(subFramesArray[c].parentStory);
		}
	}
	var isSpoiled = false;
	for (c = 0; c < subStoriesArray.length; c++) {
		var currentIndex = 0;
		var currentChar;
		var isMarkerChar;
        var storyContents = subStoriesArray[c].contents;

        var indexes = [];
        var offset = 0;
        var matchLength = 1;
        var result;
        var str = storyContents;
        var isSameLength = (str.length == currentStory.characters.length);
        if (isSameLength) {
            while ((result = str.search(/[\u0003\uFFFC]/)) !== -1) {
                indexes.push(offset + result);
                offset += result + matchLength;
                str = str.slice(result + matchLength);
            }
        }
		while (currentIndex < subStoriesArray[c].characters.length) {
            if (isSameLength) {
                var foundNext = false;
                var isAlreadyIndex = false;
                for (var j = 0; j < indexes.length; j++) {
                    if (indexes[j] >= currentIndex) {
                        if (indexes[j] > currentIndex) {
                            currentIndex = indexes[j];
                            foundNext = true;
                        }
                        else {
                            isAlreadyIndex = true;
                        }
                        break;
                    }
                }
                if (!foundNext && !isAlreadyIndex) {
                    currentIndex = currentStory.characters.length - 1;
                }
            }
			currentChar = subStoriesArray[c].characters.item(currentIndex);
            isMarkerChar = false;
            if (currentChar.contents == SpecialCharacters.END_NESTED_STYLE) { //It's an end marker chracter
                isMarkerChar = true;
                isSpoiled = true;
            }
			else if (currentChar.textFrames.length > 0) {
				if (currentChar.textFrames.item(0).extractLabel("LS_TAG") == "snippet start") { //It's a start marker chracter
					isMarkerChar = true;
                    if (isWithFixing && currentChar.textFrames.item(0).extractLabel("checked") == "yes") {
                        fixNesting (currentChar, nestedMarkers);
                        currentIndex = nestedMarkers[nestedMarkers.length-1][0].index;
                        state = "";
                        if (nestedMarkers[nestedMarkers.length-1][0].textFrames.length > 0)
                            state = nestedMarkers[nestedMarkers.length-1][0].textFrames.item(0).extractLabel("LS_STATE");
                        if (state == "spoiled" || state == "lonely")
                            isSpoiled = true;
                    }
                    else {
                        var lastCharIndex = new Array;
                        lastCharIndex.push(0);
                        if (!surveyAndCheck(currentChar, lastCharIndex, nestedMarkers, isWithFixing))
                            isSpoiled = true;
                        currentIndex = lastCharIndex[0];
                    }
				}
			}
			if (!isMarkerChar) {
				if (!surveyAndCheckSubStories(currentChar, nestedMarkers))
					isSpoiled = true;
			}
			currentIndex ++;
		}
	}
	return !isSpoiled;
}

function logDebug(message) {
    var desktopPath = Folder.desktop.fsName;
    var logFile = new File(desktopPath + "/tree_shade_debug.log");
    if (logFile.open("a")) {
        logFile.writeln(message);
        logFile.close();
    }
}

function tsPlaceLiveSnippetFromUser (selectedObject) {
    tsPlaceLiveSnippet (selectedObject, null, null, null, null, null, false, true, [], [], -1, -1, []);
}

function tsPlaceLiveSnippet (selectedObject, targetPage, dynamicSnippetCode, targetDocument, snippetFile, snippetContent, isToRemoveMarks, isPreviousHorizontal, pageItemsIDs, cellsDeep, deepIndex, cellsIndex, importedReplace) {
    if (selectedObject == null) {
        if (targetPage) {
            var createdTextFrame = targetPage.textFrames.add();
            createdTextFrame.geometricBounds = getMarginBounds (targetDocument, targetPage);
            selectedObject = createdTextFrame.insertionPoints.item (-1);
        }
        else {
            return false;
        }
    }
	var targetText;
	switch (selectedObject.constructor.name){
		case "InsertionPoint":
		case "Character":		
		case "Word":
		case "Paragraph":
		case "Text":
		case "TextColumn":
		case "TextStyleRange":
			targetText = selectedObject;
			break;
		case "FormField":
		case "GraphicLine":
		case "Group":
		case "Oval":
		case "Polygon":
		case "Rectangle":
		case "TextFrame":	
		case "Table":
		case "Cell":
			targetText = downLevel (selectedObject);
			if (targetText == null) {
				return false;
			}
			break;
		default:
			alert("Please select a text or object"); 
			return false;
    }
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
    if (!targetDocument) {
        targetDocument = app.activeDocument;
    }
    var isFromUser = false;
    var dynamicContentPhrase = null;
    var stylePhrase = null;
    if (!dynamicSnippetCode) {
        isFromUser = true;
        dynamicSnippetCode = targetDocument.extractLabel ("LS_DEFAULT_DYNAMIC");
        dynamicSnippetCode = tsGetText (["Live Snippet Identifier (Branch and Name):",
                                             "Use <.> for the snippet file parent folder name, add more dots for parent of parent...",
                                             "To reference to the document instead of the snippet file add dollar sign like this '<.$>'.",
                                             "Use also: <Increment>, <Container_File_Branch>, <Container_File_Name>, " +
                                             "<Doc_Name>, <Doc_Order>, <Page_Num>, <Applied_Section_Prefix>, <Applied_Section_Marker>," +
                                             "<Master_Prefix>, <Master_Name>, <Master_Full_Name>, <Char_Order>, <Anchored_Order>," +
                                             "<Para_Order>, <Para_First>, <Para_Before>, <Para_This>, <Para_After>, <Para_Style>," +
                                             "<Col_Head>, <Col_Num>, <Row_Head>, <Row_Num>, <Shift1>.",
                                             "",
                                             "The following accepts abbreviation and used for date, time and random code:",
                                             "<Day_Seconds>, <Day_Minutes>, <Day_Hours>, <Weekday_Order>, <Weekday_Name>, <Weekday_Abbreviation>,",
                                             "<Month_Date>, <Month_Order>, <Month_Name>, <Month_Abbreviation>, <Year_Full>,",
                                             "<Year_Abbreviation>, <Time_Decimal>, <Time_Hexadecimal>, <Time_Shortest>,",
                                             "<Random_Style_Long> Where 'Style' a combination of symbols 'a' for small letters,",
                                             "'A' capital letters, 'h' small hexadecimal, 'H' large hexadecimal and '1' for numbers.",
                                             "For example <Random_1A_8> means random 8 characters consists of capital letters and numbers",
                                             ""
                                            ], dynamicSnippetCode, false);
        if (!dynamicSnippetCode) {
            exit();
        }
        targetDocument.insertLabel ("LS_DEFAULT_DYNAMIC", dynamicSnippetCode);
    }
    var isIndividuals = null;
    var allParts = dynamicSnippetCode.split ("::");
    dynamicSnippetCode = allParts[0];
    if (dynamicSnippetCode.slice (-1) == "$") {
        isIndividuals = false;
        dynamicSnippetCode = dynamicSnippetCode.slice (0, -1);
    }
    if (isFromUser) {
        dynamicSnippetCode = dynamicSnippetCode.replace (/\\/g, "/");
        if (dynamicSnippetCode.indexOf ("/") == 0) {
            dynamicSnippetCode = dynamicSnippetCode.slice (1);
            var docPath = preparePath (targetDocument.filePath.fsName.replace(/\\/g, "/"));
            docPath = docPath.replace (workshopPath, "");
            if (docPath != "") {
                docPath = docPath.slice (1);
                docPathSplitted = docPath.split ("/");
                for (var dps = 0; dps < docPathSplitted.length; dps++) {
                    if (dynamicSnippetCode.indexOf (docPathSplitted[dps] + "/") == 0) {
                        dynamicSnippetCode = dynamicSnippetCode.replace (docPathSplitted[dps] + "/", "");
                    }
                    else {
                        break;
                    }
                }
            }
        }
        else if (dynamicSnippetCode.indexOf ("./") == 0) {
            dynamicSnippetCode = dynamicSnippetCode.slice (2);
        }
    }
    if (allParts.length > 1) {
        if (allParts[1]) {
            var dynamicPartSplitted = allParts[1].split (":");
            if (dynamicPartSplitted.length > 1 && dynamicPartSplitted[1].search (/Para|Char|Obj|Height|Width|Area|Rotate|Shear|Short|Long|Merge|Cell|Row|Column|Table|Font|Style|Color|Size|Justification|characterDirection|paragraphDirection/i) != -1) {
                stylePhrase = allParts[1];
            }
            else {
                dynamicContentPhrase = allParts[1];
            }
        }
        allParts.splice (0, 2);
        if (stylePhrase) {
            allParts.unshift (stylePhrase);
        }
        if (allParts.length > 0) {
            stylePhrase = allParts.join ("::");
        }
    }
    createMarkersStyles(targetDocument);
    var pathAndName = solveSnippetCode (targetDocument, targetText.constructor.name == "InsertionPoint"? targetText : targetText.insertionPoints.item (0), dynamicSnippetCode, null);
    if (pathAndName[1] == "") {
        pathAndName[1] = "0001";
    }
    var textStory = targetText.insertionPoints.item(0).parent;
    var startIndex = targetText.insertionPoints.item(0).index;
    insertMarkers (targetDocument, targetText, pathAndName[0], pathAndName[1], -1, isIndividuals);
    var lastCharIndex = new Array;
	var nestedMarkers = new Array;
	lastCharIndex.push(0);
    surveyAndCheck (textStory.characters.item(startIndex), lastCharIndex, nestedMarkers, false);
    if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(nestedMarkers[0][0].textFrames.item(0));
        if (isFromUser)
            alert ("The instance has no end.");
        exit();
    }
	else if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
        reflectSpoiled (nestedMarkers);
        if (isFromUser)
		    alert ("The instance is spoiled.");
        exit();
    }
    var isDynamic = false;
    if (dynamicSnippetCode.lastIndexOf ("/") != -1) {
        if (dynamicSnippetCode.slice (0, dynamicSnippetCode.lastIndexOf ("/")) != pathAndName[0] || pathAndName[1] != dynamicSnippetCode.slice (dynamicSnippetCode.lastIndexOf ("/") + 1)) {
            isDynamic = true;
        }
    }
    else if (dynamicSnippetCode != pathAndName[0]) {
        isDynamic = true;
    }
    if (isDynamic) {
        nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_PREFIX_NUMBER", dynamicSnippetCode);
    }
    if (pathAndName.length > 2) {
        nestedMarkers[0][0].textFrames.item(0).insertLabel ("LS_INCREMENT_STATE", pathAndName[2]);
        nestedMarkers[0][0].textFrames.item(0).insertLabel ("LS_INCREMENT", pathAndName[3]);
    }
    if (dynamicContentPhrase) {
        nestedMarkers[0][0].textFrames.item(0).insertLabel ("LS_SUM", "DYNAMIC");
        nestedMarkers[0][0].textFrames.item(0).insertLabel ("LS_DYNAMIC", dynamicContentPhrase);
        nestedMarkers[0][0].textFrames.item(0).insertLabel ("LS_DYNAMIC_INLINE", "YES");
    }
    if (stylePhrase) {
        nestedMarkers[0][0].textFrames.item(0).insertLabel("TS_LS_STYLE", stylePhrase);
    }
    var snippetsCacheList = new Array;
    return updateInstance (targetDocument, snippetFile, nestedMarkers, true, dynamicContentPhrase, null, snippetContent, isToRemoveMarks, snippetsCacheList, isPreviousHorizontal, pageItemsIDs, cellsDeep, deepIndex, cellsIndex, importedReplace);
}

function getPathAndName (targetDocument, targetFile) {
    if (targetFile instanceof File) {
        var docPath = preparePath (targetDocument.filePath.fsName.replace(/\\/g, "/"));
        docPath = docPath + "/";
        var snippetFullPath = targetFile.fsName.replace(/\\/g, "/");

        while (docPath.indexOf ("/") != -1 && snippetFullPath.indexOf ("/") != -1 && docPath.slice (0, docPath.indexOf ("/")) == snippetFullPath.slice (0, snippetFullPath.indexOf ("/"))) {
            docPath = docPath.slice (docPath.indexOf ("/")+1);
            snippetFullPath = snippetFullPath.slice (snippetFullPath.indexOf ("/")+1);
        }
        var prefix = (snippetFullPath.indexOf ("/") != -1)? snippetFullPath.slice (0, snippetFullPath.lastIndexOf ("/")) : "";
        var name = snippetFullPath.slice (snippetFullPath.lastIndexOf ("/") + 1);
        name = getPureName (name, false);
        return prefix + "/" + name;
    }
    else {
        return '#';
    }
}

function isInsideProject(targetDocument, targetFile) {
    if (!(targetFile instanceof File) || !(targetDocument instanceof Document)) {
        return false;
    }
    var docPath = preparePath(targetDocument.filePath.fsName.replace(/\\/g, "/")) + "/";
    var snippetFullPath = targetFile.fsName.replace(/\\/g, "/");
    return snippetFullPath.indexOf(docPath) === 0;
}


function getLiveSnippetContentAndSUM (nestedMarkers, startIndex) {
    var liveSnippetContent = "";
    var sum = 0;
	var level = 0;
	for (var c = startIndex; c < nestedMarkers.length; c++) {
        if (nestedMarkers[c][0].contents == SpecialCharacters.END_NESTED_STYLE) {
            level--;
            if (level == 1) {
                for (var ce = 0; ce < nestedMarkers[c][1].length; ce++) {
                    sum += (liveSnippetContent.length + ce) * nestedMarkers[c][1].charCodeAt (ce);
                }
                liveSnippetContent += nestedMarkers[c][1];
            }
            else if (level == 0) {
                break;
            }
        }
        else {
            if (level == 0) {
                for (var cf = 0; cf < nestedMarkers[c][1].length; cf++) {
                    sum += (liveSnippetContent.length + cf) * nestedMarkers[c][1].charCodeAt (cf);
                }
                liveSnippetContent += nestedMarkers[c][1];
            }
            else if (level == 1) {
                liveSnippetContent += "<TS_LS>" + nestedMarkers[c][0].textFrames.item(0).extractLabel("LS_ID") + "<TS_LS>";
            }
            level++;
        }
    }
    var returnedArr = new Array;
    returnedArr.push (liveSnippetContent);
    returnedArr.push (sum.toString ());
    return returnedArr;
}

function getSubnestedMarkers (nestedMarkers, startIndex) {
    var subnestedMarkers = new Array;
    var level = 0;
	for (var c = startIndex; c < nestedMarkers.length; c++) {
        if (nestedMarkers[c][0].contents == SpecialCharacters.END_NESTED_STYLE) {
            level--;
            subnestedMarkers.push (nestedMarkers[c]);
            if (level == 0) {
                break;
            }
        }
        else {
            subnestedMarkers.push (nestedMarkers[c]);
            level++;
        }
    }
    return subnestedMarkers;
}

function checkDocument (targetDocument) {
    checkAllLiveSnippets (targetDocument, false, false);
}

function checkAllLiveSnippets (targetDocument, isToSubmit, isToUpdate) {
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
	createMarkersStyles(targetDocument);
    snippetFilePath_ID_List = new Array; //reset to speed assigning files IDs
	var startFramesArray = new Array;
	var c;
    var executionOrders = new Array; //executionOrders[x][0] the order, executionOrders[x][1] counter
    for (var p = 0; p < targetDocument.pages.length; p++) {
        for (var mtf = 0; mtf < targetDocument.pages[p].allPageItems.length; mtf++) {
            if (targetDocument.pages[p].allPageItems[mtf].extractLabel("LS_TAG") == "snippet start") {
                var executeOrder = 1;
                var liveSnippetID = targetDocument.pages[p].allPageItems[mtf].extractLabel("LS_ID");
                if (liveSnippetID.slice (0, liveSnippetID.lastIndexOf ("/")) == "#") {
                    var lsOrder = parseInt (liveSnippetID.slice (liveSnippetID.lastIndexOf ("/") + 1), 10);
                    if (lsOrder != null) {
                        executeOrder = lsOrder;
                    }
                }
                var insertingIndex = 0;
                var orderIndex = executionOrders.length;
                var isOrderIndexNew = true;
                for (var eo = 0; eo < executionOrders.length; eo++) {
                    if (executionOrders[eo][0] >= executeOrder) {
                        orderIndex = eo;
                        if (executionOrders[eo][0] == executeOrder) {
                            isOrderIndexNew = false;
                        }
                        break;
                    }
                }
                if (isOrderIndexNew) {
                    executionOrders.splice (orderIndex, 0, 1);
                    executionOrders[orderIndex]= [executeOrder, 0];
                }
                for (eo = 0; eo <= orderIndex; eo++) {
                    insertingIndex += executionOrders[eo][1];
                }
                startFramesArray.splice (insertingIndex, 0, targetDocument.pages[p].allPageItems[mtf]);
                executionOrders[orderIndex][1]++;
                targetDocument.pages[p].allPageItems[mtf].insertLabel("checked", "");
                targetDocument.pages[p].allPageItems[mtf].insertLabel("LS_STATE", "lonely");
            }
        }
    }

	var nestedMarkers = new Array;
	for (c = 0; c < startFramesArray.length; c++) {
		var checkedFrame = startFramesArray[c];
		if (checkedFrame.extractLabel("checked") == "yes")
            continue;
		var lastCharIndex = new Array; //It's stored in an array with one variable to be used as returned value
        lastCharIndex.push(0);
        surveyAndCheck(checkedFrame.parent, lastCharIndex, nestedMarkers, true);
    }
	for (c = 0; c < startFramesArray.length; c++) {
		startFramesArray[c].insertLabel("checked", "");
	}
    var level = 0;
    var nestedMarkersOrderedByStories = new Array;
    //nestedMarkersOrderedByStories[x] list of nested live snippets in a stoxy x
	for (c = nestedMarkers.length - 1; c >= 0; c--) {
        if (nestedMarkers[c][0].contents == SpecialCharacters.END_NESTED_STYLE) {
            level++;
        }
        else {
            if (nestedMarkers[c][0].textFrames.item(0).extractLabel("LS_STATE") != "lonely") {
                level--;
            }
            else {
                changeStateAppearance(nestedMarkers[c][0].textFrames.item(0));
            }
            if (level == 0) {
                var storyIndexInList = -1;
                for (var s = 0; s < nestedMarkersOrderedByStories.length; s++) {
                    if (nestedMarkers[c][0].parent == nestedMarkersOrderedByStories[s][0][0][0].parent) {
                        storyIndexInList = s;
                        break;
                    }
                }
                if (storyIndexInList == -1) {
                    storyIndexInList = nestedMarkersOrderedByStories.length;
                    nestedMarkersOrderedByStories.push ([]);
                }
                var nestedIndex = nestedMarkersOrderedByStories[storyIndexInList].length;
                for (var nmo = 0; nmo < nestedMarkersOrderedByStories[storyIndexInList].length; nmo++) {
                    if (nestedMarkers[c][0].index < nestedMarkersOrderedByStories[storyIndexInList][nmo][0][0].index) {
                        nestedIndex = nmo;
                        break;
                    }
                }
                var subNestedMarkers = getSubnestedMarkers (nestedMarkers, c);
                nestedMarkersOrderedByStories[storyIndexInList].splice (nestedIndex, 0, subNestedMarkers);
            }
        }
    }
    var submittedLiveSnippets = new Array;
    var bridgeTalkBody = "";
    var bodyNewLine = "";
    var snippetsCacheList = new Array;
    for (var sn = 0; sn < nestedMarkersOrderedByStories.length; sn++) {
        for (c = nestedMarkersOrderedByStories[sn].length - 1; c >= 0; c--) {
            var startTextFrame = nestedMarkersOrderedByStories[sn][c][0][0].textFrames.item(0);
            var isItCurrent = false;
            if (currentLiveSnippetStartChar == nestedMarkersOrderedByStories[sn][c][0][0])
                isItCurrent = true;
            switch (startTextFrame.extractLabel("LS_STATE")) {
                case "spoiled":
                    changeStateAppearance(startTextFrame); 
                    break;
                case "normal":
                    var isSubmittedOrNeedToSubmit = false;
                    if (!isItCurrent) {
                        if (startTextFrame.paragraphs.item(0).appliedCharacterStyle != targetDocument.characterStyles.itemByName("Live Snippet Need to Submit")) {
                            if (startTextFrame.extractLabel("LS_SUM") != "DYNAMIC") {
                                //Check DynamicPathAndName
                                var liveSnippetID = startTextFrame.extractLabel("LS_ID");
                                if (liveSnippetID.indexOf ("#.") == 0) {
                                    continue;
                                }
                                var dynamicSnippetCode = startTextFrame.extractLabel("LS_PREFIX_NUMBER");
                                var isToChange = false;
                                if (dynamicSnippetCode) {
                                    var pathAndName = solveSnippetCode (targetDocument, nestedMarkersOrderedByStories[sn][c][0][0].parent.insertionPoints.item (nestedMarkersOrderedByStories[sn][c][0][0].index), dynamicSnippetCode, null);
                                    if (pathAndName[1] == "") {
                                        pathAndName[1] = "0001";
                                    }
                                    if (pathAndName[0] || pathAndName[1]) {
                                        liveSnippetIDSplitted = new Array;
                                        liveSnippetIDSplitted.push (liveSnippetID.slice (0, liveSnippetID.lastIndexOf ("/")));
                                        liveSnippetIDSplitted.push (liveSnippetID.slice (liveSnippetID.lastIndexOf ("/") + 1));
                                        if (pathAndName[0]) {
                                            if (liveSnippetIDSplitted[0] != pathAndName[0]) {
                                                isToChange = true;
                                            }
                                        }
                                        if (pathAndName[1]) {
                                            if (liveSnippetIDSplitted[1] != pathAndName[1]) {
                                                isToChange = true;
                                            }
                                        }
                                        if (isToChange) {
                                            changeMarkers (nestedMarkersOrderedByStories[sn][c], pathAndName[0], pathAndName[1]);
                                            startTextFrame.insertLabel("LS_TSID", "");
                                        }
                                    }
                                }

                                var liveSnippetContentAndSUM = getLiveSnippetContentAndSUM (nestedMarkersOrderedByStories[sn][c], 0);
                                if (startTextFrame.extractLabel("LS_SUM") != liveSnippetContentAndSUM[1]) {
                                    if (isToSubmit) {
                                        var isAlreadySubmitted = false;
                                        for (var sls = 0; sls < submittedLiveSnippets.length; sls++) {
                                            if (submittedLiveSnippets[sls] == startTextFrame.extractLabel("LS_ID")) {
                                                isAlreadySubmitted = true;
                                                break;
                                            }
                                        }
                                        if (!isAlreadySubmitted) {
                                            isSubmittedOrNeedToSubmit = true;
                                            submitInstance (targetDocument, nestedMarkersOrderedByStories[sn][c]);
                                            submittedLiveSnippets.push (startTextFrame.extractLabel("LS_ID"));
                                        }
                                    }
                                    else {
                                        isSubmittedOrNeedToSubmit = true;
                                        startTextFrame.insertLabel("LS_STATE", "need to submit");
                                        changeStateAppearance(startTextFrame);
                                    }
                                }
                            }
                        }
                    }
                    if (!isSubmittedOrNeedToSubmit) {
                        if (startTextFrame.paragraphs.item(0).appliedCharacterStyle != targetDocument.characterStyles.itemByName("Live Snippet Modified")) {
                            if (isToUpdate && (startTextFrame.extractLabel("LS_SUM") == "DYNAMIC" || startTextFrame.extractLabel("TS_LS_STYLE"))) {
                                updateInstance (targetDocument, null, nestedMarkersOrderedByStories[sn][c], true, null, null, null, false, snippetsCacheList, true, [], [], -1, -1, []);
                            }
                            else if (isToUpdate && startTextFrame.extractLabel("LS_PREFIX_NUMBER")) {
                                updateInstance (targetDocument, null, nestedMarkersOrderedByStories[sn][c], false, null, null, null, false, snippetsCacheList, true, [], [], -1, -1, []);
                            }
                            else {
                                var snippetFile = findTargetFile (targetDocument, startTextFrame.parent);
                                if (snippetFile == false) {
                                    continue;
                                }
                                else if (!(snippetFile instanceof File)) {
                                    var storedTSID = startTextFrame.extractLabel ("LS_TSID");
                                    var tracedPathAndName = "";
                                    if (storedTSID) {
                                        var fileID_Path_newID = [storedTSID];
                                        getPath (fileID_Path_newID);
                                        if (fileID_Path_newID[1] && File (workshopPath + fileID_Path_newID[1]).exists) {
                                            snippetFile = new File (workshopPath + fileID_Path_newID[1]);
                                            if (isInsideProject (targetDocument, snippetFile)) {
                                                removeIndex_snippetFilePath_ID_List (storedTSID);
                                                tracedPathAndName = getPathAndName (targetDocument, snippetFile);
                                                var pathAndName = ["", ""];
                                                pathAndName[0] = tracedPathAndName.slice (0, tracedPathAndName.lastIndexOf ("/"));
                                                pathAndName[1] = tracedPathAndName.slice (tracedPathAndName.lastIndexOf ("/") + 1);
                                                changeMarkers (nestedMarkersOrderedByStories[sn][c], pathAndName[0], pathAndName[1]);
                                            }
                                        }
                                    }
                                    if (tracedPathAndName == "") {
                                        startTextFrame.insertLabel("LS_STATE", "missing");
                                        changeStateAppearance(startTextFrame); //Changing the appearance
                                        //targetDocument.insertLabel ("TS_DELAY_CHECKIN", "YES"); 
                                        continue;
                                    }
                                }
                                var isFileNested = (File.decode (snippetFile.name).slice(File.decode (snippetFile.name).length - 5) == ".IDMS");
                                if (isFileNested) {
                                    updateInstance (targetDocument, snippetFile, nestedMarkersOrderedByStories[sn][c], false, null, null, null, false, snippetsCacheList, true, [], [], -1, -1, []);
                                }
                                else { 
                                    var markerLS_TSID = startTextFrame.extractLabel("LS_TSID");
                                    var fidIndex = getIndex_snippetFilePath_ID_List (snippetFile);
                                    if (fidIndex == -1) {
                                        fidIndex = snippetFilePath_ID_List.length - 1;
                                        if (snippetFilePath_ID_List[fidIndex][1] == "") {
                                            targetDocument.insertLabel ("TS_DELAY_CHECKIN", "YES");
                                            //update file state
                                            bridgeTalkBody += bodyNewLine + "tsUpdateFileState (new File (\"" + snippetFile.fsName.replace(/\\/g, "/") + "\"), -2, -1);";
                                            bodyNewLine = "\n";
                                        }
                                    }
                                    if ((parseInt(startTextFrame.extractLabel("LS_DATE"), 10) < snippetFile.modified.getTime ()) || !markerLS_TSID || (markerLS_TSID != snippetFilePath_ID_List[fidIndex][1])) {
                                        if (isToUpdate && !isItCurrent) {
                                            updateInstance (targetDocument, snippetFile, nestedMarkersOrderedByStories[sn][c], true, null, null, null, false, snippetsCacheList, true, [], [], -1, -1, []);
                                        }
                                        else {
                                            startTextFrame.insertLabel("LS_STATE", "modified");
                                            changeStateAppearance(startTextFrame); 
                                        }
                                    }
                                    else { //The file is updated
                                        startTextFrame.insertLabel("LS_STATE", "updated");
                                        changeStateAppearance(startTextFrame);
                                    }
                                }
                            } 
                        }
                        else {
                            if (isToUpdate && !isItCurrent) {
                                updateInstance (targetDocument, null, nestedMarkersOrderedByStories[sn][c], true, null, null, null, false, snippetsCacheList, true, [], [], -1, -1, []);
                            }
                        }
                    }
                    break;
            }
        }
    }
    //update file state
    var targetBridge = BridgeTalk.getSpecifier ("bridge");
    if (targetBridge && targetBridge.appStatus != "not installed") {
        var talkBridge = new BridgeTalk;
        talkBridge.target = targetBridge;
        talkBridge.body = bridgeTalkBody;
        talkBridge.send ();
    }

	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
	return true;
}

function selectInstance (selectedPortionAndTargetWay) {
    var selectedPortion = selectedPortionAndTargetWay[0];
    var targetWay = selectedPortionAndTargetWay[1]
	var lastCharIndex = new Array;
	var nestedMarkers = new Array;
	if (!getContainerSnippet (selectedPortion, lastCharIndex, nestedMarkers, "first marker")) { //"first marker"  or "first snippet marker" or "base marker"
		return false; 
	}
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(startChar.textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
	else if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
		reflectSpoiled (nestedMarkers);
		alert ("The instance is spoiled");
		return false;
	}
	var firstChar;
	var lastChar;
    switch (targetWay) {
        case "content":
            firstChar = nestedMarkers[0][0].parent.characters.item(nestedMarkers[0][0].index+1);
            lastChar = nestedMarkers[0][0].parent.characters.item(nestedMarkers[nestedMarkers.length-1][0].index-1);
            break;
        case "whole_instance":
            firstChar = nestedMarkers[0][0].parent.characters.item(nestedMarkers[0][0].index);
            lastChar = nestedMarkers[0][0].parent.characters.item(nestedMarkers[nestedMarkers.length-1][0].index);
            break;
        case "to_start_border":
            firstChar = nestedMarkers[0][0].parent.characters.item(nestedMarkers[0][0].index+1);
            lastChar = nestedMarkers[0][0].parent.characters.item(selectedPortion.insertionPoints[selectedPortion.insertionPoints.length-1].index-1);
            break;
        case "to_end_border":
            firstChar = nestedMarkers[0][0].parent.characters.item(selectedPortion.insertionPoints.item(0).index);
            lastChar = nestedMarkers[0][0].parent.characters.item(nestedMarkers[nestedMarkers.length-1][0].index-1);
            break;
        case "start_marker":
            nestedMarkers[0][0].textFrames.item(0).select();
            exit();
            break;
        case "end_marker":
            nestedMarkers[nestedMarkers.length-1][0].select();
            exit();
            break;
    }
	nestedMarkers[0][0].parent.characters.itemByRange(firstChar, lastChar).select();
}

function selectFirstChild (selectedPortion) {
	var myDoc = app.activeDocument;
	var lastCharIndex = new Array;
	var nestedMarkers = new Array;
	if (!getContainerSnippet (selectedPortion, lastCharIndex, nestedMarkers, "first marker")) { //"first marker"  or "first snippet marker" or "base marker"
		return false; 
	}

	createMarkersStyles(myDoc);
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(startChar.textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
	else if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
		reflectSpoiled (nestedMarkers);
		alert ("The instance is spoiled");
		return false;
	}
	if (nestedMarkers.length > 2) {
		var firstChar = nestedMarkers[1][0].parent.characters.item(nestedMarkers[1][0].index+1);
		var lastChar = nestedMarkers[1][0].parent.characters.item(nestedMarkers[getEndIndex(1, nestedMarkers)][0].index-1);
		nestedMarkers[1][0].parent.characters.itemByRange(firstChar, lastChar).select();
	}
}

function selectNextBrother(selectedPortion) {
	var myDoc = app.activeDocument;
	var lastCharIndex = new Array;
	var nestedMarkers = new Array;
	if (!getContainerSnippet (selectedPortion, lastCharIndex, nestedMarkers, "first marker")) { //"first marker"  or "first snippet marker" or "base marker"
		return false; 
	}

	createMarkersStyles(myDoc);
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(startChar.textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
	else if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
		reflectSpoiled (nestedMarkers);
		alert ("The instance is spoiled");
		return false;
	}
	var currentName = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_ID");
	var baseStart = null;
	if (nestedMarkers[0][0].index > 0) {
		baseStart = getStartChar (nestedMarkers[0][0].parent.characters.item(nestedMarkers[0][0].index - 1));
	}
	else {
		if (nestedMarkers[0][0].parent.constructor.name == "Cell")
			baseStart = getStartChar (nestedMarkers[0][0].parent);
		else
			baseStart = getStartChar (nestedMarkers[0][0].parentTextFrames[0]);
	}
	if (baseStart == null) 
		return false;
	nestedMarkers.length = 0;
	surveyAndCheck (baseStart, lastCharIndex, nestedMarkers, false);
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(startChar.textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
	else if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
		reflectSpoiled (nestedMarkers);
		alert ("The instance is spoiled");
		return false;
	}
	var childMarkers = new Array;
	var level = 0;
	var c;
	for (c = 0; c < nestedMarkers.length; c++) {
        if (nestedMarkers[c][0].contents == SpecialCharacters.END_NESTED_STYLE) {
			level--;
			if (level == 1)
				childMarkers.push(nestedMarkers[c][0]);
        }
        else {
			if (level == 1)
				childMarkers.push(nestedMarkers[c][0]);
			level++;
        }
	}
	var position = null;
	for (c = 0; c < childMarkers.length; c+=2) {
		if (childMarkers[c].textFrames.item(0).extractLabel("LS_ID") == currentName) {
			if (c < childMarkers.length - 3)
				position = c+2;
			break;
		}
	}
	if (position == null) {
		return false;
	}
	
	var firstChar = childMarkers[position].parent.characters.item(childMarkers[position].index+1);
	var lastChar = childMarkers[position].parent.characters.item(childMarkers[position+1].index-1);
	childMarkers[position].parent.characters.itemByRange(firstChar, lastChar).select();
}

function selectPreviousBrother(selectedPortion) {
	var myDoc = app.activeDocument;
	var lastCharIndex = new Array;
	var nestedMarkers = new Array;
	if (!getContainerSnippet (selectedPortion, lastCharIndex, nestedMarkers, "first marker")) { //"first marker"  or "first snippet marker" or "base marker"
		return false; 
	}

	createMarkersStyles(myDoc);
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(startChar.textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
	else if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
		reflectSpoiled (nestedMarkers);
		alert ("The instance is spoiled");
		return false;
	}
	var currentName = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_ID");
	var baseStart = null;
	if (nestedMarkers[0][0].index > 0) {
		baseStart = getStartChar (nestedMarkers[0][0].parent.characters.item(nestedMarkers[0][0].index - 1));
	}
	else {
		if (nestedMarkers[0][0].parent.constructor.name == "Cell")
			baseStart = getStartChar (nestedMarkers[0][0].parent);
		else
			baseStart = getStartChar (nestedMarkers[0][0].parentTextFrames[0]);
	}
	if (baseStart == null) 
		return false;
	nestedMarkers.length = 0;
	surveyAndCheck (baseStart, lastCharIndex, nestedMarkers, false);
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(startChar.textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
	else if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
		reflectSpoiled (nestedMarkers);
		alert ("The instance is spoiled");
		return false;
	}
	var childMarkers = new Array;
	var level = 0;
	var c;
	for (c = 0; c < nestedMarkers.length; c++) {
        if (nestedMarkers[c][0].contents == SpecialCharacters.END_NESTED_STYLE) {
			level--;
			if (level == 1)
				childMarkers.push(nestedMarkers[c][0]);
        }
        else {
			if (level == 1)
				childMarkers.push(nestedMarkers[c][0]);
			level++;
        }
	}
	var position = null;
	for (c = 0; c < childMarkers.length; c+=2) {
		if (childMarkers[c].textFrames.item(0).extractLabel("LS_ID") == currentName) {
			if (c > 1)
				position = c-2;
			break;
		}
	}
	if (position == null) {
		return false;
	}
	
	var firstChar = childMarkers[position].parent.characters.item(childMarkers[position].index+1);
	var lastChar = childMarkers[position].parent.characters.item(childMarkers[position+1].index-1);
	childMarkers[position].parent.characters.itemByRange(firstChar, lastChar).select();
}

function reflectHorizontally (selectedObject) {
    if (selectedObject.constructor.name == "TextFrame") {
        if (selectedObject.extractLabel("LS_TAG") == "snippet start") {
            reflectMarkerHorizontally (selectedObject);
            return true;
        }
    }
	var lastCharIndex = new Array;
	var nestedMarkers = new Array;
	if (!getContainerSnippet (selectedObject, lastCharIndex, nestedMarkers, "first marker")) { //"first marker"  or "first snippet marker" or "base marker"
		return false; 
	}
    if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely")
        return false;
    reflectMarkerHorizontally (nestedMarkers[0][0].textFrames.item(0));
    return true;
}

function reflectVertically (selectedObject) {
    if (selectedObject.constructor.name == "TextFrame") {
        if (selectedObject.appliedObjectStyle.name == "Snippet End Marker") {
            reflectMarkerVertically (selectedObject);
            return true;
        }
    }
    return false;
}

function solveSnippetCode (targetDocument, startInsertionPoint, dynamicSnippetCode, parentID) {
    dynamicSnippetCode = tsSolveDateTimeRandomPortion (dynamicSnippetCode);
    dynamicSnippetCode = solveAbbreviation (dynamicSnippetCode);
    var pathAndName = [null, null];

    var nameDotsIndex = dynamicSnippetCode.search (/<\.+\$?>/);
    while (nameDotsIndex != -1) {
        var deepCount = 3;
        var dotsFirstMatch = dynamicSnippetCode.match (/<\.+\$?>/)[0];
        var docPath = preparePath (targetDocument.filePath.fsName.replace(/\\/g, "/"));
        var targetParent = Folder (workshopPath + docPath);
        if (dotsFirstMatch[dotsFirstMatch.length - 2] == '$') {
            deepCount = 4;
        }
        for (var pml = 0; pml < dotsFirstMatch.length - deepCount; pml++) {
            if (!targetParent.parent) {
                break;
            }
            targetParent = targetParent.parent;
        }
        var parentFolderName = "";
        if (targetParent && targetParent.name) {
            parentFolderName = File.decode (targetParent.name);
        }
        dynamicSnippetCode = dynamicSnippetCode.replace (dotsFirstMatch, parentFolderName);
        nameDotsIndex = dynamicSnippetCode.search (/<\.+\$?>/);
    }

    if (dynamicSnippetCode.search (/<Container_File_Branch>/i) != -1) {
        if (!parentID) {
            var lastCharIndex = new Array;
            var parentNestedMarkers = new Array;
            var baseStart = null;
            if (startInsertionPoint.index > 0) {
                baseStart = getStartChar (startInsertionPoint.parent.characters.item(startInsertionPoint.index - 1));
            }
            else {
                if (startInsertionPoint.parent.constructor.name == "Cell")
                    baseStart = getStartChar (startInsertionPoint.parent);
                else
                    baseStart = getStartChar (startInsertionPoint.parentTextFrames[0]);
            }
            if (baseStart != null) {
                if (getContainerSnippet (baseStart , lastCharIndex, parentNestedMarkers, "base marker")) { //"first marker"  or "first snippet marker" or "base marker"
                    parentID = parentNestedMarkers[0][0].textFrames.item(0).extractLabel("LS_ID");
                }
            }
        }
        if (parentID) {
            dynamicSnippetCode = dynamicSnippetCode.replace(/<Container_File_Branch>/gi, parentID.slice (0, parentID.lastIndexOf ("/")));
        }
    }
    if (dynamicSnippetCode.search (/<Container_File_Name>/i) != -1) {
        if (!parentID) {
            var lastCharIndex = new Array;
            var parentNestedMarkers = new Array;
            var baseStart = null;
            if (startInsertionPoint.index > 0) {
                baseStart = getStartChar (startInsertionPoint.parent.characters.item(startInsertionPoint.index - 1));
            }
            else {
                if (startInsertionPoint.parent.constructor.name == "Cell")
                    baseStart = getStartChar (startInsertionPoint.parent);
                else
                    baseStart = getStartChar (startInsertionPoint.parentTextFrames[0]);
            }
            if (baseStart != null) {
                if (getContainerSnippet (baseStart , lastCharIndex, parentNestedMarkers, "base marker")) { //"first marker"  or "first snippet marker" or "base marker"
                    parentID = parentNestedMarkers[0][0].textFrames.item(0).extractLabel("LS_ID");
                }
            }
        }
        if (parentID) {
            dynamicSnippetCode = dynamicSnippetCode.replace(/<Container_File_Name>/gi, parentID.slice (parentID.lastIndexOf ("/") + 1));
        }
    }
    if (dynamicSnippetCode.search (/<Doc_Name>/i) != -1) {
        var pureName = getPureName (targetDocument.name, false);
        dynamicSnippetCode = dynamicSnippetCode.replace(/<Doc_Name>/gi, pureName);
    }
    if (dynamicSnippetCode.search (/<Applied_Section_Prefix>/i) != -1) {
        var liveSnippetPage = startInsertionPoint.parentTextFrames[0].parentPage;
        if (liveSnippetPage) {
            if (liveSnippetPage.appliedSection) {
                dynamicSnippetCode = dynamicSnippetCode.replace(/<Applied_Section_Prefix>/gi, liveSnippetPage.appliedSection.sectionPrefix);
            }
        }
    }
    if (dynamicSnippetCode.search (/<Applied_Section_Marker>/i) != -1) {
        var liveSnippetPage = startInsertionPoint.parentTextFrames[0].parentPage;
        if (liveSnippetPage) {
            if (liveSnippetPage.appliedSection) {
                dynamicSnippetCode = dynamicSnippetCode.replace(/<Applied_Section_Marker>/gi, liveSnippetPage.appliedSection.marker);
            }
        }
    }
    if (dynamicSnippetCode.search (/<Master_Prefix>/) != -1) {
        var liveSnippetPage = startInsertionPoint.parentTextFrames[0].parentPage;
        if (liveSnippetPage) {
            dynamicSnippetCode = dynamicSnippetCode.replace(/<Master_Prefix>/gi, liveSnippetPage.appliedMaster.namePrefix);
        }
    }
    if (dynamicSnippetCode.search (/<Master_Name>/i) != -1) {
        var liveSnippetPage = startInsertionPoint.parentTextFrames[0].parentPage;
        if (liveSnippetPage) {
            dynamicSnippetCode = dynamicSnippetCode.replace(/<Master_Name>/gi, liveSnippetPage.appliedMaster.baseName);
        }
    }
    if (dynamicSnippetCode.search (/<Master_Full_Name>/i) != -1) {
        var liveSnippetPage = startInsertionPoint.parentTextFrames[0].parentPage;
        if (liveSnippetPage) {
            dynamicSnippetCode = dynamicSnippetCode.replace(/<Master_Full_Name>/gi, liveSnippetPage.appliedMaster.name);
        }
    }
    if (dynamicSnippetCode.search (/<Char_Order>/i) != -1) {
        var anchorChar  = getAnchoredChar (startInsertionPoint);
        if (anchorChar) {
            var thisPara = anchorChar.paragraphs.item (0);
            var charOrder = thisPara.characters.itemByRange(thisPara.characters.firstItem(), anchorChar).characters.length - 1;
            if (charOrder) {
                dynamicSnippetCode = dynamicSnippetCode.replace(/<Char_Order>/gi, tsFillZeros(charOrder, 4));
            }
        }
    }
    if (dynamicSnippetCode.search (/<Anchored_Order>/i) != -1) {
        var anchorChar  = getAnchoredChar (startInsertionPoint);
        if (anchorChar) {
            var thisPara = anchorChar.paragraphs.item (0);
            var anchoredChars = thisPara.characters.itemByRange(thisPara.characters.firstItem(), anchorChar).contents.toString ();
            var charOrder = -1;
            for (var i = 0; i < anchoredChars.length; i++) {
                if (anchoredChars.charCodeAt(i) == 65532) {
                    charOrder++;
                }
            }
            if (charOrder) {
                dynamicSnippetCode = dynamicSnippetCode.replace(/<Anchored_Order>/gi, tsFillZeros(charOrder, 4));
            }
        }
    }
    if (dynamicSnippetCode.search (/<Para_Order>/i) != -1) {
        var downChar  = downLevel (startInsertionPoint.parentTextFrames[0]);
        var firstPara;
        var basePara;
        var mainStory;
        if (downChar) {
            mainStory = downChar.parentStory;
            firstPara = mainStory.paragraphs.firstItem ();
            basePara = downChar.paragraphs.item (0);
        }
        else {
            mainStory = startInsertionPoint.parentStory;
            firstPara = mainStory.paragraphs.firstItem ();
            if (startInsertionPoint.parent.constructor.name == "Cell") {
                basePara = startInsertionPoint.parent.parentRow.parent.storyOffset.paragraphs.item (0);
            }
            else {
                basePara = startInsertionPoint.paragraphs.item (0);
            }
        }
        if (firstPara) {
            var paraOrder = mainStory.paragraphs.itemByRange (firstPara, basePara).paragraphs.length;
            dynamicSnippetCode = dynamicSnippetCode.replace(/<Para_Order>/gi, tsFillZeros (paraOrder, 4));
        }
    }
    if (dynamicSnippetCode.search (/<Para_First>/i) != -1) {
        var downChar  = downLevel (startInsertionPoint.parentTextFrames[0]);
        var firstPara;
        if (downChar) {
            firstPara = downChar.parentStory.paragraphs.firstItem ();
        }
        else {
            firstPara = startInsertionPoint.parentStory.paragraphs.firstItem ();
        }
        if (firstPara) {
            var paraContents = getParaPlainContents (firstPara);
            if (paraContents) {
                dynamicSnippetCode = dynamicSnippetCode.replace(/<Para_First>/gi, paraContents);
            }
        }
    }
    if (dynamicSnippetCode.search (/<Para_Before>/i) != -1) {
        var beforePara = startInsertionPoint.parent.paragraphs.previousItem (startInsertionPoint.paragraphs.item (0));
        if (beforePara) {
            var paraContents = getParaPlainContents (beforePara);
            if (paraContents) {
                dynamicSnippetCode = dynamicSnippetCode.replace(/<Para_Before>/gi, paraContents);
            }
        }
    }
    if (dynamicSnippetCode.search (/<Para_This>/i) != -1) {
        var thisPara = startInsertionPoint.paragraphs.item (0);
        if (thisPara) {
            var paraContents = getParaPlainContents (thisPara);
            if (paraContents) {
                dynamicSnippetCode = dynamicSnippetCode.replace(/<Para_This>/gi, paraContents);
            }
        }
    }
    if (dynamicSnippetCode.search (/<Para_After>/i) != -1) {
        var afterPara = startInsertionPoint.parent.paragraphs.nextItem (startInsertionPoint.paragraphs.item (0));
        if (afterPara) {
            var paraContents = getParaPlainContents (afterPara);
            if (paraContents) {
                dynamicSnippetCode = dynamicSnippetCode.replace(/<Para_After>/gi, paraContents);
            }
        }
    }
    if (dynamicSnippetCode.search (/<Para_Style>/i) != -1) {
        dynamicSnippetCode = dynamicSnippetCode.replace(/<Para_Style>/gi, startInsertionPoint.paragraphs.item (0).appliedParagraphStyle.name);
    }
    if (dynamicSnippetCode.search (/<Doc_Order>/i) != -1) {
        var docFileOrder = 1;
        var docParentFolder = new Folder (targetDocument.filePath.fsName.replace(/\\/g, "/"));
        var allDocs = docParentFolder.getFiles ("*.indd");
        sortFilesList (docParentFolder, allDocs, false);
        for (var ad = 0; ad < allDocs.length; ad++) {
            if (File.decode (allDocs[ad].name) == File.decode (targetDocument.name)) {
                docFileOrder = ad + 1;
                break;
            }
        }
        dynamicSnippetCode = dynamicSnippetCode.replace(/<Doc_Order>/gi, tsFillZeros (docFileOrder, 4));
    }
    if (dynamicSnippetCode.search (/<Page_Num>/i) != -1) {
        var liveSnippetPage = startInsertionPoint.parentTextFrames[0].parentPage;
        if (liveSnippetPage) {
            dynamicSnippetCode = dynamicSnippetCode.replace(/<Page_Num>/gi, tsFillZerosIfAllDigits (liveSnippetPage.name, 4));
        }
    }
    if (dynamicSnippetCode.search (/<Col_Head>/i) != -1) {
        if (startInsertionPoint.parent.constructor.name == "Cell") {
            var firstCellInColumn = startInsertionPoint.parent.parentColumn.cells.firstItem ();
            if (firstCellInColumn) {
                var targetPara = firstCellInColumn.paragraphs.length > 0 ? firstCellInColumn.paragraphs.item (0) : null;
                if (targetPara) {
                    var paraContents = getParaPlainContents (targetPara);
                    if (paraContents) {
                        dynamicSnippetCode = dynamicSnippetCode.replace(/<Col_Head>/gi, paraContents);
                    } 
                }
            } 
        }
    }
    if (dynamicSnippetCode.search (/<Col_Num>/i) != -1) {
        var parentCell = getFarParentCell (startInsertionPoint);
        if (parentCell) {
            var currentColumnNum = parentCell.parentColumn.index + 1;
            dynamicSnippetCode = dynamicSnippetCode.replace(/<Col_Num>/gi, tsFillZeros (currentColumnNum, 4)); 
        }
    }
    if (dynamicSnippetCode.search (/<Row_Head>/i) != -1) {
        if (startInsertionPoint.parent.constructor.name == "Cell") {
            var firstCellInRow = startInsertionPoint.parent.parentRow.cells.firstItem ();
            if (firstCellInRow) {
                var targetPara = firstCellInRow.paragraphs.length > 0 ? firstCellInRow.paragraphs.item (0) : null;
                if (targetPara) {
                    var paraContents = getParaPlainContents (targetPara);
                    if (paraContents) {
                        dynamicSnippetCode = dynamicSnippetCode.replace(/<Row_Head>/gi, paraContents);
                    } 
                }
            } 
        }
    }
    if (dynamicSnippetCode.search (/<Row_Num>/i) != -1) {
        var parentCell = getFarParentCell (startInsertionPoint);
        if (parentCell) {
            var currentRowNum = (parentCell.parentRow.index + 1) - parentCell.parentRow.parent.headerRowCount;
            dynamicSnippetCode = dynamicSnippetCode.replace(/<Row_Num>/gi, tsFillZeros (currentRowNum, 4)); 
        }
    }
    if (dynamicSnippetCode.indexOf ("/") == -1) {
        pathAndName[0] = "";
        pathAndName[1] = dynamicSnippetCode;
    }
    else {
        pathAndName[0] = dynamicSnippetCode.slice (0, dynamicSnippetCode.lastIndexOf ("/"));
        pathAndName[1] = dynamicSnippetCode.slice (dynamicSnippetCode.lastIndexOf ("/") + 1);
    }
    pathAndName[0] = solveShift (pathAndName[0]);
    pathAndName[1] = solveShift (pathAndName[1]);

    if (pathAndName[0] != "#" && (pathAndName[0].search (/<Increment>/i) != -1 || pathAndName[1].search (/<Increment>/i) != -1)) {
        var theIncrement = "0001";
        var docID = targetDocument.metadataPreferences.getProperty(treeShadeNamespace, "TS_ID");
        if (docID) {
            var liveSnippetPage = startInsertionPoint.parentTextFrames[0].parentPage;
            if (liveSnippetPage) {
                var currentIncrementState = docID + liveSnippetPage.id;
                var nestedMarker = null;
                var isToIncrement = false;
                if (startInsertionPoint.parent.characters.item(startInsertionPoint.index).textFrames.length > 0) {
                    nestedMarker = startInsertionPoint.parent.characters.item(startInsertionPoint.index).textFrames.item(0);
                    var oldIncrementState = nestedMarker.extractLabel("LS_INCREMENT_STATE");
                    if (currentIncrementState != oldIncrementState) {
                        isToIncrement = true;
                    }
                    else {
                        theIncrement = nestedMarker.extractLabel("LS_INCREMENT");
                    }
                }
                else {
                    isToIncrement = true;
                }
                if (isToIncrement) {
                    var incrementBase = "";
                    var incrementPrefix = "";
                    var incrementSuffix = "";
                    var fileRest = "";
                    if (pathAndName[0].search (/<Increment>/i) != -1) {
                        var incrementSplit = pathAndName[0].split (/<Increment>/i);
                        var suggestedIncrementBase = incrementSplit[0];
                        if (suggestedIncrementBase.lastIndexOf ("/") != -1) {
                            incrementBase = suggestedIncrementBase.slice (0, suggestedIncrementBase.lastIndexOf ("/"));
                            incrementPrefix = suggestedIncrementBase.slice (suggestedIncrementBase.lastIndexOf ("/") + 1);
                        }
                        if (incrementSplit[1].indexOf ("/") != -1)
                            incrementSuffix = incrementSplit[1].slice (0, incrementSplit[1].indexOf ("/"));
                        else
                            incrementSuffix = incrementSplit[1];
                    }
                    else {
                        incrementBase = pathAndName[0];
                        var incrementSplit = pathAndName[1].split (/<Increment>/i);
                        incrementPrefix = incrementSplit[0];
                        incrementSuffix = incrementSplit[1];
                        fileRest = "*";
                    }
                    var incrementFolder = getLiveSnippetFolder (targetDocument, incrementBase, null, null);
                    var lastIncrement = 0;
                    var incrementFoldersOrFiles = incrementFolder.getFiles (incrementPrefix + "????" + incrementSuffix + fileRest);
                    if (pathAndName[0].search (/<Increment>/i) != -1) {
                        for (var iff = incrementFoldersOrFiles.length - 1; iff >= 0; iff--) {
                            if (!isUnhiddenFolder (incrementFoldersOrFiles[iff]))
                            incrementFoldersOrFiles.splice (iff, 1);
                        }
                    }
                    else {
                        for (var iff = incrementFoldersOrFiles.length - 1; iff >= 0; iff--) {
                            if (!isUnhiddenFile (incrementFoldersOrFiles[iff]))
                            incrementFoldersOrFiles.splice (iff, 1);
                        }
                    }
                    if (incrementFoldersOrFiles.length > 0) {
                        var fn = File.decode (incrementFoldersOrFiles[incrementFoldersOrFiles.length - 1].name);
                        lastIncrement = fn.slice (incrementPrefix.length, incrementPrefix.length + 4);
                        lastIncrement = parseInt (lastIncrement, 10);
                        if (!lastIncrement) {
                            lastIncrement = 0;
                        }
                    }
                    if (pathAndName[0].search (/<Increment>/i) != -1) {
                        var newFolder = null;
                        do {
                            lastIncrement++;
                            newFolder = new Folder (incrementFolder.fsName.replace(/\\/g, "/") + "/" + incrementPrefix + tsFillZeros (lastIncrement, 4) + incrementSuffix);
                        } while (newFolder.exists);
                    }
                    else {
                        var newFiles = null;
                        do {
                            lastIncrement++;
                            newFiles = incrementFolder.getFiles (incrementPrefix + tsFillZeros (lastIncrement, 4) + incrementSuffix + "*");
                        } while (newFiles.length != 0);
                    }
                    theIncrement = tsFillZeros (lastIncrement, 4);
                    if (nestedMarker) {
                        nestedMarker.insertLabel("LS_INCREMENT_STATE", currentIncrementState);
                        nestedMarker.insertLabel("LS_INCREMENT", theIncrement);
                    }
                    else {
                        pathAndName.push (currentIncrementState);
                        pathAndName.push (theIncrement);
                    }
                }
            }
        }
        pathAndName[0] = pathAndName[0].replace(/<Increment>/gi, theIncrement);
        pathAndName[1] = pathAndName[1].replace(/<Increment>/gi, theIncrement);
    }
    return pathAndName;
}

function getParaPlainContents (para) {
    var paraContents = para.contents;
    return getPlainContents (paraContents);
}

function getPlainContents (paraContents) {
    if (paraContents) {
        if (paraContents[paraContents.length - 1] == "\r")
            paraContents = paraContents.slice (0, paraContents.length - 1);
        if (paraContents) {
            var cleanParaContents = "";
            var isToStart = false;
            for (var im = 0; im <paraContents.length; im++) {
                if (!isToStart) {
                    if (paraContents.charCodeAt(im) != 32 && paraContents.charCodeAt(im) != 65279) {
                        isToStart = true;
                    }
                }
                if (isToStart) {
                    if (paraContents.charCodeAt(im) == 10 || paraContents.charCodeAt(im) == 13) {
                        break;
                    }
                    else if (paraContents.charCodeAt(im) != 65279 && paraContents.charCodeAt(im) != 65532) { // && paraContents.charCodeAt(im) != 65532 - for anchored
                        cleanParaContents += paraContents[im];
                    }
                }
            }
            paraContents = cleanParaContents;
            return paraContents;
        }
    }
    return "";
}

function dtfCalculateRange (range, dtfFunction, table) {
    range = range.slice (1, -1);
    var cellsArray = new Array;
    var joinSeparator = "";
    if (range.indexOf (")(") == -1) {
        cellsArray.push (range);
    }
    else {
        cellsArray = range.split (")(");
        if (cellsArray.length > 2) {
            if (cellsArray.length > 3) {
                return null;
            }
            if (dtfFunction == "join") {
                joinSeparator = cellsArray[2];
                cellsArray.splice (2, 1);
            }
            else {
                cellsArray[2] = cellsArray[2].replace(/\s/g, "");
                if (cellsArray[2] == "sum" || cellsArray[2] == "subtract" || cellsArray[2] == "divide" || cellsArray[2] == "multiply" || cellsArray[2] == "average" || cellsArray[2] == "count" || cellsArray[2] == "join") {
                    dtfFunction = cellsArray[2];
                    cellsArray.splice (2, 1);
                }
            }
        }
    }
    for (var ca = 0; ca < cellsArray.length; ca++) {
        cellsArray[ca] = cellsArray[ca].replace(/\s/g, "");
        cellsArray[ca] = cellsArray[ca].split (",");
        if (cellsArray[ca].length != 2) return null;
        if (!cellsArray[ca][0] || !cellsArray[ca][1]) return null;
        for (var xy = 0; xy < cellsArray[ca].length; xy++) {
            if (cellsArray[ca][xy].indexOf ("-") != -1 || cellsArray[ca][xy].indexOf ("+") != -1) {
                var opeSplitted = null;
                var isPlus = true;
                if (cellsArray[ca][xy].indexOf ("+") != -1) {
                    opeSplitted = cellsArray[ca][xy].split ("+");
                }
                else {
                    isPlus = false;
                    opeSplitted = cellsArray[ca][xy].split ("-");
                }
                if (opeSplitted.length != 2) return null;
                var opeA = parseInt (opeSplitted[0], 10);
                var opeB = parseInt (opeSplitted[1], 10);
                if (opeA == null || opeB == null) return null;
                cellsArray[ca][xy] = isPlus? (opeA + opeB) : (opeA - opeB);
            }
            else {
                cellsArray[ca][xy] = parseInt (cellsArray[ca][xy], 10);
                if (!cellsArray[ca][xy]) return null;
            }
        }
    }
    var cellsRows = new Array;
    for (var cr = 0; cr < (cellsArray[cellsArray.length - 1][1] - cellsArray[0][1] + 1); cr++) {
        cellsRows.push (new Array);
        for (var cc = 0; cc < (cellsArray[cellsArray.length - 1][0] - cellsArray[0][0] + 1); cc++) {
            cellsRows[cr].push ([cellsArray[0][0] + cc, cellsArray[0][1] + cr]);
        }
    }
    if (cellsRows.length == 0)
        return null;
    if (cellsRows[0].length == 0)
        return null;
    for (var y = 0; y < cellsRows.length; y++) {
        for (var x = 0; x < cellsRows[y].length; x++) {
            cellsRows[y][x].push ("");
        }
    }
    if (cellsArray.length == 1) {
        if (cellsRows[0].length == 2) {
            cellsRows[0].splice (1, 1);
        }
    }
    for (var tci = 0; tci < table.cells.length; tci++) {
        var testedColIndex = table.cells[tci].parentColumn.index + 1;
        var testedRowIndex = table.cells[tci].parentRow.index + 1 - table.cells[tci].parentRow.parent.headerRowCount;
        if (testedColIndex - cellsRows[0][0][0] >= 0 && testedColIndex - cellsRows[0][0][0] < cellsRows[0].length) {
            if (testedRowIndex - cellsRows[0][0][1] >= 0 && testedRowIndex - cellsRows[0][0][1] < cellsRows.length) {
                var theValue = "";
                if (table.cells[tci].paragraphs.length > 0) {
                    theValue = getParaPlainContents (table.cells[tci].paragraphs.item (0));
                }
                cellsRows[testedRowIndex - cellsRows[0][0][1]][testedColIndex - cellsRows[0][0][0]][2] = theValue;
            }
        }
    }
    return calcCellsRows (cellsRows, dtfFunction, joinSeparator);
}

function solveDynamicTableFormula (dtfPhrase, table) {
    dtfPhrase = dtfPhrase.slice (1, -1);
    var dtfFunction = dtfPhrase.slice (0, dtfPhrase.indexOf ("("));
    dtfFunction = dtfFunction.replace(/\s/g, "");
    dtfFunction = dtfFunction.toLowerCase ();
    var ranges = dtfPhrase.slice (dtfPhrase.indexOf ("("));
    ranges = ranges.split ("&");
    var cellsRows = new Array;
    cellsRows.push (new Array);
    if (!ranges) return null;
    for (var rng = 0; rng < ranges.length; rng++) {
        if (ranges[rng].indexOf (",") != -1) {
            ranges[rng] = dtfCalculateRange (ranges[rng], dtfFunction, table);
        }
        if (ranges[rng] != null) {
            ranges[rng] = ranges[rng].replace(/\(/g, "");
            ranges[rng] = ranges[rng].replace(/\)/g, "");
            cellsRows[0].push (["", "", ranges[rng]]);
        }
    }
    return calcCellsRows (cellsRows, dtfFunction, "");
}

function calcCellsRows (cellsRows, dtfFunction, joinSeparator) {
    switch(dtfFunction) {
        case "sum":
        case "average":
        case "count":
            var sum = 0;
            var count = 0;
            var all = 0;
            for (var y = 0; y < cellsRows.length; y++) {
                for (var x = 0; x < cellsRows[y].length; x++) {
                    all++;
                    if (dtfFunction == "count") {
                        if (cellsRows[y][x][2]) count++;
                    }
                    else {
                        var toBeAdded = parseFloat (cellsRows[y][x][2]);
                        if (toBeAdded)
                            sum += toBeAdded;
                    }
                }
            }
            if (dtfFunction == "count") return count;
            if (dtfFunction == "average") sum = sum / all;
            sum = sum.toFixed(2);
            if (sum.slice (-3) == ".00") {
                sum = sum.slice (0, -3);
            }
            return sum;
        case "subtract":
        case "divide":
        case "multiply":
            var result = null;
            var side1 = parseFloat (cellsRows[0][0][2]);
            if (cellsRows[0].length == 1) {
                result = side1.toFixed (2);
                if (result.slice (-3) == ".00") {
                    result = result.slice (0, -3);
                }
                return result;
            }
            var side2 = parseFloat (cellsRows[cellsRows.length - 1][cellsRows[cellsRows.length - 1].length - 1][2]);
            if (dtfFunction == "subtract") result = side1 - side2;
            else if (dtfFunction == "divide") result = (side1 / side2);
            else result = side1 * side2;
            result = result.toFixed (2);
            if (result.slice (-3) == ".00") {
                result = result.slice (0, -3);
            }
            return result;
        case "join":
            var resultStr = "";
            var beforeToAdd = "";
            for (var x = 0; x < cellsRows.length; x++) {
                for (var y = 0; y < cellsRows[x].length; y++) {
                    if (cellsRows[x][y][2]) {
                        resultStr += beforeToAdd + cellsRows[x][y][2];
                        beforeToAdd = joinSeparator;
                    }
                }
            }
            return resultStr;
        default:
            return null;
          // code block
    }
}

function solveAbbreviation (text) {
    text = text.replace(/<FB>/gi, "<File_Branch>"); 
    text = text.replace(/<FN([^>]*)>/gi, "<File_Name$1>");
    text = text.replace(/<OFN([^>]*)>/gi, "<Original_File_Name$1>");
    text = text.replace(/<PFN([^>]*)>/gi, "<Previous_File_Name$1>");
    text = text.replace(/<CFB>/gi, "<Container_File_Branch>");
    text = text.replace(/<CFN>/gi, "<Container_File_Name>"); 
    text = text.replace(/<FL>/gi, "<File_Level>");
    text = text.replace(/<FD>/gi, "<File_Display>");
    text = text.replace(/<DN>/gi, "<Doc_Name>");
    text = text.replace(/<DO>/gi, "<Doc_Order>");
    text = text.replace(/<PN>/gi, "<Page_Num>");
    text = text.replace(/<ASP>/gi, "<Applied_Section_Prefix>");
    text = text.replace(/<ASM>/gi, "<Applied_Section_Marker>");
    text = text.replace(/<MP>/gi, "<Master_Prefix>");
    text = text.replace(/<MN>/gi, "<Master_Name>");
    text = text.replace(/<MFN>/gi, "<Master_Full_Name>");
    text = text.replace(/<CO>/gi, "<Char_Order>");
    text = text.replace(/<AO>/gi, "<Anchored_Order>");
    text = text.replace(/<PO>/gi, "<Para_Order>");
    text = text.replace(/<PF>/gi, "<Para_First>");
    text = text.replace(/<PB>/gi, "<Para_Before>");
    text = text.replace(/<PT>/gi, "<Para_This>");
    text = text.replace(/<PA>/gi, "<Para_After>");
    text = text.replace(/<PS>/gi, "<Para_Style>");
    text = text.replace(/<CS>/gi, "<Char_Style>");
    text = text.replace(/<CH>/gi, "<Col_Head>");
    text = text.replace(/<CN>/gi, "<Col_Num>");
    text = text.replace(/<RH>/gi, "<Row_Head>");
    text = text.replace(/<RN>/gi, "<Row_Num>");
    text = text.replace(/<NA>/gi, "<Name_At>");
    text = text.replace(/<RIE>/gi, "<Remove_If_Empty>");
    text = text.replace(/<S(\-?\d+>)/gi, "<Shift$1");
    return text;
}

function getFarParentCell (parentTest) {
    var parentCell = null;
    do {
        var isParentFrame = false;
        if ((parentTest.constructor.name == "InsertionPoint" || parentTest.constructor.name == "Character") && parentTest.parentTextFrames.length > 0) {
            if (parentTest.parentTextFrames[0].parent.constructor.name != "Spread" && parentTest.parentTextFrames[0].parent.constructor.name != "MasterSpread") {
                parentTest = parentTest.parentTextFrames[0].parent;
                isParentFrame = true;
            }
        }
        if (!isParentFrame) {
            parentTest = parentTest.parent;
        }
        if (parentTest && parentTest.constructor.name == "Cell") {
            parentCell = parentTest;
        }
    } while (parentTest && parentTest.constructor.name != "Document");
    return parentCell;
}

function getFarParentTextFrame (parentTest) {
    do {
        var isParentFrame = false;
        if ((parentTest.constructor.name == "InsertionPoint" || parentTest.constructor.name == "Character") && parentTest.parentTextFrames.length > 0) {
            if (parentTest.parentTextFrames[0].parent.constructor.name != "Spread" && parentTest.parentTextFrames[0].parent.constructor.name != "MasterSpread") {
                parentTest = parentTest.parentTextFrames[0].parent;
                isParentFrame = true;
            }
            else {
                return parentTest.parentTextFrames[0];
            }
        }
        if (!isParentFrame) {
            parentTest = parentTest.parent;
        }
    } while (parentTest && parentTest.constructor.name != "Document");
    return null;
}

function solveDynamicLiveSnippetPhrase (targetDocument, snippetFile, startInsertionPoint, liveSnippetID, dynamicLiveSnippetPath, parentID, oldParaStyleName, oldCharStyleName, isTargetFolders) { 
    var docPath = preparePath (targetDocument.filePath.fsName.replace(/\\/g, "/"));
    var rootPath = Folder (workshopPath).parent.fsName.replace(/\\/g, "/");
    docPath = docPath.replace (workshopPath, "");
    dynamicLiveSnippetPath = tsSolveDateTimeRandomPortion (dynamicLiveSnippetPath);
    dynamicLiveSnippetPath = solveAbbreviation (dynamicLiveSnippetPath);
    if (dynamicLiveSnippetPath.indexOf ("//") == 0) {
        dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace("//", rootPath + "/");
    }
    var prefixDotsIndex = dynamicLiveSnippetPath.search (/\/?\.+\$?\//);
    while (prefixDotsIndex != -1) {
        var targetParent = null;
        var deepCount = 2;
        var isRelative = false;
        var dotsFirstMatch = dynamicLiveSnippetPath.match (/\/?\.+\$?\//)[0];
        if (dotsFirstMatch[dotsFirstMatch.length - 2] == '$') {
            targetParent = Folder (workshopPath + docPath);
            deepCount = 3;
        }
        else {
            if (!snippetFile)
                break;
            targetParent = snippetFile.parent;
        }
        if (dotsFirstMatch[0] == "/") {
            isRelative = true;
            deepCount++;
        }
        for (var pml = 0; pml < dotsFirstMatch.length - deepCount; pml++) {
            if (!targetParent.parent) {
                break;
            }
            targetParent = targetParent.parent;
        }
        var parentPath = targetParent.fsName.replace(/\\/g, "/") + "/";
        if (isRelative) {
            parentPath = parentPath.replace (workshopPath, "");
        }
        dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace (dotsFirstMatch, parentPath);
        prefixDotsIndex = dynamicLiveSnippetPath.search (/\/?\.+\$?\//);
    }
    var nameDotsIndex = dynamicLiveSnippetPath.search (/<\.+\$?>/);
    while (nameDotsIndex != -1) {
        var deepCount = 3;
        var dotsFirstMatch = dynamicLiveSnippetPath.match (/<\.+\$?>/)[0];
        var targetParent = null;
        if (dotsFirstMatch[dotsFirstMatch.length - 2] == '$') {
            targetParent = Folder (workshopPath + docPath);
            deepCount = 4;
        }
        else {
            if (!snippetFile)
                break;
            targetParent = snippetFile.parent;
        }
        for (var pml = 0; pml < dotsFirstMatch.length - deepCount; pml++) {
            if (!targetParent.parent) {
                break;
            }
            targetParent = targetParent.parent;
        }
        var parentFolderName = "";
        if (targetParent && targetParent.name) {
            parentFolderName = File.decode (targetParent.name);
        }
        dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace (dotsFirstMatch, parentFolderName);
        nameDotsIndex = dynamicLiveSnippetPath.search (/<\.+\$?>/);
    }

    if (dynamicLiveSnippetPath.search (/<Col_Num>/i) != -1) {
        var parentCell = getFarParentCell (startInsertionPoint);
        if (parentCell) {
            var currentColumnNum = parentCell.parentColumn.index + 1;
            dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Col_Num>/gi, tsFillZeros (currentColumnNum, 4));
        }
    }
    if (dynamicLiveSnippetPath.search (/<Row_Num>/i) != -1) {
        var parentCell = getFarParentCell (startInsertionPoint);
        if (parentCell) {
            var currentRowNum = (parentCell.parentRow.index + 1) - parentCell.parentRow.parent.headerRowCount;
            dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace (/<Row_Num>/gi, tsFillZeros (currentRowNum, 4)); 
        }
    }
    var dynamicTableFormulaArray = null;
    do {
        dynamicTableFormulaArray = dynamicLiveSnippetPath.match (/<(Sum|Subtract|Divide|Multiply|Average|Count|Join)[^<>]+>/gi);
        if (dynamicTableFormulaArray) {
            var dtfResult = "";
            var dtfTargetCell = startInsertionPoint.parent;
            if (dtfTargetCell.constructor.name == "Cell") {
                for (var dtf = 0; dtf < dynamicTableFormulaArray.length; dtf++) {
                    var dtfResult = solveDynamicTableFormula (dynamicTableFormulaArray[dtf], dtfTargetCell.parentColumn.parent);
                    dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace (dynamicTableFormulaArray[dtf], dtfResult);
                }
            }
            else {
                for (var dtf = 0; dtf < dynamicTableFormulaArray.length; dtf++) {
                    dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace (dynamicTableFormulaArray[dtf], "");
                }
            }
        }
    }while (dynamicTableFormulaArray);

    if (dynamicLiveSnippetPath.search (/<File_Branch>/i) != -1) {
        var liveSnippetIDSplitted = new Array;
        liveSnippetIDSplitted.push (liveSnippetID.slice (0, liveSnippetID.lastIndexOf ("/")));
        liveSnippetIDSplitted.push (liveSnippetID.slice (liveSnippetID.lastIndexOf ("/") + 1));        
        dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<File_Branch>/gi, liveSnippetIDSplitted[0]);
    }
    if (dynamicLiveSnippetPath.search (/<File_Name[^>]*>/i) != -1) {
        if (snippetFile) {
            if (snippetFile instanceof File) {
                if (snippetFile.exists) {
                    var pureName = File.decode (snippetFile.name);
                    dynamicLiveSnippetPath = solveSnippetName (dynamicLiveSnippetPath, "File_Name", pureName);
                }
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<File_Display>/i) != -1) {
        if (snippetFile) {
            if (snippetFile instanceof File) {
                if (snippetFile.exists) {
                    var pureName = File.decode (snippetFile.name);
                    dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<File_Display>/gi, getPureDisplayName (pureName));
                }
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Original_File_Name[^>]*>/i) != -1) {
        if (snippetFile) {
            if (snippetFile instanceof File) {
                if (snippetFile.exists) {
                    var fileID = getTSID (snippetFile);
                    if (fileID) {
                        var originalID = readFile (File (dataPath + "/IDs" + fileID + "/Original ID"));
                        if (originalID) {
                            var fileID_Path_newID = [originalID];
                            getPath (fileID_Path_newID);
                            if (fileID_Path_newID[1]) {
                                var originalFileFullPath = workshopPath + fileID_Path_newID[1];
                                var originalFile = new File (originalFileFullPath);
                                if (originalFile.exists) {
                                    var pureName = File.decode (originalFile.name);
                                    dynamicLiveSnippetPath = solveSnippetName (dynamicLiveSnippetPath, "Original_File_Name", pureName);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Previous_File_Name[^>]*>/i) != -1) {
        if (snippetFile) {
            if (snippetFile instanceof File) {
                if (snippetFile.exists) {
                    var fileID = getTSID (snippetFile);
                    if (fileID) {
                        var previousID = readFile (File (dataPath + "/IDs" + fileID + "/Previous ID"));
                        if (previousID) {
                            var fileID_Path_newID = [previousID];
                            getPath (fileID_Path_newID);
                            if (fileID_Path_newID[1]) {
                                var previousFileFullPath = workshopPath + fileID_Path_newID[1];
                                var previousFile = new File (previousFileFullPath);
                                if (previousFile.exists) {
                                    var pureName = File.decode (previousFile.name);
                                    dynamicLiveSnippetPath = solveSnippetName (dynamicLiveSnippetPath, "Previous_File_Name", pureName);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Container_File_Branch>/i) != -1) {
        if (!parentID) {
            var lastCharIndex = new Array;
            var parentNestedMarkers = new Array;
            if (getContainerSnippet (startInsertionPoint, lastCharIndex, parentNestedMarkers, "base marker")) { //"first marker"  or "first snippet marker" or "base marker"
                parentID = parentNestedMarkers[0][0].textFrames.item(0).extractLabel("LS_ID");
            }
        }
        if (parentID) {
            dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Container_File_Branch>/gi, parentID.slice (0, parentID.lastIndexOf ('/')));
        }
    }
    if (dynamicLiveSnippetPath.search (/<Container_File_Name>/i) != -1) {
        if (!parentID) {
            var lastCharIndex = new Array;
            var parentNestedMarkers = new Array;
            if (getContainerSnippet (startInsertionPoint, lastCharIndex, parentNestedMarkers, "base marker")) { //"first marker"  or "first snippet marker" or "base marker"
                parentID = parentNestedMarkers[0][0].textFrames.item(0).extractLabel("LS_ID");
            }
        }
        if (parentID) {
            dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Container_File_Name>/gi, parentID.slice (parentID.lastIndexOf ('/') + 1));
        }
    }
    if (dynamicLiveSnippetPath.search (/<File_Level>/i) != -1) {
        var fileLevel = 0;
        if (liveSnippetID.lastIndexOf ("/") > 0) {
            fileLevel = liveSnippetID.split ("/").length - 1;
        }
        dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<File_Level>/gi, fileLevel);
    }
    if (dynamicLiveSnippetPath.search (/<Doc_Name>/i) != -1) {
        var pureName = getPureName (targetDocument.name, false);
        dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Doc_Name>/gi, pureName);
    }
    if (dynamicLiveSnippetPath.search (/<Doc_Order>/i) != -1) {
        var docFileOrder = 1;
        var docParentFolder = new Folder (workshopPath + docPath);
        var allDocs = docParentFolder.getFiles ("*.indd");
        sortFilesList (docParentFolder, allDocs, false);
        for (var ad = 0; ad < allDocs.length; ad++) {
            if (File.decode (allDocs[ad].name) == File.decode (targetDocument.name)) {
                docFileOrder = ad + 1;
                break;
            }
        }
        dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Doc_Order>/gi, tsFillZerosIfAllDigits (docFileOrder.toString (), 4));
    }
    if (dynamicLiveSnippetPath.search (/<Page_Num>/i) != -1) {
        if (startInsertionPoint.parentTextFrames.length) {
            var liveSnippetPage = startInsertionPoint.parentTextFrames[0].parentPage;
            if (liveSnippetPage) {
                dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Page_Num>/gi, tsFillZerosIfAllDigits (liveSnippetPage.name, 4));
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Applied_Section_Prefix>/i) != -1) {
        if (startInsertionPoint.parentTextFrames.length) {
            var liveSnippetPage = startInsertionPoint.parentTextFrames[0].parentPage;
            if (liveSnippetPage) {
                if (liveSnippetPage.appliedSection) {
                    dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Applied_Section_Prefix>/gi, liveSnippetPage.appliedSection.sectionPrefix);
                }
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Applied_Section_Marker>/i) != -1) {
        if (startInsertionPoint.parentTextFrames.length) {
            var liveSnippetPage = startInsertionPoint.parentTextFrames[0].parentPage;
            if (liveSnippetPage) {
                if (liveSnippetPage.appliedSection) {
                    dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Applied_Section_Marker>/gi, liveSnippetPage.appliedSection.marker);
                }
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Master_Prefix>/i) != -1) {
        if (startInsertionPoint.parentTextFrames.length) {
            var liveSnippetPage = startInsertionPoint.parentTextFrames[0].parentPage;
            if (liveSnippetPage) {
                dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Master_Prefix>/gi, liveSnippetPage.appliedMaster.namePrefix);
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Master_Name>/i) != -1) {
        if (startInsertionPoint.parentTextFrames.length) {
            var liveSnippetPage = startInsertionPoint.parentTextFrames[0].parentPage;
            if (liveSnippetPage) {
                dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Master_Name>/gi, liveSnippetPage.appliedMaster.baseName);
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Master_Full_Name>/i) != -1) {
        if (startInsertionPoint.parentTextFrames.length) {
            var liveSnippetPage = startInsertionPoint.parentTextFrames[0].parentPage;
            if (liveSnippetPage) {
                dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Master_Full_Name>/gi, liveSnippetPage.appliedMaster.name);
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Char_Order>/i) != -1) {
        var anchorChar  = getAnchoredChar (startInsertionPoint);
        if (anchorChar) {
            var thisPara = anchorChar.paragraphs.item (0);
            var charOrder = thisPara.characters.itemByRange(thisPara.characters.firstItem(), anchorChar).characters.length - 1;
            if (charOrder) {
                dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Char_Order>/gi, tsFillZeros(charOrder, 4));
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Anchored_Order>/i) != -1) {
        var anchorChar  = getAnchoredChar (startInsertionPoint);
        if (startInsertionPoint.parentTextFrames.length) {
            anchorChar = downLevel (startInsertionPoint.parentTextFrames[0]);
        }
        if (anchorChar) {
            var thisPara = anchorChar.paragraphs.item (0);
            var anchoredChars = thisPara.characters.itemByRange(thisPara.characters.firstItem(), anchorChar).contents.toString ();
            var charOrder = -1;
            for (var i = 0; i < anchoredChars.length; i++) {
                if (anchoredChars.charCodeAt(i) == 65532) {
                    charOrder++;
                }
            }
            if (charOrder) {
                dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Anchored_Order>/gi, tsFillZeros(charOrder, 4));
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Para_Order>/i) != -1) {
        var downChar  = null;
        if (startInsertionPoint.parentTextFrames.length) {
            downChar = downLevel (startInsertionPoint.parentTextFrames[0]);
        }
        var firstPara = null;
        var basePara = null;
        var mainStory = null;
        if (downChar) {
            mainStory = downChar.parentStory;
            firstPara = mainStory.paragraphs.firstItem ();
            basePara = downChar.paragraphs.item (0);
        }
        else {
            mainStory = startInsertionPoint.parentStory;
            firstPara = mainStory.paragraphs.firstItem ();
            if (startInsertionPoint.parent.constructor.name == "Cell") {
                basePara = startInsertionPoint.parent.parentRow.parent.storyOffset.paragraphs.item (0);
            }
            else if (startInsertionPoint.paragraphs.length) {
                basePara = startInsertionPoint.paragraphs.item (0);
            }
        }
        if (firstPara && basePara) {
            var paraOrder = mainStory.paragraphs.itemByRange (firstPara, basePara).paragraphs.length;
            dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Para_Order>/gi, tsFillZeros (paraOrder, 4));
        }
    }
    if (dynamicLiveSnippetPath.search (/<Para_First>/i) != -1) {
        var downChar  = null;
        if (startInsertionPoint.parentTextFrames.length) {
            downChar = downLevel (startInsertionPoint.parentTextFrames[0]);
        }
        var firstPara = null;
        if (downChar) {
            firstPara = downChar.parentStory.paragraphs.firstItem ();
        }
        else {
            firstPara = startInsertionPoint.parentStory.paragraphs.firstItem ();
        }
        if (firstPara) {
            var paraContents = getParaPlainContents (firstPara);
            if (paraContents) {
                dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Para_First>/gi, paraContents);
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Para_Before>/i) != -1) {
        if (startInsertionPoint.paragraphs.length) {
            var beforePara = startInsertionPoint.parent.paragraphs.previousItem (startInsertionPoint.paragraphs.item (0));
            if (beforePara) {
                var paraContents = getParaPlainContents (beforePara);
                if (paraContents) {
                    dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Para_Before>/gi, paraContents);
                }
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Para_This>/i) != -1) {
        if (startInsertionPoint.paragraphs.length) {
            var thisPara = startInsertionPoint.paragraphs.item (0);
            if (thisPara) {
                var paraContents = getParaPlainContents (thisPara);
                if (paraContents) {
                    dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Para_This>/gi, paraContents);
                }
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Para_After>/i) != -1) {
        if (startInsertionPoint.paragraphs.length) {
            var beforePara = startInsertionPoint.parent.paragraphs.nextItem (startInsertionPoint.paragraphs.item (0));
            if (beforePara) {
                var paraContents = getParaPlainContents (beforePara);
                if (paraContents) {
                    dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Para_After>/gi, paraContents);
                }
            }
        }
    }
    if (dynamicLiveSnippetPath.search (/<Para_Style>/i) != -1) {
        if (oldParaStyleName == null) {
            oldParaStyleName = startInsertionPoint.appliedParagraphStyle.name;
        }
        dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Para_Style>/gi, oldParaStyleName);
    }

    if (dynamicLiveSnippetPath.search (/<Char_Style>/i) != -1) {
        if (oldCharStyleName == null) {
            oldCharStyleName = startInsertionPoint.appliedCharacterStyle.name;
        }
        dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Char_Style>/gi, oldCharStyleName);
    }

    //<Col_Head>, <Row_Head>
    if (dynamicLiveSnippetPath.search (/<Col_Head>/i) != -1) {
        if (startInsertionPoint.parent.constructor.name == "Cell") {
            var firstCellInColumn = startInsertionPoint.parent.parentColumn.cells.firstItem ();
            if (firstCellInColumn) {
                var targetPara = firstCellInColumn.paragraphs.length > 0 ? firstCellInColumn.paragraphs.item (0) : null;
                if (targetPara) {
                    var paraContents = getParaPlainContents (targetPara);
                    if (paraContents) {
                        dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Col_Head>/gi, paraContents);
                    } 
                }
            } 
        }
    }
    if (dynamicLiveSnippetPath.search (/<Row_Head>/i) != -1) {
        if (startInsertionPoint.parent.constructor.name == "Cell") {
            var firstCellInColumn = startInsertionPoint.parent.parentColumn.cells.firstItem ();
            if (firstCellInColumn) {
                var targetPara = firstCellInColumn.paragraphs.length > 0 ? firstCellInColumn.paragraphs.item (0) : null;
                if (targetPara) {
                    var paraContents = getParaPlainContents (targetPara);
                    if (paraContents) {
                        var paraContents = getParaPlainContents (targetPara);
                        if (paraContents) {
                            dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace(/<Row_Head>/gi, paraContents);
                        }
                    } 
                }
            } 
        }
    }
    dynamicLiveSnippetPath = solveShift (dynamicLiveSnippetPath);
    dynamicLiveSnippetPath = solveNameAt (dynamicLiveSnippetPath);
    if (dynamicLiveSnippetPath.indexOf ("/") == -1) {
        return dynamicLiveSnippetPath;
    }
    if (dynamicLiveSnippetPath.indexOf (rootPath) == 0) {
        dynamicLiveSnippetPath = dynamicLiveSnippetPath.replace (rootPath, "");
    }
    var returnedArray = new Array;
    searchForDynamicLiveSnippets (Folder (rootPath), dynamicLiveSnippetPath, returnedArray, isTargetFolders);
    return returnedArray;
}

function solveSnippetName (thePhrase, thePart, targetName) {
    var findRegex = new RegExp("<" + thePart + "[^>]*>", "i");
    var regexFlagi = new RegExp("<" + thePart + "[^>]*>", "i");
    var fnDotIndex = thePhrase.search (findRegex);
    if (fnDotIndex != -1) {
        targetName = getPureName (targetName, false);
        while (fnDotIndex != -1) {
            var fnDotFirstMatch = thePhrase.match (findRegex)[0];
            var newName = targetName;
            if (newName != "") {
                var sepAndNum = fnDotFirstMatch.slice (thePart.length + 1, -1);
                var sepAndNumSplitted = sepAndNum.split ("_");
                var theSep = null;
                if (sepAndNumSplitted.length > 1) {
                    theSep = sepAndNumSplitted[1];
                }
                var theNumOrNewSep = null;
                if (sepAndNumSplitted.length > 2) {
                    theNumOrNewSep = sepAndNumSplitted[2];
                }
                if (theSep != null) {
                    var nameParts = newName.split (theSep);
                    if (theNumOrNewSep) {
                        var numPhrase = theNumOrNewSep.match (/-?\d+/);
                        if (numPhrase != null && numPhrase[0].length == theNumOrNewSep.length) {
                            numPhrase = parseInt (theNumOrNewSep, 10);
                        }
                        else {
                            numPhrase = null;
                        }
                        if (numPhrase != null) {
                            if (numPhrase == 0) {
                                numPhrase = 1;
                            }
                            if (numPhrase < 0) {
                                numPhrase = nameParts.length - Math.abs (numPhrase) + 1;
                            }
                            if (numPhrase > nameParts.length) {
                                numPhrase = nameParts.length;
                            }
                            newName = nameParts[numPhrase - 1];
                        }
                        else {
                            newName = nameParts.join (theNumOrNewSep);
                        }
                    }
                    else {
                        newName = nameParts.join ("");
                    }
                }
            }
            thePhrase = thePhrase.replace (regexFlagi, newName); 
            fnDotIndex = thePhrase.search (findRegex);
        }
    }
    return thePhrase;
}

function sortFilesList (targetFolder, filesList, isTargetFolders) {
    var tsSortFile = new File (targetFolder.fsName.replace(/\\/g, "/") + "/ts_sort.txt");
    if (tsSortFile.exists) {
        var sortFileContent = readFile (tsSortFile);
        if (sortFileContent) {
            var sortList = '';
            if (sortFileContent.indexOf ('\r\n') != -1) {
                sortList = sortFileContent.split ('\r\n');
            }
            else {
                sortList = sortFileContent.split ('\n');
            }
            if (sortList) {
                var sortIndex = 0;
                for (var sl = 0; sl < sortList.length; sl++) {
                    var pureFileName = sortList[sl];
                    for (var fl = sortIndex; fl < filesList.length; fl++) {
                        var testedName = File.decode (filesList[fl].name);
                        if (filesList[fl] instanceof File) {
                            if (isTargetFolders) {
                                continue;
                            }
                            testedName = getPureName (testedName, false);
                        }
                        else {
                            if (isTargetFolders == false) {
                                continue;
                            }
                        }
                        if (testedName == pureFileName) {
                            if (fl != sortIndex) {
                                var swapCell = filesList[fl];
                                filesList[fl] = filesList[sortIndex];
                                filesList[sortIndex] = swapCell;
                            }
                            sortIndex++;
                            break;
                        }
                    }
                }
                return true;
            }
        }
    }
    return false;
}

function searchForDynamicLiveSnippets (targetFolder, dynamicLiveSnippetPath, returnedArray, isTargetFolders) {
    if (dynamicLiveSnippetPath.indexOf ("/") == 0) {
        dynamicLiveSnippetPath = dynamicLiveSnippetPath.slice (1);
    }
    if (dynamicLiveSnippetPath.indexOf ("/") != -1) {
        var currentPortion = dynamicLiveSnippetPath.slice (0, dynamicLiveSnippetPath.indexOf ("/"));
        var allFolders = targetFolder.getFiles (currentPortion);
        sortFilesList (targetFolder, allFolders, isTargetFolders);
        for (var af = 0; af < allFolders.length; af++) {
            if (isCSFolder (allFolders[af])) {
                searchForDynamicLiveSnippets (allFolders[af], dynamicLiveSnippetPath.slice (dynamicLiveSnippetPath.indexOf ("/")), returnedArray, isTargetFolders);
            }
        }
    }
    else {
        var allFiles = targetFolder.getFiles (dynamicLiveSnippetPath);
        sortFilesList (targetFolder, allFiles, isTargetFolders);
        for (var al = 0; al < allFiles.length; al++) {
            if (isTargetFolders) {
                if (isCSFolder (allFiles[al])) {
                    returnedArray.push (allFiles[al]);
                }
            } else if (isTargetFolders == false) {
                if (isUnhiddenFile (allFiles[al]) && File.decode (allFiles[al].name) != "ts_sort.txt") {
                    returnedArray.push (allFiles[al]);
                }
            }
            else {
                if ((isCSFolder (allFiles[al]) || isUnhiddenFile (allFiles[al])) && File.decode (allFiles[al].name) != "ts_sort.txt") {
                    returnedArray.push (allFiles[al]);
                }
            }
        }
    }
}

function removeIndex_snippetFilePath_ID_List (liveSnippetTSID) {
    for (var fid = snippetFilePath_ID_List.length - 1; fid >= 0; fid--) {
        if (liveSnippetTSID == snippetFilePath_ID_List[fid][1]) {
            snippetFilePath_ID_List.splice (fid, 1);
            return true;
        }
    }
    return false;
}

function getIndex_snippetFilePath_ID_List (snippetFile) {
    var fidIndex = -1;
    for (var fid = snippetFilePath_ID_List.length - 1; fid >= 0; fid--) {
        if (snippetFile.fsName.replace(/\\/g, "/") == snippetFilePath_ID_List[fid][0]) {
            fidIndex = fid;
            break;
        }
    }
    if (fidIndex == -1) {
        var liveSnippetTSID = getTSID (snippetFile);
        if (!liveSnippetTSID) {
            liveSnippetTSID = "";
            addToBridgeScanningQueue (snippetFile.parent.fsName.replace(/\\/g, "/"));
        }
        snippetFilePath_ID_List.push ([snippetFile.fsName.replace(/\\/g, "/"), liveSnippetTSID]);
    }
    return fidIndex;
}

function tsSolveDateTimeRandomPortion (dateTimeRandom) {
    //DS <Day_Seconds>
    //DM <Day_Minutes>
    //DH <Day_Hours>
    //WO <Weekday_Order>
    //WN <Weekday_Name>
    //WA <Weekday_Abbreviation>
    //MD <Month_Date>
    //MO <Month_Order>
    //MN <Month_Name>
    //MA <Month_Abbreviation>
    //YF <Year_Full>
    //YA <Year_Abbreviation>
    //TN <Time_Numbers>
    //TH <Time_Hexadecimal>
    //TM <Time_Minimum>
    //R <Random_Style_Long>
    //Use in Style 'a' small letters, 'A' capital letters, 'h' small hexadecimal, 'H' large hexadecimal, '1' numbers

    dateTimeRandom = dateTimeRandom.replace(/<Day_Seconds>/gi,           "<DS>");
    dateTimeRandom = dateTimeRandom.replace(/<Day_Minutes>/gi,           "<DM>");
    dateTimeRandom = dateTimeRandom.replace(/<Day_Hours>/gi,             "<DH>");
    dateTimeRandom = dateTimeRandom.replace(/<Weekday_Order>/gi,         "<WO>");
    dateTimeRandom = dateTimeRandom.replace(/<Weekday_Name>/gi,          "<WN>");
    dateTimeRandom = dateTimeRandom.replace(/<Weekday_Abbreviation>/gi,  "<WA>");
    dateTimeRandom = dateTimeRandom.replace(/<Month_Date>/gi,            "<MD>");
    dateTimeRandom = dateTimeRandom.replace(/<Month_Order>/gi,           "<MO>");
    dateTimeRandom = dateTimeRandom.replace(/<Month_Name>/gi,            "<MN>");
    dateTimeRandom = dateTimeRandom.replace(/<Month_Abbreviation>/gi,    "<MA>");
    dateTimeRandom = dateTimeRandom.replace(/<Year_Full>/gi,             "<YF>");
    dateTimeRandom = dateTimeRandom.replace(/<Year_Abbreviation>/gi,     "<YA>");
    dateTimeRandom = dateTimeRandom.replace(/<Time_Numbers>/gi,          "<TN>");
    dateTimeRandom = dateTimeRandom.replace(/<Time_Hexadecimal>/gi,      "<TH>");
    dateTimeRandom = dateTimeRandom.replace(/<Time_Minimum>/gi,          "<TM>");
    dateTimeRandom = dateTimeRandom.replace(/<Random_([^<>]+>)/gi,       "<R_$1");
    dateTimeRandom = dateTimeRandom.replace(/<R>/gi,                     "<R_A1_8>");

    var matches = null;
    var theDate = new Date();
    do {
        matches = dateTimeRandom.match (/<(DS|DM|DH|WO|WN|WA|MD|MO|MN|MA|YF|YA|TN|TH|TM|R_)[^<>]*>/gi);
        if (matches) {
            var firstMatch = matches[0];
            var firstThree = firstMatch.slice (0, 3).toUpperCase();
            switch(firstThree) {
                case "<DS": case "<ds":
                    var daySeconds = tsFillZeros (theDate.getSeconds(), 2);
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, daySeconds);
                    break;
                case "<DM": case "<dm":
                    var dayMinutes = tsFillZeros (theDate.getMinutes(), 2);
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, dayMinutes);
                    break;
                case "<DH": case "<dh":
                    var dayHours = tsFillZeros (theDate.getHours(), 2);
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, dayHours);
                    break;
                case "<WO": case "<wo":
                    var weekdayOrder = (theDate.getDay() + 1).toString ();
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, weekdayOrder);
                    break;
                case "<WN": case "<wn":
                    var weekdayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
                    var weekdayName = weekdayNames[theDate.getDay()];
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, weekdayName);
                    break;
                case "<WA": case "<wa":
                    var weekdayAbbs = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
                    var weekdayAbb = weekdayAbbs[theDate.getDay()];
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, weekdayAbb);
                    break;
                case "<MD": case "<md":
                    var monthDate = tsFillZeros (theDate.getDate(), 2);
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, monthDate);
                    break;
                case "<MO": case "<mo":
                    var monthOrder = tsFillZeros (theDate.getMonth() + 1, 2);
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, monthOrder);
                    break;
                case "<MN": case "<mn":
                    var monthsNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
                    var monthName = monthsNames[theDate.getMonth()];
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, monthName);
                    break;
                case "<MA": case "<ma":
                    var monthNameAbbs = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                    var monthNameAbb = monthNameAbbs[theDate.getMonth()];
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, monthNameAbb);
                    break;
                case "<YF": case "<yf":
                    var yearFull = theDate.getFullYear().toString ();
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, yearFull);
                    break;
                case "<YA": case "<ya":
                    var yearAbb = theDate.getFullYear().toString ().slice (2);
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, yearAbb);
                    break;
                case "<TN": case "<tn":
                    var timeNumbers = theDate.getTime ().toString ();
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, timeNumbers);
                    break;
                case "<TH": case "<th":
                    var timeHexa = theDate.getTime ().toString (16);
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, timeHexa);
                    break;
                case "<TM": case "<tm":
                    var timeShort = theDate.getTime ().toString (36);
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, timeShort);
                    break;
                case "<R_": case "<r_":
                    var randomPhrases = firstMatch.slice (3, -1).split ("_");
                    var capital = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                    var small = "abcdefghijklmnopqrstuvwxyz";
                    var numbers = "0123456789";
                    var capitalHexa = "0123456789ABCDEF";
                    var smallHexa = "0123456789abcdef";
                    var combination = "";
                    var theLong = 4;
                    if (randomPhrases.length > 0) {
                        if (randomPhrases[0].indexOf ('A') != -1) {
                            combination += capital;
                        }
                        if (randomPhrases[0].indexOf ('a') != -1) {
                            combination += small;
                        }
                        if (randomPhrases[0].indexOf ('1') != -1) {
                            combination += numbers;
                        }
                        if (randomPhrases[0].indexOf ('H') != -1) {
                            combination += capitalHexa;
                        }
                        if (randomPhrases[0].indexOf ('h') != -1) {
                            combination += smallHexa;
                        }
                    }
                    if (randomPhrases.length > 1) {
                        var longGet = parseInt (randomPhrases[1], 10);
                        if (longGet && longGet > 0) {
                            theLong = longGet;
                        }
                    }
                    var randomCode = tsMakeRandom (theLong, combination);
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, randomCode);
                    break;
                default:
                    dateTimeRandom = dateTimeRandom.replace (firstMatch, "");
            }
        }
    } while (matches);
    return dateTimeRandom;
}

function updateInstance (targetDocument, snippetFile, nestedMarkers, isForcibly, dynamicLiveSnippetPhrase, parentID, liveSnippetContent, isToRemoveMarks, snippetsCacheList, isPreviousHorizontal, pageItemsIDs, cellsDeep, deepIndex, cellsIndex, importedReplace) {
    deepIndex++;
    cellsDeep.splice (deepIndex, 0, null);
    var c;
    var level;
    var isColumnsSeparator = false;
    var liveSnippetID = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_ID");
    var liveSnippetIDSplitted = null;
    var pathAndName = null;
    var isToUpdateNested = false;
    var dynamicSnippetCode = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_PREFIX_NUMBER");
    var isToChange = false;
    var inlineDynamicMark = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_DYNAMIC_INLINE");
    var dynamicStylePhrases = new Array;
    var expandValues = [1, 1, 0, 0]; //[rowsSpan, columnsSpan, rowAdded, columnAdded]
    var cellsDirection = '|';
    if (!dynamicLiveSnippetPhrase)
        dynamicLiveSnippetPhrase = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_DYNAMIC");
    if (dynamicSnippetCode) {
        pathAndName = solveSnippetCode (targetDocument, nestedMarkers[0][0].parent.insertionPoints.item (nestedMarkers[0][0].index), dynamicSnippetCode, parentID);
        if (pathAndName[1] == "") {
            pathAndName[1] = "0001";
        }
        if (pathAndName[0] || pathAndName[1]) {
            liveSnippetIDSplitted = new Array;
            liveSnippetIDSplitted.push (liveSnippetID.slice (0, liveSnippetID.lastIndexOf ("/")));
            liveSnippetIDSplitted.push (liveSnippetID.slice (liveSnippetID.lastIndexOf ("/") + 1));
            if (pathAndName[0]) {
                if (liveSnippetIDSplitted[0] != pathAndName[0]) {
                    isToChange = true;
                }
            }
            if (pathAndName[1]) {
                if (liveSnippetIDSplitted[1] != pathAndName[1]) {
                    isToChange = true;
                }
            }
            if (isToChange) {
                snippetFile = null;
                changeMarkers (nestedMarkers, pathAndName[0], pathAndName[1]);
                nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_TSID", "");
            }
        }
    }
    else {
        pathAndName = new Array;
        pathAndName.push (liveSnippetID.slice (0, liveSnippetID.lastIndexOf ("/")));
        pathAndName.push (liveSnippetID.slice (liveSnippetID.lastIndexOf ("/") + 1));
    }
    if (!snippetFile && pathAndName[0] != "#")
        snippetFile = findTargetFile (targetDocument, nestedMarkers[0][0]);
    if (!(snippetFile instanceof File) && pathAndName[0] != "#") {
        if (dynamicLiveSnippetPhrase && inlineDynamicMark == "") {
            //create file
            liveSnippetID = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_ID");
            var newLiveSnippetIDSplitted = new Array;
            newLiveSnippetIDSplitted.push (liveSnippetID.slice (0, liveSnippetID.lastIndexOf ("/")));
            newLiveSnippetIDSplitted.push (liveSnippetID.slice (liveSnippetID.lastIndexOf ("/") + 1));
            var liveSnippetsFolder = getLiveSnippetFolder (targetDocument, newLiveSnippetIDSplitted[0], null, null);
            snippetFile = createLiveSnippetFile (targetDocument, liveSnippetsFolder, nestedMarkers, dynamicLiveSnippetPhrase);
            if (snippetFile instanceof File) {
                if (!snippetFile.exists) {
                    snippetFile = "Couldn't Create Dynamic Live Snippet File!";
                }
            }
            else {
                snippetFile = "Couldn't Create Dynamic Live Snippet File!";
            }    
        }

        if (!(snippetFile instanceof File)) {
            if (dynamicSnippetCode) {
                var oldFile = null;
                if (liveSnippetIDSplitted) {
                    var liveSnippetFolder = getLiveSnippetFolder (targetDocument, liveSnippetIDSplitted[0], null, null);
                    if (liveSnippetFolder.exists) {
                        oldFile = liveSnippetFolder.getFiles (liveSnippetIDSplitted[1] + "*");
                        if (oldFile.length > 0) {
                            var lastCreatedIndex = 0;
                            for (var sf = 1; sf < oldFile.length; sf++) {
                                if (oldFile[sf].created.getTime () > oldFile[lastCreatedIndex].created.getTime ()) {
                                    lastCreatedIndex = sf;
                                }
                            }
                            oldFile = oldFile[lastCreatedIndex];
                        }
                        else {
                            oldFile = null;	
                        }
                    }
                }
                if (oldFile) {
                    if (dynamicLiveSnippetPhrase) {
                        var liveSnippetFolder = getLiveSnippetFolder (targetDocument, pathAndName[0], null, null);
                        snippetFile = new File (liveSnippetFolder.fsName.replace(/\\/g, "/") + "/" + pathAndName[1] + oldFile.fsName.replace(/\\/g, "/").slice (oldFile.fsName.replace(/\\/g, "/").lastIndexOf (".")));
                        oldFile.copy (snippetFile);
                    }
                    else {
                        snippetFile = submitInstance (targetDocument, nestedMarkers);
                    }
                }
                else {
                    snippetFile = submitInstance (targetDocument, nestedMarkers);
                    if (!snippetFile) {
                        nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_STATE", "missing");
                        changeStateAppearance(nestedMarkers[0][0].textFrames.item(0)); //Changing the appearance
                        return snippetFile;
                    }
                }
            }
            else {
                var storedTSID = nestedMarkers[0][0].textFrames.item(0).extractLabel ("LS_TSID");
                var tracedPathAndName = "";
                if (storedTSID) {
                    var fileID_Path_newID = [storedTSID];
                    getPath (fileID_Path_newID);
                    if (fileID_Path_newID[1] && File (workshopPath + fileID_Path_newID[1]).exists) {
                        snippetFile = new File (workshopPath + fileID_Path_newID[1]);
                        if (isInsideProject (targetDocument, snippetFile)) {
                            tracedPathAndName = getPathAndName (targetDocument, snippetFile);
                            pathAndName[0] = tracedPathAndName.slice (0, tracedPathAndName.lastIndexOf ("/"));
                            pathAndName[1] = tracedPathAndName.slice (tracedPathAndName.lastIndexOf ("/") + 1);
                            isToChange = true;
                            changeMarkers (nestedMarkers, pathAndName[0], pathAndName[1]);
                        }
                    }
                }
                if (tracedPathAndName == "") {
                    nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_STATE", "missing");
                    changeStateAppearance(nestedMarkers[0][0].textFrames.item(0)); //Changing the appearance
                    return snippetFile;
                }
            }
        }
    }
    var isModified = false;
    //check dynamic live snippet
    if (!isForcibly && pathAndName[0] != "#")
        isModified = (parseInt(nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_DATE"), 10) < snippetFile.modified.getTime ());
    var oldLiveSnippetID = nestedMarkers[0][0].textFrames.item(0).extractLabel ("LS_TSID");
    var isIDMS = false;
    if (pathAndName[0] != "#") {
        if (isForcibly || isToChange || (!oldLiveSnippetID) || isModified) {
            var fidIndex = getIndex_snippetFilePath_ID_List (snippetFile); 
            if (fidIndex == -1) {
                fidIndex = snippetFilePath_ID_List.length - 1;
            }
            if (oldLiveSnippetID != snippetFilePath_ID_List[fidIndex][1]) {
                nestedMarkers[0][0].textFrames.item(0).insertLabel ("LS_TSID", snippetFilePath_ID_List[fidIndex][1]);
            }
        }
        //isToUpdateNested
        isIDMS = (File.decode (snippetFile.name).slice(-5) == ".IDMS");        
    }
    if (isIDMS) {
        isToUpdateNested = true;
    }
    if (isForcibly || isModified || isToChange) {
        var instanceStory = nestedMarkers[0][0].parent;
        var startIndex = nestedMarkers[0][0].index; //Saving the index because myFirstChar object will be destoyed by pasting the text
        var lastCharIndex = new Array;
        lastCharIndex.push(0);
        var oldParaStyleName = nestedMarkers[0][0].appliedParagraphStyle.name;
        var oldCharStyleName = nestedMarkers[0][0].appliedCharacterStyle.name;
        var tableCellsStyles = new Array; //[stylePhrase, cellStyleFile, insertionPoint, cellStyleID]
        var testDynamicStylePhrase = nestedMarkers[0][0].textFrames.item(0).extractLabel("TS_LS_STYLE");
        var replaceDynamicPhrase = nestedMarkers[0][0].textFrames.item(0).extractLabel("TS_LS_REPLACE");
        var cellsList = new Array;
        var isHorizontal = false;
        var isStock = false;
        var isCellsNested = false;
        var isItNotNested = false;
        if (replaceDynamicPhrase) {
            replaceDynamicPhrase = replaceDynamicPhrase.split (":");
            for (var rdp = 0; rdp < replaceDynamicPhrase.length; rdp++) {
                replaceDynamicPhrase[rdp] = solveDynamicLiveSnippetPhrase (targetDocument, snippetFile, nestedMarkers[0][0].parent.insertionPoints.item (nestedMarkers[0][0].index), liveSnippetID, replaceDynamicPhrase[rdp], parentID, oldParaStyleName, oldCharStyleName, false);
                replaceDynamicPhrase[rdp] = replaceDynamicPhrase[rdp].replace (/Colon/gi, ":");
                replaceDynamicPhrase[rdp] = replaceDynamicPhrase[rdp].replace(/\\r/gi, "\r");
                replaceDynamicPhrase[rdp] = replaceDynamicPhrase[rdp].replace(/\\t/gi, "\t");
                replaceDynamicPhrase[rdp] = replaceDynamicPhrase[rdp].replace(/\\n/gi, "\n");
            }
            if ((replaceDynamicPhrase.length % 2) != 0) {
                replaceDynamicPhrase.push ('');
            }
        }
        else {
            replaceDynamicPhrase = [];
        }
        replaceDynamicPhrase = replaceDynamicPhrase.concat (importedReplace);
        if (testDynamicStylePhrase) {
            var testDynamicStylePhraseSplitted = testDynamicStylePhrase.split ("::");
            for (var tdsp = testDynamicStylePhraseSplitted.length - 1; tdsp >= 0; tdsp--) {
                dynamicStylePhrases.push (testDynamicStylePhraseSplitted[tdsp]);
            }
        }
        if (isIDMS) {
            instanceStory.characters.itemByRange(instanceStory.characters.item (startIndex + 1), instanceStory.characters.item (nestedMarkers[nestedMarkers.length-1][0].index - 1)).contents = "";
            var targetPage = null;
            if (nestedMarkers[0][0].parentTextFrames.length > 0) {
                targetPage = nestedMarkers[0][0].parentTextFrames[0].parentPage;
            }
            else {
                targetPage = nestedMarkers[0][0].parentStory.textFrames.lastItem ().parentPage;
            }
            var scoopFrame = targetPage.place (snippetFile)[0];
            if (scoopFrame instanceof TextFrame && scoopFrame.characters.length > 1) {
                oldParaStyleName = scoopFrame.characters.firstItem ().appliedParagraphStyle.name;
                oldCharStyleName = scoopFrame.characters.firstItem ().appliedCharacterStyle.name;
                scoopFrame.characters.itemByRange (1, -2).duplicate(LocationOptions.atBeginning, instanceStory.insertionPoints.item (startIndex + 1));    
            }
            else {
                isToUpdateNested = false;
            }
            scoopFrame.remove();
        }
        else {
            if (pathAndName[0] != "#" && !liveSnippetContent && File.decode (snippetFile.name).slice (-4).toLowerCase() == ".txt")
                liveSnippetContent = readFile (snippetFile);
            if (liveSnippetContent) {
                if (liveSnippetContent.indexOf (tsDynamicSnippetTag) == 0) {
                    dynamicLiveSnippetPhrase = liveSnippetContent.slice (tsDynamicSnippetTag.length);
                    liveSnippetContent = "";
                    inlineDynamicMark = "";
                    if (dynamicLiveSnippetPhrase.indexOf (":\n") != -1) {
                        liveSnippetContent = dynamicLiveSnippetPhrase.slice (dynamicLiveSnippetPhrase.indexOf (":\n") + 2);
                        dynamicLiveSnippetPhrase = dynamicLiveSnippetPhrase.slice (0, dynamicLiveSnippetPhrase.indexOf (":\n"));
                    }
                }
                else if (inlineDynamicMark == "") {
                    dynamicLiveSnippetPhrase = "";
                }
            }
            else if (inlineDynamicMark == "") {
                dynamicLiveSnippetPhrase = "";
            }
            if (dynamicLiveSnippetPhrase) {
                var twoDotsSplit = dynamicLiveSnippetPhrase.split ("::");
                for (var tds = twoDotsSplit.length - 1; tds > 0; tds--) {
                    if (twoDotsSplit[tds].search (/Replace\:/i) == 0) {
                        var explicitReplace = twoDotsSplit[tds].replace (/Replace\:/i, '');
                        explicitReplace = explicitReplace.split (":");
                        for (var er = 0; er < explicitReplace.length; er++) {
                            explicitReplace[er] = solveDynamicLiveSnippetPhrase (targetDocument, snippetFile, nestedMarkers[0][0].parent.insertionPoints.item (nestedMarkers[0][0].index), liveSnippetID, explicitReplace[er], parentID, oldParaStyleName, oldCharStyleName, false);
                            explicitReplace[er] = explicitReplace[er].replace (/Colon/gi, ":");
                            explicitReplace[er] = explicitReplace[er].replace(/\\r/gi, "\r");
                            explicitReplace[er] = explicitReplace[er].replace(/\\t/gi, "\t");
                            explicitReplace[er] = explicitReplace[er].replace(/\\n/gi, "\n");
                        }
                        if (explicitReplace.length > 0) {
                            if ((explicitReplace.length % 2) != 0) {
                                explicitReplace.push ('no_text');
                            }
                            replaceDynamicPhrase = replaceDynamicPhrase.concat (explicitReplace);
                        }
                    }
                    else {
                        dynamicStylePhrases.push (twoDotsSplit[tds]);
                    }
                }
                var dynamicLiveSnippetPhraseSplitted = twoDotsSplit[0].split (":");
                var dynamicLiveSnippetPath = dynamicLiveSnippetPhraseSplitted[0];
                var dynamicLiveSnippetMethod = dynamicLiveSnippetPhraseSplitted.length > 1? dynamicLiveSnippetPhraseSplitted[1] : "Place";
                var dynamicLiveSnippetSeparator = dynamicLiveSnippetPhraseSplitted.length > 2? dynamicLiveSnippetPhraseSplitted[2] : "";
                var firstDivider = "";
                var secondDivider = "";
                var keyTwo = null;
                var keyValue = "";
                var isKeyValueNumber = false;
                var isTableSeparator = false;
                var isRowsSeparator = false;
                var isThereOuterFrame = false;
                var isFrameSeparator = false;
                var isCellsSeparator = false;
                var isFileSplitRepeat = false;
                var isTableRepeat = false;
                var listItemsCount = 0;
                if (dynamicLiveSnippetSeparator.search (/^\[.+\]$/i) != -1) {
                    dynamicLiveSnippetSeparator = dynamicLiveSnippetSeparator.slice (1, -1);
                    isThereOuterFrame = true;
                }
                if (dynamicLiveSnippetSeparator.search (/Table/i) != -1) {
                    isTableSeparator = true;
                    dynamicLiveSnippetSeparator = dynamicLiveSnippetSeparator.replace(/Table/gi, "");
                }
                else if (dynamicLiveSnippetSeparator.search (/Rows/i) != -1) {
                    isRowsSeparator = true;
                    dynamicLiveSnippetSeparator = dynamicLiveSnippetSeparator.replace(/Rows/gi, "");
                }
                else if (dynamicLiveSnippetSeparator.search (/Columns/i) != -1) {
                    isColumnsSeparator = true;
                    dynamicLiveSnippetSeparator = dynamicLiveSnippetSeparator.replace(/Columns/gi, "");
                }
                else if (dynamicLiveSnippetSeparator.search (/Frames/i) != -1) {
                    isFrameSeparator = true;
                    dynamicLiveSnippetSeparator = dynamicLiveSnippetSeparator.replace(/Frames/gi, "");
                }
                else if (dynamicLiveSnippetSeparator.search (/Cells/i) != -1) {
                    isCellsSeparator = true;
                    if (dynamicLiveSnippetSeparator.search (/\-/i) != -1) {
                        isHorizontal = true;
                        cellsDirection = '-';
                    }
                    else if (dynamicLiveSnippetSeparator.search (/\=/i) != -1) {
                        isHorizontal = isPreviousHorizontal;
                        cellsDirection = '=';
                    }
                    else if (dynamicLiveSnippetSeparator.search (/\+/i) != -1) {
                        isHorizontal = !isPreviousHorizontal;
                        cellsDirection = '+';
                    }
                    if (dynamicLiveSnippetSeparator.search (/Stock/i) != -1) {
                        isStock = true;
                    }
                    if (dynamicLiveSnippetSeparator.search (/Nested/i) != -1) {
                        isCellsNested = true;
                    }
                }
                var isFoldersExplicit = false;
                if (dynamicLiveSnippetMethod.search (/Folders/i) != -1) {
                    isFoldersExplicit = true;
                    dynamicLiveSnippetMethod = dynamicLiveSnippetMethod.replace (/\s*Folders\s*/gi, "");
                }
                if (dynamicLiveSnippetMethod.search (/Repeat/i) != -1 && (isRowsSeparator || isColumnsSeparator)) {
                    isTableRepeat = true;
                    dynamicLiveSnippetMethod = "Data";
                }
                var isTargetFolders = false;
                if (isFoldersExplicit && (isTableSeparator || isRowsSeparator || isColumnsSeparator)) {
                    isTargetFolders = true;
                }
                if (dynamicLiveSnippetMethod.search (/Key|Data/i) != -1) {
                    isTargetFolders = false;
                    if (isTableSeparator || isRowsSeparator || isColumnsSeparator) {
                        if (dynamicLiveSnippetMethod.search (/Key/i) != -1) {
                            dynamicLiveSnippetMethod = "KeyTable";
                        }
                    }
                    else if (dynamicLiveSnippetMethod.search (/Key/i) != -1 && dynamicLiveSnippetPhraseSplitted.length == 6) {
                        keyTwo = dynamicLiveSnippetPhraseSplitted.splice (4, 2);
                        keyTwo[0] = keyTwo[0].replace (/Colon/gi, ":");
                        keyTwo[0] = keyTwo[0].replace(/\\r/gi, "\r");
                        keyTwo[0] = keyTwo[0].replace(/\\t/gi, "\t");
                        keyTwo[0] = keyTwo[0].replace(/\\n/gi, "\n");
                        keyTwo[1] = parseInt (keyTwo[1], 10);
                        if (!keyTwo[1]) {
                            keyTwo[1] = 1;
                        }
                    }
                }
                dynamicLiveSnippetSeparator = dynamicLiveSnippetSeparator.replace (/Colon/gi, ":");
                dynamicLiveSnippetSeparator = dynamicLiveSnippetSeparator.replace(/\\r/gi, "\r");
                dynamicLiveSnippetSeparator = dynamicLiveSnippetSeparator.replace(/\\t/gi, "\t");
                dynamicLiveSnippetSeparator = dynamicLiveSnippetSeparator.replace(/\\n/gi, "\n");
                if (dynamicLiveSnippetPhraseSplitted.length > 3) {
                    if (dynamicLiveSnippetPhraseSplitted.length == 4 && dynamicLiveSnippetMethod.search (/Key/i) != -1) {
                        keyValue = dynamicLiveSnippetPhraseSplitted[3];
                        isKeyValueNumber = true;
                    }
                    else {
                        firstDivider = dynamicLiveSnippetPhraseSplitted[3];
                        firstDivider = firstDivider.replace (/Colon/gi, ":");
                        firstDivider = firstDivider.replace(/\\r/gi, "\r");
                        firstDivider = firstDivider.replace(/\\t/gi, "\t");
                        firstDivider = firstDivider.replace(/\\n/gi, "\n");
                    }
                    if (dynamicLiveSnippetMethod.search (/Repeat/i) != -1) {
                        isFileSplitRepeat = true;
                    }
                }
                if (dynamicLiveSnippetPhraseSplitted.length > 4) {
                        if (dynamicLiveSnippetMethod.search (/Key/i) != -1) {
                            keyValue = dynamicLiveSnippetPhraseSplitted[4];
                        }
                        else {
                            secondDivider = dynamicLiveSnippetPhraseSplitted[4];
                            secondDivider = secondDivider.replace (/Colon/gi, ":");
                            secondDivider = secondDivider.replace(/\\r/gi, "\r");
                            secondDivider = secondDivider.replace(/\\t/gi, "\t");
                            secondDivider = secondDivider.replace(/\\n/gi, "\n");
                        }
                }
                var dynamicLiveSnippetFiles = solveDynamicLiveSnippetPhrase (targetDocument, snippetFile, nestedMarkers[0][0].parent.insertionPoints.item (nestedMarkers[0][0].index), liveSnippetID, dynamicLiveSnippetPath, parentID, oldParaStyleName, oldCharStyleName, isTargetFolders);
                var theKeyOrDataContent = null;
                var dataTableHead = new Array;
                if (dynamicLiveSnippetMethod.search (/Key|Data/i) != -1 || isFileSplitRepeat) {
                    if (dynamicLiveSnippetFiles == "#") {
                        theKeyOrDataContent = liveSnippetContent;
                    }
                    else if (dynamicLiveSnippetFiles instanceof Array) {
                        if (dynamicLiveSnippetFiles.length > 0) {
                            if (dynamicLiveSnippetFiles[0] instanceof File) {
                                if (File.decode (dynamicLiveSnippetFiles[0].name).search (/\.(txt|csv|tsv)$/i) != -1) {
                                    var isAlreadyKeyContent = false;
                                    for (var kfl = 0; kfl < snippetsCacheList.length; kfl++) {
                                        if (dynamicLiveSnippetFiles[0].fsName.replace(/\\/g, "/") == snippetsCacheList[kfl][0]) {
                                            theKeyOrDataContent = snippetsCacheList[kfl][1];
                                            isAlreadyKeyContent = true;
                                            break;
                                        }
                                    }
                                    if (!isAlreadyKeyContent) {
                                        if (pathAndName[0] != "#" && snippetFile.fsName.replace(/\\/g, "/") == dynamicLiveSnippetFiles[0].fsName.replace(/\\/g, "/") && liveSnippetContent) {
                                            theKeyOrDataContent = liveSnippetContent;
                                        }
                                        else {
                                            theKeyOrDataContent = readFile (dynamicLiveSnippetFiles[0]);
                                            if (theKeyOrDataContent.indexOf (tsDynamicSnippetTag) == 0) {
                                                if (theKeyOrDataContent.indexOf (":\n") != -1) {
                                                    theKeyOrDataContent = theKeyOrDataContent.slice (theKeyOrDataContent.indexOf (":\n") + 2);
                                                }
                                            }
                                        }
                                        snippetsCacheList.push ([dynamicLiveSnippetFiles[0].fsName.replace(/\\/g, "/"), theKeyOrDataContent]);
                                    }
                                }
                            }
                        }
                    }
                }
                if (theKeyOrDataContent) {
                    if (isFileSplitRepeat) {
                        listItemsCount = theKeyOrDataContent.split (firstDivider).length;
                        dynamicLiveSnippetFiles = new Array;
                    }
                    else if (dynamicLiveSnippetMethod.search (/Key/i) != -1) {
                        if (keyValue) {
                            keyValue = solveDynamicLiveSnippetPhrase (targetDocument, snippetFile, nestedMarkers[0][0].parent.insertionPoints.item (nestedMarkers[0][0].index), liveSnippetID, keyValue, parentID, oldParaStyleName, oldCharStyleName, false);
                            if (keyValue instanceof Array) {
                                keyValue = "";
                            }
                            else if (isKeyValueNumber) {
                                var tempKeyValue = null;
                                var valueNums = keyValue.match (/\d+/g);
                                if (valueNums && valueNums.length > 0) {
                                    tempKeyValue = parseInt (valueNums[0], 10);
                                }
                                if (tempKeyValue) {
                                    keyValue = tempKeyValue;
                                }
                                else {
                                    keyValue = 1;
                                }
                            }
                        }
                        if (dynamicLiveSnippetSeparator == "") {
                            dynamicLiveSnippetSeparator = "\n-\n";
                        }
                        if (firstDivider == "") {
                            firstDivider = ":\n";
                        }
                        theKeyOrDataContent = theKeyOrDataContent.split (dynamicLiveSnippetSeparator);
                        if (!isKeyValueNumber) {
                            for (var kdc = 0; kdc < theKeyOrDataContent.length; kdc++) {
                                theKeyOrDataContent[kdc] = theKeyOrDataContent[kdc].split (firstDivider);
                            }
                        }
                        dynamicLiveSnippetFiles = "";
                        if (!isKeyValueNumber) {
                            for (var kdc = 0; kdc < theKeyOrDataContent.length; kdc++) {
                                if (theKeyOrDataContent[kdc].length > 1) {
                                    if (theKeyOrDataContent[kdc][0] == keyValue) {
                                        dynamicLiveSnippetFiles = theKeyOrDataContent[kdc][1];
                                        var kdci = 2;
                                        while (kdci < theKeyOrDataContent[kdc].length) {
                                            dynamicLiveSnippetFiles += firstDivider + theKeyOrDataContent[kdc][kdci];
                                            kdci++;
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                        else if (theKeyOrDataContent.length > (keyValue - 1)) {
                            dynamicLiveSnippetFiles = theKeyOrDataContent[keyValue - 1];
                        }
                    }
                    else {
                        if (isTableSeparator || isRowsSeparator || isColumnsSeparator) {
                            if (firstDivider == "") {
                                firstDivider = "\n";
                            }
                            if (secondDivider == "") {
                                secondDivider = ", ";
                            }
                            theKeyOrDataContent = theKeyOrDataContent.split (firstDivider);
                            for (var kdc = 0; kdc < theKeyOrDataContent.length; kdc++) {
                                theKeyOrDataContent[kdc] = theKeyOrDataContent[kdc].split (secondDivider);
                            }
                            dynamicLiveSnippetFiles = theKeyOrDataContent;
                            if (isTableSeparator) {
                                if (dynamicLiveSnippetFiles.length > 0) {
                                    dataTableHead = dynamicLiveSnippetFiles[0];
                                    dynamicLiveSnippetFiles.splice (0, 1);
                                }
                            }
                        }
                        else {
                            dynamicLiveSnippetFiles = "";
                        }
                    }
                }
                else if (dynamicLiveSnippetMethod.search (/KeyTable/i) != -1) {
                    if (firstDivider == "") {
                        firstDivider = "\n-\n";
                    }
                    if (secondDivider == "") {
                        secondDivider = ":\n";
                    }
                }
                if (dynamicLiveSnippetFiles == "#") {
                    if (dynamicLiveSnippetMethod.search (/KeyTable/i) != -1) {
                        if (dynamicLiveSnippetSeparator) {
                            dynamicLiveSnippetFiles = liveSnippetContent.split (dynamicLiveSnippetSeparator);
                        }
                        else {
                            dynamicLiveSnippetFiles = [liveSnippetContent];
                        }
                    }
                    else {
                        if (firstDivider) {
                            dynamicLiveSnippetFiles = liveSnippetContent.split (firstDivider);
                        }
                        else {
                            dynamicLiveSnippetFiles = [liveSnippetContent];
                        }
                    }
                    if (isTableSeparator || isRowsSeparator || isColumnsSeparator) {
                        if (dynamicLiveSnippetMethod.search (/KeyTable/i) == -1) {
                            dynamicLiveSnippetMethod = "Data";
                        }
                    }
                }
                if (!listItemsCount && dynamicLiveSnippetMethod.search (/Repeat/i) != -1) {
                    if (dynamicLiveSnippetFiles instanceof Array) {
                        listItemsCount = dynamicLiveSnippetFiles.length;
                    }
                    else {
                        listItemsCount = parseInt (dynamicLiveSnippetFiles);
                        if (!listItemsCount) {
                            listItemsCount = 0;
                        }
                    }
                    dynamicLiveSnippetFiles = new Array;
                }
                if (dynamicLiveSnippetFiles instanceof Array) {
                    if (isTableSeparator || isRowsSeparator || isColumnsSeparator) { //All for table but in different methods!
                        var isAdded = false;
                        var isWidthStyle = false;
                        if (isRowsSeparator || isColumnsSeparator || instanceStory.characters.item (startIndex + 1).tables.length == 0) {
                            instanceStory.characters.itemByRange(instanceStory.characters.item (startIndex + 1), instanceStory.characters.item (nestedMarkers[nestedMarkers.length-1][0].index - 1)).contents = "";
                            instanceStory.insertionPoints.item (startIndex + 1).tables.add (); 
                            isAdded = true;
                        }
                        else {
                            var endBorderIndex = nestedMarkers[nestedMarkers.length-1][0].index - 1;
                            if (startIndex + 2 < endBorderIndex)
                                instanceStory.characters.itemByRange(instanceStory.characters.item (startIndex + 2), instanceStory.characters.item (endBorderIndex)).contents = "";
                        }
                        var targetTable = instanceStory.characters.item (startIndex + 1).tables.item (0);
                        if (isAdded) {
                            for (var tc = targetTable.columns.length - 1; tc > 0; tc--) {
                                targetTable.columns.item (tc).remove ();
                            }
                            for (var tr = targetTable.rows.length - 1; tr > 0; tr--) {
                                targetTable.rows.item (tr).remove ();
                            }
                        }
                        var toRemoveIndex = -1;
                        for (var ri = targetTable.rows.length - 1; ri >= 0; ri--) {
                            if (targetTable.rows.item (ri).rowType != RowTypes.HEADER_ROW || !isTableSeparator) {
                                if (toRemoveIndex != -1 && isTableSeparator) {
                                    //targetTable.rows.item (toRemoveIndex).remove ();
                                }
                                toRemoveIndex = ri;
                            }
                            else {
                                break;
                            }
                        }
                        if (toRemoveIndex != -1) {
                            if (toRemoveIndex == 0 && isTableSeparator) {
                                targetTable.rows.add (LocationOptions.AT_BEGINNING).rowType = RowTypes.HEADER_ROW;
                                toRemoveIndex = 1;
                            }
                            var headerRow = null;
                            if (isTableSeparator) {
                                headerRow = targetTable.rows.item (toRemoveIndex - 1);
                            }
                            var columnHeadIndexPairs = new Array;
                            var unusedIndexes = new Array;
                            var emptyColumns = new Array;
                            if (isTableSeparator) {
                                for (var rci = 0; rci < headerRow.cells.length; rci++) {
                                    var targetPara = headerRow.cells.item (rci).paragraphs.length > 0 ? headerRow.cells.item (rci).paragraphs.item (0) : null;
                                    if (targetPara) {
                                        var paraContents = getParaPlainContents (targetPara);
                                        if (paraContents) {
                                            columnHeadIndexPairs.push ([paraContents, rci]);
                                            unusedIndexes.push (rci);
                                        }
                                        else {
                                            emptyColumns.push (rci);
                                        }
                                    }
                                    else {
                                        emptyColumns.push (rci);
                                    }                                     
                                }
                            }                              
                            var counter = 0;
                            var rowsShift = [];
                            var rowsSpans = [];
                            var columnsShift = [];
                            var columnsSpans = [];
                            var maxRowIndex = 0;
                            var maxColumnIndex = 0;
                            var rowOrColumnStylesSplitted = new Array;
                            var rowOrColumnStylesFiles = new Array;
                            for (var dlsf = 0; dlsf < dynamicLiveSnippetFiles.length; dlsf++) { //if "Files" included in Method it will be Files
                                var autoPrefix = "";
                                if (pathAndName[0] != "#" && isTableSeparator && dynamicLiveSnippetMethod.search (/Data/i) == -1 && dynamicLiveSnippetMethod.search (/KeyTable/i) == -1 && inlineDynamicMark == "") {
                                    if (dynamicLiveSnippetFiles[dlsf].parent.fsName.replace(/\\/g, "/").indexOf (snippetFile.parent.fsName.replace(/\\/g, "/")) == 0) {
                                        var folderSubNames = dynamicLiveSnippetFiles[dlsf].parent.fsName.replace(/\\/g, "/").replace (snippetFile.parent.fsName.replace(/\\/g, "/"), "");
                                        var folderName = File.decode (dynamicLiveSnippetFiles[dlsf].name);
                                        var digitsPortion = folderName.match (/\d+/g);
                                        if (digitsPortion) {
                                            if (digitsPortion.length > 0) {
                                                if (parseInt (digitsPortion[0], 10) == (dlsf + 1)) {
                                                    if (digitsPortion[0].length < 5) {
                                                        var replaceWith = "<RN>";
                                                        switch(digitsPortion[0].length) {
                                                            case 3:
                                                                replaceWith = "<RN><S0>";
                                                                break;
                                                            case 2:
                                                                replaceWith = "<RN><S00>";
                                                                break;
                                                            case 1:
                                                                replaceWith = "<RN><S000>";
                                                                break;
                                                        }
                                                        autoPrefix = "<PSP>" + folderSubNames + "/" + folderName.replace (digitsPortion[0], replaceWith);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                //Get cells Folders by reading the row snippet not by this:
                                var cellsFolders = null;
                                var areFiles = false;
                                if (dynamicLiveSnippetMethod.search (/KeyTable/i) == -1) {
                                    rowOrColumnStylesSplitted = new Array;
                                    rowOrColumnStylesFiles = new Array;
                                }
                                if (dynamicLiveSnippetMethod.search (/Data/i) != -1) {
                                    cellsFolders = dynamicLiveSnippetFiles[dlsf];
                                }
                                else if (dynamicLiveSnippetMethod.search (/KeyTable/i) != -1) {
                                    if (dynamicLiveSnippetFiles[dlsf] instanceof File) {
                                        var testCellsFolders = readFile (dynamicLiveSnippetFiles[dlsf]);
                                        if (testCellsFolders) {
                                            if (testCellsFolders.indexOf (tsDynamicSnippetTag) == 0) {
                                                var rowLiveSnippetPhrase = null;
                                                if (testCellsFolders.indexOf (":\n") != -1) {
                                                    rowLiveSnippetPhrase = testCellsFolders.slice (tsDynamicSnippetTag.length, testCellsFolders.indexOf (":\n"));
                                                    testCellsFolders = testCellsFolders.slice (testCellsFolders.indexOf (":\n") + 2);
                                                }
                                                else {
                                                    rowLiveSnippetPhrase = testCellsFolders.slice (tsDynamicSnippetTag.length);
                                                    testCellsFolders = "";
                                                }
                                                if (rowLiveSnippetPhrase) {
                                                    var rowTwoDotsSplit = rowLiveSnippetPhrase.split ("::");
                                                    for (var rtds = 0; rtds < rowTwoDotsSplit.length; rtds++) {
                                                        rowLiveSnippetPhrase = rowTwoDotsSplit[rtds].split (":");
                                                        if (rowLiveSnippetPhrase.length > 1 && rowLiveSnippetPhrase[1].search (/Para|Char|Obj|Height|Width|Area|Rotate|Shear|Short|Long|Merge|Cell|Row|Column|Table|Font|Style|Color|Size|Justification|characterDirection|paragraphDirection/i) != -1) {
                                                            rowOrColumnStylesSplitted.push (rowLiveSnippetPhrase);
                                                            rowOrColumnStylesFiles.push (dynamicLiveSnippetFiles[dlsf]);
                                                        }
                                                    }
                                                }
                                                if (testCellsFolders != "") {
                                                    cellsFolders = testCellsFolders.split (firstDivider);
                                                }
                                                else {
                                                    continue;
                                                }
                                            }
                                            else {
                                                cellsFolders = testCellsFolders.split (firstDivider);
                                            }
                                        }
                                    }
                                    else {
                                        cellsFolders = dynamicLiveSnippetFiles[dlsf].split (firstDivider);
                                    }
                                }
                                else {
                                    var rowLiveSnippetFile = [];
                                    if (isTargetFolders) {
                                        rowLiveSnippetFile = dynamicLiveSnippetFiles[dlsf].getFiles ("-*");
                                    }
                                    else {
                                        rowLiveSnippetFile = [dynamicLiveSnippetFiles[dlsf]];
                                    }
                                    var rowLiveSnippetPhraseSplitted = null;
                                    if (rowLiveSnippetFile.length > 0) {
                                        rowLiveSnippetFile = rowLiveSnippetFile[0];
                                        if (rowLiveSnippetFile.name.slice (-4).toLowerCase() == ".txt") {
                                            var rowSnippetContent = readFile (rowLiveSnippetFile);
                                            if (rowSnippetContent.indexOf (tsDynamicSnippetTag) == 0) {
                                                var rowLiveSnippetPhrase = null;
                                                if (rowSnippetContent.indexOf (":\n") != -1) {
                                                    rowLiveSnippetPhrase = rowSnippetContent.slice (tsDynamicSnippetTag.length, rowSnippetContent.indexOf (":\n"));
                                                }
                                                else {
                                                    rowLiveSnippetPhrase = rowSnippetContent.slice (tsDynamicSnippetTag.length);
                                                }
                                                if (rowLiveSnippetPhrase) {
                                                    var rowTwoDotsSplit = rowLiveSnippetPhrase.split ("::");
                                                    for (var rtds = 0; rtds < rowTwoDotsSplit.length; rtds++) {
                                                        rowLiveSnippetPhrase = rowTwoDotsSplit[rtds].split (":");
                                                        if (rtds == 0) {
                                                            if (rowLiveSnippetPhrase.length > 1) {
                                                                rowLiveSnippetPhraseSplitted = rowLiveSnippetPhrase;                                                            }
                                                        }
                                                        else {
                                                            if (rowLiveSnippetPhrase.length > 1 && rowLiveSnippetPhrase[1].search (/Para|Char|Obj|Height|Width|Area|Rotate|Shear|Short|Long|Merge|Cell|Row|Column|Table|Font|Style|Color|Size|Justification|characterDirection|paragraphDirection/i) != -1) {
                                                                rowOrColumnStylesSplitted.push (rowLiveSnippetPhrase);
                                                                rowOrColumnStylesFiles.push (rowLiveSnippetFile);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (!rowLiveSnippetPhraseSplitted && isTargetFolders) {
                                        var rowLiveSnippetFiles = dynamicLiveSnippetFiles[dlsf].getFiles ("*.txt");
                                        for (var flf = 0; flf < rowLiveSnippetFiles.length; flf++) {
                                            if (!rowLiveSnippetFiles[flf].hidden && rowLiveSnippetFiles[flf].name[0] != '-') {
                                                var rowSnippetContent = readFile (rowLiveSnippetFiles[flf]);
                                                if (rowSnippetContent.indexOf (tsDynamicSnippetTag) == 0) {
                                                    var rowLiveSnippetPhrase = null;
                                                    if (rowSnippetContent.indexOf (":\n") != -1) {
                                                        rowLiveSnippetPhrase = rowSnippetContent.slice (tsDynamicSnippetTag.length, rowSnippetContent.indexOf (":\n"));
                                                    }
                                                    else {
                                                        rowLiveSnippetPhrase = rowSnippetContent.slice (tsDynamicSnippetTag.length);
                                                    }
                                                    if (rowLiveSnippetPhrase) {
                                                        var rowTwoDotsSplit = rowLiveSnippetPhrase.split ("::");
                                                        var isToBreak = false;
                                                        for (var rtds = 0; rtds < rowTwoDotsSplit.length; rtds++) {
                                                            rowLiveSnippetPhrase = rowTwoDotsSplit[rtds].split (":");
                                                            if (rowLiveSnippetPhrase.length > 1) {
                                                                if (rowLiveSnippetPhrase[1] == "Files" || rowLiveSnippetPhrase[1] == "Folders") {
                                                                    rowLiveSnippetPhraseSplitted = rowLiveSnippetPhrase;
                                                                    rowLiveSnippetFile = dynamicLiveSnippetFiles[dlsf];
                                                                    if (!isRowsSeparator && !isColumnsSeparator) {
                                                                        isToBreak = true;
                                                                        break;
                                                                    }
                                                                }
                                                                else if (rowLiveSnippetPhrase[1].search (/Para|Char|Obj|Height|Width|Area|Rotate|Shear|Short|Long|Merge|Cell|Row|Column|Table|Font|Style|Color|Size|Justification|characterDirection|paragraphDirection/i) != -1) {
                                                                    rowOrColumnStylesSplitted.push (rowLiveSnippetPhrase);
                                                                    rowOrColumnStylesFiles.push (rowLiveSnippetFiles[flf]);
                                                                }
                                                            }
                                                        }
                                                        if (isToBreak) {
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (rowLiveSnippetPhraseSplitted) {
                                        var isFoldersList = false;
                                        if (rowLiveSnippetPhraseSplitted[1] == "Folders") {
                                            isFoldersList = true;
                                        }
                                        var rowLiveSnippetFiles = solveDynamicLiveSnippetPhrase (targetDocument, rowLiveSnippetFile, nestedMarkers[0][0].parent.insertionPoints.item (nestedMarkers[0][0].index), liveSnippetID, rowLiveSnippetPhraseSplitted[0], liveSnippetID, null, null, isFoldersList);
                                        if (rowLiveSnippetFiles instanceof Array) {
                                            cellsFolders = rowLiveSnippetFiles;
                                            areFiles = !isFoldersList;
                                        }
                                    }
                                    if (!cellsFolders && isTargetFolders) {
                                        cellsFolders = dynamicLiveSnippetFiles[dlsf].getFiles (isCSFolder);
                                        if (cellsFolders.length == 0) {
                                            cellsFolders = dynamicLiveSnippetFiles[dlsf].getFiles (isUnhiddenFile);
                                            areFiles = true;
                                        }
                                    }
                                }
                                var rowIndex = toRemoveIndex + counter;

                                for (var cfi = 0; cfi < cellsFolders.length; cfi++) {
                                    rowIndex = toRemoveIndex + counter;
                                    if (isTableRepeat) {
                                        break;
                                    }
                                    var columnIndex = isColumnsSeparator? dlsf : cfi;
                                    var fileHead = "";
                                    var cellFiles = null;
                                    if (dynamicLiveSnippetMethod.search (/Data/i) != -1) {
                                        fileHead = dataTableHead[cfi];
                                        cellFiles = [cellsFolders[cfi]];
                                    }
                                    else if (dynamicLiveSnippetMethod.search (/KeyTable/i) != -1) {
                                        var cellsFoldersSplitted = cellsFolders[cfi].split (secondDivider);
                                        if (cellsFoldersSplitted.length > 1) {
                                            fileHead = cellsFoldersSplitted[0];
                                        }
                                        cellFiles = [cellsFoldersSplitted[cellsFoldersSplitted.length - 1]];
                                    }
                                    else {
                                        fileHead = File.decode (cellsFolders[cfi].name);
                                        if (areFiles) {
                                            fileHead = getPureDisplayName (fileHead);
                                        }
                                    }
                                    var headIndex = -1;
                                    if (isTableSeparator && fileHead != "") {
                                        for (var hipi = 0; hipi < columnHeadIndexPairs.length; hipi++) {
                                            if (fileHead == columnHeadIndexPairs[hipi][0]) {
                                                headIndex = columnHeadIndexPairs[hipi][1];
                                                break; 
                                            }
                                        }
                                        if (headIndex == -1) {
                                            if (emptyColumns.length > 0) {
                                                headIndex = emptyColumns[0];
                                                emptyColumns.splice (0, 1);
                                            }
                                            else {
                                                var headIndex = 1;
                                                for (hipi = 0; hipi < columnHeadIndexPairs.length; hipi++) {
                                                    if (columnHeadIndexPairs[hipi][1] >= headIndex) {
                                                        headIndex = columnHeadIndexPairs[hipi][1] + 1;
                                                    }
                                                }
                                                targetTable.columns.add (LocationOptions.AFTER, targetTable.columns.item (headIndex - 1));
                                            }
                                            columnHeadIndexPairs.push ([fileHead, headIndex]);
                                            headerRow.cells.item (headIndex).contents = fileHead;
                                        }
                                        else {
                                            for (var ui = 0; ui < unusedIndexes.length; ui++) {
                                                if (unusedIndexes[ui] == headIndex) {
                                                    unusedIndexes.splice (ui, 1);
                                                    break;
                                                }
                                            }
                                        }
                                        columnIndex = headIndex;
                                    }
                                    else {
                                        if (isColumnsSeparator) {
                                            rowIndex = toRemoveIndex + cfi;
                                        }
                                        else {
                                            columnIndex = cfi;
                                        }
                                    }
                                    var rsi = rowIndex;
                                    while (rowIndex >= rowsShift.length) {
                                        rowsShift.push (0);
                                        rowsSpans.push (1);
                                    }
                                    for (var rsh = 0; rsh < rsi; rsh++) {
                                        rowIndex += rowsShift[rsh];
                                    }
                                    while (rowIndex >= targetTable.rows.length || (isTableSeparator && targetTable.rows.item (rowIndex).rowType != RowTypes.BODY_ROW)) { //targetTable.rows.item (rowIndex).rowType != RowTypes.BODY_ROW
                                        targetTable.rows.add (LocationOptions.AFTER, targetTable.rows.item (rowIndex - 1));
                                    }
                                    if (!isAdded) {
                                        targetTable.rows.item (rowIndex).unmerge ();
                                    }
                                    var csi = columnIndex;
                                    while (columnIndex >= columnsShift.length) {
                                        columnsShift.push (0);
                                        columnsSpans.push (1);
                                    }
                                    for (var csh = 0; csh < csi; csh++) {
                                        columnIndex += columnsShift[csh];
                                    }
                                    var columnsCountToAdd = columnIndex - targetTable.columns.length;
                                    while (columnsCountToAdd-- >= 0) {
                                        targetTable.columns.add (LocationOptions.AT_END);
                                    }
                                    if (rowIndex > maxRowIndex) maxRowIndex = rowIndex;
                                    if (columnIndex > maxColumnIndex) maxColumnIndex = columnIndex;
                                    var targetCell = getCellAt (targetTable, rowIndex, columnIndex, false);
                                    if (rowsSpans[rsi] > targetCell.rows.length) {
                                        var rowsCountDown = rowsSpans[rsi] - targetCell.rows.length;
                                        var nextCountUp = 1;
                                        while (rowsCountDown--) {
                                            var nextCell = getCellAt (targetTable, rowIndex + nextCountUp++, columnIndex, false);
                                            targetCell.merge (nextCell);
                                        }
                                    }
                                    if (columnsSpans[csi] > targetCell.columns.length) {
                                        var columnsCountDown = columnsSpans[csi] - targetCell.columns.length;
                                        var nextCountUp = 1;
                                        while (columnsCountDown--) {
                                            var nextCell = getCellAt (targetTable, rowIndex, columnIndex + nextCountUp++, false);
                                            targetCell.merge (nextCell);
                                        }
                                    }
                                    if (dynamicLiveSnippetMethod.search (/Data/i) == -1 && dynamicLiveSnippetMethod.search (/KeyTable/i) == -1) {
                                        if (areFiles) {
                                            if (dynamicLiveSnippetMethod.search (/Place/i) == -1) {
                                                targetCell.contents = "";
                                            }
                                            cellFiles = [cellsFolders[cfi]];
                                        }
                                        else {
                                            targetCell.contents = "";
                                            cellFiles = cellsFolders[cfi].getFiles (isUnhiddenFile);
                                        }
                                        if (autoPrefix) {
                                            if (cellsFolders[cfi].fsName.replace(/\\/g, "/").indexOf (dynamicLiveSnippetFiles[dlsf].fsName.replace(/\\/g, "/") + "/") != -1) {
                                                var secSubPortion = cellsFolders[cfi].parent.fsName.replace(/\\/g, "/").replace (dynamicLiveSnippetFiles[dlsf].fsName.replace(/\\/g, "/"), "");
                                                if (areFiles) 
                                                    fileHead = "";
                                                else
                                                    fileHead = "/" + fileHead;
                                            } 
                                            else {
                                                autoPrefix = "";
                                            }
                                        }
                                    }
                                    var isPlacedBefore = false;
                                    var forDynamicStylePhrases = new Array;
                                    var forTableCellsStyles = new Array;
                                    for (var cf = 0; cf < cellFiles.length; cf++) {
                                        var isToPlaceAsSnippet = null;
                                        if (dynamicLiveSnippetMethod.search (/Data|KeyTable/i) != -1) {
                                            if (cellFiles[cf].indexOf (tsDynamicSnippetTag) == 0) {
                                                var thisPathAndName = getPathAndName (targetDocument, dynamicLiveSnippetFiles[dlsf]);
                                                var rowLiveSnippetPhrase = null;
                                                if (cellFiles[cf].indexOf (":\n") != -1) {
                                                    rowLiveSnippetPhrase = cellFiles[cf].slice (tsDynamicSnippetTag.length, cellFiles[cf].indexOf (":\n"));
                                                }
                                                else {
                                                    rowLiveSnippetPhrase = cellFiles[cf].slice (tsDynamicSnippetTag.length);
                                                }
                                                if (rowLiveSnippetPhrase) {
                                                    var rowTwoDotsSplit = rowLiveSnippetPhrase.split ("::");
                                                    for (var rtds = 0; rtds < rowTwoDotsSplit.length; rtds++) {
                                                        rowLiveSnippetPhrase = rowTwoDotsSplit[rtds].split (":");
                                                        var isKeyStyle = false;
                                                        if (rowLiveSnippetPhrase.length > 1) {
                                                            if (rowLiveSnippetPhrase[1].search (/Para|Char|Obj|Height|Width|Area|Rotate|Shear|Short|Long|Merge|Cell|Row|Column|Table|Font|Style|Color|Size|Justification|characterDirection|paragraphDirection/i) != -1) {
                                                                isKeyStyle = true;
                                                                tableCellsStyles.push ([rowLiveSnippetPhrase.join (":"), dynamicLiveSnippetFiles[dlsf], targetCell.insertionPoints.item (-1), thisPathAndName]);
                                                            }
                                                        }
                                                        if (!isKeyStyle) {
                                                            targetCell.contents = "";
                                                            var tempReturned = tsPlaceLiveSnippet (targetCell.insertionPoints.item (-1), null, thisPathAndName, targetDocument, dynamicLiveSnippetFiles[dlsf], theContent, true, true, pageItemsIDs, cellsDeep, deepIndex, cellsIndex, replaceDynamicPhrase);
                                                            if (tempReturned instanceof Array) {
                                                                var returnedStyles = tempReturned[0];
                                                                if (returnedStyles.length > 0) {
                                                                    dynamicStylePhrases = dynamicStylePhrases.concat (returnedStyles);
                                                                }
                                                                var returnedShift = tempReturned[1];
                                                                rowsShift[rsi] += returnedShift[2];
                                                                if (returnedShift[0] > rowsSpans[rsi]) {
                                                                    rowsSpans[rsi] = returnedShift[0];
                                                                }
                                                                columnsShift[csi] += returnedShift[3];
                                                                if (returnedShift[1] > columnsSpans[csi]) {
                                                                    columnsSpans[csi] = returnedShift[1];
                                                                }
                                                            }
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                            else if (cellFiles[cf] != "none") {
                                                targetCell.contents = "";
                                                for (var rdp = 0; rdp < replaceDynamicPhrase.length - 1; rdp+=2) {
                                                    var findText = replaceDynamicPhrase[rdp];
                                                    var findRegex = new RegExp(findText,"g");
                                                    cellFiles[cf] = cellFiles[cf].replace (findRegex, replaceDynamicPhrase[rdp + 1]);
                                                }
                                                targetCell.insertionPoints.item (-1).contents = cellFiles[cf];
                                            }
                                        }
                                        else if (dynamicLiveSnippetMethod.search (/Names/i) != -1) {
                                            var fileDisplayName = File.decode (cellFiles[cf].name);
                                            fileDisplayName = getPureDisplayName (fileDisplayName);
                                            for (var rdp = 0; rdp < replaceDynamicPhrase.length - 1; rdp+=2) {
                                                var findText = replaceDynamicPhrase[rdp];
                                                var findRegex = new RegExp(findText,"g");
                                                fileDisplayName = fileDisplayName.replace (findRegex, replaceDynamicPhrase[rdp + 1]);
                                            }
                                            targetCell.insertionPoints.item (-1).contents = fileDisplayName;
                                        }
                                        else {
                                            var fileDisplayName = File.decode (cellFiles[cf].name);
                                            var checkTxtContent = null;
                                            var cellLiveSnippetPhrase = null;
                                            var cellStyleMethod = null;
                                            var newModification = (cellFiles[cf].modified.getTime ()).toString ();
                                            if (fileDisplayName.slice(-4).toLowerCase() == ".txt") {
                                                if (areFiles && dynamicLiveSnippetMethod.search (/Place/i) != -1) {
                                                    var cellContentLabel = targetTable.extractLabel ("TS_LS_CELL_CONTENTS_" + columnIndex + "_" + rowIndex);
                                                    if (cellContentLabel) {
                                                        var isToSubmitCell = false;
                                                        if (cellContentLabel != targetCell.contents) {
                                                            if (targetCell.contents != "") {
                                                                cellContentLabel = targetCell.contents;
                                                                targetTable.insertLabel ("TS_LS_CELL_CONTENTS_" + columnIndex + "_" + rowIndex, cellContentLabel);
                                                                isToSubmitCell = true;
                                                            }
                                                            else if (cellContentLabel != "no_text") {
                                                                cellContentLabel = "no_text";
                                                                targetTable.insertLabel ("TS_LS_CELL_CONTENTS_" + columnIndex + "_" + rowIndex, cellContentLabel);
                                                                isToSubmitCell = true;
                                                            }
                                                            if (isToSubmitCell) {
                                                                cellFiles[cf].encoding = "UTF8";
                                                                writeFile (cellFiles[cf], cellContentLabel);
                                                                
                                                                //Mark the snippet instrance with the modification date of the snippet file.
                                                                targetTable.insertLabel ("TS_LS_CELL_DATE_" + columnIndex + "_" + rowIndex, newModification);
                                                                //Order Bridge to checkin the file
                                                                var checkInMessageFolder = new Folder (dataPath + "/Messages/To CheckIn Live Snippet");
                                                                if (!checkInMessageFolder.exists) {
                                                                    checkInMessageFolder.create ();
                                                                }
                                                                var cCellSubmit = 1;
                                                                while (File (checkInMessageFolder.fsName.replace(/\\/g, "/") + "/" + newModification + cCellSubmit).exists) {
                                                                    cCellSubmit++;
                                                                }
                                                                writeEncodedFile (File (checkInMessageFolder.fsName.replace(/\\/g, "/") + "/" + newModification + cCellSubmit), cellFiles[cf].fsName.replace(/\\/g, "/"));
                                                                continue;
                                                            }
                                                        }
                                                        if (!isToSubmitCell) {
                                                            if (newModification == targetTable.extractLabel ("TS_LS_CELL_DATE_" + columnIndex + "_" + rowIndex)) {
                                                                continue;
                                                            }
                                                        }
                                                    }
                                                }
                                                checkTxtContent = readFile (cellFiles[cf]);
                                                if (checkTxtContent.indexOf (tsDynamicSnippetTag) == 0) {
                                                    cellLiveSnippetPhrase = checkTxtContent.slice (tsDynamicSnippetTag.length);
                                                    var cellPhraseSplitted = cellLiveSnippetPhrase.split (":");
                                                    if (cellPhraseSplitted.length > 1) {
                                                        if (cellPhraseSplitted[1].search (/Para|Char|Obj|Height|Width|Area|Rotate|Shear|Short|Long|Merge|Cell|Row|Column|Table|Font|Style|Color|Size|Justification|characterDirection|paragraphDirection/i) != -1) {
                                                            cellStyleMethod = cellPhraseSplitted[1];
                                                        }
                                                    }
                                                    isToPlaceAsSnippet = true;
                                                }
                                                else if (dynamicLiveSnippetMethod.search (/Snippets/i) != -1) {
                                                    isToPlaceAsSnippet = true;
                                                }
                                                else {
                                                    isToPlaceAsSnippet = false;
                                                }
                                            }
                                            else if (fileDisplayName.slice(fileDisplayName.lastIndexOf (".")) == ".IDMS") {
                                                isToPlaceAsSnippet = true;
                                            }
                                            if (isPlacedBefore && !cellStyleMethod) {
                                                targetCell.insertionPoints.item (-1).contents = dynamicLiveSnippetSeparator;
                                            }
                                            if (isToPlaceAsSnippet) {
                                                var thisPathAndName = null;
                                                if (autoPrefix) {
                                                    var numberPortion = fileDisplayName;
                                                    if (fileDisplayName.indexOf (".") != -1) {
                                                        numberPortion = fileDisplayName.slice (0, fileDisplayName.lastIndexOf ("."));
                                                    }
                                                    thisPathAndName = autoPrefix + secSubPortion + fileHead + "/" + numberPortion;
                                                }
                                                if (!thisPathAndName)
                                                    thisPathAndName = getPathAndName (targetDocument, cellFiles[cf]);
                                                if (cellStyleMethod) {
                                                    if (cellStyleMethod.search (/Inner/i) != -1 || cellStyleMethod.indexOf ("^") != -1) {
                                                        forDynamicStylePhrases.push (cellLiveSnippetPhrase);
                                                    }
                                                    else {
                                                        forTableCellsStyles.push ([cellLiveSnippetPhrase, cellFiles[cf], targetCell.insertionPoints.item (-1), thisPathAndName]);
                                                    }
                                                }
                                                else {
                                                    var tempReturned = tsPlaceLiveSnippet (targetCell.insertionPoints.item (-1), null, thisPathAndName, targetDocument, cellFiles[cf], checkTxtContent, (dynamicLiveSnippetMethod.search (/Snippets/i) == -1), !isColumnsSeparator, pageItemsIDs, cellsDeep, deepIndex, cellsIndex, replaceDynamicPhrase);
                                                    if (tempReturned instanceof Array) {
                                                        var returnedStyles = tempReturned[0];
                                                        if (returnedStyles.length > 0) {
                                                            dynamicStylePhrases = dynamicStylePhrases.concat (returnedStyles);
                                                        }
                                                        var returnedShift = tempReturned[1];
                                                        rowsShift[rsi] += returnedShift[2];
                                                        if (returnedShift[0] > rowsSpans[rsi]) {
                                                            rowsSpans[rsi] = returnedShift[0];
                                                        }
                                                        columnsShift[csi] += returnedShift[3];
                                                        if (returnedShift[1] > columnsSpans[csi]) {
                                                            columnsSpans[csi] = returnedShift[1];
                                                        }
                                                    }
                                                    isPlacedBefore = true;
                                                }
                                            }
                                            else if (isToPlaceAsSnippet == false) {
                                                if (areFiles) {
                                                    if (checkTxtContent == "no_text") {
                                                        targetCell.contents = "";
                                                    }
                                                    else {
                                                        for (var rdp = 0; rdp < replaceDynamicPhrase.length - 1; rdp+=2) {
                                                            var findText = replaceDynamicPhrase[rdp];
                                                            var findRegex = new RegExp(findText,"g");
                                                            checkTxtContent = checkTxtContent.replace (findRegex, replaceDynamicPhrase[rdp + 1]);
                                                        }
                                                        targetCell.contents = checkTxtContent;
                                                    }
                                                    targetTable.insertLabel ("TS_LS_CELL_CONTENTS_" + columnIndex + "_" + rowIndex, checkTxtContent);
                                                    targetTable.insertLabel ("TS_LS_CELL_DATE_" + columnIndex + "_" + rowIndex, newModification);
                                                }
                                                else {
                                                    if (checkTxtContent == "no_text") {
                                                        targetCell.insertionPoints.item (-1).contents = ""; 
                                                    }
                                                    else {
                                                        for (var rdp = 0; rdp < replaceDynamicPhrase.length - 1; rdp+=2) {
                                                            var findText = replaceDynamicPhrase[rdp];
                                                            var findRegex = new RegExp(findText,"g");
                                                            checkTxtContent = checkTxtContent.replace (findRegex, replaceDynamicPhrase[rdp + 1]);
                                                        }
                                                        targetCell.insertionPoints.item (-1).contents = checkTxtContent;
                                                    }
                                                    isPlacedBefore = true;
                                                }
                                            }
                                            else {
                                                try {
                                                    targetCell.insertionPoints.item (-1).place (cellFiles[cf])[0];
                                                }
                                                catch (e) {}
                                                isPlacedBefore = true;
                                            }
                                        }
                                    }
                                    for (var fdsp = forDynamicStylePhrases.length - 1; fdsp >= 0; fdsp--) {
                                        dynamicStylePhrases.push (forDynamicStylePhrases[fdsp]);
                                    }
                                    for (var ftcs = forTableCellsStyles.length - 1; ftcs >= 0; ftcs--) {
                                        tableCellsStyles.push (forTableCellsStyles[ftcs]);
                                    }
                                    if ((isColumnsSeparator && dlsf == dynamicLiveSnippetFiles.length - 1) || (!isColumnsSeparator && cfi == cellsFolders.length - 1)) {
                                        for (var rcs = rowOrColumnStylesSplitted.length - 1; rcs >= 0; rcs--) {
                                            if (rowOrColumnStylesSplitted[rcs][1].search (/Inner/i) != -1 || rowOrColumnStylesSplitted[rcs][1].indexOf ("^") != -1) {
                                                dynamicStylePhrases.push (rowOrColumnStylesSplitted[rcs].join (":"));
                                            }
                                            else {
                                                var thisPathAndName = getPathAndName (targetDocument, rowOrColumnStylesFiles[rcs]);
                                                tableCellsStyles.push ([rowOrColumnStylesSplitted[rcs].join (":"), rowOrColumnStylesFiles[rcs], targetCell.insertionPoints.item (-1), thisPathAndName]);
                                            }
                                        }
                                        if (dynamicLiveSnippetMethod.search (/KeyTable/i) != -1) {
                                            rowOrColumnStylesSplitted = new Array;
                                            rowOrColumnStylesFiles = new Array;
                                        }
                                    }
                                }
                                if (isTableRepeat) {
                                    if (isRowsSeparator) {
                                        if (rowIndex > maxRowIndex) maxRowIndex = rowIndex;
                                        while (rowIndex >= targetTable.rows.length) {
                                            targetTable.rows.add (LocationOptions.AFTER, targetTable.rows.item (rowIndex - 1));
                                        }
                                        for (var ce = 0; ce < targetTable.rows.item (rowIndex).cells.length; ce++) {
                                            if (dlsf > 0 ) {
                                                if (targetTable.rows.item (rowIndex - 1).cells.item (ce).paragraphs.length > 0 && targetTable.rows.item (rowIndex).cells.item (ce).contents == "") {
                                                    targetTable.rows.item (rowIndex - 1).cells.item (ce).paragraphs.everyItem ().duplicate (LocationOptions.AFTER, targetTable.rows.item (rowIndex).cells.item (ce).insertionPoints.item (-1));
                                                } 
                                            }
                                            var copiedParas = targetTable.rows.item (rowIndex).cells.item (ce).paragraphs;
                                            updateParagraphs (targetDocument, copiedParas, liveSnippetID, snippetsCacheList, cellsFolders, !isColumnsSeparator);
                                        }
                                    }
                                }
                                else if (isTableSeparator) {
                                    if (dlsf > 0) {
                                        for (var ui = 0; ui < unusedIndexes.length; ui++) {
                                            if (targetTable.rows.item (rowIndex - 1).cells.item (unusedIndexes[ui]).paragraphs.length > 0 && targetTable.rows.item (toRemoveIndex + counter).cells.item (unusedIndexes[ui]).contents == "")
                                                targetTable.rows.item (rowIndex - 1).cells.item (unusedIndexes[ui]).paragraphs.everyItem ().duplicate (LocationOptions.AFTER, targetTable.rows.item (toRemoveIndex + counter).cells.item (unusedIndexes[ui]).insertionPoints.item (-1));
                                        }
                                    }
                                    for (var ui = 0; ui < unusedIndexes.length; ui++) {
                                        var copiedParas = targetTable.rows.item (rowIndex).cells.item (unusedIndexes[ui]).paragraphs;
                                        for (var cp = 0; cp < copiedParas.length; cp++) {
                                            if (copiedParas[cp].characters.length > 0) {
                                                if (copiedParas[cp].characters.item(0).textFrames.length > 0) {
                                                    if (copiedParas[cp].characters.item(0).textFrames.item (0).extractLabel("LS_TAG") == "snippet start") {
                                                        var bodyLastCharIndex = new Array;
                                                        var bodyNestedMarkers = new Array;
                                                        bodyLastCharIndex.push(0);
                                                        surveyAndCheck (copiedParas[cp].characters.item(0), bodyLastCharIndex, bodyNestedMarkers, false);
                                                        if (bodyNestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
                                                            continue;
                                                        }
                                                        else if (bodyNestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
                                                            continue;
                                                        }
                                                        updateInstance (targetDocument, null, bodyNestedMarkers, true, null, liveSnippetID, null, false, keyFilesList, true, pageItemsIDs, cellsDeep, deepIndex, cellsIndex, replaceDynamicPhrase);    
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                counter++;
                            }
                            if (isColumnsSeparator || isRowsSeparator) {
                                if (!isRowsSeparator) {
                                    for (var tc = targetTable.columns.length - 1; tc > 0; tc--) {
                                        if (tc > maxColumnIndex) {
                                            targetTable.columns.item (tc).remove ();
                                        }
                                        else {
                                            break;
                                        }
                                    }
                                }
                                var removingStarted = false;
                                for (var tr = targetTable.rows.length - 1; tr > 0; tr--) {
                                    if (targetTable.rows.item (tr).rowType != RowTypes.BODY_ROW) {
                                        continue;
                                    }
                                    if (tr > maxRowIndex) {
                                        if (removingStarted && targetTable.rows.item (tr - 1).rowType != RowTypes.BODY_ROW) {
                                            break;
                                        }
                                        targetTable.rows.item (tr).remove ();
                                        removingStarted = true;
                                    }
                                    else {
                                        break;
                                    }
                                }
                            }
                            if (isTableSeparator) {
                                for (var hf = 0; hf < targetTable.rows.length; hf++) {
                                    if (targetTable.rows.item (hf).rowType == RowTypes.HEADER_ROW || targetTable.rows.item (hf).rowType == RowTypes.FOOTER_ROW) {
                                        var sourceCell = null;
                                        for (var hfc = 0; hfc < targetTable.rows.item (hf).cells.length; hfc++) {
                                            var theCell = targetTable.rows.item (hf).cells.item (hfc);
                                            if (theCell.paragraphs.length == 0) {
                                                if (sourceCell) {
                                                    sourceCell.paragraphs.everyItem ().duplicate (LocationOptions.AFTER, theCell.insertionPoints.item (-1));
                                                }
                                            }
                                            else {
                                                sourceCell = theCell;
                                            }
                                            for (var hfcp = 0; hfcp < theCell.paragraphs.length; hfcp++) {
                                                if (theCell.paragraphs[hfcp].characters.length > 0) {
                                                    if (theCell.paragraphs[hfcp].characters.item(0).textFrames.length > 0) {
                                                        if (theCell.paragraphs[hfcp].characters.item(0).textFrames.item (0).extractLabel("LS_TAG") == "snippet start") {
                                                            var tableLastCharIndex = new Array;
                                                            var tableNestedMarkers = new Array;
                                                            tableLastCharIndex.push(0);
                                                            surveyAndCheck (theCell.paragraphs[hfcp].characters.item(0), tableLastCharIndex, tableNestedMarkers, false);
                                                            if (tableNestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
                                                                continue;
                                                            }
                                                            else if (tableNestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
                                                                continue;
                                                            }
                                                            updateInstance (targetDocument, null, tableNestedMarkers, true, null, liveSnippetID, null, false, keyFilesList, true, pageItemsIDs, cellsDeep, deepIndex, cellsIndex, replaceDynamicPhrase);    
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else if (dynamicLiveSnippetMethod.search (/Repeat/i) != -1) {
                        if (nestedMarkers[nestedMarkers.length-1][0].index - startIndex > 1) {
                            isToUpdateNested = true;
                            if (listItemsCount == 0) {
                                instanceStory.characters.itemByRange(instanceStory.characters.item (startIndex + 1), instanceStory.characters.item (nestedMarkers[nestedMarkers.length-1][0].index - 1)).contents = "";
                            }
                            else {
                                instanceStory.characters.itemByRange(instanceStory.characters.item (startIndex + 2), instanceStory.characters.item (nestedMarkers[nestedMarkers.length-1][0].index - 1)).contents = ""; //delete all snippet content except first character
                                while (listItemsCount-- > 1) {
                                    instanceStory.characters.item (startIndex + 1).duplicate(LocationOptions.atBeginning, instanceStory.insertionPoints.item(startIndex + 1));
                                    instanceStory.insertionPoints.item (startIndex + 2).contents = dynamicLiveSnippetSeparator;
                                }
                            }
                        }
                    }
                    else if (dynamicLiveSnippetMethod.search (/Place|Snippets/i) != -1) {
                        var placeList = new Array;
                        var actualPlacedCount = 0;
                        for (var dlsf = dynamicLiveSnippetFiles.length - 1; dlsf >= 0; dlsf--) {
                            var fileDisplayName = '';
                            var isToPlaceAsSnippet = null;
                            var checkTxtContent = null;
                            var checkLiveSnippetPhrase = null;
                            var isStyleMethod = false;
                            if (!(dynamicLiveSnippetFiles[dlsf] instanceof File)) {
                                checkTxtContent = dynamicLiveSnippetFiles[dlsf];
                                isToPlaceAsSnippet = false;
                                if (checkTxtContent.indexOf (tsDynamicSnippetTag) == 0) {
                                    isToPlaceAsSnippet = true;
                                    dynamicLiveSnippetFiles[dlsf] = snippetFile;
                                }
                            }
                            if ((dynamicLiveSnippetFiles[dlsf] instanceof File) || isToPlaceAsSnippet) {
                                fileDisplayName = File.decode (dynamicLiveSnippetFiles[dlsf].name);
                                if (fileDisplayName.search (/\.(txt|csv|tsv|md)$/i) != -1) {
                                    if (checkTxtContent == null) {
                                        if (pathAndName[0] != "#" && snippetFile.fsName.replace(/\\/g, "/") == dynamicLiveSnippetFiles[dlsf].fsName.replace(/\\/g, "/") && liveSnippetContent) {
                                            checkTxtContent = liveSnippetContent;
                                        }
                                        else {
                                            checkTxtContent = readFile (dynamicLiveSnippetFiles[dlsf]);
                                        }
                                    }
                                    if (checkTxtContent.indexOf (tsDynamicSnippetTag) == 0) {
                                        var possibleStyleBody = '';
                                        if (checkTxtContent.indexOf (":\n") != -1) {
                                            checkLiveSnippetPhrase = checkTxtContent.slice (tsDynamicSnippetTag.length, checkTxtContent.indexOf (":\n"));
                                            possibleStyleBody = checkTxtContent.slice (checkTxtContent.indexOf (":\n") + 2);
                                        }
                                        else {
                                            checkLiveSnippetPhrase = checkTxtContent.slice (tsDynamicSnippetTag.length);
                                        }
                                        var checkPhraseSplitted = checkLiveSnippetPhrase.split (":");
                                        if (checkPhraseSplitted.length > 1) {
                                            if (checkPhraseSplitted[1].search (/Para|Char|Obj|Height|Width|Area|Rotate|Shear|Short|Long|Merge|Cell|Row|Column|Table|Font|Style|Color|Size|Justification|characterDirection|paragraphDirection/i) != -1) {
                                                isStyleMethod = true;
                                                if (possibleStyleBody != '') {
                                                    checkLiveSnippetPhrase = checkLiveSnippetPhrase.replace (/#/g, possibleStyleBody);
                                                }
                                            }
                                        }
                                        isToPlaceAsSnippet = true;
                                    }
                                    else if (dynamicLiveSnippetMethod.search (/Snippets/i) != -1) {
                                        isToPlaceAsSnippet = true;
                                    }
                                    else {
                                        isToPlaceAsSnippet = false;
                                    }
                                }
                                else if (fileDisplayName.slice(fileDisplayName.lastIndexOf (".")) == ".IDMS") {
                                    isToPlaceAsSnippet = true;
                                }
                            }
                            var thisIndex = actualPlacedCount;
                            if (isStyleMethod && thisIndex > 0) {
                                thisIndex--;
                            }
                            placeList.unshift ([
                                fileDisplayName,
                                isToPlaceAsSnippet,
                                checkTxtContent,
                                checkLiveSnippetPhrase,
                                isStyleMethod,
                                thisIndex
                            ]);
                            if (!isStyleMethod) {
                                actualPlacedCount++;
                            }
                        }
                        if (actualPlacedCount == 0) {
                            actualPlacedCount++;
                        }
                        if (isCellsSeparator && !isCellsNested && instanceStory.insertionPoints.item (startIndex + 1).parent.constructor.name == "Cell") {
                            isItNotNested = true;
                            var containerCell = instanceStory.insertionPoints.item (startIndex + 1).parent;
                            containerCell.contents = '';
                            cellsList = createAndGetCells (containerCell.insertionPoints.item (-1), isCellsNested, actualPlacedCount, isHorizontal, isStock, expandValues);
                            
                        }
                        else {
                            instanceStory.characters.itemByRange(instanceStory.characters.item (startIndex + 1), instanceStory.characters.item (nestedMarkers[nestedMarkers.length-1][0].index - 1)).contents = "";
                            if (isCellsSeparator) {
                                cellsList = createAndGetCells (instanceStory.insertionPoints.item (startIndex + 1), isCellsNested, actualPlacedCount, isHorizontal, isStock, expandValues);
                            }
                        }
                        if (isCellsSeparator && cellsList.length > 0) {
                            var cellsFixedList = new Array;
                            for (var csl = 0; csl < cellsList.length; csl++) {
                                cellsFixedList.push (0);
                            }
                            var cellsChildList = new Array;
                            for (var csl = 0; csl < cellsList.length; csl++) {
                                cellsChildList.push (null);
                            }
                            cellsDeep[deepIndex] = [0, isHorizontal, cellsDirection, cellsList, cellsFixedList, cellsIndex, cellsChildList];
                            if (deepIndex > 0) {
                                if (cellsDeep[deepIndex - 1] != null) {
                                    if (cellsDeep[deepIndex - 1][6].length > cellsIndex) {
                                        cellsDeep[deepIndex - 1][6][cellsIndex] = cellsDeep[deepIndex];
                                    }
                                }
                            }
                            if (isHorizontal) {
                                for (var csl = 0; csl < cellsList.length; csl++) {
                                    cellsDeep[deepIndex][0] += cellsList[csl].width;
                                }
                            }
                            else {
                                for (var csl = 0; csl < cellsList.length; csl++) {
                                    cellsDeep[deepIndex][0] += cellsList[csl].height;
                                }
                            }
                        }
                        var isPlacedBefore = false;
                        var outerFrame = null;
                        if (isCellsSeparator) { //reorder placing so dynamic at last
                            var theEnd = 0;
                            for (var dlsf = dynamicLiveSnippetFiles.length - 1; dlsf >= theEnd; dlsf--) {
                                if (placeList[dlsf][1]) {
                                    theEnd++;
                                    dynamicLiveSnippetFiles.unshift (dynamicLiveSnippetFiles[dlsf]);
                                    placeList.unshift (placeList[dlsf]);
                                    dlsf++;
                                    dynamicLiveSnippetFiles.splice (dlsf, 1);
                                    placeList.splice (dlsf, 1);
                                }
                            }
                        }
                        for (var dlsf = dynamicLiveSnippetFiles.length - 1; dlsf >= 0; dlsf--) {
                            var fileDisplayName = placeList[dlsf][0];
                            var isToPlaceAsSnippet = placeList[dlsf][1];
                            var checkTxtContent = placeList[dlsf][2];
                            var checkLiveSnippetPhrase = placeList[dlsf][3];
                            var isStyleMethod = placeList[dlsf][4];
                            var cellIndex = placeList[dlsf][5];
                            var theCellIndex = actualPlacedCount - cellIndex - 1;
                            var targetInsertionPoint = null;
                            if (!isCellsSeparator) {
                                theCellIndex = cellsIndex;
                                if (isThereOuterFrame) {
                                    if (outerFrame == null) {
                                        outerFrame = instanceStory.insertionPoints.item (startIndex + 1).textFrames.add ();
                                        var isOuterBoundFailed = false;
                                        var theOuterBounds = null;
                                        try {
                                            theOuterBounds = outerFrame.geometricBounds.join (",");
                                        }
                                        catch (e) {
                                            isOuterBoundFailed = true;
                                        }
                                        if (!isOuterBoundFailed) {
                                            theOuterBounds = theOuterBounds.split (",");
                                            for (var tbs = 0; tbs < 4; tbs++) 
                                                theOuterBounds[tbs] = parseFloat (theOuterBounds[tbs]);
                                            var theOuterBoundsWidth = theOuterBounds[3] - theOuterBounds[1];
                                            var theOuterBoundsHeight = theOuterBounds[2] - theOuterBounds[0];
                                            outerFrame.geometricBounds = [theOuterBounds[0], theOuterBounds[1], theOuterBounds[2] + theOuterBoundsHeight, theOuterBounds[3] + theOuterBoundsWidth];
                                        }
                                        pageItemsIDs.push ([outerFrame.id, true, false]);
                                    }
                                    targetInsertionPoint = outerFrame.insertionPoints.item (0);
                                }
                                else {
                                    targetInsertionPoint = instanceStory.insertionPoints.item (startIndex + 1);
                                }
                                if (!isStyleMethod) {
                                    if (isPlacedBefore) {
                                        if (isThereOuterFrame) {
                                            outerFrame.insertionPoints.item (0).contents = dynamicLiveSnippetSeparator;
                                            if (dynamicLiveSnippetSeparator.search (/\r/) != -1) {
                                                if (outerFrame.insertionPoints.item (0).appliedParagraphStyle.name.search(/\^$/) != -1) {
                                                    outerFrame.insertionPoints.item (0).appliedParagraphStyle = targetDocument.paragraphStyles.itemByName("[Basic Paragraph]");
                                                }
                                            }
                                        }
                                        else {
                                            targetInsertionPoint.contents = dynamicLiveSnippetSeparator;
                                            if (dynamicLiveSnippetSeparator.search (/\r/) != -1) {
                                                if (targetInsertionPoint.appliedParagraphStyle.name.search(/\^$/) != -1) {
                                                    targetInsertionPoint.appliedParagraphStyle = targetDocument.paragraphStyles.itemByName("[Basic Paragraph]");
                                                }
                                            }
                                        }
                                    }
                                    else {
                                        if (isThereOuterFrame) {
                                            if (outerFrame.insertionPoints.item (0).appliedParagraphStyle.name.search(/\^$/) != -1) {
                                                outerFrame.insertionPoints.item (0).appliedParagraphStyle = targetDocument.paragraphStyles.itemByName("[Basic Paragraph]");
                                            }
                                        }
                                        else {
                                            if (targetInsertionPoint.appliedParagraphStyle.name.search(/\^$/) != -1) {
                                                targetInsertionPoint.appliedParagraphStyle = targetDocument.paragraphStyles.itemByName("[Basic Paragraph]");
                                            }
                                        }
                                    }
                                }
                            }
                            else {
                                targetInsertionPoint = cellsList[theCellIndex].insertionPoints.item (-1);
                            }
                            if (isFrameSeparator && !isStyleMethod && isToPlaceAsSnippet != null) {
                                var listFrame = targetInsertionPoint.textFrames.add ();
                                var isBoundFailed = false;
                                var theBounds = null;
                                try {
                                    theBounds = listFrame.geometricBounds.join (",");
                                }
                                catch (e) {
                                    isBoundFailed = true;
                                }
                                if (!isBoundFailed) {
                                    theBounds = theBounds.split (",");
                                    for (var tbs = 0; tbs < 4; tbs++) 
                                        theBounds[tbs] = parseFloat (theBounds[tbs]);
                                    var theBoundsWidth = theBounds[3] - theBounds[1];
                                    var theBoundsHeight = theBounds[2] - theBounds[0];
                                    listFrame.geometricBounds = [theBounds[0], theBounds[1], theBounds[2] + theBoundsHeight, theBounds[3] + theBoundsWidth];
                                }
                                targetInsertionPoint = listFrame.insertionPoints.item (-1);
                                pageItemsIDs.push ([listFrame.id, true, false]);
                            }
                            else if (!isCellsSeparator && !isThereOuterFrame) {
                                targetInsertionPoint = instanceStory.insertionPoints.item (startIndex + 1);
                            }
                            if (isToPlaceAsSnippet) {
                                var thisPathAndName = getPathAndName (targetDocument, dynamicLiveSnippetFiles[dlsf]);
                                if (isStyleMethod) {
                                    var checkLiveSnippetPhrases = checkLiveSnippetPhrase.split ("::");
                                    for (var clsp = checkLiveSnippetPhrases.length - 1; clsp >= 0; clsp--) {
                                        var thisStyleSplitted = checkLiveSnippetPhrases[clsp].split (":");
                                        if (thisStyleSplitted.length > 1) {
                                            var isInner = true; //isCellsSeparator isFrameSeparator
                                            if (thisStyleSplitted[1].search (/Outer/i) != -1) {
                                                isInner = false;
                                            }
                                            else if (thisStyleSplitted[1].search (/Inner/i) != -1) {
                                                isInner = true;
                                            }
                                            if (isInner || thisStyleSplitted[1].indexOf ("^") != -1) {
                                                dynamicStylePhrases.push (checkLiveSnippetPhrases[clsp]);
                                            }
                                            else {
                                                applyDynamicStyle (checkLiveSnippetPhrases[clsp], targetDocument, dynamicLiveSnippetFiles[dlsf], targetInsertionPoint, false, thisPathAndName, liveSnippetID, null, null, false, [], cellsDeep, deepIndex, theCellIndex);
                                            }
                                        }
                                    }
                                }
                                else {
                                    var tempReturned = tsPlaceLiveSnippet (targetInsertionPoint, null, thisPathAndName, targetDocument, dynamicLiveSnippetFiles[dlsf], checkTxtContent, (dynamicLiveSnippetMethod.search (/Place/i) != -1), isHorizontal, pageItemsIDs, cellsDeep, deepIndex, theCellIndex, replaceDynamicPhrase);
                                    if (tempReturned instanceof Array) {
                                        var returnedStyles = tempReturned[0];
                                        if (returnedStyles.length > 0) {
                                            dynamicStylePhrases = dynamicStylePhrases.concat (returnedStyles);
                                        }
                                        if (tempReturned[1][0] > 1) {
                                            tempReturned[1][0]--;
                                            expandValues[0] += tempReturned[1][0];
                                        }
                                        if (tempReturned[1][1] > 1) {
                                            tempReturned[1][1]--;
                                            expandValues[1] += tempReturned[1][1];
                                        }
                                        expandValues[2] += tempReturned[1][2];
                                        expandValues[3] += tempReturned[1][3];
                                    }
                                    isPlacedBefore = true;
                                }
                            }
                            else if (isToPlaceAsSnippet == false) {
                                for (var rdp = 0; rdp < replaceDynamicPhrase.length - 1; rdp+=2) {
                                    var findText = replaceDynamicPhrase[rdp];
                                    var findRegex = new RegExp(findText,"g");
                                    checkTxtContent = checkTxtContent.replace (findRegex, replaceDynamicPhrase[rdp + 1]);
                                }
                                targetInsertionPoint.contents = checkTxtContent;
                                isPlacedBefore = true;
                            }
                            else {
                                if (fileDisplayName.slice(fileDisplayName.lastIndexOf (".")) == ".idms") {
                                    isToUpdateNested = true;
                                }
                                if (isCellsSeparator && cellsList[theCellIndex].contents == "") {
                                    cellsList[theCellIndex].contents = " ";
                                }
                                var preIndex = targetInsertionPoint.index;
                                try {
                                    targetInsertionPoint.place (dynamicLiveSnippetFiles[dlsf])[0];
                                }
                                catch (e) {
                                }
                                if (preIndex + 1 == targetInsertionPoint.index) {
                                    if (targetInsertionPoint.parent.characters.item (preIndex).pageItems.length > 0) {
                                        targetInsertionPoint.parent.characters.item (preIndex).pageItems[0].appliedObjectStyle = targetDocument.objectStyles.itemByName("Snippet Temporarily Anchored");
                                        if (isCellsSeparator) {
                                            pageItemsIDs.push ([targetInsertionPoint.parent.characters.item (preIndex).pageItems[0].id, false, true, targetInsertionPoint.parent.characters.item (preIndex).pageItems[0]]);
                                        }
                                        else {
                                            pageItemsIDs.push ([targetInsertionPoint.parent.characters.item (preIndex).pageItems[0].id, false, false]);
                                        }
                                    }
                                }
                                isPlacedBefore = true;
                            }
                        }
                    }
                    else if (dynamicLiveSnippetMethod.search (/Names/i) != -1) {
                        instanceStory.characters.itemByRange(instanceStory.characters.item (startIndex + 1), instanceStory.characters.item (nestedMarkers[nestedMarkers.length-1][0].index - 1)).contents = "";
                        for (var dlsf = dynamicLiveSnippetFiles.length - 1; dlsf >= 0; dlsf--) {
                            if (dlsf < (dynamicLiveSnippetFiles.length - 1)) {
                                instanceStory.insertionPoints.item (startIndex + 1).contents = dynamicLiveSnippetSeparator;
                            }
                            var pureName = File.decode (dynamicLiveSnippetFiles[dlsf].name);
                            pureName = getPureDisplayName (pureName);
                            for (var rdp = 0; rdp < replaceDynamicPhrase.length - 1; rdp+=2) {
                                var findText = replaceDynamicPhrase[rdp];
                                var findRegex = new RegExp(findText,"g");
                                pureName = pureName.replace (findRegex, replaceDynamicPhrase[rdp + 1]);
                            }
                            instanceStory.insertionPoints.item (startIndex + 1).contents = pureName;
                        }
                    }
                }
                else {
                    if (keyTwo) {
                        dynamicLiveSnippetFiles = dynamicLiveSnippetFiles.split (keyTwo[0]);
                        if (keyTwo[1] - 1 < dynamicLiveSnippetFiles.length) {
                            dynamicLiveSnippetFiles = dynamicLiveSnippetFiles[keyTwo[1] - 1];
                        }
                        else {
                            dynamicLiveSnippetFiles = "";
                        }
                    }
                    for (var rdp = 0; rdp < replaceDynamicPhrase.length - 1; rdp+=2) {
                        var findText = replaceDynamicPhrase[rdp];
                        var findRegex = new RegExp(findText,"g");
                        dynamicLiveSnippetFiles = dynamicLiveSnippetFiles.replace (findRegex, replaceDynamicPhrase[rdp + 1]);
                    }
                    instanceStory.characters.itemByRange(instanceStory.characters.item (startIndex + 1), instanceStory.characters.item (nestedMarkers[nestedMarkers.length-1][0].index - 1)).contents = dynamicLiveSnippetFiles;
                }
            }
            else if (liveSnippetContent != null) {
                instanceStory.characters.itemByRange(instanceStory.characters.item (startIndex + 1), instanceStory.characters.item (nestedMarkers[nestedMarkers.length-1][0].index - 1)).contents = "";
                if (liveSnippetContent != "no_text") {
                    for (var rdp = 0; rdp < replaceDynamicPhrase.length - 1; rdp+=2) {
                        var findText = replaceDynamicPhrase[rdp];
                        var findRegex = new RegExp(findText,"g");
                        liveSnippetContent = liveSnippetContent.replace (findRegex, replaceDynamicPhrase[rdp + 1]);
                    }
                    instanceStory.insertionPoints.item (startIndex + 1).contents = liveSnippetContent;
                }
            }
        }
        if (isItNotNested) {
            if (dynamicStylePhrases.length > 0) {
                for (var dsp = dynamicStylePhrases.length - 1; dsp >= 0; dsp--) {
                    var dspSplitted = dynamicStylePhrases[dsp].split (":");
                    if (dspSplitted.length > 1 && dspSplitted[1].indexOf ("^") != -1) {
                        dspSplitted[1] = dspSplitted[1].replace ("^", "");
                        dynamicStylePhrases[dsp] = dspSplitted.join (":");
                    }
                    else {
                        var targetIndex = 0;
                        var isPositive = null;
                        var isFromLast = false;
                        var isEvenOnly = null;
                        var appliedCounter = 0;
                        if (dspSplitted[1].search (/For[+-]?\d+[A-Z]?/i) != -1) {
                            var theNum = dspSplitted[1].match (/For[+-]?\d+[A-Z]?/i);
                            if (theNum != null && theNum.length > 0) {
                                theNum = theNum[0].slice (3);
                                if (theNum[0] == '+' || theNum[0] == '-') {
                                    if (theNum[0] == '+') {
                                        isPositive = true;
                                    }
                                    else {
                                        isPositive = false;
                                    }
                                    theNum = theNum.slice (1);
                                }
                                targetIndex = parseInt (theNum, 10) - 1;
                                if (theNum.search (/L/i) != -1) {
                                    isFromLast = true;
                                }
                                if (theNum.search (/E/i) != -1) {
                                    isEvenOnly = true;
                                }
                                else if (theNum.search (/O/i) != -1) {
                                    isEvenOnly = false;
                                }
                            }
                        }
                        else {
                            isFromLast = true;
                        }
                        if (isFromLast) {
                            targetIndex = cellsList.length - targetIndex - 1;
                        }
                        for (var csl = 0; csl < cellsList.length; csl++) {
                            if (targetIndex != null) {
                                if (isPositive == true && csl < targetIndex) {
                                    continue;
                                }
                                else if (isPositive == false && csl > targetIndex) {
                                    continue;
                                }
                                else if (isPositive == null && csl != targetIndex) {
                                    continue;
                                }
                            }
                            appliedCounter++;
                            if (isEvenOnly == true && (appliedCounter % 2 == 1)) {
                                continue;
                            }
                            else if (isEvenOnly == false && (appliedCounter % 2 == 0)) {
                                continue;
                            }
                            if (cellsList[csl].insertionPoints.length > 0) {
                                var cellInsertionPoint = cellsList[csl].insertionPoints.item (0);
                                var textRange = null;
                                if (cellsList[csl].characters.length > 0) {
                                    textRange = cellsList[csl].characters.itemByRange (0, -1);
                                }
                                applyDynamicStyle (dynamicStylePhrases[dsp], targetDocument, snippetFile, cellInsertionPoint, textRange, liveSnippetID, parentID, null, null, false, pageItemsIDs, cellsDeep, deepIndex, cellsIndex);
                            }
                        }
                        dynamicStylePhrases.splice (dsp, 1);
                    }
                }
            }
            return [dynamicStylePhrases, expandValues];
        }
        nestedMarkers = new Array;
        surveyAndCheck(instanceStory.characters.item(startIndex), lastCharIndex, nestedMarkers, false);
        nestedMarkers[0][0].parent.characters.itemByRange(nestedMarkers[0][0].parent.characters.item(nestedMarkers[0][0].index), nestedMarkers[0][0].parent.characters.item(nestedMarkers[nestedMarkers.length-1][0].index)).clearOverrides(OverrideType.ALL);
        if (dynamicStylePhrases.length > 0) {
            var textRange = instanceStory.characters.itemByRange(instanceStory.characters.item (startIndex + 1), instanceStory.characters.item (nestedMarkers[nestedMarkers.length-1][0].index - 1));
            for (var dsp = dynamicStylePhrases.length - 1; dsp >= 0; dsp--) {
                var dspSplitted = dynamicStylePhrases[dsp].split (":");
                if (dspSplitted.length > 1 && dspSplitted[1].indexOf ("^") != -1) {
                    dspSplitted[1] = dspSplitted[1].replace ("^", "");
                    dynamicStylePhrases[dsp] = dspSplitted.join (":");
                }
                else {
                    var isStill = applyDynamicStyle (dynamicStylePhrases[dsp], targetDocument, snippetFile, nestedMarkers[0][0].parent.insertionPoints.item (nestedMarkers[0][0].index), textRange, liveSnippetID, parentID, oldParaStyleName, oldCharStyleName, false, pageItemsIDs, cellsDeep, deepIndex, cellsIndex);
                    if (isStill == false) {
                        return false;
                    }
                    dynamicStylePhrases.splice (dsp, 1);
                }
            }
        }
        if (tableCellsStyles.length > 0) {
            for (var tcs = tableCellsStyles.length - 1; tcs >= 0; tcs--) {
                applyDynamicStyle (tableCellsStyles[tcs][0], targetDocument, tableCellsStyles[tcs][1], tableCellsStyles[tcs][2], false, tableCellsStyles[tcs][3], liveSnippetID, null, null, false, [], cellsDeep, deepIndex, cellsIndex);
            }
        }
        if (isToRemoveMarks) {
            nestedMarkers[nestedMarkers.length-1][0].remove();
        }
        else {
            if (pathAndName[0] != "#" && inlineDynamicMark == "") {
                nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_DATE", snippetFile.modified.getTime ().toString ());
            }
            if (!dynamicLiveSnippetPhrase) {
                var liveSnippetContentAndSUM = getLiveSnippetContentAndSUM (nestedMarkers, 0);
                nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_SUM", liveSnippetContentAndSUM[1]);
                nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_DYNAMIC", "");
                nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_DYNAMIC_INLINE", "");
            }
            else {
                nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_SUM", "DYNAMIC");
                nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_DYNAMIC", dynamicLiveSnippetPhrase);
                nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_DYNAMIC_INLINE", inlineDynamicMark);
            }
            nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_STATE", "updated");
            changeStateAppearance(nestedMarkers[0][0].textFrames.item(0));
        }
    }
    if (isToUpdateNested) {
        c = nestedMarkers.length - 1;
        level = 0;
        var subInstanceMarkers = new Array;
        var subIsForcibly;
        while (c-- > 1) {
            subIsForcibly = isForcibly;
            if (nestedMarkers[c][0].contents == SpecialCharacters.END_NESTED_STYLE) {
                level++;
                if (level == 1) {
                    subInstanceMarkers.length = 0;
                }
                subInstanceMarkers.unshift(nestedMarkers[c]);
            }
            else {
                subInstanceMarkers.unshift(nestedMarkers[c]);
                if (level == 1) {
                    if (subInstanceMarkers[0][0].textFrames.item(0).extractLabel("LS_SUM") == "DYNAMIC") {
                        subIsForcibly = true;
                    }
                    updateInstance (targetDocument, null, subInstanceMarkers, subIsForcibly, null, liveSnippetID, null, false, snippetsCacheList, !isColumnsSeparator, pageItemsIDs, cellsDeep, deepIndex, cellsIndex, replaceDynamicPhrase);
                }
                level--;
            }
        }
    }
    if (isToRemoveMarks) {
        nestedMarkers[0][0].remove();
    }
    if (dynamicStylePhrases.length > 0) {
        return [dynamicStylePhrases, [0, 0, 0, 0]]; //The second array for cells shift
    }
    return true;
}

function createAndGetCells (theInsertionPoint, isCellsNested, theCount, isHorizontal, isStock, expandValues) {
    var returnedCells = [];
    var parentCell = null;
    if (theInsertionPoint.parent.constructor.name != "Cell") {
        isCellsNested = true;
    }
    if (isCellsNested) {
        var addedTable = theInsertionPoint.tables.add ();
        for (var tc = addedTable.columns.length - 1; tc > 0; tc--) {
            addedTable.columns.item (tc).remove ();
        }
        for (var tr = addedTable.rows.length - 1; tr > 0; tr--) {
            addedTable.rows.item (tr).remove ();
        }
        parentCell = addedTable.rows.item (0).cells.item (0);
        parentCell.width = 100; //to avoid bounds error when cell is small
    }
    else {
        parentCell = theInsertionPoint.parent;
    }
    var theTable = parentCell.parentRow.parent;
    var totalCount = theCount;
    var d1Index = parentCell.parentRow.index;
    var d2Index = parentCell.parentColumn.index;
    parentCell.unmerge ();
    isToSwitchIndexes = false;
    var dimension1 = "rows";
    var dimension2 = "columns";
    var HorOrVerEnum = HorizontalOrVertical.HORIZONTAL;
    if (isHorizontal) {
        dimension1 = "columns";
        dimension2 = "rows";
        d1Index = parentCell.parentColumn.index;
        d2Index = parentCell.parentRow.index;
        HorOrVerEnum = HorizontalOrVertical.VERTICAL;
        isToSwitchIndexes = true;
    }
    var d1Count = parentCell[dimension1].length;
    var d2Count = parentCell[dimension2].length;
    for (var r = 0; r < d1Count; r++) {
        for (var c = 0; c < d2Count - 1; c++) {
            var toMergeCell = getCellAt (theTable, d1Index + r, d2Index, isToSwitchIndexes);
            var mergeD2 = getCellAt (theTable, d1Index + r, d2Index + c + 1, isToSwitchIndexes);
            toMergeCell.merge (mergeD2);
        }
    }
    var difference = theCount - d1Count;
    if (!isHorizontal) {
        expandValues[0] = theCount;
    }
    else {
        expandValues[1] = theCount;
    }
    if (difference > 0) {
        if (!isHorizontal) {
            expandValues[2] = difference;
        }
        else {
            expandValues[3] = difference;
        }
        var netDiv = (difference - difference % d1Count) / d1Count;
        var remains = difference % d1Count;
        var frc = netDiv + remains;
        var rrc = netDiv;
        if (isStock) {
            frc = d1Count;
            rrc = 0;
        }
        var isFirstDone = true;
        var counter = d1Count;
        while (theCount-- > d1Count) {
            var theCell = getCellAt (theTable, d1Index + counter - 1, d2Index, isToSwitchIndexes);
            if (isToSwitchIndexes) {
                theCell.width *= 2;
                var minimumWidth = (theCell.textLeftInset + theCell.textRightInset + 1) * 2;
                if (theCell.width < minimumWidth) {
                    theCell.width = minimumWidth;
                }
            }
            for (var ty = 0; ty < 3; ty++) {
                var failed = false;
                try {
                    theCell.split (HorOrVerEnum);
                }
                catch (e) {
                    failed = true;
                    if (isToSwitchIndexes) {
                        theCell.width *= 2;
                    }
                    else {
                        theCell.height *= 2;
                    }
                }
                if (!failed) {
                    break;
                }
            }
            if (isFirstDone) {
                if (frc > 1) {
                    frc--;
                }
                else {
                    if (counter > 1) {
                        counter--;
                    }
                    isFirstDone = false;
                }
            }
            else {
                if (rrc > 1) {
                    rrc--;
                }
                else {
                    if (counter > 1) {
                        counter--;
                    }
                    rrc = netDiv;
                }
            }
        }
    }
    else {
        difference = difference * -1;
        var netDiv = (difference - difference % theCount) / theCount;
        var remains = difference % theCount;
        var frc = netDiv + remains;
        var rrc = netDiv;
        if (isStock) {
            frc = difference;
            rrc = 1;
        }
        var isFirstDone = true;
        var counter = d1Count;
        while (d1Count-- > theCount) {
            var theCell = getCellAt (theTable, d1Index + counter - 1, d2Index, isToSwitchIndexes);
            var previousCell = getCellAt (theTable, d1Index + counter - 2, d2Index, isToSwitchIndexes);
            theCell.merge (previousCell);
            if (isFirstDone) {
                if (frc > 1) {
                    frc--;
                    if (counter > 2) {
                        counter--;
                    }
                }
                else {
                    if (counter > 2) {
                        counter -= 2;
                    }
                    isFirstDone = false;
                }
            }
            else {
                if (rrc > 1) {
                    rrc--;
                    if (counter > 2) {
                        counter--;
                    }
                }
                else {
                    if (counter > 2) {
                        counter -= 2;
                    }
                    rrc = netDiv;
                }
            }
        }
    }
    var FirstIndex = -1;
    for (var v = 0; v < theTable[dimension2][d2Index].cells.length; v++) {
        if (isToSwitchIndexes) {
            if (FirstIndex == -1 && theTable[dimension2][d2Index].cells[v].parentColumn.index == d1Index) {
                FirstIndex = v;
            }
        }
        else {
            if (FirstIndex == -1 && theTable[dimension2][d2Index].cells[v].parentRow.index == d1Index) {
                FirstIndex = v;
            }
        }
        if (FirstIndex >= 0 && v < FirstIndex + totalCount) {
            returnedCells.push (theTable[dimension2][d2Index].cells[v]);
        }
    }
    return returnedCells;
}

function getCellAt (theTable, d1Index, d2Index, isToSwitchIndexes) {
    var theCell = null;
    if (isToSwitchIndexes) {
        var temp = d1Index;
        d1Index = d2Index;
        d2Index = temp;
    }
    if (d1Index < theTable.rows.length) {
        for (var c = 0; c < theTable.rows[d1Index].cells.length; c++) {
            if (theTable.rows[d1Index].cells[c].parentColumn.index == d2Index) {
                theCell = theTable.rows[d1Index].cells[c];
                break;
            }
        }
    }
    return theCell;
}

function updateParagraphs (targetDocument, targetParas, parentID, snippetsCacheList, cellsFolders, isPreviousHorizontal) {
    var firstCharTextFrames = new Array;
    for (var cp = 0; cp < targetParas.length; cp++) {
        for (var api = 0; api < targetParas[cp].allPageItems.length; api++) {
            if (targetParas[cp].allPageItems[api].extractLabel("LS_TAG") == "snippet start") {
                firstCharTextFrames.push (targetParas[cp].allPageItems[api]);
            }
        }
    }
    for (var fctf = firstCharTextFrames.length - 1; fctf >= 0; fctf--) {
        if (firstCharTextFrames[fctf] != null) {
            var bodyLastCharIndex = new Array;
            var bodyNestedMarkers = new Array;
            bodyLastCharIndex.push(0);
            surveyAndCheck (firstCharTextFrames[fctf].parent, bodyLastCharIndex, bodyNestedMarkers, false);
            if (bodyNestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
                continue;
            }
            else if (bodyNestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
                continue;
            }
            var liveSnippetContent = null;
            var liveSnippetID = bodyNestedMarkers[0][0].textFrames.item(0).extractLabel("LS_ID");
            var liveSnippetIDSplitted = new Array;
            liveSnippetIDSplitted.push (liveSnippetID.slice (0, liveSnippetID.lastIndexOf ("/")));
            liveSnippetIDSplitted.push (liveSnippetID.slice (liveSnippetID.lastIndexOf ("/") + 1));
            if (liveSnippetIDSplitted[0] == "#") {
                var snippetNum = parseInt (liveSnippetIDSplitted[1], 10);
                if (snippetNum > 0 && snippetNum <= cellsFolders.length) {
                    liveSnippetContent = cellsFolders[snippetNum - 1];
                }
            }
            updateInstance (targetDocument, null, bodyNestedMarkers, true, null, parentID, liveSnippetContent, false, snippetsCacheList, isPreviousHorizontal, [], [], -1, -1, []);
        }
    }
}

function findTargetFile (targetDocument, liveSnippetMarker) {
    var liveSnippetID = liveSnippetMarker.textFrames.item(0).extractLabel("LS_ID");
    var liveSnippetIDSplitted = new Array;
    liveSnippetIDSplitted.push (liveSnippetID.slice (0, liveSnippetID.lastIndexOf ("/")));
    if (liveSnippetIDSplitted[0] == "#") {
        return false;
    }
    liveSnippetIDSplitted.push (liveSnippetID.slice (liveSnippetID.lastIndexOf ("/") + 1));
    var liveSnippetFiles = new Array;
    getLiveSnippetFolder (targetDocument, liveSnippetIDSplitted[0], liveSnippetIDSplitted[1], liveSnippetFiles);
    if (liveSnippetFiles.length == 0) { //There is no file with this number
        return "There is no file.";	
    }
    else {
        var lastCreatedIndex = 0;
        for (var sf = 1; sf < liveSnippetFiles.length; sf++) {
            if (liveSnippetFiles[sf].created.getTime () > liveSnippetFiles[lastCreatedIndex].created.getTime ()) {
                lastCreatedIndex = sf;
            }
        }
        //No need to delete other files, it's risky to do that!
        /*for (var sfd = 0; sfd < liveSnippetFiles.length; sfd++) {
            if (sfd == lastCreatedIndex)
                continue;
            var fileID = getTSID (liveSnippetFiles[sfd]);
            if (fileID) {
                var deleteMessageFolder = new Folder (dataPath + "/Messages/To Delete");
                if (!deleteMessageFolder.exists) {
                    deleteMessageFolder.create ();
                }
                var messageFileName = fileID.replace (/\//g,"\.");
                messageFileName = messageFileName.slice (1);
                writeEncodedFile (File (deleteMessageFolder.fsName.replace(/\\/g, "/") + "/" + messageFileName), liveSnippetFiles[sfd].fsName.replace(/\\/g, "/"));
            }
            else {
                addToBridgeScanningQueue (liveSnippetFiles[sfd].parent.fsName.replace(/\\/g, "/"));
            }
        }*/
        return liveSnippetFiles[lastCreatedIndex];
	}
}

function sortTextItems (textRange, pageItemsIDs, isOnlyFrames) {
    var indexes = new Array;
    var textItemsPairs = new Array;
    for (var tris = 0; tris < textRange.pageItems.length; tris++) {
        indexes.push (tris);
    }
    for (var pii = pageItemsIDs.length - 1; pii >= 0; pii--) {
        if (isOnlyFrames && !pageItemsIDs[pii][1]) {
            continue;
        }
        else if (isOnlyFrames == false && pageItemsIDs[pii][1]) {
            continue;
        }
        if (pageItemsIDs[pii][2]) {
            textItemsPairs.push ([pageItemsIDs[pii][3], pageItemsIDs[pii][1]]);
        }
        else {
            for (var ind = indexes.length - 1; ind >= 0; ind--) {
                if (pageItemsIDs[pii][0] == textRange.pageItems[indexes[ind]].id) {
                    textItemsPairs.push ([textRange.pageItems[indexes[ind]], pageItemsIDs[pii][1]]);
                    indexes.splice (ind, 1);
                    break;
                }
            }
        }
    }
    return textItemsPairs;
}

function applyDynamicStyle (dynamicStylePhrase, targetDocument, snippetFile, startInsertionPoint, textRange, liveSnippetID, parentID, oldParaStyleName, oldCharStyleName, isTargetFolders, pageItemsIDs, cellsDeep, deepIndex, cellsIndex) {
    var dynamicStylePhraseSplitted = dynamicStylePhrase.split (":");
    var dynamicStylePath = dynamicStylePhraseSplitted[0];
    var basedOnStyle = "";
    if (dynamicStylePath.search (/<BO>|<Based_On>/i) != -1) {
        var dynamicStylePathSplitted = dynamicStylePath.split (/<BO>|<Based_On>/i);
        dynamicStylePath = dynamicStylePathSplitted[0];
        if (dynamicStylePathSplitted.length > 0) {
            basedOnStyle = dynamicStylePathSplitted[1];
        }
    }
    var dynamicStyleMethod = dynamicStylePhraseSplitted.length > 1? dynamicStylePhraseSplitted[1] : "Para";
    var dynamicStyleSeparator = dynamicStylePhraseSplitted.length > 2? dynamicStylePhraseSplitted[2] : " ";
    if (dynamicStyleSeparator == "Colon") {
        dynamicStyleSeparator = ":";
    }
    if (dynamicStyleMethod.search (/Outer/i) != -1) {
        textRange = false;
    }
    var newStyleName = "";
    var dynamicStyleFiles = solveDynamicLiveSnippetPhrase (targetDocument, snippetFile, startInsertionPoint, liveSnippetID, dynamicStylePath, parentID, oldParaStyleName, oldCharStyleName, isTargetFolders);
    if (dynamicStyleFiles instanceof Array) {
        if (dynamicStyleSeparator == "\r" || dynamicStyleSeparator == "\n") {
            dynamicStyleSeparator = " ";
        }
        for (var dsf = dynamicStyleFiles.length - 1; dsf >= 0; dsf--) {
            if (dsf < (dynamicStyleFiles.length - 1)) {
                newStyleName += dynamicStyleSeparator;
            }
            newStyleName += readFile (dynamicStyleFiles[dsf]);
        }
    }
    else {
        newStyleName = dynamicStyleFiles;
    }
    var newBasedOnStyleName = "";
    if (basedOnStyle != "") {
        dynamicStyleFiles = solveDynamicLiveSnippetPhrase (targetDocument, snippetFile, startInsertionPoint, liveSnippetID, basedOnStyle, parentID, oldParaStyleName, oldCharStyleName, isTargetFolders);
        if (dynamicStyleFiles instanceof Array) {
            if (dynamicStyleSeparator == "\r" || dynamicStyleSeparator == "\n") {
                dynamicStyleSeparator = " ";
            }
            for (var dsf = dynamicStyleFiles.length - 1; dsf >= 0; dsf--) {
                if (dsf < (dynamicStyleFiles.length - 1)) {
                    newBasedOnStyleName += dynamicStyleSeparator;
                }
                newBasedOnStyleName += readFile (dynamicStyleFiles[dsf]);
            }
        }
        else {
            newBasedOnStyleName = dynamicStyleFiles;
        }
    }
    if (newStyleName) {
        var isRIE = false;
        var removeStyle = "";
        if (newStyleName.search(/<RIE>|<Remove_If_Empty>/i) !== -1) {
            isRIE = true;
            var tempSplit = newStyleName.split(/<RIE>|<Remove_If_Empty>/i);
            newStyleName = tempSplit[0];
            if (tempSplit.length > 1) {
                removeStyle = tempSplit[1];
            }
        }
        if (dynamicStyleMethod.search (/Height|Width|Area|Short|Long/i) != -1) {
            if (cellsDeep[deepIndex] != null && dynamicStyleMethod.search (/Obj/i) == -1) {
                /*[0, isHorizontal, cellsDirection, cellsList, cellsFixedList, cellsIndex, cellsChildList]*/
                if (dynamicStyleMethod.search (/Short|Long/i) != -1) {
                    if (cellsDeep[deepIndex][1]) {
                        if (dynamicStyleMethod.search (/Long/i) != -1) {
                            dynamicStyleMethod.replace (/Long/i, "Width");
                        }
                        else {
                            dynamicStyleMethod.replace (/Short/i, "Height");
                        }
                    }
                    else {
                        if (dynamicStyleMethod.search (/Long/i) != -1) {
                            dynamicStyleMethod.replace (/Long/i, "Height");
                        }
                        else {
                            dynamicStyleMethod.replace (/Short/i, "Width");
                        }
                    }
                }
                var isForHeight = (dynamicStyleMethod.search (/Height/i) != -1);
                var isPercent = false;
                if (newStyleName.indexOf ("%") != -1) {
                    newStyleName = newStyleName.replace ("%", "");
                    isPercent = true;
                }
                var newValue = parseFloat (newStyleName);
                if (newValue) {
                    if (isPercent) {
                        var parentValue = null;
                        var linedIndex = -1;
                        for (var csd = deepIndex - 1; csd >= 0; csd--) {
                            if (cellsDeep[csd] != null) {
                                if (cellsDeep[csd][1] && !isForHeight) {
                                    linedIndex = csd;
                                    break;
                                }
                                if (!cellsDeep[csd][1] && isForHeight) {
                                    linedIndex = csd;
                                    break;
                                }
                            }
                            else {
                                break;
                            }
                        }
                        if (linedIndex != -1) {
                            parentValue = cellsDeep[linedIndex][0];
                        }
                        else {
                            if (startInsertionPoint.parent.constructor.name == "Cell") {
                                if (isForHeight) {
                                    parentValue = parseFloat (startInsertionPoint.parent.height) - startInsertionPoint.parent.textTopInset - startInsertionPoint.parent.textBottomInset;
                                }
                                else {
                                    parentValue = parseFloat (startInsertionPoint.parent.width) - startInsertionPoint.parent.textLeftInset - startInsertionPoint.parent.textRightInset; 
                                }
                            }
                            else if (startInsertionPoint.parentTextFrames.length) {
                                if (isForHeight) {
                                    var theFrameBounds = startInsertionPoint.parentTextFrames[0].geometricBounds.join (",");
                                    theFrameBounds = theFrameBounds.split (",");
                                    parentValue = parseFloat (theFrameBounds[2]) - parseFloat (theFrameBounds[0]);
                                }
                                else if (startInsertionPoint.parentTextFrames.length) {
                                    parentValue = startInsertionPoint.parentTextFrames[0].textFramePreferences.textColumnFixedWidth;
                                }
                            }
                        }
                        if (parentValue) {
                            newValue = parentValue * (newValue / 100);
                        }
                    }
                    if (!isPercent) {
                        if (deepIndex > 0 && cellsDeep[deepIndex - 1] != null)
                        if (cellsIndex >= 0 && cellsIndex < cellsDeep[deepIndex - 1][4].length) {
                            cellsDeep[deepIndex - 1][4][cellsIndex] = newValue;
                        }
                    }
                    var subCellsList = [0, []];
                    getSubCells (subCellsList, cellsDeep[deepIndex], false, !isForHeight);
                    newValue -= subCellsList[0];
                    if (newValue > 0) {
                        if (isForHeight) {
                            var totalOldHeight = 0;
                            for (var scl = 0; scl < subCellsList[1].length; scl++) {
                                if (!subCellsList[1][scl][1]) {
                                    totalOldHeight += parseFloat (subCellsList[1][scl][0].height);
                                }
                            }
                            if (totalOldHeight > 0) {
                                var changeRatio = newValue / totalOldHeight;
                                for (var scl = 0; scl < subCellsList[1].length; scl++) {
                                    if (!subCellsList[1][scl][1]) {
                                        try {
                                            subCellsList[1][scl][0].height = parseFloat (subCellsList[1][scl][0].height) * changeRatio;
                                        }
                                        catch (e) {
                                        }
                                    }
                                }
                            }
                        }
                        else {
                            var totalOldWidth = 0;
                            for (var scl = 0; scl < subCellsList[1].length; scl++) {
                                if (!subCellsList[1][scl][1]) {
                                    totalOldWidth += parseFloat (subCellsList[1][scl][0].width);
                                }
                            }
                            if (totalOldWidth > 0) {
                                var changeRatio = newValue / totalOldWidth;
                                for (var scl = 0; scl < subCellsList[1].length; scl++) {
                                    if (!subCellsList[1][scl][1]) {
                                        try {
                                            subCellsList[1][scl][0].width = parseFloat (subCellsList[1][scl][0].width) * changeRatio;
                                        }
                                        catch (e) {
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else if (textRange != false) {
                if (textRange != null && textRange.tables.length > 0 && dynamicStyleMethod.search (/Width/i) != -1 && dynamicStyleMethod.search (/Obj/i) == -1) {
                    var isVirtualPercent = false;
                    if (newStyleName.toLowerCase () == "fit") {
                        newStyleName = '100%|Fit';
                    }
                    var virtualParentValue = null;
                    if (newStyleName.search (/\|/) != -1) {
                        var nsnSplitted = newStyleName.split ('|');
                        virtualParentValue = nsnSplitted[0].replace (/(\s+$)|(^\s+)/g, '');
                        if (nsnSplitted.length > 1) {
                            newStyleName = nsnSplitted[1];
                        }
                        else {
                            newStyleName = 'Fit';
                        }
                        if (virtualParentValue.search (/\%$/) != -1) {
                            virtualParentValue = virtualParentValue.replace (/\%/g, '');
                            isVirtualPercent = true;
                        }
                        virtualParentValue = parseFloat (virtualParentValue);
                    }
                    var parentValue = null;
                    if (startInsertionPoint.parent.constructor.name == "Cell") {
                        parentValue = parseFloat (startInsertionPoint.parent.width) - startInsertionPoint.parent.textLeftInset - startInsertionPoint.parent.textRightInset; 
                    }
                    else if (startInsertionPoint.parentTextFrames.length) {
                        parentValue = startInsertionPoint.parentTextFrames[0].textFramePreferences.textColumnFixedWidth;
                    }
                    if (virtualParentValue != null) {
                        if (isVirtualPercent) {
                            parentValue = parentValue * (virtualParentValue / 100);
                        }
                        else {
                            parentValue = virtualParentValue;
                        }
                    }
                    if (newStyleName.toLowerCase () == "fit") {
                        for (var trt = 0; trt < textRange.tables.length; trt++) {
                            var thisTable = textRange.tables[trt];
                            //reduce width till all near overflows
                            if (isResetWidth) {
                                for (var ttc = 0; ttc < thisTable.columns.length; ttc++) {
                                    var thisColumn = thisTable.columns[ttc];
                                    try {thisColumn.width = minColWidth;} catch (e) {}
                                    thisColumn.recompose ();
                                }
                            }
                            var isOverflowsDone = true;
                            do {
                                isOverflowsDone = true;
                                for (var ttc = 0; ttc < thisTable.columns.length; ttc++) {
                                    var thisColumn = thisTable.columns[ttc];
                                    var isColumnOverflowsDone = false;
                                    do {
                                        if (thisColumn.width < parentValue) {
                                            isColumnOverflowsDone = false;
                                            for (var cs = 0; cs < thisColumn.cells.length; cs++) {
                                                var ofs = thisColumn.cells[cs].overflows.toString ();
                                                if (ofs == "true") {
                                                    isColumnOverflowsDone = true;
                                                    break;
                                                }
                                            }
                                            if (isColumnOverflowsDone) {
                                                try {
                                                    thisColumn.width = parseFloat (thisColumn.width) + stepsAmount;
                                                } catch (e) {
                                                    isColumnOverflowsDone = false;
                                                }
                                                thisColumn.recompose ();
                                            }
                                        }
                                        else {
                                            isColumnOverflowsDone = false;
                                        }
                        
                                    } while (isColumnOverflowsDone);
                                }
                            } while (!isOverflowsDone);
                            var isNeedToAdjust = true;
                            var diff = parentValue - thisTable.width;
                            if (isExpandColumnsHaveExtraLines && diff > 0) {
                                //using extra space if there is
                                //calculate columns needing for extra space
                                var extraArr = [];
                                for (var ttc = 0; ttc < thisTable.columns.length; ttc++) {
                                    extraArr.push ([]);
                                }
                                for (var ttc = 0; ttc < thisTable.columns.length; ttc++) {
                                    var thisColumn = thisTable.columns[ttc];
                                    for (var cc = 0; cc < thisColumn.cells.length; cc++) {
                                        var thisCell = thisColumn.cells[cc];
                                        var extraLinesCount = 0;
                                        for (var cp = 0; cp < thisCell.paragraphs.length; cp++) {
                                            var thisPara = thisCell.paragraphs[cp];
                                            for (var cpl = 0; cpl < thisPara.lines.length - 1; cpl++) {
                                                var lineText = '';
                                                if (thisPara.lines[cpl].contents) {
                                                    lineText = thisPara.lines[cpl].contents.toString ();
                                                    if (lineText.charCodeAt (lineText.length - 1) != 10) {
                                                        extraLinesCount++;
                                                    }
                                                }
                                            }
                                        }
                                        if (extraLinesCount != 0) {
                                            for (var ccol = 0; ccol < thisCell.columns.length; ccol++) {
                                                extraArr[thisCell.columns[ccol].index].push (extraLinesCount/thisCell.columns.length);
                                            }
                                        }
                                    }
                                }
                                //distribute
                                totalRatio = 0;
                                var maxCols = [];
                                for (var ea = 0; ea < extraArr.length; ea++) {
                                    maxCols.push (0);
                                    for (var mc = 0; mc < extraArr[ea].length; mc++) {
                                        if (extraArr[ea][mc] > maxCols[ea]) {
                                            maxCols[ea] = extraArr[ea][mc];
                                        }
                                    }
                                    totalRatio += maxCols[ea];
                                }
                                if (totalRatio > 0) {
                                    var ratioFactor = (diff/totalRatio);
                                    for (var ttc = 0; ttc < thisTable.columns.length; ttc++) {
                                        if (maxCols[ttc] != 0) {
                                            try {thisTable.columns[ttc].width = parseFloat (thisTable.columns[ttc].width) + (ratioFactor * maxCols[ttc]);} catch (e) {break;}
                                        }
                                    }
                                    isNeedToAdjust = false;
                                }
                            }
                            if (isNeedToAdjust && isToAdjust) {
                                //fit width to parent textFrame or cell
                                var ratio = parentValue / thisTable.width;
                                for (var ttc = 0; ttc < thisTable.columns.length; ttc++) {
                                    try {thisTable.columns[ttc].width = parseFloat (thisTable.columns[ttc].width) * ratio;} catch (e) {break;}
                                    thisTable.columns[ttc].recompose ();
                                }
                            }
                        }
                    }
                    else {
                        var widthValues = newStyleName.split (",");
                        var fromEndValues = new Array;
                        var isEnd = false;
                        var starValue = '%';
                        for (var wv = widthValues.length - 1; wv >= 0; wv--) {
                            widthValues[wv] = widthValues[wv].replace (/(\s+$)|(^\s+)/g, '');
                            if (widthValues[wv] == '') {
                                widthValues.splice (wv, 1);
                                continue;
                            }
                            if (widthValues[wv].search (/\*/) != -1) {
                                widthValues[wv] = widthValues[wv].replace (/\*/g, '');
                                if (widthValues[wv] != '') {
                                    starValue = widthValues[wv];
                                }
                                isEnd = true;
                                widthValues.splice (wv, 1);
                                continue;
                            }
                            if (isEnd) {
                                fromEndValues.unshift (widthValues[wv]);
                                widthValues.splice (wv, 1);
                                continue;
                            }
                        }
                        for (var trt = 0; trt < textRange.tables.length; trt++) {
                            var thisTable = textRange.tables[trt];
                            var restIndexes = new Array;
                            var restValue = parentValue;
                            for (var ttc = 0; ttc < thisTable.columns.length; ttc++) {
                                var thisValue = null;
                                if (ttc >= (thisTable.columns.length - fromEndValues.length)) {
                                    thisValue = fromEndValues[ttc - thisTable.columns.length - fromEndValues.length];
                                }
                                else if (widthValues.length > ttc) {
                                    thisValue = widthValues[ttc];
                                }
                                if (thisValue == "=") {
                                    restIndexes.push ([ttc, '=']);
                                    continue;
                                }
                                if (thisValue == "%") {
                                    restIndexes.push ([ttc, '%']);
                                    continue;
                                }
                                if (thisValue == null) {
                                    restIndexes.push ([ttc, starValue]);
                                    continue;
                                }
                                if (thisValue.length > 0) {
                                    var isPercent = false;
                                    if (thisValue.indexOf ("%") != -1) {
                                        thisValue = thisValue.replace ("%", "");
                                        isPercent = true;
                                    }
                                    thisValue = parseFloat (thisValue);
                                    if (thisValue) {
                                        if (thisValue > 0) {
                                            if (isPercent && parentValue) {
                                                var assignedValue = parentValue * (thisValue / 100);
                                                restValue -= assignedValue;
                                                try {thisTable.columns[ttc].width = assignedValue;} catch (e) {}
                                            }
                                            else {
                                                try {thisTable.columns[ttc].width = thisValue;} catch (e) {}
                                                restValue -= thisValue;
                                            }
                                        }
                                    }
                                }
                            }
                            var restTotal = 0;
                            for (var rix = 0; rix < restIndexes.length; rix++) {
                                restTotal += parseFloat (thisTable.columns[restIndexes[rix][0]].width);
                            }
                            for (var rix = 0; rix < restIndexes.length; rix++) {
                                restIndexes[rix].push (thisTable.columns[restIndexes[rix][0]].width / restTotal);
                            }
                            var toEqualTotal = 0;
                            var toEqualCount = 0;
                            for (var rix = 0; rix < restIndexes.length; rix++) {
                                if (restIndexes[rix][1] == '=') {
                                    toEqualTotal += restIndexes[rix][2];
                                    toEqualCount++;
                                }
                            }
                            if (toEqualCount > 1) {
                                for (var rix = 0; rix < restIndexes.length; rix++) {
                                    if (restIndexes[rix][1] == '=') {
                                        restIndexes[rix][2] = toEqualTotal / toEqualCount;
                                    }
                                }
                            }
                            for (var rix = 0; rix < restIndexes.length; rix++) {
                                try {thisTable.columns[restIndexes[rix][0]].width = restValue * restIndexes[rix][2];} catch (e) {}
                            }
                        }
                    }
                }
                else if (textRange != null) {
                    var isPercent = false;
                    var asIndex = -1;
                    var asIndexPositive = null;
                    if (dynamicStyleMethod.search (/As[+-]?\d+/i) != -1) {
                        asIndex = dynamicStyleMethod.match (/As[+-]?\d+/i);
                        if (asIndex != null && asIndex.length > 0) {
                            asIndex = asIndex[0].slice (2);
                            if (asIndex[0] == '+' || asIndex[0] == '-') {
                                if (asIndex[0] == '+') {
                                    asIndexPositive = true;
                                }
                                else {
                                    asIndexPositive = false;
                                }
                                asIndex = asIndex.slice (1);
                            }
                            if (asIndexPositive == null) {
                                asIndex = parseInt (asIndex, 10) - 1;
                            }
                            else {
                                asIndex = parseInt (asIndex, 10);
                            }
                        }
                    }
                    if (newStyleName.indexOf ("%") != -1) {
                        newStyleName = newStyleName.replace ("%", "");
                        isPercent = true;
                    }
                    var stylePureValue = parseFloat (newStyleName);
                    var targetIndex = null;
                    var isPositive = null;
                    var isFromLast = false;
                    var isEvenOnly = null;
                    var appliedCounter = 0;
                    var isOnlyFrames = null;
                    if (dynamicStyleMethod.search (/For[+-]?\d+[A-Z]?/i) != -1) {
                        var theNum = dynamicStyleMethod.match (/For[+-]?\d+[A-Z]?/i);
                        if (theNum != null && theNum.length > 0) {
                            theNum = theNum[0].slice (3);
                            if (theNum[0] == '+' || theNum[0] == '-') {
                                if (theNum[0] == '+') {
                                    isPositive = true;
                                }
                                else {
                                    isPositive = false;
                                }
                                theNum = theNum.slice (1);
                            }
                            targetIndex = parseInt (theNum, 10) - 1;
                            if (theNum.search (/L/i) != -1) {
                                isFromLast = true;
                            }
                            if (theNum.search (/T/i) != -1) {
                                isOnlyFrames = true;
                            }
                            else if (theNum.search (/P/i) != -1) {
                                isOnlyFrames = false;
                            }
                            if (theNum.search (/E/i) != -1) {
                                isEvenOnly = true;
                            }
                            else if (theNum.search (/O/i) != -1) {
                                isEvenOnly = false;
                            }
                        }
                    }
                    var textItemsPairs = sortTextItems (textRange, pageItemsIDs, null);
                    if (isFromLast) {
                        var framesCount = 0;
                        for (var tris = 0; tris < textItemsPairs.length; tris++) {
                            if (textItemsPairs[tris][1]) {
                                framesCount++;
                            }
                        }
                        if (isOnlyFrames) {
                            targetIndex = framesCount - targetIndex - 1;
                        }
                        else if (isOnlyFrames == false) {
                            targetIndex = textItemsPairs.length - framesCount - targetIndex - 1;
                        }
                        else {
                            targetIndex = textItemsPairs.length - targetIndex - 1;
                        }
                    }
                    var targetCounter = -1;
                    if (targetIndex == -1 && isPositive == false) {
                        if (startInsertionPoint.parentTextFrames.length) {
                            targetIndex = textItemsPairs.length;
                            textItemsPairs.push([startInsertionPoint.parentTextFrames[0], true]);
                            isPositive = true;
                        }
                    }
                    for (var tris = 0; tris < textItemsPairs.length; tris++) {
                        if (textItemsPairs[tris][1]) {
                            if (isOnlyFrames) {
                                targetCounter++;
                            }
                            else if (isOnlyFrames == false) {
                                continue;
                            }
                            else {
                                targetCounter++;
                            }
                        }
                        else {
                            if (isOnlyFrames) {
                                continue;
                            }
                            else if (isOnlyFrames == false) {
                                targetCounter++;
                            }
                            else {
                                targetCounter++;
                            }
                        }
                        if (targetIndex != null) {
                            if (isPositive == true && targetCounter < targetIndex) {
                                continue;
                            }
                            else if (isPositive == false && targetCounter > targetIndex) {
                                continue;
                            }
                            else if (isPositive == null && targetCounter != targetIndex) {
                                continue;
                            }
                        }
                        appliedCounter++;
                        if (isEvenOnly == true && (appliedCounter % 2 == 1)) {
                            continue;
                        }
                        else if (isEvenOnly == false && (appliedCounter % 2 == 0)) {
                            continue;
                        }
                        var newValue = stylePureValue;
                        var theBounds = null;
                        try {
                            theBounds = textItemsPairs[tris][0].geometricBounds.join (",");
                        }
                        catch (err) {
                            continue;
                        }
                        theBounds = theBounds.split (",");
                        for (var tbs = 0; tbs < 4; tbs++) 
                            theBounds[tbs] = parseFloat (theBounds[tbs]);
                        var isForHeight = true;
                        if (dynamicStyleMethod.search (/Short|Long|Area/i) != -1) {
                            var toCompareWidth = theBounds[3] - theBounds[1];
                            var toCompareHeight = theBounds[2] - theBounds[0];
                            if (dynamicStyleMethod.search (/Area/i) != -1) {
                                var ratio = toCompareWidth / toCompareHeight;
                                var newValue = Math.sqrt(newValue * ratio);
                                isForHeight = false;
                            }
                            else if (dynamicStyleMethod.search (/Short/i) != -1) {
                                if (toCompareWidth < toCompareHeight) {
                                    isForHeight = false;
                                }
                            }
                            else {
                                if (toCompareWidth > toCompareHeight) {
                                    isForHeight = false;
                                }
                            }
                        }
                        else if (dynamicStyleMethod.search (/Width/i) != -1) {
                            isForHeight = false;
                        }
                        if (isPercent) {
                            var parentValue = null;
                            var objectInsertionPoint = textItemsPairs[tris][0].parent;
                            if (objectInsertionPoint instanceof Array) {
                                objectInsertionPoint = objectInsertionPoint[0];
                            }
                            if (objectInsertionPoint.constructor.name == "Character") {
                                objectInsertionPoint = objectInsertionPoint.parent.insertionPoints.item (objectInsertionPoint.index);
                            }
                            if (asIndex != -1) {
                                var actualIndex = asIndex;
                                if (asIndexPositive != null) {
                                    if (asIndexPositive == true) {
                                        actualIndex = tris + asIndex;
                                    }
                                    else {
                                        actualIndex = tris - asIndex;
                                    }
                                    if (actualIndex < 0 || actualIndex >= textItemsPairs.length) {
                                        continue;
                                    }
                                }
                                var asBounds = null;
                                try {
                                    asBounds = textItemsPairs[actualIndex][0].geometricBounds.join (",");
                                }
                                catch (err) {
                                    continue;
                                }
                                asBounds = asBounds.split (",");
                                for (var asbs = 0; asbs < 4; asbs++) 
                                    asBounds[asbs] = parseFloat (asBounds[asbs]);
                                if (isForHeight) {
                                    parentValue = asBounds[2] - asBounds[0];
                                }
                                else {
                                    parentValue = asBounds[3] - asBounds[1];
                                }
                            }
                            else if (objectInsertionPoint.parent.constructor.name == "Cell") {
                                if (isForHeight) {
                                    parentValue = parseFloat (objectInsertionPoint.parent.height) - objectInsertionPoint.parent.textTopInset - objectInsertionPoint.parent.textBottomInset;
                                }
                                else {
                                    parentValue = parseFloat (objectInsertionPoint.parent.width) - objectInsertionPoint.parent.textLeftInset - objectInsertionPoint.parent.textRightInset; 
                                }
                            }
                            else if (objectInsertionPoint.parentTextFrames.length) {
                                if (isForHeight) {
                                    var theFrameBounds = objectInsertionPoint.parentTextFrames[0].geometricBounds.join (",");
                                    theFrameBounds = theFrameBounds.split (",");
                                    parentValue = parseFloat (theFrameBounds[2]) - parseFloat (theFrameBounds[0]);
                                }
                                else {
                                    parentValue = objectInsertionPoint.parentTextFrames[0].textFramePreferences.textColumnFixedWidth;
                                }
                            }
                            if (parentValue) {
                                newValue = parentValue * (newValue / 100);
                            }
                        }
                        if (newValue) {
                            var tIndex = 1;
                            var oIndex = 0;
                            if (isForHeight) {
                                tIndex = 0;
                                oIndex = 1;
                            }
                            var isInAcceptedRange = false;
                            if (theBounds[tIndex + 2] == (theBounds[tIndex] + newValue)) {
                                isInAcceptedRange = true;
                            }
                            else if (dynamicStyleMethod.search (/Min/i) != -1 || dynamicStyleMethod.search (/Max/i) != -1) {
                                if (dynamicStyleMethod.search (/Min/i) != -1) {
                                    if (theBounds[tIndex + 2] > (theBounds[tIndex] + newValue)) {
                                        isInAcceptedRange = true;
                                    }
                                }
                                else {
                                    if (theBounds[tIndex + 2] < (theBounds[tIndex] + newValue)) {
                                        isInAcceptedRange = true;
                                    }
                                }
                            }
                            if (!isInAcceptedRange) {
                                if (dynamicStyleMethod.search (/Relative/i) != -1 || dynamicStyleMethod.search (/Area/i) != -1) {
                                    var secDim = (theBounds[oIndex + 2] - theBounds[oIndex]) * newValue / (theBounds[tIndex + 2] - theBounds[tIndex]);
                                    theBounds[oIndex + 2] = theBounds[oIndex] + secDim;
                                }
                                theBounds[tIndex + 2] = theBounds[tIndex] + newValue;
                                textItemsPairs[tris][0].geometricBounds = [theBounds[0], theBounds[1], theBounds[2], theBounds[3]];
                            }
                        }
                    }
                }
            }
            else {
                if (newStyleName.toLowerCase () == "fit") {
                    var targetObject = null;
                    if (startInsertionPoint.parentTextFrames && startInsertionPoint.parentTextFrames.length > 0) {
                        targetObject = startInsertionPoint.parentTextFrames[0];
                    }
                    if (targetObject != null) {
                        var bounds = targetObject.geometricBounds.join (",");
                        bounds = bounds.split (",");
                        for (var bds = 0; bds < 4; bds++) 
                            bounds[bds] = parseFloat (bounds[bds]);
                        if (dynamicStyleMethod.search (/Group/i) != -1 && targetObject.parent instanceof Group) {
                            var theItems = targetObject.parent.pageItems;
                            for (var ti = 0; ti < theItems.length; ti++) {
                                if (theItems[ti].id != targetObject.id) {
                                    var theItemBounds = theItems[ti].geometricBounds.join (",");
                                    var theItemBounds = theItemBounds.split (",");
                                    for (var bds = 0; bds < 4; bds++) 
                                        theItemBounds[bds] = parseFloat (theItemBounds[bds]);
                                    var newBounds = [bounds[0], bounds[1], bounds[2], bounds[3]];
                                    if (dynamicStyleMethod.search (/Height/i) != -1) {
                                        newBounds[1] = theItemBounds[1];
                                        newBounds[3] = theItemBounds[3];
                                    }
                                    else if (dynamicStyleMethod.search (/Width/i) != -1) {
                                        newBounds[0] = theItemBounds[0];
                                        newBounds[2] = theItemBounds[2];
                                    }
                                    theItems[ti].geometricBounds = [newBounds[0], newBounds[1], newBounds[2], newBounds[3]];
                                }
                            }
                        }
                        if (dynamicStyleMethod.search (/Page/i) != -1 && targetObject.parentPage) {
                            var errorHappened = false;
                            var isForWidth = false;
                            var isForHeight = false;
                            if (dynamicStyleMethod.search (/Height/i) != -1) {
                                isForHeight = true;
                            }
                            if (dynamicStyleMethod.search (/Width/i) != -1) {
                                isForWidth = true;
                            }
                            var visibleBounds = null;
                            try {
                                visibleBounds = targetObject.visibleBounds.join (",");
                            }
                            catch (err) {
                                errorHappened = true;
                            }
                            if (!errorHappened) {
                                visibleBounds = visibleBounds.split (",");
                                var newWidth = visibleBounds[3] - visibleBounds[1];
                                var newHeight = visibleBounds[2] - visibleBounds[0];
                                var originalVisibleWidth = newWidth;
                                var originalVisibleHeight = newHeight;
                                var unitHorFactor = 1;
                                if (targetDocument.viewPreferences.horizontalMeasurementUnits != MeasurementUnits.POINTS) {
                                    switch (targetDocument.viewPreferences.horizontalMeasurementUnits) {
                                        case MeasurementUnits.MILLIMETERS:
                                            unitHorFactor = 2.835;
                                            break;
                                        case MeasurementUnits.CENTIMETERS:
                                            unitHorFactor = 28.346;
                                            break;
                                        case MeasurementUnits.INCHES:
                                            unitHorFactor = 72;
                                            break;
                                        case MeasurementUnits.PICAS:
                                            unitHorFactor = 12;
                                            break;
                                    }
                                }
                                var unitVerFactor = 1;
                                if (targetDocument.viewPreferences.verticalMeasurementUnits != MeasurementUnits.POINTS) {
                                    switch (targetDocument.viewPreferences.verticalMeasurementUnits) {
                                        case MeasurementUnits.MILLIMETERS:
                                            unitVerFactor = 2.835;
                                            break;
                                        case MeasurementUnits.CENTIMETERS:
                                            unitVerFactor = 28.346;
                                            break;
                                        case MeasurementUnits.INCHES:
                                            unitVerFactor = 72;
                                            break;
                                        case MeasurementUnits.PICAS:
                                            unitVerFactor = 12;
                                            break;
                                    }
                                }
                                if (!isForHeight) {
                                    newHeight = targetObject.parentPage.bounds[2] - targetObject.parentPage.bounds[0];
                                }
                                if (!isForWidth) {
                                    newWidth = targetObject.parentPage.bounds[3] - targetObject.parentPage.bounds[1];
                                }
                                try {
                                    targetObject.parentPage.resize (CoordinateSpaces.SPREAD_COORDINATES, 
                                        AnchorPoint.TOP_LEFT_ANCHOR, 
                                        ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH, 
                                        [newWidth * unitHorFactor, newHeight * unitVerFactor]);
                                }
                                catch (err) {
                                    errorHappened = true;
                                }
                                if (!errorHappened) {
                                    targetObject.visibleBounds = [(newHeight - originalVisibleHeight) / 2, (newWidth - originalVisibleWidth) / 2, (newHeight + originalVisibleHeight) / 2, (newWidth + originalVisibleWidth) / 2];
                                }
                            }
                        }
                    }
                }
                else if (startInsertionPoint.parent.constructor.name == "Cell") {
                    var previousDeepIndex = -1;
                    if (deepIndex > 0) {
                        previousDeepIndex = deepIndex - 1;
                        if (cellsDeep[previousDeepIndex] != null) {
                            if (dynamicStyleMethod.search (/Short|Long/i) != -1) {
                                //[0, isHorizontal, cellsDirection, cellsList]
                                if (cellsDeep[previousDeepIndex][1]) {
                                    if (dynamicStyleMethod.search (/Long/i) != -1) {
                                        dynamicStyleMethod.replace (/Long/i, "Width");
                                    }
                                    else {
                                        dynamicStyleMethod.replace (/Short/i, "Height");
                                    }
                                }
                                else {
                                    if (dynamicStyleMethod.search (/Long/i) != -1) {
                                        dynamicStyleMethod.replace (/Long/i, "Height");
                                    }
                                    else {
                                        dynamicStyleMethod.replace (/Short/i, "Width");
                                    }
                                }
                            }
                        }
                    }
                    var isForHeight = (dynamicStyleMethod.search (/Height/i) != -1);
                    var isPercent = false;
                    if (newStyleName.indexOf ("%") != -1) {
                        newStyleName = newStyleName.replace ("%", "");
                        isPercent = true;
                    }
                    var newValue = parseFloat (newStyleName);
                    if (newValue) {
                        if (isPercent) {
                            var parentValue = null;
                            if (previousDeepIndex != -1 && cellsDeep[previousDeepIndex] != null && (cellsDeep[previousDeepIndex][1] != isForHeight)) {
                                parentValue = cellsDeep[previousDeepIndex][0];
                            }
                            else {
                                var parentTable = startInsertionPoint.parent;
                                while (parentTable != null && parentTable.constructor.name != "Table") {
                                    parentTable = parentTable.parent;
                                }
                                if (parentTable.rows.length == 1 && parentTable.columns.length == 1) {
                                    var tableInsertionPoint = parentTable.storyOffset;
                                    if (tableInsertionPoint.parent.constructor.name == "Cell") {
                                        if (isForHeight) {
                                            parentValue = parseFloat (tableInsertionPoint.parent.height) - tableInsertionPoint.parent.textTopInset - startInsertionPoint.parent.textBottomInset;
                                        }
                                        else {
                                            parentValue = parseFloat (tableInsertionPoint.parent.width) - tableInsertionPoint.parent.textLeftInset - startInsertionPoint.parent.textRightInset; 
                                        }
                                    }
                                    else if (tableInsertionPoint.parentTextFrames.length) {
                                        if (isForHeight) {
                                            var theFrameBounds = tableInsertionPoint.parentTextFrames[0].geometricBounds.join (",");
                                            theFrameBounds = theFrameBounds.split (",");
                                            parentValue = parseFloat (theFrameBounds[2]) - parseFloat (theFrameBounds[0]);
                                        }
                                        else {
                                            parentValue = tableInsertionPoint.parentTextFrames[0].textFramePreferences.textColumnFixedWidth;
                                        }
                                    }
                                }
                                else {
                                    if (isForHeight) {
                                        parentValue = parseFloat (parentTable.height);
                                    }
                                    else {
                                        parentValue = parseFloat (parentTable.width);
                                    }
                                }
                            }
                            if (parentValue) {
                                newValue = parentValue * (newValue / 100);
                            }
                        }
                        if (!isPercent) {
                            if (previousDeepIndex != -1 && cellsDeep[previousDeepIndex] != null && (cellsDeep[previousDeepIndex][1] != isForHeight)) {
                                if (cellsIndex >= 0 && cellsIndex < cellsDeep[previousDeepIndex][4].length) {
                                    cellsDeep[previousDeepIndex][4][cellsIndex] = newValue;
                                }
                            }
                        }
                        if (isForHeight) {
                            try {
                                startInsertionPoint.parent.height = newValue;
                            }
                            catch (e) {
                            }
                        }
                        else {
                            try {
                                startInsertionPoint.parent.width = newValue;
                            }
                            catch (e) {
                            }
                        }
                    }
                }
                else {
                    var targetObject = null;
                    if (startInsertionPoint.parentTextFrames && startInsertionPoint.parentTextFrames.length > 0) {
                        targetObject = startInsertionPoint.parentTextFrames[0];
                    }
                    if (targetObject != null) {
                        var isPercent = false;
                        if (newStyleName.indexOf ("%") != -1) {
                            newStyleName = newStyleName.replace ("%", "");
                            isPercent = true;
                        }
                        var newValue = parseFloat (newStyleName);
                        if (isPercent) {
                            if (targetObject.parent.constructor.name == "Character") {
                                var frameInsertionPoint = targetObject.parent.parent.insertionPoints.item (targetObject.parent.index);
                                if (frameInsertionPoint.parent.constructor.name == "Cell") {
                                    if (isForHeight) {
                                        parentValue = parseFloat (frameInsertionPoint.parent.height) - frameInsertionPoint.parent.textTopInset - startInsertionPoint.parent.textBottomInset;
                                    }
                                    else {
                                        parentValue = parseFloat (frameInsertionPoint.parent.width) - frameInsertionPoint.parent.textLeftInset - startInsertionPoint.parent.textRightInset; 
                                    }
                                }
                                else if (frameInsertionPoint.parentTextFrames.length) {
                                    if (isForHeight) {
                                        var theFrameBounds = frameInsertionPoint.parentTextFrames[0].geometricBounds.join (",");
                                        theFrameBounds = theFrameBounds.split (",");
                                        parentValue = parseFloat (theFrameBounds[2]) - parseFloat (theFrameBounds[0]);
                                    }
                                    else {
                                        parentValue = frameInsertionPoint.parentTextFrames[0].textFramePreferences.textColumnFixedWidth;
                                    }
                                }
                            }
                            else if (targetObject.parentPage != null) {
                                if (isForHeight) {
                                    parentValue = targetObject.parentPage.bounds[2] - targetObject.parentPage.bounds[0];
                                }
                                else {
                                    parentValue = targetObject.parentPage.bounds[3] - targetObject.parentPage.bounds[1];
                                }
                            }
                        }
                        var bounds = targetObject.geometricBounds.join (",");
                        bounds = bounds.split (",");
                        for (var bds = 0; bds < 4; bds++) 
                            bounds[bds] = parseFloat (bounds[bds]);
                        if (dynamicStyleMethod.search (/Height/i) != -1) {
                            targetObject.geometricBounds = [bounds[0], bounds[1], bounds[0] + newValue, bounds[3]];    
                        }
                        else {
                            targetObject.textFramePreferences.textColumnFixedWidth = newValue;
                        }   
                    }
                }
            }
        }
        else {
            if (dynamicStyleMethod.search (/Para/i) != -1) {
                var targetText = startInsertionPoint;
                if (textRange != false) {
                    if (isRIE) {
                        if (textRange == null || textRange.contents == "") {
                            if (removeStyle != "") {
                                newStyleName = removeStyle;
                            }
                            else {
                                var thisPara = startInsertionPoint.paragraphs.item (0);
                                var previousPara = null;
                                if (thisPara.characters.item (-1).contents != "\r") {
                                    previousPara = thisPara.parent.paragraphs.previousItem(thisPara);
                                }
                                if (previousPara != null) {
                                    thisPara.parent.characters.itemByRange (previousPara.characters.item (-1), thisPara.characters.item (-1)).remove ();
                                }
                                else {
                                    thisPara.remove ();
                                }
                                return false;
                            }
                        }
                    }
                    if (textRange != null) {
                        targetText = textRange;
                    }
                }
                if (newStyleName) {
                    var newStyle = null;
                    if (dynamicStyleMethod.search (/Code/i) != -1 || newStyleName.search (/\/\/Code/i) == 0 || newStyleName.search (/\/\*code\*\//i) == 0) {
                        newStyle = newStyleName;
                    }
                    else if (newStyleName) {
                        newStyle = targetDocument.paragraphStyles.itemByName (newStyleName);
                        if (newStyle == null) { // Create new style
                            newStyle = targetDocument.paragraphStyles.add({name: newStyleName});
                            if (newBasedOnStyleName != "") {
                                var newBasedOnStyle = targetDocument.paragraphStyles.itemByName (newBasedOnStyleName);
                                if (newBasedOnStyle == null) {
                                    newBasedOnStyle = targetDocument.paragraphStyles.add({name: newBasedOnStyleName});
                                }
                                if (newStyle.name != newBasedOnStyle.name) {
                                    newStyle.basedOn = newBasedOnStyle;
                                }
                            }
                            var expectedCellStyle = targetDocument.cellStyles.itemByName (newStyleName);
                            if (expectedCellStyle != null) {
                                expectedCellStyle.appliedParagraphStyle = newStyle;
                            }
                            var expectedObjectStyle = targetDocument.objectStyles.itemByName (newStyleName);
                            if (expectedObjectStyle != null) {
                                expectedObjectStyle.appliedParagraphStyle = newStyle;
                            }
                        }    
                    }
                    if (newStyle != null) {
                        var targetIndex = 0;
                        var isPositive = true;
                        var isFromLast = false;
                        var isEvenOnly = null;
                        var appliedCounter = 0;
                        if (dynamicStyleMethod.search (/For[+-]?\d+[A-Z]?/i) != -1) {
                            var theNum = dynamicStyleMethod.match (/For[+-]?\d+[A-Z]?/i);
                            if (theNum != null && theNum.length > 0) {
                                theNum = theNum[0].slice (3);
                                if (theNum[0] == '+' || theNum[0] == '-') {
                                    if (theNum[0] == '+') {
                                        isPositive = true;
                                    }
                                    else {
                                        isPositive = false;
                                    }
                                    theNum = theNum.slice (1);
                                }
                                else {
                                    isPositive = null;
                                }
                                targetIndex = parseInt (theNum, 10) - 1;
                                if (theNum.search (/L/i) != -1) {
                                    isFromLast = true;
                                }
                                if (theNum.search (/E/i) != -1) {
                                    isEvenOnly = true;
                                }
                                else if (theNum.search (/O/i) != -1) {
                                    isEvenOnly = false;
                                }
                            }
                        }
                        if (isFromLast) {
                            targetIndex = targetText.paragraphs.length - targetIndex - 1;
                        }
                        var parasArr = [];
                        for (var trs = 0; trs < targetText.paragraphs.length; trs++) {
                            if (targetIndex != null) {
                                if (isPositive == true && trs < targetIndex) {
                                    continue;
                                }
                                else if (isPositive == false && trs > targetIndex) {
                                    continue;
                                }
                                else if (isPositive == null && trs != targetIndex) {
                                    continue;
                                }
                            }
                            appliedCounter++;
                            if (isEvenOnly == true && (appliedCounter % 2 == 1)) {
                                continue;
                            }
                            if (isEvenOnly == false && (appliedCounter % 2 == 0)) {
                                continue;
                            }
                            if (targetText.paragraphs[trs].appliedParagraphStyle) {
                                if (targetText.paragraphs[trs].appliedParagraphStyle.constructor.name === "Array") {
                                    if (targetText.paragraphs[trs].appliedParagraphStyle.length > 0) {
                                        if (targetText.paragraphs[trs].appliedParagraphStyle[0].name) {
                                            if (targetText.paragraphs[trs].appliedParagraphStyle[0].name.search(/\^$/) != -1) {
                                                var isToContinue = true;
                                                if (newStyleName.search(/\^$/) != -1) {
                                                    var toBeAppliedCount = newStyleName.match (/\^+$/)[0].length;
                                                    var destinationCount = targetText.paragraphs[trs].appliedParagraphStyle[0].name.match (/\^+$/)[0].length;
                                                    if (toBeAppliedCount > destinationCount) {
                                                        isToContinue = false;
                                                    }
                                                }
                                                if (isToContinue) {
                                                    continue;
                                                }
                                            }
                                        }
                                    }
                                } else if (targetText.paragraphs[trs].appliedParagraphStyle.name) {
                                    if (targetText.paragraphs[trs].appliedParagraphStyle.name.search(/\^$/) != -1) {
                                        var isToContinue = true;
                                        if (newStyleName.search(/\^$/) != -1) {
                                            var toBeAppliedCount = newStyleName.match (/\^+$/)[0].length;
                                            var destinationCount = targetText.paragraphs[trs].appliedParagraphStyle.name.match (/\^+$/)[0].length;
                                            if (toBeAppliedCount > destinationCount) {
                                                isToContinue = false;
                                            }
                                        }
                                        if (isToContinue) {
                                            continue;
                                        }
                                    }
                                }
                            }
                            if (dynamicStyleMethod.search (/Code/i) != -1 || newStyleName.search (/\/\/Code/i) == 0 || newStyleName.search (/\/\*code\*\//i) == 0) {
                                parasArr.push (targetText.paragraphs[trs]);
                            }
                            else {
                                targetText.paragraphs[trs].appliedParagraphStyle = newStyle;
                            }
                        }
                        if (parasArr.length > 0) {
                            var wrappedCode =
                                "var paras = context[0];\n" +
                                "var doc = context[1];\n" +
                                "var file = context[2];\n" +
                                "var level = context[3];\n" +
                                "var frame = null;\n" +
                                newStyleName;
                            var customFun = new Function("context", wrappedCode);
                            var codeFile = null;
                            if (dynamicStyleFiles instanceof Array) {
                                if (dynamicStyleFiles.length > 0) {
                                    codeFile = dynamicStyleFiles[0];
                                }
                            }
                            var fileLevel = 0;
                            if (liveSnippetID.lastIndexOf ("/") > 0) {
                                fileLevel = liveSnippetID.split ("/").length - 1;
                            }
                            app.doScript (customFun, ScriptLanguage.JAVASCRIPT, [parasArr, targetDocument, codeFile, fileLevel], UndoModes.ENTIRE_SCRIPT, "User Code");
    
                        }
                    }
                }
            }
            if (dynamicStyleMethod.search (/Char/i) != -1) {
                var targetText = startInsertionPoint;
                if (textRange != false && textRange != null) {
                    targetText = textRange;
                }
                var newStyle = targetDocument.characterStyles.itemByName (newStyleName);
                if (newStyle == null) { // Create new style
                    newStyle = targetDocument.characterStyles.add({name: newStyleName});
                    if (newBasedOnStyleName != "") {
                        var newBasedOnStyle = targetDocument.characterStyles.itemByName (newBasedOnStyleName);
                        if (newBasedOnStyle == null) {
                            newBasedOnStyle = targetDocument.characterStyles.add({name: newBasedOnStyleName});
                        }
                        if (newStyle.name != newBasedOnStyle.name) {
                            newStyle.basedOn = newBasedOnStyle;
                        }
                    }
                }
                if (newStyle != null) {
                    targetText.appliedCharacterStyle = newStyle;
                }
            }
            if (dynamicStyleMethod.search (/Obj/i) != -1) {
                if (isRIE && textRange != false) {
                    if (textRange == null || textRange.contents == "") {
                        if (removeStyle != "") {
                            newStyleName = removeStyle;
                        }
                        else {
                            if (dynamicStyleMethod.search (/Location/i) != -1) {
                                var parentFrame = startInsertionPoint.parentTextFrames[0];
                                var theParent = parentFrame.parent;
                                do {
                                    if (theParent.constructor.name == "Character") {
                                        if (theParent.paragraphs.length > 0) {
                                            var toRemovePara = theParent.paragraphs.item(0);
                                            var previousPara = null;
                                            if (toRemovePara.characters.item (-1).contents != "\r") {
                                                previousPara = toRemovePara.parent.paragraphs.previousItem(toRemovePara);
                                            }
                                            toRemovePara.remove ();
                                            if (previousPara != null) {
                                                previousPara.characters.item (-1).remove ();
                                            }
                                            return false;
                                        }
                                    }
                                    theParent = theParent.parent;
                                } while (theParent.constructor.name != "Spread" && theParent.constructor.name != "MasterSpread" && theParent.constructor.name != "Page" && theParent.constructor.name != "Document");
                            }
                            parentFrame.remove ();
                            return false;
                        }
                    }
                }
                var newStyle = null;
                if (dynamicStyleMethod.search (/Code/i) != -1 || newStyleName.search (/\/\/Code/i) == 0 || newStyleName.search (/\/\*code\*\//i) == 0) {
                    newStyle = newStyleName;
                }
                else if (newStyleName) {
                    newStyle = targetDocument.objectStyles.itemByName (newStyleName);
                    if (newStyle == null) { // Create new style
                        newStyle = targetDocument.objectStyles.add({name: newStyleName});
                        if (newBasedOnStyleName != "") {
                            var newBasedOnStyle = targetDocument.objectStyles.itemByName (newBasedOnStyleName);
                            if (newBasedOnStyle == null) {
                                newBasedOnStyle = targetDocument.objectStyles.add({name: newBasedOnStyleName});
                            }
                            if (newStyle.name != newBasedOnStyle.name) {
                                newStyle.basedOn = newBasedOnStyle;
                            }
                        }
                        var expectedParagraphStyle = targetDocument.paragraphStyles.itemByName (newStyleName);
                        if (expectedParagraphStyle != null) {
                            newStyle.appliedParagraphStyle = expectedParagraphStyle;
                        }
                    }
                }
                if (textRange != false) {
                    if (newStyle != null && textRange != null) {
                        var targetIndex = null;
                        var isPositive = null;
                        var isFromLast = false;
                        var isEvenOnly = null;
                        var appliedCounter = 0;
                        var isOnlyFrames = null;
                        if (dynamicStyleMethod.search (/For[+-]?\d+[A-Z]?/i) != -1) {
                            var theNum = dynamicStyleMethod.match (/For[+-]?\d+[A-Z]?/i);
                            if (theNum != null && theNum.length > 0) {
                                theNum = theNum[0].slice (3);
                                if (theNum[0] == '+' || theNum[0] == '-') {
                                    if (theNum[0] == '+') {
                                        isPositive = true;
                                    }
                                    else {
                                        isPositive = false;
                                    }
                                    theNum = theNum.slice (1);
                                }
                                targetIndex = parseInt (theNum, 10) - 1;
                                if (theNum.search (/L/i) != -1) {
                                    isFromLast = true;
                                }
                                if (theNum.search (/T/i) != -1) {
                                    isOnlyFrames = true;
                                }
                                else if (theNum.search (/P/i) != -1) {
                                    isOnlyFrames = false;
                                }
                                if (theNum.search (/E/i) != -1) {
                                    isEvenOnly = true;
                                }
                                else if (theNum.search (/O/i) != -1) {
                                    isEvenOnly = false;
                                }
                            }
                        }
                        var textItemsPairs = sortTextItems (textRange, pageItemsIDs, isOnlyFrames);
                        if (isFromLast) {
                            targetIndex = textItemsPairs.length - targetIndex - 1;
                        }
                        for (var tris = 0; tris < textItemsPairs.length; tris++) {
                            if (targetIndex != null) {
                                if (isPositive == true && tris < targetIndex) {
                                    continue;
                                }
                                else if (isPositive == false && tris > targetIndex) {
                                    continue;
                                }
                                else if (isPositive == null && tris != targetIndex) {
                                    continue;
                                }
                            }
                            appliedCounter++;
                            if (isEvenOnly == true && (appliedCounter % 2 == 1)) {
                                continue;
                            }
                            else if (isEvenOnly == false && (appliedCounter % 2 == 0)) {
                                continue;
                            }
                            if (dynamicStyleMethod.search (/Code/i) != -1 || newStyleName.search (/\/\/Code/i) == 0 || newStyleName.search (/\/\*code\*\//i) == 0) {
                                var wrappedCode =
                                    "var frame = context[0];\n" +
                                    "var doc = context[1];\n" +
                                    "var file = context[2];\n" +
                                    "var level = context[3];\n" +
                                    "var paras = null;\n" +
                                    newStyleName;
                                var customFun = new Function("context", wrappedCode);
                                var codeFile = null;
                                if (dynamicStyleFiles instanceof Array) {
                                    if (dynamicStyleFiles.length > 0) {
                                        codeFile = dynamicStyleFiles[0];
                                    }
                                }
                                var fileLevel = 0;
                                if (liveSnippetID.lastIndexOf ("/") > 0) {
                                    fileLevel = liveSnippetID.split ("/").length - 1;
                                }
                                app.doScript (customFun, ScriptLanguage.JAVASCRIPT, [textItemsPairs[tris][0], targetDocument, codeFile, fileLevel], UndoModes.ENTIRE_SCRIPT, "User Code");
                            }
                            else {
                                textItemsPairs[tris][0].appliedObjectStyle = newStyle;
                            }
                        }
                    }
                }
                else if (newStyle != null) {
                    var parentFrame = null;
                    if (startInsertionPoint.parentTextFrames && startInsertionPoint.parentTextFrames.length > 0) {
                        parentFrame = startInsertionPoint.parentTextFrames[0];
                    }
                    if (parentFrame != null) {
                        if (dynamicStyleMethod.search (/Code/i) != -1) {
                            var wrappedCode =
                                "var frame = context[0];\n" +
                                "var doc = context[1];\n" +
                                "var file = context[2];\n" +
                                "var level = context[3];\n" +
                                "var paras = null;\n" +
                                newStyle;
                            var customFun = new Function("context", wrappedCode);
                            var codeFile = null;
                            if (dynamicStyleFiles instanceof Array) {
                                if (dynamicStyleFiles.length > 0) {
                                    codeFile = dynamicStyleFiles[0];
                                }
                            }
                            var fileLevel = 0;
                            if (liveSnippetID.lastIndexOf ("/") > 0) {
                                fileLevel = liveSnippetID.split ("/").length - 1;
                            }
                            app.doScript (customFun, ScriptLanguage.JAVASCRIPT, [parentFrame, targetDocument, codeFile, fileLevel], UndoModes.ENTIRE_SCRIPT, "User Code");
                        }
                        else {
                            parentFrame.appliedObjectStyle = newStyle;
                        }
                    }
                }
            }
            if (dynamicStyleMethod.search (/Table/i) != -1) {
                if (textRange != false && isRIE) {
                    if (textRange == null || textRange.contents == "") {
                        if (removeStyle != "") {
                            newStyleName = removeStyle;
                        }
                        else {
                            startInsertionPoint.parent.parentRow.parent.remove ();
                            return false;
                        }
                    }
                }
                if (newStyleName) {
                    var newStyle = targetDocument.tableStyles.itemByName (newStyleName);
                    if (newStyle == null) { // Create new style
                        newStyle = targetDocument.tableStyles.add({name: newStyleName});
                        if (newBasedOnStyleName != "") {
                            var newBasedOnStyle = targetDocument.tableStyles.itemByName (newBasedOnStyleName);
                            if (newBasedOnStyle == null) {
                                newBasedOnStyle = targetDocument.tableStyles.add({name: newBasedOnStyleName});
                            }
                            if (newStyle.name != newBasedOnStyle.name) {
                                newStyle.basedOn = newBasedOnStyle;
                            }
                        }
                        var bodyCellStyleName = newStyleName + dynamicStyleSeparator + "Body";
                        var expectedBodyCellStyle = targetDocument.cellStyles.itemByName (bodyCellStyleName);
                        if (expectedBodyCellStyle == null) { // Create new style
                            expectedBodyCellStyle = targetDocument.cellStyles.add({name: bodyCellStyleName});
                            if (newBasedOnStyleName != "") {
                                var newBasedOnStyle = targetDocument.cellStyles.itemByName (newBasedOnStyleName + dynamicStyleSeparator + "Body");
                                if (newBasedOnStyle != null) {
                                    expectedBodyCellStyle.basedOn = newBasedOnStyle;
                                }
                            }
                            var expectedParaStyle = targetDocument.paragraphStyles.itemByName (bodyCellStyleName);
                            if (expectedParaStyle == null) {
                                expectedParaStyle = targetDocument.paragraphStyles.add({name: bodyCellStyleName});
                                if (newBasedOnStyleName != "") {
                                    var newBasedOnStyle = targetDocument.paragraphStyles.itemByName (newBasedOnStyleName + dynamicStyleSeparator + "Body");
                                    if (newBasedOnStyle != null) {
                                        expectedParaStyle.basedOn = newBasedOnStyle;
                                    }
                                }
                            }
                            if (expectedParaStyle != null) {
                                expectedBodyCellStyle.appliedParagraphStyle = expectedParaStyle;
                            }
                        }
                        if (expectedBodyCellStyle != null) {
                            newStyle.bodyRegionCellStyle = expectedBodyCellStyle;
                        }
                        var headerCellStyleName = newStyleName + dynamicStyleSeparator + "Header";
                        var expectedHeaderCellStyle = targetDocument.cellStyles.itemByName (headerCellStyleName);
                        if (expectedHeaderCellStyle == null) { // Create new style
                            expectedHeaderCellStyle = targetDocument.cellStyles.add({name: headerCellStyleName});
                            if (newBasedOnStyleName != "") {
                                var newBasedOnStyle = targetDocument.cellStyles.itemByName (newBasedOnStyleName + dynamicStyleSeparator + "Header");
                                if (newBasedOnStyle != null) {
                                    expectedHeaderCellStyle.basedOn = newBasedOnStyle;
                                }
                            }
                            var expectedParaStyle = targetDocument.paragraphStyles.itemByName (headerCellStyleName);
                            if (expectedParaStyle == null) {
                                expectedParaStyle = targetDocument.paragraphStyles.add({name: headerCellStyleName});
                                if (newBasedOnStyleName != "") {
                                    var newBasedOnStyle = targetDocument.paragraphStyles.itemByName (newBasedOnStyleName + dynamicStyleSeparator + "Header");
                                    if (newBasedOnStyle != null) {
                                        expectedParaStyle.basedOn = newBasedOnStyle;
                                    }
                                }
                            }
                            if (expectedParaStyle != null) {
                                expectedHeaderCellStyle.appliedParagraphStyle = expectedParaStyle;
                            }
                        }
                        if (expectedHeaderCellStyle != null) {
                            newStyle.headerRegionCellStyle = expectedHeaderCellStyle;
                        }
                        var footerCellStyleName = newStyleName + dynamicStyleSeparator + "Footer";
                        var expectedFooterCellStyle = targetDocument.cellStyles.itemByName (footerCellStyleName);
                        if (expectedFooterCellStyle == null) { // Create new style
                            expectedFooterCellStyle = targetDocument.cellStyles.add({name: footerCellStyleName});
                            if (newBasedOnStyleName != "") {
                                var newBasedOnStyle = targetDocument.cellStyles.itemByName (newBasedOnStyleName + dynamicStyleSeparator + "Footer");
                                if (newBasedOnStyle != null) {
                                    expectedFooterCellStyle.basedOn = newBasedOnStyle;
                                }
                            }
                            var expectedParaStyle = targetDocument.paragraphStyles.itemByName (footerCellStyleName);
                            if (expectedParaStyle == null) {
                                expectedParaStyle = targetDocument.paragraphStyles.add({name: footerCellStyleName});
                                if (newBasedOnStyleName != "") {
                                    var newBasedOnStyle = targetDocument.paragraphStyles.itemByName (newBasedOnStyleName + dynamicStyleSeparator + "Footer");
                                    if (newBasedOnStyle != null) {
                                        expectedParaStyle.basedOn = newBasedOnStyle;
                                    }
                                }
                            }
                            if (expectedParaStyle != null) {
                                expectedFooterCellStyle.appliedParagraphStyle = expectedParaStyle;
                            }
                        }
                        if (expectedFooterCellStyle != null) {
                            newStyle.footerRegionCellStyle = expectedFooterCellStyle;
                        }
                        var leftCellStyleName = newStyleName + dynamicStyleSeparator + "Left";
                        var expectedLeftCellStyle = targetDocument.cellStyles.itemByName (leftCellStyleName);
                        if (expectedLeftCellStyle == null) { // Create new style
                            expectedLeftCellStyle = targetDocument.cellStyles.add({name: leftCellStyleName});
                            if (newBasedOnStyleName != "") {
                                var newBasedOnStyle = targetDocument.cellStyles.itemByName (newBasedOnStyleName + dynamicStyleSeparator + "Left");
                                if (newBasedOnStyle != null) {
                                    expectedLeftCellStyle.basedOn = newBasedOnStyle;
                                }
                            }
                            var expectedParaStyle = targetDocument.paragraphStyles.itemByName (leftCellStyleName);
                            if (expectedParaStyle == null) {
                                expectedParaStyle = targetDocument.paragraphStyles.add({name: leftCellStyleName});
                                if (newBasedOnStyleName != "") {
                                    var newBasedOnStyle = targetDocument.paragraphStyles.itemByName (newBasedOnStyleName + dynamicStyleSeparator + "Left");
                                    if (newBasedOnStyle != null) {
                                        expectedParaStyle.basedOn = newBasedOnStyle;
                                    }
                                }
                            }
                            if (expectedParaStyle != null) {
                                expectedLeftCellStyle.appliedParagraphStyle = expectedParaStyle;
                            }
                        }
                        if (expectedLeftCellStyle != null) {
                            newStyle.leftColumnRegionCellStyle = expectedLeftCellStyle;
                        }
                        var rightCellStyleName = newStyleName + dynamicStyleSeparator + "Right";
                        var expectedRightCellStyle = targetDocument.cellStyles.itemByName (rightCellStyleName);
                        if (expectedRightCellStyle == null) { // Create new style
                            expectedRightCellStyle = targetDocument.cellStyles.add({name: rightCellStyleName});
                            if (newBasedOnStyleName != "") {
                                var newBasedOnStyle = targetDocument.cellStyles.itemByName (newBasedOnStyleName + dynamicStyleSeparator + "Right");
                                if (newBasedOnStyle != null) {
                                    expectedRightCellStyle.basedOn = newBasedOnStyle;
                                }
                            }
                            var expectedParaStyle = targetDocument.paragraphStyles.itemByName (rightCellStyleName);
                            if (expectedParaStyle == null) {
                                expectedParaStyle = targetDocument.paragraphStyles.add({name: rightCellStyleName});
                                if (newBasedOnStyleName != "") {
                                    var newBasedOnStyle = targetDocument.paragraphStyles.itemByName (newBasedOnStyleName + dynamicStyleSeparator + "Right");
                                    if (newBasedOnStyle != null) {
                                        expectedParaStyle.basedOn = newBasedOnStyle;
                                    }
                                }
                            }
                            if (expectedParaStyle != null) {
                                expectedRightCellStyle.appliedParagraphStyle = expectedParaStyle;
                            }
                        }
                        if (expectedRightCellStyle != null) {
                            newStyle.rightColumnRegionCellStyle = expectedRightCellStyle;
                        }
                    }
                    if (newStyle != null) {
                        if (textRange != false) {
                            if (textRange != null) {
                                for (var trt = 0; trt < textRange.tables.length; trt++) {
                                    textRange.tables[trt].appliedTableStyle = newStyle;
                                    textRange.tables[trt].cells.everyItem ().clearCellStyleOverrides (true);
                                    textRange.tables[trt].cells.everyItem ().paragraphs.everyItem ().clearOverrides (OverrideType.ALL);
                                }
                            }
                        }
                        else {
                            startInsertionPoint.parent.parentRow.parent.appliedTableStyle = newStyle;
                            startInsertionPoint.parent.parentRow.parent.cells.everyItem ().clearCellStyleOverrides (true);
                            startInsertionPoint.parent.parentRow.parent.cells.everyItem ().paragraphs.everyItem ().clearOverrides (OverrideType.ALL);
                        }
                    }
                }
            }
            if (newStyleName == "Header" && dynamicStyleMethod.search (/Row/i) != -1) {
                if (startInsertionPoint.parent.constructor.name == "Cell") {
                    var fatCell = startInsertionPoint.parent.rows.item (0).cells.item (0);
                    for (var cfr = 1; cfr < fatCell.rows.item (0).cells.length; cfr++) {
                        if (fatCell.rows.item (0).cells.item (cfr).rows.length > fatCell.rows.length) {
                            fatCell = fatCell.rows.item (0).cells.item (cfr);
                        }
                    }
                    try {
                        var parentTable = fatCell;
                        while (parentTable != null && parentTable.constructor.name != "Table") {
                            parentTable = parentTable.parent;
                        }
                        parentTable.rows.itemByRange (0, fatCell.rows.item (-1).index).rowType = RowTypes.HEADER_ROW;
                    }
                    catch (e) {
    
                    }
                }
            }
            else if (newStyleName == "Footer" && dynamicStyleMethod.search (/Row/i) != -1) {
                if (startInsertionPoint.parent.constructor.name == "Cell") {
                    var fatCell = startInsertionPoint.parent.rows.item (-1).cells.item (0);
                    for (var cfr = 1; cfr < fatCell.rows.item (-1).cells.length; cfr++) {
                        if (fatCell.rows.item (-1).cells.item (cfr).rows.length > fatCell.rows.length) {
                            fatCell = fatCell.rows.item (-1).cells.item (cfr);
                        }
                    }
                    try {
                        var parentTable = fatCell;
                        while (parentTable != null && parentTable.constructor.name != "Table") {
                            parentTable = parentTable.parent;
                        }
                        parentTable.rows.itemByRange (fatCell.rows.item (0).index, fatCell.rows.item (-1).index).rowType = RowTypes.FOOTER_ROW;
                    }
                    catch (e) {
                        
                    }
                }
            }
            else if (dynamicStyleMethod.search (/Cell/i) != -1 || dynamicStyleMethod.search (/Row/i) != -1 || dynamicStyleMethod.search (/Column/i) != -1) {
                var newStyle = targetDocument.cellStyles.itemByName (newStyleName);
                var styleNameSuffix = dynamicStyleSeparator + "Body";
                var isTableDefaultStyle = false;
                if (newStyle == null) { // Create new style
                    newStyle = targetDocument.cellStyles.add({name: newStyleName});
                    if (newBasedOnStyleName != "") {
                        var newBasedOnStyle = targetDocument.cellStyles.itemByName (newBasedOnStyleName);
                        if (newBasedOnStyle == null) {
                            newBasedOnStyle = targetDocument.cellStyles.add({name: newBasedOnStyleName});
                        }
                        if (newStyle.name != newBasedOnStyle.name) {
                            newStyle.basedOn = newBasedOnStyle;
                        }
                    }
                    var expectedParaStyle = targetDocument.paragraphStyles.itemByName (newStyleName);
                    if (expectedParaStyle != null) {
                        expectedCellStyle.appliedParagraphStyle = expectedParaStyle;
                    }
                    if (newStyleName.indexOf (styleNameSuffix) > 0 && newStyleName.indexOf (styleNameSuffix) == (newStyleName.length - styleNameSuffix.length)) {
                        var expectedTableStyle = targetDocument.tableStyles.itemByName (newStyleName.slice (0, newStyleName.indexOf (styleNameSuffix)));
                        if (expectedTableStyle != null) {
                            expectedTableStyle.bodyRegionCellStyle = newStyle;
                            isTableDefaultStyle = true;
                        } 
                    }
                    styleNameSuffix = dynamicStyleSeparator + "Header";
                    if (newStyleName.indexOf (styleNameSuffix) > 0 && newStyleName.indexOf (styleNameSuffix) == (newStyleName.length - styleNameSuffix.length)) {
                        var expectedTableStyle = targetDocument.tableStyles.itemByName (newStyleName.slice (0, newStyleName.indexOf (styleNameSuffix)));
                        if (expectedTableStyle != null) {
                            expectedTableStyle.headerRegionCellStyle = newStyle;
                            isTableDefaultStyle = true;
                        } 
                    }
                    styleNameSuffix = dynamicStyleSeparator + "Footer";
                    if (newStyleName.indexOf (styleNameSuffix) > 0 && newStyleName.indexOf (styleNameSuffix) == (newStyleName.length - styleNameSuffix.length)) {
                        var expectedTableStyle = targetDocument.tableStyles.itemByName (newStyleName.slice (0, newStyleName.indexOf (styleNameSuffix)));
                        if (expectedTableStyle != null) {
                            expectedTableStyle.footerRegionCellStyle = newStyle;
                            isTableDefaultStyle = true;
                        } 
                    }
                    styleNameSuffix = dynamicStyleSeparator + "Left";
                    if (newStyleName.indexOf (styleNameSuffix) > 0 && newStyleName.indexOf (styleNameSuffix) == (newStyleName.length - styleNameSuffix.length)) {
                        var expectedTableStyle = targetDocument.tableStyles.itemByName (newStyleName.slice (0, newStyleName.indexOf (styleNameSuffix)));
                        if (expectedTableStyle != null) {
                            expectedTableStyle.leftColumnRegionCellStyle = newStyle;
                            isTableDefaultStyle = true;
                        } 
                    }
                    styleNameSuffix = dynamicStyleSeparator + "Right";
                    if (newStyleName.indexOf (styleNameSuffix) > 0 && newStyleName.indexOf (styleNameSuffix) == (newStyleName.length - styleNameSuffix.length)) {
                        var expectedTableStyle = targetDocument.tableStyles.itemByName (newStyleName.slice (0, newStyleName.indexOf (styleNameSuffix)));
                        if (expectedTableStyle != null) {
                            expectedTableStyle.rightColumnRegionCellStyle = newStyle;
                            isTableDefaultStyle = true;
                        } 
                    }
                }
                if (startInsertionPoint.parent.constructor.name == "Cell") {
                    styleNameSuffix = dynamicStyleSeparator + "Header";
                    if (newStyleName.indexOf (styleNameSuffix) > 0 && newStyleName.indexOf (styleNameSuffix) == (newStyleName.length - styleNameSuffix.length)) {
                        for (var thr = 0; thr <= startInsertionPoint.parent.parentRow.index; thr++) {
                            if (startInsertionPoint.parent.parentRow.parent.rows.item(thr).rowType != RowTypes.HEADER_ROW) {
                                if (thr == startInsertionPoint.parent.parentRow.parent.rows.length - 1) {
                                    startInsertionPoint.parent.parentRow.parent.rows.add (LocationOptions.AT_END);
                                }
                                startInsertionPoint.parent.parentRow.parent.rows.item(thr).rowType = RowTypes.HEADER_ROW;
                            }
                        }
                    }
                    styleNameSuffix = dynamicStyleSeparator + "Footer";
                    if (newStyleName.indexOf (styleNameSuffix) > 0 && newStyleName.indexOf (styleNameSuffix) == (newStyleName.length - styleNameSuffix.length)) {
                        for (var tfr = startInsertionPoint.parent.parentRow.parent.rows.length - 1; tfr >= startInsertionPoint.parent.parentRow.index; tfr--) {
                            if (startInsertionPoint.parent.parentRow.parent.rows.item(tfr).rowType != RowTypes.FOOTER_ROW) {
                                startInsertionPoint.parent.parentRow.parent.rows.item(tfr).rowType = RowTypes.FOOTER_ROW;
                            }
                        }
                    }
                }
                if (dynamicStyleMethod.search (/Cell/i) != -1) {
                    if (startInsertionPoint.parent.constructor.name == "Cell") {
                        if (isRIE && textRange != false) {
                            if (textRange == null || textRange.contents == "") {
                                dynamicStyleMethod = dynamicStyleMethod.replace ("Cell", "Merge");
                            }
                        }
                        if (newStyle != null && !isTableDefaultStyle) {
                            startInsertionPoint.parent.appliedCellStyle = newStyle;
                        }
                    }
                }
                else if (dynamicStyleMethod.search (/Row/i) != -1) {
                    if (startInsertionPoint.parent.constructor.name == "Cell") {
                        if (isRIE && textRange != false) {
                            if (textRange == null || textRange.contents == "") {
                                if (removeStyle != "") {
                                    newStyle = targetDocument.cellStyles.itemByName (removeStyle);
                                }
                                else {
                                    startInsertionPoint.parent.parentRow.remove ();
                                    return false;
                                }
                            }
                        }
                        if (newStyle != null && !isTableDefaultStyle) {
                            var theRow = startInsertionPoint.parent.parentRow;
                            theRow.cells.everyItem ().appliedCellStyle = newStyle;
                        }
                    }
                }
                else if (dynamicStyleMethod.search (/Column/i) != -1) {
                    if (startInsertionPoint.parent.constructor.name == "Cell") {
                        if (isRIE && textRange != false) {
                            if (textRange == null || textRange.contents == "") {
                                if (removeStyle != "") {
                                    newStyle = targetDocument.cellStyles.itemByName (removeStyle);
                                }
                                else {
                                    startInsertionPoint.parent.parentColumn.remove ();
                                    return false;
                                }
                            }
                        }
                        if (newStyle != null && !isTableDefaultStyle) {
                            var theColumn = startInsertionPoint.parent.parentColumn;
                            theColumn.cells.everyItem ().appliedCellStyle = newStyle;
                        }
                    }
                }
            }
        }       
        if (dynamicStyleMethod.search (/Font/i) != -1) {
            if (textRange != false && textRange != null) {
                if (textRange.contents != "") {
                    var theFont = newStyleName;
                    var theFontStyle = "";
                    if (newStyleName.indexOf (", ") != -1) {
                        var newStyleNameSplitted = newStyleName.split (", ");
                        theFont = newStyleNameSplitted[0];
                        theFontStyle = newStyleNameSplitted[1];
                    }
                    else if (newStyleName.indexOf (" | ") != -1) {
                        var newStyleNameSplitted = newStyleName.split (" | ");
                        theFont = newStyleNameSplitted[0];
                        theFontStyle = newStyleNameSplitted[1];
                    }
                    try {
                        textRange.appliedFont = theFont;
                    }
                    catch (err) {

                    }
                    if (theFontStyle != "") {
                        try {
                            textRange.fontStyle = theFontStyle;
                        }
                        catch (err) {
    
                        }
                    }
                }
            }  
        }
        if (dynamicStyleMethod.search (/Style/i) != -1) {
            if (textRange != false && textRange != null) {
                if (textRange.contents != "") {
                    try {
                        textRange.fontStyle = newStyleName;
                    }
                    catch (err) {

                    }
                }
            }  
        }
        if (dynamicStyleMethod.search (/Size/i) != -1) {
            if (textRange != false && textRange != null) {
                if (textRange.contents != "") {
                    var isPercent = false;
                    if (newStyleName.indexOf ("%") != -1) {
                        newStyleName = newStyleName.replace ("%", "");
                        newStyleName = parseFloat (newStyleName);
                        isPercent = true;
                    }
                    if (isPercent) {
                        try {
                            textRange.pointSize *= newStyleName / 100;
                        }
                        catch (err) {
    
                        }
                    }
                    else {
                        try {
                            textRange.pointSize = newStyleName;
                        }
                        catch (err) {
    
                        }
                    }
                }
            }  
        }
        if (dynamicStyleMethod.search (/Color/i) != -1) {
            if (textRange != false && textRange != null) {
                if (textRange.contents != "") {
                    try {
                        textRange.fillColor = newStyleName;
                    }
                    catch (err) {

                    }
                }
            }  
        }
        if (dynamicStyleMethod.search (/Justification/i) != -1) {
            if (textRange != false && textRange != null) {
                if (textRange.contents != "") {
                    newStyleName = newStyleName.toUpperCase();
                    switch(newStyleName) {
                        case "LEFT_ALIGN":
                            textRange.justification = Justification.LEFT_ALIGN;
                            break;
                        case "CENTER_ALIGN":
                            textRange.justification = Justification.CENTER_ALIGN;
                            break;
                        case "RIGHT_ALIGN":
                            textRange.justification = Justification.RIGHT_ALIGN;
                            break;
                        case "LEFT_JUSTIFIED":
                            textRange.justification = Justification.LEFT_JUSTIFIED;
                            break;
                        case "RIGHT_JUSTIFIED":
                            textRange.justification = Justification.RIGHT_JUSTIFIED;
                            break;
                        case "CENTER_JUSTIFIED":
                            textRange.justification = Justification.CENTER_JUSTIFIED;
                            break;
                        case "FULLY_JUSTIFIED":
                            textRange.justification = Justification.FULLY_JUSTIFIED;
                            break;
                        case "TO_BINDING_SIDE":
                            textRange.justification = Justification.TO_BINDING_SIDE;
                            break;
                        case "AWAY_FROM_BINDING_SIDE":
                            textRange.justification = Justification.AWAY_FROM_BINDING_SIDE;
                            break;
                    }
                }
            }  
        }
        if (dynamicStyleMethod.search (/characterDirection/i) != -1) {
            if (textRange != false && textRange != null) {
                if (textRange.contents != "") {
                    newStyleName = newStyleName.toUpperCase();
                    switch(newStyleName) {
                        case "DEFAULT_DIRECTION":
                            textRange.characterDirection = CharacterDirectionOptions.DEFAULT_DIRECTION;
                            break;
                        case "LEFT_TO_RIGHT_DIRECTION":
                            textRange.characterDirection = CharacterDirectionOptions.LEFT_TO_RIGHT_DIRECTION;
                            break;
                        case "RIGHT_TO_LEFT_DIRECTION":
                            textRange.characterDirection = CharacterDirectionOptions.RIGHT_TO_LEFT_DIRECTION;
                            break;
                    }
                }
            }  
        }
        if (dynamicStyleMethod.search (/paragraphDirection/i) != -1) {
            if (textRange != false && textRange != null) {
                if (textRange.contents != "") {
                    newStyleName = newStyleName.toUpperCase();
                    switch(newStyleName) {
                        case "LEFT_TO_RIGHT_DIRECTION":
                            textRange.paragraphDirection = ParagraphDirectionOptions.LEFT_TO_RIGHT_DIRECTION;
                            break;
                        case "RIGHT_TO_LEFT_DIRECTION":
                            textRange.paragraphDirection = ParagraphDirectionOptions.RIGHT_TO_LEFT_DIRECTION;
                            break;
                    }
                }
            }  
        }
        if (dynamicStyleMethod.search (/Rotate/i) != -1 || dynamicStyleMethod.search (/Shear/i) != -1) {
            if (textRange != false) {
                if (textRange != null) {
                    var newValue = parseFloat (newStyleName);
                    if (newValue) {
                        var targetIndex = null;
                        var isPositive = null;
                        var isFromLast = false;
                        var isEvenOnly = null;
                        var appliedCounter = 0;
                        var isOnlyFrames = null;
                        if (dynamicStyleMethod.search (/For[+-]?\d+[A-Z]?/i) != -1) {
                            var theNum = dynamicStyleMethod.match (/For[+-]?\d+[A-Z]?/i);
                            if (theNum != null && theNum.length > 0) {
                                theNum = theNum[0].slice (3);
                                if (theNum[0] == '+' || theNum[0] == '-') {
                                    if (theNum[0] == '+') {
                                        isPositive = true;
                                    }
                                    else {
                                        isPositive = false;
                                    }
                                    theNum = theNum.slice (1);
                                }
                                targetIndex = parseInt (theNum, 10) - 1;
                                if (theNum.search (/L/i) != -1) {
                                    isFromLast = true;
                                }
                                if (theNum.search (/T/i) != -1) {
                                    isOnlyFrames = true;
                                }
                                else if (theNum.search (/P/i) != -1) {
                                    isOnlyFrames = false;
                                }
                                if (theNum.search (/E/i) != -1) {
                                    isEvenOnly = true;
                                }
                                else if (theNum.search (/O/i) != -1) {
                                    isEvenOnly = false;
                                }
                            }
                        }
                        var textItemsPairs = sortTextItems (textRange, pageItemsIDs, isOnlyFrames);
                        if (isFromLast) {
                            targetIndex = textItemsPairs.length - targetIndex - 1;
                        }
                        if (dynamicStyleMethod.search (/Rotate/i) != -1) {
                            for (var tris = 0; tris < textItemsPairs.length; tris++) {
                                if (targetIndex != null) {
                                    if (isPositive == true && tris < targetIndex) {
                                        continue;
                                    }
                                    else if (isPositive == false && tris > targetIndex) {
                                        continue;
                                    }
                                    else if (isPositive == null && tris != targetIndex) {
                                        continue;
                                    }
                                }
                                appliedCounter++;
                                if (isEvenOnly == true && (appliedCounter % 2 == 1)) {
                                    continue;
                                }
                                else if (isEvenOnly == false && (appliedCounter % 2 == 0)) {
                                    continue;
                                }
                                textItemsPairs[tris][0].rotationAngle = newValue;
                            }
                        }
                        else {
                            for (var tris = 0; tris < textItemsPairs.length; tris++) {
                                if (targetIndex != null) {
                                    if (isPositive == true && tris < targetIndex) {
                                        continue;
                                    }
                                    else if (isPositive == false && tris > targetIndex) {
                                        continue;
                                    }
                                    else if (isPositive == null && tris != targetIndex) {
                                        continue;
                                    }
                                }
                                appliedCounter++;
                                if (isEvenOnly == true && (appliedCounter % 2 == 1)) {
                                    continue;
                                }
                                else if (isEvenOnly == false && (appliedCounter % 2 == 0)) {
                                    continue;
                                }
                                textItemsPairs[tris][0].shearAngle = newValue;
                            }
                        }
                    }
                }
            }
            else {
                var newValue = parseFloat (newStyleName);
                if (newValue) {
                    var targetObject = null;
                    if (startInsertionPoint.parentTextFrames && startInsertionPoint.parentTextFrames.length > 0) {
                        targetObject = startInsertionPoint.parentTextFrames[0];
                    }
                    if (targetObject != null) {
                        if (dynamicStyleMethod.search (/Rotate/i) != -1) {
                            targetObject.rotationAngle = newValue;
                        }
                        else {
                            targetObject.shearAngle = newValue;
                        }
                    }
                }
            }
        }
        if (dynamicStyleMethod.search (/Merge/i) != -1) {
            var theCells = new Array;
            if (textRange != false) {
                if (textRange != null) {
                    for (var trt = 0; trt < textRange.tables.length; trt++) {
                        for (var hf = 0; hf < textRange.tables[trt].rows.length; hf++) {
                            for (var hfc = 0; hfc < textRange.tables[trt].rows.item (hf).cells.length; hfc++) {
                                var theCell = textRange.tables[trt].rows.item (hf).cells.item (hfc);
                                if (theCell.contents == "") {
                                    theCells.push (theCell);
                                }
                            }
                        }
                    }
                }
            }
            else if (startInsertionPoint.parent.constructor.name == "Cell") {
                theCells.push (startInsertionPoint.parent);
            }
            for (var tc = theCells.length - 1; tc >= 0; tc--) {
                if (theCells[tc]) {
                    switch(newStyleName) {
                        case "Next":
                            if (theCells[tc].columns[0].index < (theCells[tc].columns[0].parent.columns.length - 1)) {
                                var targetCell = theCells[tc].rows[0].cells.item (theCells[tc].columns[0].index + 1);
                                if (targetCell) {
                                    theCells[tc].merge (targetCell);
                                }
                            }
                            break;
                        case "Up":
                            if (theCells[tc].rows[0].index > 0) {
                                var targetCell = theCells[tc].columns[0].cells.item (theCells[tc].rows[0].index - 1);
                                if (targetCell) {
                                    theCells[tc].merge (targetCell);
                                }
                            }
                            break;
                        case "Down":
                            if (theCells[tc].rows[0].index < (theCells[tc].rows[0].parent.rows.length - 1)) {
                                var targetCell = theCells[tc].columns[0].cells.item (theCells[tc].rows[0].index + 1);
                                if (targetCell) {
                                    theCells[tc].merge (targetCell);
                                }
                            }
                            break;
                        default:
                            if (theCells[tc].columns[0].index > 0) {
                                var targetCell = theCells[tc].rows[0].cells.item (theCells[tc].columns[0].index - 1);
                                if (targetCell) {
                                    theCells[tc].merge (targetCell);
                                }
                            }
                    }
                }
            }
        }
    }
    return true;
}

function getPureName (pureName, isToRemoveID) {
    if (pureName.indexOf (".") != -1) {
        pureName = pureName.slice (0, pureName.lastIndexOf ("."));
    }
    var testName = pureName;
    testName = pureName.replace (/\s*\[[^\]]*\]\s*/g, "");
    if (testName == '') {
        return pureName;
    }
    else {
        pureName = testName;
    }
    testName = pureName.replace (/(\s?)(ts)(_\w+)+/g, "");
    if (testName == '') {
        return pureName;
    }
    else {
        pureName = testName;
    }
    if (isToRemoveID) {
        testName = pureName.replace (/^ID[A-Z0-9]+\s*/, "");
        if (testName == '') {
            return pureName;
        }
        else {
            pureName = testName;
        }
    }
    testName = pureName.replace (/^-\s*/, "");
    if (testName == '') {
        return pureName;
    }
    else {
        pureName = testName;
    }
    return pureName;
}

function getSubCells (subCellsList, cellsDeepItem, isParentFixed, isHorizontal) {
    //cellsDeepItem =
    /*[0, isHorizontal, cellsDirection, cellsList, cellsFixedList, cellsIndex, cellsChildList]*/
    //subCellsList =
    /*[totalFixed, [[cellObj, isFixed], ...]]*/
    var testSubs = [];
    for (var csl = 0; csl < cellsDeepItem[3].length; csl++) {
        var isFixed = isParentFixed;
        if (!isParentFixed && (cellsDeepItem[1] == isHorizontal)) {
            if (cellsDeepItem[4][csl] != 0) {
                subCellsList[0] += cellsDeepItem[4][csl];
                isFixed = true;
            }
        }
        if (cellsDeepItem[6][csl] != null) {
            if (cellsDeepItem[1] == isHorizontal) {
                getSubCells (subCellsList, cellsDeepItem[6][csl], isFixed, isHorizontal);
            }
            else {
                testSubs.push ([0, []]);
                getSubCells (testSubs[testSubs.length - 1], cellsDeepItem[6][csl], isFixed, isHorizontal);
            }
        }
        else if (cellsDeepItem[1] == isHorizontal) {
            subCellsList[1].push ([cellsDeepItem[3][csl], isFixed]);
        }
    }
    if (cellsDeepItem[1] != isHorizontal) {
        var testIndex = -1;
        if (testSubs.length > 0) {
            testIndex = 0;
            for (var ts = 1; ts < testSubs.length; ts++) {
                if (testSubs[testIndex][1].length < testSubs[ts][1].length) {
                    testIndex = ts;
                }
            }
            if (testSubs[testIndex][1].length == 1) {
                for (var ts = 1; ts < testSubs.length; ts++) {
                    if (ts == testIndex) continue;
                    if (testSubs[ts][1][0][1]) {
                        testIndex = ts;
                        break;
                    }
                }
            }
        }
        if (testIndex == -1 && cellsDeepItem[3].length > 0) {
            subCellsList[1].push ([cellsDeepItem[3][0], isParentFixed]);
        }
        else {
            subCellsList[0] += testSubs[testIndex][0];
            subCellsList[1] = subCellsList[1].concat (testSubs[testIndex][1]);
        }
    }
}

function getPureDisplayName (pureName) {
    if (pureName.indexOf (".") != -1) {
        pureName = pureName.slice (0, pureName.lastIndexOf ("."));
    }

    var testName = pureName;
    
    testName = pureName.replace (/(\s?)(ts)(_\w+)+/g, "");
    if (testName == '') {
        return pureName;
    }
    else {
        pureName = testName;
    }

    if (testName.search (/\[[^\[\]]+\]$/) != -1) {
        testName = testName.match (/\[[^\[\]]+\]$/)[0].slice (1, -1);
        if (testName != '') {
            return testName;
        }
        else {
            testName = pureName;
        }
    }

    testName = pureName.replace (/\s*\[[^\]]*\]\s*/g, "");
    if (testName == '') {
        return pureName;
    }
    else {
        pureName = testName;
    }


    testName = pureName.replace (/^-\s*/, "");
    if (testName == '') {
        return pureName;
    }
    else {
        pureName = testName;
    }

    return pureName;
}

function submitSelected (dataArray) {
	var targetDocument = dataArray[0];
    var selectedObject = dataArray[1];
	var lastCharIndex = new Array;
	var nestedMarkers = new Array;
	if (!getContainerSnippet (selectedObject, lastCharIndex, nestedMarkers, "first marker")) { //"first marker"  or "base marker"
		return false; 
	}
	createMarkersStyles(targetDocument);
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(nestedMarkers[0][0].textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
	else if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "spoiled") {
		reflectSpoiled (nestedMarkers);
		alert ("The instance is spoiled");
		return false;
    }
    var liveSnippetID = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_ID");
    var dynamicSnippetCode = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_PREFIX_NUMBER");
    var isToChange = false;
    if (dynamicSnippetCode) {
        var pathAndName = solveSnippetCode (targetDocument, nestedMarkers[0][0].parent.insertionPoints.item (nestedMarkers[0][0].index), dynamicSnippetCode, null);
        if (pathAndName[1] == "") {
            pathAndName[1] = "0001";
        }
        if (pathAndName[0] || pathAndName[1]) {
            liveSnippetIDSplitted = new Array;
            liveSnippetIDSplitted.push (liveSnippetID.slice (0, liveSnippetID.lastIndexOf ("/")));
            liveSnippetIDSplitted.push (liveSnippetID.slice (liveSnippetID.lastIndexOf ("/") + 1));
            if (pathAndName[0]) {
                if (liveSnippetIDSplitted[0] != pathAndName[0]) {
                    isToChange = true;
                }
            }
            if (pathAndName[1]) {
                if (liveSnippetIDSplitted[1] != pathAndName[1]) {
                    isToChange = true;
                }
            }
            if (isToChange) {
                changeMarkers (nestedMarkers, pathAndName[0], pathAndName[1]);
                nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_TSID", "");
            }
        }
    }
    var submitFile = submitInstance (targetDocument, nestedMarkers);
	if (!submitFile) {
        if (submitFile == null) {
            alert ("It's Inline Snippet and couldn't be submitted!");
        }
        else {
            alert ("It's Dynamic Live Snippet and couldn't be submitted!");
        }
    }
    else {
        var theSnippetText = nestedMarkers[0][0].parent.characters.itemByRange(nestedMarkers[0][0], nestedMarkers[nestedMarkers.length-1][0]);
        theSnippetText.select();
    }
}

function getTextCoordinates (theText, offsetValue, myPageWidth, myPageHeight) {
    var pageBoundingsPairs = new Array;
    var scanChar = theText.characters[0];
    var X_first = null;
    var X_last = 0;
    for (var ttc = 0; ttc < theText.characters.length; ttc++) {
        //Check if current is hidden
        var currentFrame = null;
        if (!theText.characters[ttc].parentTextFrames || theText.characters[ttc].parentTextFrames.length == 0) {
            break;
        }
        else if (theText.characters[ttc].parentTextFrames[0] instanceof Array) {
            if (theText.characters[ttc].parentTextFrames[0].length == 0) {
                break;
            }
            else {
                currentFrame = theText.characters[ttc].parentTextFrames[0][0];
            }
        }
        else {
            currentFrame = theText.characters[ttc].parentTextFrames[0];
        }
        var firstChar = scanChar;
        var lastChar = null;
        if (ttc < theText.characters.length - 1) { //Not the last char
            var nextFrame = null;
            if (theText.characters[ttc + 1].parentTextFrames && theText.characters[ttc + 1].parentTextFrames.length != 0) {
                if (theText.characters[ttc + 1].parentTextFrames[0] instanceof Array) {
                    if (theText.characters[ttc + 1].parentTextFrames[0].length != 0) {
                        nextFrame = theText.characters[ttc + 1].parentTextFrames[0][0];
                    }
                }
                else {
                    nextFrame = theText.characters[ttc + 1].parentTextFrames[0];
                }
            }
            if (nextFrame) {
                if (currentFrame.id != nextFrame.id) {
                    lastChar = theText.characters[ttc];
                    scanChar = theText.characters[ttc + 1];
                }
            }
            else {
                lastChar = theText.characters[ttc];
            }
        }
        else {
            lastChar = theText.characters[ttc];
        }
        var toExmine1 = parseFloat (theText.characters[ttc].horizontalOffset.toString ());
        var toExmine2 = parseFloat (theText.characters[ttc].endHorizontalOffset.toString ());
        if (toExmine1 > toExmine2) {
            var swapCell = toExmine1;
            toExmine1 = toExmine2;
            toExmine2 = swapCell;
        }
        if (X_first == null || X_first > toExmine1) {
            X_first = toExmine1;
        } 
        if (X_last < toExmine2) {
            X_last = toExmine2;
        }
        if (lastChar) {
            var expectedParent = null;
            if (firstChar.parent instanceof Array) {
                expectedParent = firstChar.parent[0];
            }
            else {
                expectedParent = firstChar.parent;
            }
            var thePortion = expectedParent.texts.itemByRange (firstChar, lastChar);
            var firstCharFrame = firstChar.parentTextFrames[0];
            if (firstCharFrame instanceof Array) {
                firstCharFrame = firstCharFrame[0];
            }
            var thePage = firstCharFrame.parentPage;
            if (thePage) {
                endOffset = 0;
                var theAscent = parseFloat (thePortion.ascent.toString ());
                var theDescent = parseFloat (thePortion.descent.toString ());
                var Y_top = parseFloat (thePortion.baseline.toString ()) - theAscent;
                var Y_bottom = parseFloat (thePortion.endBaseline.toString ()) + theDescent;
                if (offsetValue != null) {
                    X_first -= offsetValue; 
                    X_last += offsetValue;
                    Y_top -= offsetValue;
                    Y_bottom += offsetValue;
                }
                if (myPageWidth) {
                    X_first = (X_first / myPageWidth) * 100;
                    X_last = (X_last / myPageWidth) * 100;
                }
                if (myPageHeight) {
                    Y_top = (Y_top / myPageHeight) * 100;
                    Y_bottom = (Y_bottom / myPageHeight) * 100;
                }
                X_first = X_first.toFixed(2);
                if (X_first.slice (-3) == ".00") {
                    X_first = X_first.slice (0, -3);
                }
                X_last = X_last.toFixed(2);
                if (X_last.slice (-3) == ".00") {
                    X_last = X_last.slice (0, -3);
                }
                Y_top = Y_top.toFixed(2);
                if (Y_top.slice (-3) == ".00") {
                    Y_top = Y_top.slice (0, -3);
                }
                Y_bottom = Y_bottom.toFixed(2);
                if (Y_bottom.slice (-3) == ".00") {
                    Y_bottom = Y_bottom.slice (0, -3);
                }
                pageBoundingsPairs.push ([thePage.name, [Y_top, X_first, Y_bottom, X_last]]);
                X_first = 0;
                X_last = 0;
            } 
        }
    }
    return pageBoundingsPairs;
}

function submitInstance (targetDocument, nestedMarkers) {
    if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_SUM") != "DYNAMIC") {
        nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_STATE", "updated");
        changeStateAppearance(nestedMarkers[0][0].textFrames.item(0));
        var liveSnippetID = nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_ID");
        var liveSnippetIDSplitted = new Array;
        liveSnippetIDSplitted.push (liveSnippetID.slice (0, liveSnippetID.lastIndexOf ("/")));
        if (liveSnippetIDSplitted[0] == "#") {
            return null;
        }
        liveSnippetIDSplitted.push (liveSnippetID.slice (liveSnippetID.lastIndexOf ("/") + 1));
        var liveSnippetFolder = getLiveSnippetFolder (targetDocument, liveSnippetIDSplitted[0], liveSnippetIDSplitted[1], []);
        var liveSnippetFile = createLiveSnippetFile (targetDocument, liveSnippetFolder, nestedMarkers, "");
        return liveSnippetFile;
    }
    else {
        return false;
    }
}

function getMarginBounds (targetDocument, targetPage){
	var targetPageWidth = targetDocument.documentPreferences.pageWidth;
	var targetPageHeight = targetDocument.documentPreferences.pageHeight;
	if (targetPage.side == PageSideOptions.leftHand){
		var myX2 = targetPage.marginPreferences.left;
		var myX1 = targetPage.marginPreferences.right;
	}
	else {
		var myX1 = targetPage.marginPreferences.left;
		var myX2 = targetPage.marginPreferences.right;
	}
	var myY1 = targetPage.marginPreferences.top;
	var myX2 = targetPageWidth - myX2;
	var myY2 = targetPageHeight - targetPage.marginPreferences.bottom;
	return [myY1, myX1, myY2, myX2];
}

function unlinkInstance (dataArray) {
    //[targetDocument, selectedObject] 
    var targetDocument = dataArray[0];
    var selectedObject = dataArray[1];
	createMarkersStyles(targetDocument);
	var lastCharIndex = new Array;
	var nestedMarkers = new Array;
	if (!getContainerSnippet (selectedObject, lastCharIndex, nestedMarkers, "first marker")) { //"first marker"  or "base marker"
		return false; 
	}
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(startChar.textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
    var textStory = nestedMarkers[0][0].parent;
    var startIndex = nestedMarkers[0][0].index;
    var lastIndex = nestedMarkers[nestedMarkers.length-1][0].index;
    nestedMarkers[nestedMarkers.length-1][0].remove();
    nestedMarkers[0][0].remove();
    textStory.characters.itemByRange(startIndex, lastIndex-2).select();
}

function showInfo (dataArray) {
    //[targetDocument, selectedObject] 
    var targetDocument = dataArray[0];
    var selectedObject = dataArray[1];
	createMarkersStyles(targetDocument);
	var lastCharIndex = new Array;
	var nestedMarkers = new Array;
	if (!getContainerSnippet (selectedObject, lastCharIndex, nestedMarkers, "first marker")) { //"first marker"  or "base marker"
		return false; 
	}
	if (nestedMarkers[0][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
		changeStateAppearance(startChar.textFrames.item(0));
		alert ("The instance has no end");
		return false;
	}
    alert (messageNestedMarkers (nestedMarkers));
}

function getStartChar (inletObject) {
	var startChar = null;
	var currentStory = null;
	var currentIndex = null;
	var currentChar = null;
	switch (inletObject.constructor.name){
		case "InsertionPoint":
			currentStory = inletObject.parent;
			currentIndex = inletObject.index - 1;
			currentChar = inletObject.parent.characters.item(currentIndex); //if the insertion point index is zero then the currentChar when be null, don't warry because it will not be used
			break;
		case "Character":
			currentStory = inletObject.parent;
			currentIndex = inletObject.index;
			currentChar = inletObject;
			break;		
		case "Word":
		case "Paragraph":
		case "Text":
		case "TextColumn":
		case "TextStyleRange":
			currentStory = inletObject.characters.item(0).parent;
			currentIndex = inletObject.characters.item(0).index;
			currentChar = inletObject.characters.item(0);
			break;
		case "FormField":
		case "GraphicLine":
		case "Group":
		case "Oval":
		case "Polygon":
		case "Rectangle":
		case "TextFrame":
		case "Table":
		case "Cell":
			currentChar = downLevel(inletObject);
			if (currentChar != null) {
				currentStory = currentChar.parent;
				currentIndex = currentChar.index;
			}
			else
				return startChar; //or null
			break;
		default:
			return null;
	}
	//scan the characters range from currentChar to the beginning of the current sub story
	do {
		var level = 0;
		for (var c = currentIndex; c >= 0; c--) {
            if (currentStory.characters.item(c).contents == SpecialCharacters.END_NESTED_STYLE) { // It's an end.
                level++;
            }
            else if (currentStory.characters.item(c).textFrames.length > 0) {
				if (currentStory.characters.item(c).textFrames.item(0).extractLabel("LS_TAG") == "snippet start") { // It's a start. 
                    if (level == 0) {
                        startChar = currentStory.characters.item(c);
                        break;
                    }
                    else {
                        //It's not the base level start.
                    }
                    level--;
				}
			}
		}
		if (startChar)
			break;
		if (currentStory.constructor.name == "Cell")
			currentChar = downLevel(currentStory);
		else
			currentChar = downLevel(currentStory.insertionPoints.item(0).parentTextFrames[0]);
		if (currentChar) {
			currentIndex = currentChar.index;
			currentStory = currentChar.parent;
		}
	} while (currentChar);
	return startChar;
}

function downLevel(currentObject) {
	var rootChar = null;
	if (currentObject.constructor.name == "Cell") {
		currentObject = currentObject.parentRow.parent;
	}
	if (currentObject.constructor.name == "Table") {
		rootChar = currentObject.storyOffset.parent.characters.item(currentObject.storyOffset.index);
	}
	else {
		do {
			currentObject = currentObject.parent;
			if (currentObject.constructor.name == "Character") {
				rootChar = currentObject;
				break;
			}
		} while (currentObject.constructor.name != "Spread" && currentObject.constructor.name != "MasterSpread" && currentObject.constructor.name != "Page" && currentObject.constructor.name != "Document");
	}
	return rootChar;
}

function changeStateAppearance(MarkerTextFrame) { // Change the appearance of the anchored marker to reflect if the instance is Updated, Modified , Missing, or Spoiled
	var myDoc = MarkerTextFrame.parentStory.parent;
	switch (MarkerTextFrame.extractLabel("LS_STATE")){
		case "updated":
			MarkerTextFrame.paragraphs.item(0).appliedCharacterStyle = myDoc.characterStyles.itemByName("[None]");
			MarkerTextFrame.paragraphs.item(0).clearOverrides(OverrideType.ALL);
			MarkerTextFrame.nonprinting = true;
			break;
		case "modified":
			MarkerTextFrame.paragraphs.item(0).appliedCharacterStyle = myDoc.characterStyles.itemByName("Live Snippet Modified");
			MarkerTextFrame.paragraphs.item(0).clearOverrides(OverrideType.ALL);
			MarkerTextFrame.nonprinting = false; //So the wrong font style will appear in the preflight panel.
			break;
		case "missing":
			MarkerTextFrame.paragraphs.item(0).appliedCharacterStyle = myDoc.characterStyles.itemByName("Live Snippet Missing");
			MarkerTextFrame.paragraphs.item(0).clearOverrides(OverrideType.ALL);
			MarkerTextFrame.nonprinting = false; //So the wrong font style will appear in the preflight panel.
			break;
		case "spoiled":
			MarkerTextFrame.paragraphs.item(0).appliedCharacterStyle = myDoc.characterStyles.itemByName("Live Snippet Spoiled");
			MarkerTextFrame.paragraphs.item(0).clearOverrides(OverrideType.ALL);
			MarkerTextFrame.nonprinting = false; //So the wrong font style will appear in the preflight panel.
			break;
		case "lonely":
			MarkerTextFrame.paragraphs.item(0).appliedCharacterStyle = myDoc.characterStyles.itemByName("Live Snippet Lonely");
			MarkerTextFrame.paragraphs.item(0).clearOverrides(OverrideType.ALL);
			MarkerTextFrame.nonprinting = false; //So the wrong font style will appear in the preflight panel.
			break;
		case "need to submit":
			MarkerTextFrame.paragraphs.item(0).appliedCharacterStyle = myDoc.characterStyles.itemByName("Live Snippet Need to Submit");
			MarkerTextFrame.paragraphs.item(0).clearOverrides(OverrideType.ALL);
			MarkerTextFrame.nonprinting = false; //So the wrong font style will appear in the preflight panel.
			break;
		default:
			alert("Wrong in Live Snippet Script Code " + $.line);
	}
}

function messageNestedMarkers(nestedMarkers) {
	var structureMessage = "Info.\r";
	var indentSpan;
	var c;
	var level = 0;
	for (c = 0; c < nestedMarkers.length; c++) {
        if (nestedMarkers[c][0].contents == SpecialCharacters.END_NESTED_STYLE) {
			/*indentSpan = "";
			for (var t = 1; t < level; t++) indentSpan += "\t";
			var markerLabel = nestedMarkers[c][0].textFrames.item(0).extractLabel("LS_ID");
			if (nestedMarkers[c][0].textFrames.item(0).extractLabel("LS_ID").indexOf("live_snippet") != -1)
				var markerLabel = parseInt(nestedMarkers[c][0].textFrames.item(0).extractLabel("LS_ID").substr(8, 4), 10);
			structureMessage += indentSpan + nestedMarkers[c][0].textFrames.item(0).extractLabel("LS_STATE") + " end " + markerLabel+"\r";*/
			level--;
        }
        else {
			level++;
			indentSpan = "";
			for (var t = 1; t < level; t++) indentSpan += "\t";
			var markerLabel = "snippet "+ nestedMarkers[c][0].textFrames.item(0).extractLabel("LS_ID");
			structureMessage += indentSpan + markerLabel+ "\r";
			if (nestedMarkers[c][0].textFrames.item(0).extractLabel("LS_STATE") == "lonely")
				level--;
        }
	}
	return structureMessage;
}

function reflectSpoiled (nestedMarkers) {
	var c;
	for (c = 0; c < nestedMarkers.length; c++) {
        if (nestedMarkers[c][0].contents == SpecialCharacters.END_NESTED_STYLE) {

        }
        else {
			switch (nestedMarkers[c][0].textFrames.item(0).extractLabel("LS_STATE")) {
				case "lonely":
				case "spoiled":
					changeStateAppearance(nestedMarkers[c][0].textFrames.item(0)); //Changing the appearance of the text frame to reflect the actual state
					break;
				case "normal":
					break;
			}
        }
	}
}

function newLiveSnippet (dataArray){ //[targetDocument, selectedObject]
	var targetDocument = dataArray[0];
    var selectedObject = dataArray[1];
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
    var targetText = null;
    var isForInsertionPoint = false;
    var portions = new Array;
    var paraCount = 0;
    var isCellsOrTableSelected = false;
    var cellsList = new Array;
	switch (selectedObject.constructor.name){
        case "InsertionPoint":
			if (selectedObject.paragraphs.length > 0) {
                if (selectedObject.paragraphs[0].characters.length == 1) {
                    targetText = selectedObject;
                    isForInsertionPoint = true;
                }
                else {
                    targetText = selectedObject.paragraphs.item(0);
                }
			}
			else {
                targetText = selectedObject;
                isForInsertionPoint = true;
            }
			break;
        case "Character":
        case "Word":
        case "Paragraph":
        case "Text":
        case "TextColumn":
        case "TextStyleRange":
			targetText = selectedObject;
			break;
        case "FormField":
        case "GraphicLine":
        case "Group":
        case "Oval":
        case "Polygon":
        case "Rectangle":
        case "TextFrame":
			targetText = downLevel (selectedObject);
			if (targetText == null) {
				return false;
			}
			break;
        case "Table":
        case "Cell":
            isCellsOrTableSelected = true;
            for (var tt = 0; tt < selectedObject.cells.length; tt++) {
                if (selectedObject.cells[tt].paragraphs.length > 0) {
                    paraCount += selectedObject.cells[tt].paragraphs.length;
                    portions.push (selectedObject.cells[tt].paragraphs);
                    cellsList.push (selectedObject.cells[tt]);
                }
                    
            }
			targetText = null;
			break;
		default:
			alert("Please select a live snippet instance."); 
			return false;
	}
    if (portions.length == 0 && targetText == null) {
        return false;
    }
    if (targetText != null) {
        if (targetText.paragraphs.length > 1) {
            paraCount += targetText.paragraphs.length;
            portions.push (targetText.paragraphs);
        }
    }
    var isRemoveReturn = null;
    var isIndividuals = null;
    if (portions.length > 0) {
        if (paraCount > 1) {
            isIndividuals = confirm ("Do you want as Individuals?");
            if (!isIndividuals) {
                if (isCellsOrTableSelected) {
                    isRemoveReturn = false;
                    for (var por = 0; por < portions.length; por++) {
                        var wholeCellText = portions[por][0].parent.texts.itemByRange (portions[por][0].parent.characters.item (0), portions[por][0].parent.characters.item (-1));
                        portions[por] = [wholeCellText];
                    }
                }
                else {
                    isRemoveReturn = false;
                    portions = [[targetText]];
                }
            }
        }
        if (isRemoveReturn == null) {
            isRemoveReturn = confirm ("Without Return Character?");
        }
    }
    else {
        portions.push ([targetText]);
        if (!isForInsertionPoint && (selectedObject.constructor.name == "InsertionPoint")) {
            isRemoveReturn = true;
        }
    }
    var dynamicSnippetCode = targetDocument.extractLabel ("LS_DEFAULT_DYNAMIC");
    dynamicSnippetCode = tsGetText (["Live Snippet Phrase:",
                                         "Use <.> for the snippet file parent folder name, add more dots for parent of parent...",
                                         "To reference to the document instead of the snippet file add dollar sign like this '<.$>'.",
                                         "Use also: <Increment>, <Container_File_Branch>, <Container_File_Name>,",
                                         "<Doc_Name>, <Doc_Order>, <Page_Num>, <Applied_Section_Prefix>,",
                                         "<Applied_Section_Marker>, <Master_Prefix>, <Master_Name>, <Master_Full_Name>, <Para_Order>,",
                                         "<Para_First>, <Para_Before>, <Para_This>, <Para_After>, <Para_Style>, <Col_Head>, <Col_Num>,",
                                         "<Row_Head>, <Row_Num>, <Shift1>.",
                                         "",
                                         "The following accepts abbreviation and used for date, time and random code:",
                                         "<Day_Seconds>, <Day_Minutes>, <Day_Hours>, <Weekday_Order>, <Weekday_Name>, <Weekday_Abbreviation>,",
                                         "<Month_Date>, <Month_Order>, <Month_Name>, <Month_Abbreviation>, <Year_Full>,",
                                         "<Year_Abbreviation>, <Time_Decimal>, <Time_Hexadecimal>, <Time_Shortest>,",
                                         "<Random_Style_Long> Where 'Style' a combination of symbols 'a' for small letters,",
                                         "'A' capital letters, 'h' small hexadecimal, 'H' large hexadecimal and '1' for numbers.",
                                         "For example <Random_1A_8> means random 8 characters consists of capital letters and numbers",
                                         ""
                                        ], dynamicSnippetCode, false);
    if (!dynamicSnippetCode) {
        exit();
    }
    if (dynamicSnippetCode.indexOf ("./") == 0) {
        dynamicSnippetCode = dynamicSnippetCode.slice (2);
    }
    else if (dynamicSnippetCode.indexOf ("/") == 0) {
        dynamicSnippetCode = dynamicSnippetCode.slice (1);
    }
    targetDocument.insertLabel ("LS_DEFAULT_DYNAMIC", dynamicSnippetCode);
    createMarkersStyles(targetDocument);
    var foldersCountersList = new Array;
    //foldersCountersList[fc][0] folder full path
    //foldersCountersList[fc][1] counter number
    //foldersCountersList[fc][2] list of [targetText, pathAndName] pairs
    var dynamicContentPhrase = "";
    if (dynamicSnippetCode.indexOf ("::") != -1) {
        dynamicContentPhrase = dynamicSnippetCode.slice (dynamicSnippetCode.indexOf ("::") + 2);
        dynamicSnippetCode = dynamicSnippetCode.slice (0, dynamicSnippetCode.indexOf ("::"));
    }
    for (var por = portions.length - 1; por >= 0; por--) {
        for (var pp = portions[por].length - 1; pp >= 0; pp--) {
            targetText = portions[por][pp];
            var solvingInsertionPoint = null;
            if (!isIndividuals && isCellsOrTableSelected) {
                solvingInsertionPoint = cellsList[por].insertionPoints.item (0);
            }
            else {
                solvingInsertionPoint = targetText.insertionPoints.item (0);
            }
            var pathAndName = solveSnippetCode (targetDocument, solvingInsertionPoint, dynamicSnippetCode, null);
            var folderIndex = -1;
            if (pathAndName[0] == "#") {
                for (var fc = 0; fc < foldersCountersList.length; fc++) {
                    if (foldersCountersList[fc][0] == "#") {
                        folderIndex = fc;
                        break;
                    }
                }
                if (folderIndex == -1) {
                    foldersCountersList.unshift (["#", 1, [[targetText, pathAndName]]]);
                }
                else {
                    foldersCountersList[folderIndex][2].unshift ([targetText, pathAndName]);
                    foldersCountersList[folderIndex][1]++;
                }
            }
            else {
                var liveSnippetsFolder = getLiveSnippetFolder (targetDocument, pathAndName[0], null, null);
                for (var fc = 0; fc < foldersCountersList.length; fc++) {
                    if (foldersCountersList[fc][0] == liveSnippetsFolder.fsName.replace(/\\/g, "/")) {
                        folderIndex = fc;
                        break;
                    }
                }
                if (folderIndex == -1) {
                    foldersCountersList.unshift ([liveSnippetsFolder.fsName.replace(/\\/g, "/"), parseInt (getNewLiveSnippetDigits (liveSnippetsFolder), 10), [[targetText, pathAndName]]]);
                }
                else {
                    foldersCountersList[folderIndex][2].unshift ([targetText, pathAndName]);
                    foldersCountersList[folderIndex][1]++;
                }
            }
        }
    }
    var isToMakeDynamic = null;
    var isToDynamicContent = null;
    if (isForInsertionPoint || dynamicContentPhrase != "") {
        isToDynamicContent = true;
    }
    else {
        isToDynamicContent = false;
    }
    for (var fc = foldersCountersList.length - 1; fc >= 0; fc--) {
        for (var pairs = foldersCountersList[fc][2].length - 1; pairs >= 0; pairs--) {
            targetText = foldersCountersList[fc][2][pairs][0];
            var pathAndName = foldersCountersList[fc][2][pairs][1];
            if (pathAndName[1] == "") {
                pathAndName[1] = tsFillZeros (foldersCountersList[fc][1], 4);
                foldersCountersList[fc][1]--;
            }
            var trimIndex = -1;
            if (isRemoveReturn) {
                if (targetText.characters.item (-1).contents == "\r") {
                    trimIndex = -2;
                }
            }
            var startMarkerFrame = insertMarkers (targetDocument, targetText, pathAndName[0], pathAndName[1], trimIndex, isIndividuals);
            var nestedMarkers = new Array;
            var lastCharIndex = new Array;
            lastCharIndex.push(0);
            surveyAndCheck (startMarkerFrame.parent, lastCharIndex, nestedMarkers, false); 
            if (pathAndName.length > 2) {
                nestedMarkers[0][0].textFrames.item(0).insertLabel ("LS_INCREMENT_STATE", pathAndName[2]);
                nestedMarkers[0][0].textFrames.item(0).insertLabel ("LS_INCREMENT", pathAndName[3]);
            }
            var isDynamic = false;
            if (dynamicSnippetCode.lastIndexOf ("/") != -1) {
                if (dynamicSnippetCode.slice (0, dynamicSnippetCode.lastIndexOf ("/")) != pathAndName[0] || pathAndName[1] != dynamicSnippetCode.slice (dynamicSnippetCode.lastIndexOf ("/") + 1)) {
                    isDynamic = true;
                }
            }
            else if (dynamicSnippetCode != pathAndName[0]) {
                isDynamic = true;
            }
            if (isDynamic) {
                if (isToMakeDynamic == null) {
                    isToMakeDynamic = confirm ("Set Dynamic Identifier (Branch and Name)?");
                }
                if (isToMakeDynamic) {
                    nestedMarkers[0][0].textFrames.item(0).insertLabel("LS_PREFIX_NUMBER", dynamicSnippetCode);
                }
            }
            if (nestedMarkers[0][0].paragraphDirection == ParagraphDirectionOptions.RIGHT_TO_LEFT_DIRECTION) {
                reflectHorizontally (nestedMarkers[0][0].textFrames.item(0));
            }
            if (isToDynamicContent) {
                newDynamicLiveSnippet (targetDocument, nestedMarkers, dynamicContentPhrase, true, false);
            }
            else if (pathAndName[0] != "#") {
                createLiveSnippetFile (targetDocument, new Folder (foldersCountersList[fc][0]), nestedMarkers, "");
            }
        }
    }
    return true;
}

function getLiveSnippetFolder (targetDocument, liveSnippetPrefix, liveSnippetNumber, liveSnippetFiles) {
    var firstExistFolder = null;
    var docPath = preparePath (targetDocument.filePath.fsName.replace(/\\/g, "/"));
    var newPathPortion = "";
    if (docPath != workshopPath) {
        newPathPortion = docPath.replace (workshopPath, "");
    }
    var isfinalRound = false;
    var prefixWithSlash = "";
    if (liveSnippetPrefix) {
        if (liveSnippetPrefix[0] != "/") {
            prefixWithSlash = "/" + liveSnippetPrefix;
        } else {
            prefixWithSlash = liveSnippetPrefix;
        }
    }
    do {
        var testFolder = new Folder (workshopPath + newPathPortion + prefixWithSlash);
        if (testFolder.exists) {
            if (!firstExistFolder) {
                firstExistFolder = testFolder;
            }
            if (liveSnippetNumber) {
                var testedLiveSnippetFiles = testFolder.getFiles (liveSnippetNumber + "*");
                if (testedLiveSnippetFiles.length > 0) {
                    for (var tlsf = 0; tlsf < testedLiveSnippetFiles.length; tlsf++) {
                        if ((testedLiveSnippetFiles[tlsf] instanceof File) && (testedLiveSnippetFiles[tlsf].name.slice (-4).toLowerCase() == ".txt" || testedLiveSnippetFiles[tlsf].name.slice (-5) == ".IDMS")) {
                            liveSnippetFiles.push (testedLiveSnippetFiles[tlsf]);
                        }
                    }
                    return testFolder;
                }
            }
            else {
                return testFolder;
            }
        }
        if (newPathPortion.indexOf ("/") != -1) {
            newPathPortion = newPathPortion.slice (0, newPathPortion.lastIndexOf ("/"));
        }
        else {
            isfinalRound = true;
        }
    } while (!isfinalRound);
    if (firstExistFolder) {
        return firstExistFolder;
    }
    else {
        var createdFolder = new Folder (docPath + prefixWithSlash);
        createdFolder.create ();
        return (createdFolder);
    }
}

function getNewLiveSnippetDigits (liveSnippetsFolder) {
	//Producing the new number of the snippet to be appended to the end of the file name.
    var newNumber  = 1;
    if (!liveSnippetsFolder.exists) {
        liveSnippetsFolder.create ();
    }
    else {
        var allFiles = liveSnippetsFolder.getFiles(isUnhiddenFile);
        if (allFiles.length > 0) {
            var lastOldFile = allFiles[allFiles.length - 1];
            var testNewNumber = parseInt (File.decode (lastOldFile.name), 10);
            if (testNewNumber)
                newNumber = testNewNumber + 1;
        }
        
    }
    return tsFillZeros (newNumber, 4);
}

function fixNesting(startChar, nestedMarkers) {
	var c;
    var startIndex = 0
    var amount = 0;
	var snippetID = startChar.textFrames.item(0).extractLabel("LS_ID");
	for (c = 0; c < nestedMarkers.length; c++) {
		if (nestedMarkers[c][0] == startChar) {
			startIndex = c;
			if (startChar.textFrames.item(0).extractLabel("LS_STATE") == "lonely") {
				amount = 1;
				break;
			}
            var i = c + 1;
            var level = 0;
			while (i < nestedMarkers.length) {
				if (nestedMarkers[i][0].contents == SpecialCharacters.END_NESTED_STYLE) {
					if (nestedMarkers[i][0].textFrames.item(0).extractLabel("LS_ID") == snippetID) {
                        if (level == 0) {
                            amount = i - startIndex + 1;
                            break;
                        }
                        level--;
                    }
                    else {
                        level++;
                    }
				}
				else {
				}
				i++;
			}
			break;
		}
	}
    var shiftedSubInstance = new Array;
    if (amount > 0) {
        shiftedSubInstance = nestedMarkers.splice(startIndex, amount);
    }
	for (c = 0; c < shiftedSubInstance.length; c++) {
		nestedMarkers.push(shiftedSubInstance[c]);
	}
	return shiftedSubInstance;
}

function getEndIndex(startMarkerIndex, nestedMarkers) { //this index related to nestedMarkers array not to the index of the character inside its story.
    var i = startMarkerIndex;
    var level = 0;
	for (i ; i < nestedMarkers.length; i++) {
        if (nestedMarkers[i][0].contents == SpecialCharacters.END_NESTED_STYLE) {
            if (level == 0) {
                return i;
            }
            level--;
        }
        else {
            if (nestedMarkers[i][0].textFrames.item(0).extractLabel("LS_STATE") != "lonely") {
                level++;
            }
        }		
    }
    return startMarkerIndex;
}

function setExceptionStyle (myTextFrame) {

	with (myTextFrame)
	{
		fillColor = app.activeDocument.colors.itemByName("Paper");
	}
}

function isActiveDocReady () {
	try {
		app.activeDocument;
	}
	catch (myError) {
		return "no_document";
	}
	try {
		app.activeDocument.fullName;
	}
	catch (myError) {
		return "not_saved";
	}
	return "ready";
}
//end live snippets

mainInDesign ();