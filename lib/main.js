// Get a reference to a CSInterface object
var csInterface = new CSInterface();
window.onload = load;
function load () {
    var hostApplication = csInterface.getSystemPath(SystemPath.HOST_APPLICATION);
    if (hostApplication.indexOf ("Bridge") == -1) {
        csInterface.evalScript("mainInDesign ()");
        csInterface.addEventListener("com.treeshade.rootset", function(evt) {
            handleRoot (evt.data);
        });
    }
    else {    
        csInterface.addEventListener("com.treeshade.copystartupresult", function(evt) {
            handleCopyStartup (evt.data);
        });

    csInterface.evalScript("installStartupScripts ('" + 
    csInterface.getSystemPath(SystemPath.EXTENSION) + "', '" + 
    csInterface.getSystemPath(SystemPath.USER_DATA) + "', '" + 
    csInterface.getSystemPath(SystemPath.MY_DOCUMENTS) + "', '" + 
    csInterface.getSystemPath(SystemPath.HOST_APPLICATION) + "', '" + 
    "LOAD" + "')");

    setInterval(function(){ 
        csInterface.evalScript("installStartupScripts ('" + 
        csInterface.getSystemPath(SystemPath.EXTENSION) + "', '" + 
        csInterface.getSystemPath(SystemPath.USER_DATA) + "', '" + 
        csInterface.getSystemPath(SystemPath.MY_DOCUMENTS) + "', '" + 
        csInterface.getSystemPath(SystemPath.HOST_APPLICATION) + "', '" + 
        "TASK" + "')");
     }, 60000);
    }
}
function handleRoot (result) {
    var p = document.getElementById("instruction");
    p.innerHTML = result;
}

function handleCopyStartup (result) {
    var copyGuideEle = document.getElementById("copyGuide");
    if (result == "Updated") {
        copyGuideEle.innerHTML = "The 'Startup Scripts' file is updated, please restart Adobe Bridge.";
    }
    else if (result == "Welcome") {
        copyGuideEle.innerHTML = "The 'Startup Scripts' file has been copied, please restart Adobe Bridge.";
    }
}
function OpenGuideBrowser()
{
    window.cep.util.openURLInDefaultBrowser("https://indd.adobe.com/view/96766e10-9225-427a-8b3a-c6544b4086b1");
}