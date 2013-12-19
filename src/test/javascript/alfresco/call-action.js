// Demonstrates calling the action
//
// This works on a sample file available from Apache POI
// https://svn.apache.org/repos/asf/poi/trunk/test-data/spreadsheet/SampleSS.xls

// NodeRef of SampleSS.xls
var n = search.findNode("workspace://SpacesStore/a154c309-dd1a-47bc-8398-ec47b1ab2b4f");
// NodeRef of the destination folder
var d = search.findNode("workspace://SpacesStore/8fcb5ddd-9b27-45e7-ad47-4e76f1194a9c");

// Setup the action
var a = actions.create("spreadsheet-excerpt");
a.parameters["destination-folder"] = d;
a.parameters["keep-sheets"] = ["First Sheet","Sheet3"];

// Have the excerpt performed
a.execute(n);
