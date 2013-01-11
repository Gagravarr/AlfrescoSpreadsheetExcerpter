AlfrescoSpreadsheetExcerpter
============================

Produce read-only excerpts of Spreadsheets (Excel .xls and .xlsx) in
Alfresco, using Apache POI.

This comes as a single AMP, SpreadsheetExcerpt.amp, which is an Alfresco
Repository Module. It provides code to perform the Excerpt, along with
a WebScript and an Action to trigger it.

Building
========
This has been tested against Alfresco Enterprise 4.1.1.

It ought to work fine on Community 4.x too

To build, simply run "ant", and the AMP will be produced in the /build/dist/ 
directory. Install this into the Alfresco Repository war as normal.

Installation
============
Once you have built the AMP, install it into the WAR using the MMT jar, and 
restart Tomcat. You'll do something like

   java -jar alfresco-mmt.jar install SpreadsheetExcerpt.amp alfresco.war

License
=======
The code is available under the LGPL v3 License, which is the same license
that Alfresco itself is made available under.
