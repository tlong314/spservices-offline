/**
 * @overview
 * A script using Web Storage APIs to simulate jQuery SPServices operations offline.
 * For training and practice purposes.
 * Based on code from the SPServices jQuery plugin by CodePlex.
 *
 * Although this file requires jQuery, note that it does not require SPServices or SharePoint, or internet connection.
 *
 * @author Tim Scott Long <tim@timlongcreative.com>, based on original code by CodePlex.
 * @requires jquery
 * @license MIT
 */

;(function($, window) {
// SharePoint pages should have this function built in; included here in case it has been removed or is inaccessible due to the current file location.
  if(typeof STSHtmlEncode === "undefined") {
    window.STSHtmlEncode = function (b){ULSA13:;if(null==b||typeof b=="undefined")return "";for(var d=new String(b),a=[],c=0,f=d.length,c=0;c<f;c++){var e=d.charAt(c);switch(e){case "<":a.push("&lt;");break;case ">":a.push("&gt;");break;case "&":a.push("&amp;");break;case '"':a.push("&quot;");break;case "'":a.push("&#39;");break;default:a.push(e)}}return a.join("")};
  }

  if(!$) {
    document.body.innerHTML = "Please be sure to include jQuery in your page before calling this file.<br>" +
    "A basic layout in your HTML will look something like this:<br><br>" +
    '&lt;script src="jquery.js"&gt;&lt;/script&gt;<br>' +
    '&lt;script src="jquery.spservices-offline.js"&gt;&lt;/script&gt;<br>' +
    '&lt;script&gt;<br><br>' +
    '/* All of your custom/practice code will go here */<br><br>' +
    '&lt;/script&gt;<br>';
 }

  var SPServicesOfflineData = localStorage.getItem("SPServicesOfflineData");
     if(SPServicesOfflineData === null || SPServicesOfflineData === "null") {
        SPServicesOfflineData = {
          webURLs: {},
          defaults: {}
        };

		localStorage.setItem("SPServicesOfflineData", JSON.stringify(SPServicesOfflineData));
     }

  // Globals and constants from SPServices

  var SLASH     = "/";
  var SCHEMASharePoint  = "http://schemas.microsoft.com/sharepoint";
  // Caching
  var promisesCache = {};
  var i = 0;           // Generic loop counter

  //*** Begin SPOffline
  var spOfflineVersion = "1.0.0";
  // IE < 8, or IE8 with closed web console
  if(typeof window.console === "undefined") {
    window.console = {
      log: function(msg) { alert("console.log output:  " + msg); },
      info: function(msg) { alert("console.info output: " + msg); },
      warn: function(msg) { alert("console.warn output: " + msg); },
      error: function(msg) { alert("console.error output: " + msg); },
      clear: function(msg) { alert("console.clear output: " + msg); },
      table: function(arr) {/* noop */},
      dir: function(obj) {/* noop */},
      dirxml: function(doc) {/* noop */},
      assert: function(val) {/* noop */},
      debug: function(val) {/* noop */}
    };
  }

  // Some basic structures to build on, SPServices, and SPServicesOffline
  $.fn.SPServicesOffline = function() {
	// This currently doesn't do anything, besides remind the user of correct syntax
    
    console.log("Please use the syntax .SPServices, rather than .SPServicesOffline");
    return this;
  };

  $.fn.SPServicesOffline.Version = function() { return spOfflineVersion; };
  $.fn.SPServices = function(options) {

  var opts = {
      webURL: $().SPServices.SPGetCurrentSite(),
      completefunc: function(xData, Status) {}
    },
    options = $.extend(opts, options);

  options = $.extend($().SPServices.defaults, options);
  
  if(!options.operation) {
   console.error("Cannot read property 0 of undefined");
   console.log("The error above often means that you forgot the \"operation\" option when calling $().SPServices");
   return;
  }
  switch(options.operation) {
   case "UpdateListItems":
    if(options.batchCmd === "Delete")
     xData = deleteListItem(options);
    else if(options.batchCmd === "New")
     xData = createListItem(options);
    else
     xData = updateListItem(options);
    break;
   case "GetListItems":
    xData = getListItems(options);
    break;
  }
  
  options.completefunc(xData, "success");
 };
 var getListItems = function(options) {
  var CAMLQuery = options.CAMLQuery,
   CAMLViewFields = options.CAMLViewFields,
   
   webURL = options.webURL || $().SPServices.SPGetCurrentSite(),
   listName = options.listName;
  
  var SPServicesOfflineData = JSON.parse( localStorage.getItem("SPServicesOfflineData") ),
   list = SPServicesOfflineData.webURLs[webURL].lists[listName];
  var colArray = [];
  for(var col in list.columns) {
   colArray.push(col);
  }
  // Create a deep copy that we can mutate
  listItems = JSON.parse( JSON.stringify( list.listItems ) );
  // list.columns.hasOwnProperty(); //*** Only keep fields in CAMLViewFields, and filter out items that don't match the caml query
  return createResponseData(list.listItems, "Read");
 };
 var deleteListItem = function(options) {
  var itemId = options.ID;
  
  if(!itemId) {
   console.error("You must provide a list item ID to delete a list item.");
  }
  //*** Delete the list item from the localStorage
  var SPServicesOfflineData = JSON.parse( localStorage.getItem("SPServicesOfflineData") ),
   list = SPServicesOfflineData.webURLs[webURL].lists[listName];
  delete list.listItems[itemId];
  
  localStorage.setItem("SPServicesOfflineData", SPServicesOfflineData);
  return createResponseData(options, "Delete");
 };
 var createListItem = function(options) {
 
  var creator = $().SPServices.SPGetCurrentUser(),
   editor = $().SPServices.SPGetCurrentUser(),
   created = (new Date()),
   modified = created,
   webURL = options.webURL || $().SPServices.SPGetCurrentSite(),
   listName = options.listName,
   valuepairs = options.valuepairs || [],
   valuePairsObj = {};
  for(var i = 0; i < valuepairs.length; i++) {
   valuePairsObj[ valuepairs[i][1] ] = valuepairs[i][2];
  }
  valuePairsObj.Creator = creator;
  valuePairsObj.Editor = editor;
  valuePairsObj.Created = created;
  valuePairsObj.Modified = modified;
  var SPServicesOfflineData = JSON.parse( localStorage.getItem("SPServicesOfflineData") ),
   list = SPServicesOfflineData.webURLs[webURL].lists[listName];
  
  list.listItems.push(valuePairsObj);
  localStorage.setItem("SPServicesOfflineData", JSON.stringify( SPServicesOfflineData) );
  options.valuepairs = options.valuepairs.concat([
   ["Creator", creator],
   ["Editor", editor],
   ["Created", created],
   ["Modified", modified]
  ]);
  return createResponseData(options, "Create");
 };
 var updateListItem = function(options) {
  var editor = $().SPServices.SPGetCurrentUser(),
   modified = created,
   
   CAMLQuery = options.CAMLQuery, //*** Wait, is this even used here? Or the line below?
   CAMLViewFields = opions.CAMLViewFields,
   listItemId = options.ID,
   valuepairs = options.valuepairs,
   valuePairsObj = {};
  for(var i = 0; i < valuepairs.length; i++) {
   valuePairsObj[ valuepairs[i][1] ] = valuepairs[i][2];
  }
  valuePairsObj.Editor = editor;
  valuePairsObj.Modified = modified;
  var SPServicesOfflineData = JSON.parse( localStorage.getItem("SPServicesOfflineData") ),
   list = SPServicesOfflineData.webURLs[webURL].lists[listName],
   item = list[listItemId];
  for(var x in valuePairsObj) {
   item[x] = valuePairsObj[x];
  }
  localStorage.setItem("SPServicesOfflineData", JSON.stringify( SPServicesOfflineData) );
  options.valuepairs = options.valuepairs.concat([
   ["Editor", editor],
   ["Modified", modified]
  ]);
  return createResponseData(options, "Update");
 };
 // Defaults added as a function in our library means that the caller can override the defaults
 // for their session by calling this function.  Each operation requires a different set of options;
 // we allow for all in a standardized way.
 $.fn.SPServices.defaults = {
  cacheXML: false,   // If true, we'll cache the XML results for the call
  operation: "",    // The Web Service operation
  webURL: "",     // URL of the target Web
  makeViewDefault: false,  // true to make the view the default view for the list
  // For operations requiring CAML, these options will override any abstractions
  CAMLViewName: "",   // View name in CAML format.
  CAMLQuery: "",    // Query in CAML format
  CAMLViewFields: "",   // View fields in CAML format
  CAMLRowLimit: 0,   // Row limit as a string representation of an integer
  CAMLQueryOptions: "<QueryOptions></QueryOptions>",  // Query options in CAML format
  // Abstractions for CAML syntax
  batchCmd: "Update",   // Method Cmd for UpdateListItems
  valuepairs: [],    // Fieldname / Fieldvalue pairs for UpdateListItems
  // As of v0.7.1, removed all options which were assigned an empty string ("")
  DestinationUrls: [],  // Array of destination URLs for copy operations
  behavior: "Version3",  // An SPWebServiceBehavior indicating whether the client supports Windows SharePoint Services 2.0 or Windows SharePoint Services 3.0: {Version2 | Version3 }
  storage: "Shared",   // A Storage value indicating how the Web Part is stored: {None | Personal | Shared}
  objectType: "List",   // objectType for operations which require it
  cancelMeeting: true,  // true to delete a meeting;false to remove its association with a Meeting Workspace site
  nonGregorian: false,  // true if the calendar is set to a format other than Gregorian;otherwise, false.
  fClaim: false,    // Specifies if the action is a claim or a release. Specifies true for a claim and false for a release.
  recurrenceId: 0,   // The recurrence ID for the meeting that needs its association removed. This parameter can be set to 0 for single-instance meetings.
  sequence: 0,    // An integer that is used to determine the ordering of updates in case they arrive out of sequence. Updates with a lower-than-current sequence are discarded. If the sequence is equal to the current sequence, the latest update are applied.
  maximumItemsToReturn: 0, // SocialDataService maximumItemsToReturn
  startIndex: 0,    // SocialDataService startIndex
  isHighPriority: false,  // SocialDataService isHighPriority
  isPrivate: false,   // SocialDataService isPrivate
  rating: 1,     // SocialDataService rating
  maxResults: 10,    // Unless otherwise specified, the maximum number of principals that can be returned from a provider is 10.
  principalType: "User",  // Specifies user scope and other information: [None | User | DistributionList | SecurityGroup | SharePointGroup | All]
  async: true,    // Allow the user to force async
  completefunc: null   // Function to call on completion
 };
 // Build SPServicesOffline user/site defaults
 var generateNewUser = function(userInfo) {
  var users = localStorage.getItem("SPServicesOfflineUsers"),
   newUser = {};
  
  if(users === null || users === "null") {
   users = [{ // Create a placeholder for the zero-index of our users array. Note that list item IDs always start at 1.
    ID: 0,
    AccountName: "DoesNotExist"
   }];
  } else {
   users = JSON.parse(users);
   // We will not generate a brand new user if the username already exists. If it does, that user account will be returned.
   var existingUser = users.find(function(element, idx, arr) { return element.AccountName === userInfo.AccountName; });
   if(existingUser) {
    return existingUser;
   }
  }
  newUser.ID = users.length;
  for(var x in userInfo) {
   newUser[x] = userInfo[x];
  }
  users.push(newUser); // Now the next index (the new user id) maps to the username, and to any other user data that we decide to add later, like user groups
  localStorage.setItem("SPServicesOfflineUsers", JSON.stringify(users));
  return users[newUser.ID];
 };
 $.fn.SPServicesOffline.SPSetCurrentUser = function(userName) {
  localStorage.setItem("SPServicesOfflineCurrentUser", userName);
  currentUserId = generateNewUser({AccountName: userName}).ID;
  alert("Your user ID in our fake SharePoint site is " + currentUserId);
 };
 $.fn.SPServicesOffline.SPSetCurrentSite = function(sitePath) {
  localStorage.setItem("SPServicesOfflineCurrentSite", sitePath);
  alert("You will be using this as the current webURL: " + sitePath);
 };
 $.fn.SPServicesOffline.SPAddList = function(options) {
  if(!options || !options.listName) {
   console.error("To create a new list with $().SPServicesOffline.AddList you must at least provide a listName in the options. See documentation.");
   return;
  }
  var webURL = options.webURL || $().SPServices.SPGetCurrentSite(),
  listName = options.listName,
  description = options.description || "",
  templateID = options.templateID || null; // This is currently unused in SPServicesOffline

  /*
   See if storage item "SPServicesOfflineData" exists.   
  */
  
  var SPServicesOfflineData = JSON.parse(localStorage.getItem("SPServicesOfflineData"));

  if(!SPServicesOfflineData.webURLs[webURL])
   SPServicesOfflineData.webURLs[webURL] = {};
  var lists = SPServicesOfflineData.webURLs[webURL].lists;
  
  if(!lists) {
   lists = SPServicesOfflineData.webURLs[webURL].lists = {};
  }
  var thisList = SPServicesOfflineData.webURLs[webURL].lists[listName];
  if(!thisList) {
   SPServicesOfflineData.webURLs[webURL].lists[listName] = { //*** Perhaps also generate a list ID here.
    columns: {
     "ID": "Number",
     "Title": "Single line of text",
     "Description": "Multiple lines of text",
     "Created": "Date and Time",
     "Creator": "Person or Group",
     "Modified": "Date and Time",
     "Editor": "Person or Group"
    },
    listItems: [],
    listType: 1,
    listViews: [],
    defaultView: "Default" // If listViews.length === 0, defaultView will always be "Default"
   };
  }
  localStorage.setItem("SPServicesOfflineData", JSON.stringify(SPServicesOfflineData));
 };
 $.fn.SPServicesOffline.SPDeleteList = function(options) {
  if(!options || !options.listName) {
   console.error("To delete an existing list with $().SPServicesOffline.DeleteList you must at least provide a listName in the options. If no webURL is provided, the default webURL will be assumed. See documentation.");
   return;
  }
  var listName = options.listName,
   webURL = options.webURL;
  if(!webURL) {
   webURL = $().SPServices.SPGetCurrentSite();
  }
  //*** Might need two JSON.parse calls here 
  var SPServicesOfflineData = JSON.parse( localStorage.getItem("SPServicesOfflineData") );
  delete SPServicesOfflineData.webURLs[webURL].lists[listName];
  
  localStorage.setItem("SPServicesOfflineData", JSON.stringify( localStorage.getItem("SPServicesOfflineData")));
 };
 $.fn.SPServicesOffline.SPAddColumns = function(options) {
  if(!options || !options.listName) {
   console.error("To add new list fields with $().SPServicesOffline.AddColumns you must at least provide a listName in the options. See documentation.");
   return;
  }
  var listName = options.listName,
   webURL = options.webURL || $().SPServices.SPGetCurrentSite();
  var SPServicesOfflineData = JSON.parse( localStorage.getItem("SPServicesOfflineData") ),
   listToEdit = SPServicesOfflineData.webURLs[webURL].lists[listName];
  for(var column in options.columns) {
   listToEdit.columns[column] = options.columns[column];
  }
  localStorage.setItem("SPServicesOfflineData", JSON.stringify( localStorage.getItem("SPServicesOfflineData")));
 };
 var currentUser = localStorage.getItem("SPServicesOfflineCurrentUser"),
  currentSite = localStorage.getItem("SPServicesOfflineCurrentSite");
 if(!currentUser) { // null if not already set
  currentUser = window.prompt("What do you want SPServicesOffline to use as your 'current user' value?");
  
  if(typeof currentUser === "string") {
	  currentUser = currentUser.replace(/"|'|`/g, ""); // No need to add your own string quotes
  }
  
  $().SPServicesOffline.SPSetCurrentUser(currentUser);
 }
 if(!currentSite) { // null if not already set
  currentSite = window.prompt("What do you want SPServicesOffline to use as your 'current site' value?");
  
  if(typeof currentSite === "string") {
	  currentSite = currentSite.replace(/"|'|`/g, ""); // No need to add your own string quotes
  }

  $().SPServicesOffline.SPSetCurrentSite(currentSite);
 }

 // Some utility methods or testing and analysis - do not attempt to use this method in production
 $.fn.SPServicesOffline.SPGetListAsArray = function(listName) {
  // just returns the given list in array format. Currently unused.
 };
 // For testing and analysis - do not attempt to use this method in production
 $.fn.SPServicesOffline.SPGetListAsTable = function(listName) {
  // just returns the given list in table format. Currently unused.
 };
 $.fn.SPServicesOffline.SPExportList = function(listName) {  
  tableToExcel(
	$.fn.SPServicesOffline.SPGetListAsTable(listName).html(), listName, listName
  );
 };
 $.fn.SPServicesOffline.SPImportList = function(table) {
   // Currently unused
 };

 // Begin building mock SPServices methods
 // options - an array of names of data-ng- to import.
 $.fn.SPServices.SPGetCurrentUser = function(options) {
  if(typeof options === "object" && options.fieldNames) {
	  var detailsObj = JSON.parse(localStorage.getItem("SPServicesOfflineCurrentUserFields") || "{}"),
		returnDetailsObj = {};

	   for(var i = 0; i < options.fieldNames.length; i++) {
		returnDetailsObj[options.fieldNames[i]] = detailsObj[options.fieldNames[i]];
	   }

	  return returnDetailsObj;
  }

  // Default is just the username
  return localStorage.getItem("SPServicesOfflineCurrentUser") || "Guest";
 };
 
 $().SPServices.SPGetCurrentSite = function() {
  return localStorage.getItem("SPServicesOfflineCurrentSite") || (window.location.href.substring(0, window.location.href.lastIndexOf("/") - 1));
 };
 
 $.fn.SPServices.SPUpdateMultipeListItems = function(options) {
  // Parse CAMLQuery
  // Use a $.when.apply array
 
 };
 /**
  * @description Creates response data (including XML documents) to be returned in mock web service calls.
  * @param {Object|array} zRowItems - A plain JS object or an array, where the key is the list item ID, and the value is an object of field names pointing to their values.
  * @param {string} serviceType - "Create", "Read", "Update", "Delete"
  * @returns {Object} A plain JS object that includes responseXML, responseText, and some ajax-related properties.
  * @private
  */
 var createResponseData = function(zRowItems, serviceType) {
  var numItemsReturned = 0;
 
  if(zRowItems instanceof Array) {
   numItemsReturned = zRowItems.length;
  } else {
   numItemsReturned = Object.keys(zRowItems).length;
  }
  /*
   Do a double loop here.
   Outer loop: loop through the zRowItems array, in the order it was returned.
   Inner loop: do a for-in loop on an attributes JS object, defining all of the attributes on that given object. If one of the attributes is undefined, don't have it show up in the XML for that item.  
  
  */
  
  // For GetListItems xData.responseText
  var getTemplate = `<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema"><soap:Body><GetListItemsResponse xmlns="http://schemas.microsoft.com/sharepoint/soap/"><GetListItemsResult><listitems xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882' xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882' xmlns:rs='urn:schemas-microsoft-com:rowset' xmlns:z='#RowsetSchema'><rs:data ItemCount="${numItemsReturned}">` + 
        
        /*
        <z:row ows_Attachments='0' ows_ID='1' ows_Author='1#;${UserName}' ow_Created='' ows_Modified='${formatDateToSp(new Date())}' />
        */
        // Basically in here we will have a loop putting data about each returned item. Either use ViewFields requested, or return all available fields for that list.
        // Only include the most common/useful list fields, and 
        
       `</rs:data></listitems></GetListItemsResult></GetListItemsResponse></soap:Body></soap:Envelope>`;
  // UpdateListItems (Update) xData.responseText
  var updateTemplate = `<?xml version="1.0" encoding="utf-8"?>
   <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soap:Body>
     <UpdateListItemsResponse xmlns="http://schemas.microsoft.com/sharepoint/soap/">
      <UpdateListItemsResult>
       <Results>
        <Result ID="1,Update">
         <ErrorCode>
          0x00000000 // This number will determine if there was an error and what type. Try to note if wrong listname, nonexistent field, etc.
         </ErrorCode>
         // There should only be one item in the response - Note that ampersands are converted to &amp; probably similar for other STSHtmlEncode characters
         // Include all of the fields in the list I suppose. Have them point to the NEW (updated) field values, not the old ones.
         <z:row ows_ContentTypeId="0x0100A94FB350658A4348848D4F8B6DAF843A" ows_FMDAnalyst="par" ows_TransactionDate="2017-06-26 12:00:00" ows_FY="2017.00000000000" ows_Appn="D&amp;CP" ows_AppnCode="2017 - NONE - 19___701130003" ows_AllotmentCode="1019" ows_TransactionType="Fob " ows_Requestor="test" ows_TrackingNum="11" ows_Amount="1.00000000000000" ows_Notes="dsfds" ows_SrcInvestment="Bureau Executive &amp; Management Support" ows_SrcOffice="ABC" ows_SrcLocation="ABC" ows_SrcProjectCode="IM011604" ows_SrcOrgCode="181010" ows_FunctionCode="1234" ows_DestInvestment="Bureau Executive &amp; Management Support" ows_DestOffice="ABC" ows_DestLocation="ABC" ows_DestProjectCode="IME02S01" ows_DestOrgCode="180000" ows_Status="Pending" ows_TransactionStatus="Analyst Review" ows_FMDNotes="tes" ows_PONotes="test" ows_ID="6" ows_ContentType="Item" ows_Modified="2017-07-06 10:54:09" ows_Created="2017-06-15 13:32:09" ows_Author="1;#Long, Timothy S" ows_Editor="1;#Long, Timothy S" ows_owshiddenversion="20" ows_WorkflowVersion="1" ows__UIVersion="512" ows__UIVersionString="1.0" ows_Attachments="0" ows__ModerationStatus="0" ows_SelectTitle="6" ows_Order="600.000000000000" ows_GUID="{145361B1-D636-4833-B88A-B3F6FB56A510}" ows_FileRef="6;#sites/bmp/SPO/FMD/AMS/DEV/DB01/Lists/TelecommunicationTracking/6_.000" ows_FileDirRef="6;#sites/bmp/SPO/FMD/AMS/DEV/DB01/Lists/TelecommunicationTracking" ows_Last_x0020_Modified="6;#2017-06-15 13:32:09" ows_Created_x0020_Date="6;#2017-06-15 13:32:09" ows_FSObjType="6;#0" ows_SortBehavior="6;#0" ows_PermMask="0x7fffffffffffffff" ows_FileLeafRef="6;#6_.000" ows_UniqueId="6;#{0B45A903-442B-46D2-BD23-90E3D21AEFAE}" ows_ProgId="6;#" ows_ScopeId="6;#{7A4EEB55-A2DD-4502-B000-B93BAC46CBD8}" ows__EditMenuTableStart="6_.000" ows__EditMenuTableStart2="6" ows__EditMenuTableEnd="6" ows_LinkFilenameNoMenu="6_.000" ows_LinkFilename="6_.000" ows_LinkFilename2="6_.000" ows_ServerUrl="/sites/bmp/SPO/FMD/AMS/DEV/DB01/Lists/TelecommunicationTracking/6_.000" ows_EncodedAbsUrl="http://irm.m.state.sbu/sites/bmp/SPO/FMD/AMS/DEV/DB01/Lists/TelecommunicationTracking/6_.000" ows_BaseName="6_" ows_MetaInfo="6;#" ows__Level="1" ows__IsCurrentVersion="1" ows_ItemChildCount="6;#0" ows_FolderChildCount="6;#0" xmlns:z="#RowsetSchema" />
        </Result>
       </Results>
      </UpdateListItemsResult>
     </UpdateListItemsResponse>
    </soap:Body>
   </soap:Envelope>`;
  // UpdateListItems (Delete) xData.responseText
  var deleteTemplate = `<?xml version="1.0" encoding="utf-8"?>
   <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soap:Body>
     <UpdateListItemsResponse xmlns="http://schemas.microsoft.com/sharepoint/soap/">
      <UpdateListItemsResult>
       <Results>
        <Result ID="1,Delete"> // 1 must describe the operation. Strangely, the list item ID (or any info) does not seem to show up in this result after successful deletion.
         <ErrorCode>
          0x00000000
         </ErrorCode>
        </Result>
       </Results>
      </UpdateListItemsResult>
     </UpdateListItemsResponse>
    </soap:Body>
   </soap:Envelope>`;
  // UpdateListItems (New) xData.responseText
  var newTemplate = `<?xml version="1.0" encoding="utf-8"?>
   <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soap:Body>
     <UpdateListItemsResponse xmlns="http://schemas.microsoft.com/sharepoint/soap/">
      <UpdateListItemsResult>
       <Results>
        <Result ID="1,New">
         <ErrorCode>
          0x00000000
         </ErrorCode>
         <ID /> // Will this always be an empty tag?
         // I suppose there will always be one z:row element in the response - make sure it includes the new ows_ID and numerous default fields (and fields with default values), plus
         // any new fields that were included in the valuepairs array (but I don't think we need to include custom fields that aren't in the valuepairs array, and don't have custom values.
         <z:row ows_ContentTypeId="0x0100A94FB350658A4348848D4F8B6DAF843A" ows_TrackingNum="42222" ows_Status="Requested" ows_TransactionStatus="New" ows_ID="58" ows_ContentType="Item" ows_Modified="2017-07-06 11:36:31" ows_Created="2017-07-06 11:36:31" ows_Author="1;#Long, Timothy S" ows_Editor="1;#Long, Timothy S" ows_owshiddenversion="1" ows_WorkflowVersion="1" ows__UIVersion="512" ows__UIVersionString="1.0" ows_Attachments="0" ows__ModerationStatus="0" ows_SelectTitle="58" ows_Order="5800.00000000000" ows_GUID="{42E9AEA3-2BED-40F8-A708-B5F58F68711A}" ows_FileRef="58;#sites/bmp/SPO/FMD/AMS/DEV/DB01/Lists/TelecommunicationTracking/58_.000" ows_FileDirRef="58;#sites/bmp/SPO/FMD/AMS/DEV/DB01/Lists/TelecommunicationTracking" ows_Last_x0020_Modified="58;#2017-07-06 11:36:31" ows_Created_x0020_Date="58;#2017-07-06 11:36:31" ows_FSObjType="58;#0" ows_SortBehavior="58;#0" ows_PermMask="0x7fffffffffffffff" ows_FileLeafRef="58;#58_.000" ows_UniqueId="58;#{98CDB37F-71FC-4358-A7B3-A8AD737FB711}" ows_ProgId="58;#" ows_ScopeId="58;#{7A4EEB55-A2DD-4502-B000-B93BAC46CBD8}" ows__EditMenuTableStart="58_.000" ows__EditMenuTableStart2="58" ows__EditMenuTableEnd="58" ows_LinkFilenameNoMenu="58_.000" ows_LinkFilename="58_.000" ows_LinkFilename2="58_.000" ows_ServerUrl="/sites/bmp/SPO/FMD/AMS/DEV/DB01/Lists/TelecommunicationTracking/58_.000" ows_EncodedAbsUrl="http://irm.m.state.sbu/sites/bmp/SPO/FMD/AMS/DEV/DB01/Lists/TelecommunicationTracking/58_.000" ows_BaseName="58_" ows_MetaInfo="58;#" ows__Level="1" ows__IsCurrentVersion="1" ows_ItemChildCount="58;#0" ows_FolderChildCount="58;#0" xmlns:z="#RowsetSchema" />
        </Result>
       </Results>
      </UpdateListItemsResult>
     </UpdateListItemsResponse>
    </soap:Body>
   </soap:Envelope>`;
  // Error response for trying to Delete or Update an item with an ID that does not exist
  var updateError = `<?xml version="1.0" encoding="utf-8"?>
   <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soap:Body>
     <UpdateListItemsResponse xmlns="http://schemas.microsoft.com/sharepoint/soap/">
      <UpdateListItemsResult>
       <Results>
        <Result ID="1,Delete">
         <ErrorCode>
          0x81020016
         </ErrorCode>
         <ErrorText>
          Item does not exist.
          The page you selected contains an item that does not exist.  It may have been deleted by another user.
         </ErrorText>
        </Result>
       </Results>
      </UpdateListItemsResult>
     </UpdateListItemsResponse>
    </soap:Body>
   </soap:Envelope>`;
  // Error response for trying to call a list that does not exist. This is preceded by a console.error:
  // "POST http:// ... the url of the lists, which is webURL + '_vti_bin/Lists.asmx' 500 (Internal Server Error)
  var getError = `<?xml version="1.0" encoding="utf-8"?>
   <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soap:Body>
     <soap:Fault>
      <faultcode>
       soap:Server
      </faultcode>
      <faultstring>
       Exception of type 'Microsoft.SharePoint.SoapServer.SoapServerException' was thrown.
      </faultstring>
      <detail>
       <errorstring xmlns="http://schemas.microsoft.com/sharepoint/soap/">
        List does not exist.
        The page you selected contains a list that does not exist.  It may have been deleted by another user.
       </errorstring>
       <errorcode xmlns="http://schemas.microsoft.com/sharepoint/soap/">
        0x82000006
       </errorcode>
      </detail>
     </soap:Fault>
    </soap:Body>
   </soap:Envelope>`;
  // This was the result of a GetListItems call with CAMLQuery set to "You suck". This was preceded by a console error:
  // "POST http:// ... the url of the lists, which is webURL + '_vti_bin/Lists.asmx' 500 (Internal Server Error)
  // On the other hand, when I used <Query>You suck</Query> it returned all the list items.
  var camlQueryError = `
   <?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
     <soap:Body>
      <soap:Fault>
       <faultcode>
        soap:Server
       </faultcode>
       <faultstring>
        Exception of type 'Microsoft.SharePoint.SoapServer.SoapServerException' was thrown.
       </faultstring>
       <detail>
        <errorstring xmlns="http://schemas.microsoft.com/sharepoint/soap/">
         Element &lt;Query&gt; of parameter query is missing or invalid.
        </errorstring>
        <errorcode xmlns="http://schemas.microsoft.com/sharepoint/soap/">
         0x82000000
        </errorcode>
       </detail>
      </soap:Fault>
     </soap:Body>
    </soap:Envelope>`;
  // This was the result of a GetListItems call, trying to get an item with an ID that obviously didn't exist (50000)
  var getWithNoReturn = `<?xml version="1.0" encoding="utf-8"?>
   <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soap:Body>
     <GetListItemsResponse xmlns="http://schemas.microsoft.com/sharepoint/soap/">
      <GetListItemsResult>
       <listitems xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882'
        xmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'
        xmlns:rs='urn:schemas-microsoft-com:rowset'
        xmlns:z='#RowsetSchema'>
        <rs:data ItemCount="0">
        </rs:data>
       </listitems>
      </GetListItemsResult>
     </GetListItemsResponse>
    </soap:Body>
   </soap:Envelope>`;
  /*
   //*** Here, first check if there is an error in the CAMLQuery.
  */
  switch(serviceType) {
   case "Read":
    var responseXML = $.parseXML(getTemplate);
    return {
     responseText: getTemplate,
     responseXML: responseXML,
     status: 200,
     statusText: "OK",
     readyState: 4,
     then: function(callback) { callback(responseXML); },
     done: function(callback) { callback(responseXML); }   
    };
    break;
   case "Update":
    var responseXML = $.parseXML(updateTemplate);
    return {
     responseText: getTemplate,
     responseXML: responseXML,
     status: 200,
     statusText: "OK",
     readyState: 4,
     then: function(callback) { callback(responseXML); },
     done: function(callback) { callback(responseXML); }   
    };
    break;
   case "Create":
    var responseXML = $.parseXML(newTemplate);
    return {
     responseText: getTemplate,
     responseXML: responseXML,
     status: 200,
     statusText: "OK",
     readyState: 4,
     then: function(callback) { callback(responseXML); },
     done: function(callback) { callback(responseXML); }   
    };
    break;
   case "Delete":
    var responseXML = $.parseXML(deleteTemplate);
    return {
     responseText: getTemplate,
     responseXML: responseXML,
     status: 200,
     statusText: "OK",
     readyState: 4,
     then: function(callback) { callback(responseXML); },
     done: function(callback) { callback(responseXML); }   
    };
    break;
   default: {}   
  }
 };
 // SPServicesOffline: These are (mostly) taken directly from the SPServices code.
 // Convert a JavaScript date to the ISO 8601 format required by SharePoint to update list items
 $.fn.SPServices.SPConvertDateToISO = function (options) {
  var opt = $.extend({}, {
   dateToConvert: new Date(),  // The JavaScript date we'd like to convert. If no date is passed, the function returns the current date/time
   dateOffset: "-05:00"   // The time zone offset requested. Default is EST
  }, options);
  //Generate ISO 8601 date/time formatted string
  var s = "";
  var d = opt.dateToConvert;
  s += d.getFullYear() + "-";
  s += pad(d.getMonth() + 1) + "-";
  s += pad(d.getDate());
  s += "T" + pad(d.getHours()) + ":";
  s += pad(d.getMinutes()) + ":";
  s += pad(d.getSeconds()) + "Z" + opt.dateOffset;
  //Return the ISO8601 date string
  return s;
 }; // End $.fn.SPServices.SPConvertDateToISO
 // An inverse function for SPConvertDateToISO - do not use this in production
 $.fn.SPServices.SPConvertISOToDate = function(datestring) {
  var splitOnT = datestring.split("T"),
   calendarDate = splitOnT[0].split("-"),
   timeAndOffset = splitOnT[1].split("Z"),
   time = timeAndOffset[0].split(":"),
   offset = timeAndOffset[1].split(":"),
   offsetHours = offset[0],
   offsetMinutes = offset[1];
  
  return new Date(calendarDate[0], calendarDate[1] - 1, calendarDate[2], parseInt(time[0], 10), parseInt(time[1], 10), parseInt(time[2], 10));
 };
 // This method for finding specific nodes in the returned XML was developed by Steve Workman. See his blog post
 // http://www.steveworkman.com/html5-2/javascript/2011/improving-javascript-xml-node-finding-performance-by-2000/
 // for performance details.
 $.fn.SPFilterNode = function(name) {
  return this.find('*').filter(function() {
   return this.nodeName === name;
  });
 }; // End $.fn.SPFilterNode

 // After running into issues with the SPServices version, we are replacing it here with code found on David Walsh's site: https://davidwalsh.name/convert-xml-json
 $.fn.SPXmlToJson = function(options) {
  
  // Create the return object
  var obj = {};
  if (xml.nodeType == 1) { // element
   // do attributes
   if (xml.attributes.length > 0) {
   obj["@attributes"] = {};
    for (var j = 0; j < xml.attributes.length; j++) {
     var attribute = xml.attributes.item(j);
     obj["@attributes"][attribute.nodeName] = attribute.nodeValue;
    }
   }
  } else if (xml.nodeType == 3) { // text
   obj = xml.nodeValue;
  }
  // do children
  if (xml.hasChildNodes()) {
   for(var i = 0; i < xml.childNodes.length; i++) {
    var item = xml.childNodes.item(i);
    var nodeName = item.nodeName;
    if (typeof(obj[nodeName]) == "undefined") {
     obj[nodeName] = xmlToJson(item);
    } else {
     if (typeof(obj[nodeName].push) == "undefined") {
      var old = obj[nodeName];
      obj[nodeName] = [];
      obj[nodeName].push(old);
     }
     obj[nodeName].push(xmlToJson(item));
    }
   }
  }
  return obj;
 };
 function attrToJson(v, objectType) {
  var colValue;
  switch (objectType) {
   case "DateTime":
   case "datetime": // For calculated columns, stored as datetime;#value
    // Dates have dashes instead of slashes: ows_Created="2009-08-25 14:24:48"
    colValue = dateToJsonObject(v);
    break;
   case "User":
    colValue = userToJsonObject(v);
    break;
   case "UserMulti":
    colValue = userMultiToJsonObject(v);
    break;
   case "Lookup":
    colValue = lookupToJsonObject(v);
    break;
   case "LookupMulti":
    colValue = lookupMultiToJsonObject(v);
    break;
   case "Boolean":
    colValue = booleanToJsonObject(v);
    break;
   case "Integer":
    colValue = intToJsonObject(v);
    break;
   case "Counter":
    colValue = intToJsonObject(v);
    break;
   case "MultiChoice":
    colValue = choiceMultiToJsonObject(v);
    break;
   case "Currency":
   case "float":       // For calculated columns, stored as float;#value
    colValue = floatToJsonObject(v);
    break;
   case "Calc":
    colValue = calcToJsonObject(v);
    break;
   default:
    // All other objectTypes will be simple strings
    colValue = stringToJsonObject(v);
    break;
  }
  return colValue;
 }
 function stringToJsonObject(s) {
  return s;
 }
 function intToJsonObject(s) {
  return parseInt(s, 10);
 }
 function floatToJsonObject(s) {
  return parseFloat(s);
 }
 function booleanToJsonObject(s) {
  var out = s === "0" ? false : true;
  return out;
 }
 function dateToJsonObject(s) {
  return new Date(s.replace(/-/g, "/"));
 }
   function userToJsonObject(s) {
        if (s.length === 0) {
   return null;
        } else {
            var thisUser = new SplitIndex(s); 
            var thisUserExpanded = thisUser.value.split(",#");
            if(thisUserExpanded.length === 1) {
                return {userId: thisUser.Id, userName: thisUser.value};
            } else {
                return {
     userId: thisUser.Id, 
     userName: thisUserExpanded[0].replace( /(,,)/g, ","), 
     loginName: thisUserExpanded[1].replace( /(,,)/g, ","), 
     email: thisUserExpanded[2].replace( /(,,)/g, ","), 
     sipAddress: thisUserExpanded[3].replace( /(,,)/g, ","), 
     title: thisUserExpanded[4].replace( /(,,)/g, ",")
    };
            }
        }
    }
 function userMultiToJsonObject(s) {
  if(s.length === 0) {
   return null;
  } else {
   var thisUserMultiObject = [];
   var thisUserMulti = s.split(";#");
   for(i=0; i < thisUserMulti.length; i=i+2) {
    var thisUser = userToJsonObject(thisUserMulti[i] + ";#" + thisUserMulti[i+1]);
    thisUserMultiObject.push(thisUser);
   }
   return thisUserMultiObject;
  }
 }
 function lookupToJsonObject(s) {
  if(s.length === 0) {
   return null;
  } else {
   var thisLookup = new SplitIndex(s);
   return {lookupId: thisLookup.id, lookupValue: thisLookup.value};
  }
 }
 function lookupMultiToJsonObject(s) {
  if(s.length === 0) {
   return null;
  } else {
   var thisLookupMultiObject = [];
   var thisLookupMulti = s.split(";#");
   for(i=0; i < thisLookupMulti.length; i=i+2) {
    var thisLookup = lookupToJsonObject(thisLookupMulti[i] + ";#" + thisLookupMulti[i+1]);
    thisLookupMultiObject.push(thisLookup);
   }
   return thisLookupMultiObject;
  }
 }
 function choiceMultiToJsonObject(s) {
  if(s.length === 0) {
   return null;
  } else {
   var thisChoiceMultiObject = [];
   var thisChoiceMulti = s.split(";#");
   for(i=0; i < thisChoiceMulti.length; i++) {
    if(thisChoiceMulti[i].length !== 0) {
     thisChoiceMultiObject.push(thisChoiceMulti[i]);
    }
   }
   return thisChoiceMultiObject;
  }
 }
 function calcToJsonObject(s) {
  if(s.length === 0) {
   return null;
  } else {
   var thisCalc = s.split(";#");
   // The first value will be the calculated column value type, the second will be the value
   return attrToJson(thisCalc[1], thisCalc[0]);
  }
 }
 
  // Utility function to show the results of a Web Service call formatted well in the browser.
 $.fn.SPServices.SPDebugXMLHttpResult = function(options) {
  var opt = $.extend({}, {
   node: null,       // An XMLHttpResult object from an ajax call
   indent: 0       // Number of indents
  }, options);
  var i;
  var NODE_TEXT = 3;
  var NODE_CDATA_SECTION = 4;
  var outString = "";
  // For each new subnode, begin rendering a new TABLE
  outString += "<table class='ms-vb' style='margin-left:" + opt.indent * 3 + "px;' width='100%'>";
  // DisplayPatterns are a bit unique, so let's handle them differently
  if(opt.node.nodeName === "DisplayPattern") {
   outString += "<tr><td width='100px' style='font-weight:bold;'>" + opt.node.nodeName +
    "</td><td><textarea readonly='readonly' rows='5' cols='50'>" + opt.node.xml + "</textarea></td></tr>";
  // A node which has no children
  } else if (!opt.node.hasChildNodes()) {
   outString += "<tr><td width='100px' style='font-weight:bold;'>" + opt.node.nodeName +
    "</td><td>" + ((opt.node.nodeValue !== null) ? checkLink(opt.node.nodeValue) : "&nbsp;") + "</td></tr>";
   if (opt.node.attributes) {
    outString += "<tr><td colspan='99'>" + showAttrs(opt.node) + "</td></tr>";
   }
  // A CDATA_SECTION node
  } else if (opt.node.hasChildNodes() && opt.node.firstChild.nodeType === NODE_CDATA_SECTION) {
   outString += "<tr><td width='100px' style='font-weight:bold;'>" + opt.node.nodeName +
    "</td><td><textarea readonly='readonly' rows='5' cols='50'>" + opt.node.parentNode.text + "</textarea></td></tr>";
  // A TEXT node
  } else if (opt.node.hasChildNodes() && opt.node.firstChild.nodeType === NODE_TEXT) {
   outString += "<tr><td width='100px' style='font-weight:bold;'>" + opt.node.nodeName +
    "</td><td>" + checkLink(opt.node.firstChild.nodeValue) + "</td></tr>";
  // Handle child nodes
  } else {
   outString += "<tr><td width='100px' style='font-weight:bold;' colspan='99'>" + opt.node.nodeName + "</td></tr>";
   if (opt.node.attributes) {
    outString += "<tr><td colspan='99'>" + showAttrs(opt.node) + "</td></tr>";
   }
   // Since the node has child nodes, recurse
   outString += "<tr><td>";
   for (i = 0;i < opt.node.childNodes.length; i++) {
    outString += $().SPServices.SPDebugXMLHttpResult({
     node: opt.node.childNodes.item(i),
     indent: opt.indent + 1
    });
   }
   outString += "</td></tr>";
  }
  outString += "</table>";
  // Return the HTML which we have built up
  return outString;
 }; // End $.fn.SPServices.SPDebugXMLHttpResult
  // Show a single attribute of a node, enclosed in a table
 //   node    The XML node
 //   opt    The current set of options
 function showAttrs(node) {
  var i;
  var out = "<table class='ms-vb' width='100%'>";
  for (i=0; i < node.attributes.length; i++) {
   out += "<tr><td width='10px' style='font-weight:bold;'>" + i + "</td><td width='100px'>" +
    node.attributes.item(i).nodeName + "</td><td>" + checkLink(node.attributes.item(i).nodeValue) + "</td></tr>";
  }
  out += "</table>";
  return out;
 } // End of function showAttrs
 //*** These might not all be necessary:
  // Find a dropdown (or multi-select) in the DOM. Returns the dropdown onject and its type:
 // S = Simple (select);C = Compound (input + select hybrid);M = Multi-select (select hybrid)
 function DropdownCtl(colName) {
  // Simple
  if((this.Obj = $("select[Title='" + colName + "']")).length === 1) {
   this.Type = "S";
  // Compound
  } else if((this.Obj = $("input[Title='" + colName + "']")).length === 1) {
   this.Type = "C";
  // Multi-select: This will find the multi-select column control in English and most other languages sites where the Title looks like 'Column Name possible values'
  } else if((this.Obj = $("select[ID$='SelectCandidate'][Title^='" + colName + " ']")).length === 1) {
   this.Type = "M";
  // Multi-select: This will find the multi-select column control on a Russian site (and perhaps others) where the Title looks like 'Выбранных значений: Column Name'
  } else if((this.Obj = $("select[ID$='SelectCandidate'][Title$=': " + colName + "']")).length === 1) {
   this.Type = "M";
  // Multi-select: This will find the multi-select column control on a German site (and perhaps others) where the Title looks like 'Mögliche Werte für &quot;Column name&quot;.'
  } else if((this.Obj = $("select[ID$='SelectCandidate'][Title$='\"" + colName + "\".']")).length === 1) {
   this.Type = "M";
  // Multi-select: This will find the multi-select column control on a Italian site (and perhaps others) where the Title looks like "Valori possibili Column name"
  } else if((this.Obj = $("select[ID$='SelectCandidate'][Title$=' " + colName + "']")).length === 1) {
   this.Type = "M";
  } else {
   this.Type = null;
  }
 } // End of function DropdownCtl

 // Find the MultiLookupPickerdata input element. The structures are slightly different in 2013 vs. prior versions.
 function MultiLookupPicker(o) {
  // Find input element that contains 'MultiLookup' and ends with 'data'. This holds all available values.
   this.MultiLookupPickerdata = o.closest("span").find("input[id*='MultiLookup'][id$='data']");
  // The ids in 2013 are different than prior versions, so we need to parse them out.
  var thisMultiLookupPickerdataId = this.MultiLookupPickerdata.attr("id");
   var thisIdEndLoc = thisMultiLookupPickerdataId.indexOf("Multi");
   var thisIdEnd = thisMultiLookupPickerdataId.substr(thisIdEndLoc);
   var thisMasterId = thisMultiLookupPickerdataId.substr(0, thisIdEndLoc) + thisIdEnd.substr(0, thisIdEnd.indexOf("_") + 1) + "m";
   this.master = window[thisMasterId];
 } // End of function MultiLookupPicker

 // Returns the selected value(s) for a dropdown in an array. Expects a dropdown object as returned by the DropdownCtl function.
 // If matchOnId is true, returns the ids rather than the text values for the selection options(s).
 function getDropdownSelected(columnSelect, matchOnId) {
  var columnSelectSelected = [];
  
  switch(columnSelect.Type) {
   case "S":
    if(matchOnId) {
     columnSelectSelected.push(columnSelect.Obj.find("option:selected").val() || []);
    } else {
     columnSelectSelected.push(columnSelect.Obj.find("option:selected").text() || []);
    }
    break;
   case "C":
    if(matchOnId) {
     columnSelectSelected.push($("input[id='"+ columnSelect.Obj.attr("optHid") + "']").val() || []);
    } else {
     columnSelectSelected.push(columnSelect.Obj.attr("value") || []);
    }    
    break;
   case "M":
    var columnSelections = columnSelect.Obj.closest("span").find("select[ID$='SelectResult']");
    $(columnSelections).find("option").each(function() {
     columnSelectSelected.push($(this).html());
    });
    break;
   default:
    break;
  }
  return columnSelectSelected;
 } // End of function getDropdownSelected
 
 // Build an error message based on passed parameters
 function errBox(func, param, msg) {
  var errMsg = "<b>Error in function</b><br/>" + func + "<br/>" +
   "<b>Parameter</b><br/>" + param + "<br/>" +
   "<b>Message</b><br/>" + msg + "<br/><br/>" +
   "<span onmouseover='this.style.cursor=\"hand\";' onmouseout='this.style.cursor=\"inherit\";' style='width=100%;text-align:right;'>Click to continue</span></div>";
  modalBox(errMsg);
 } // End of function errBox
 // Call this function to pop up a branded modal msgBox
 function modalBox(msg) {
  var boxCSS = "position:absolute;width:300px;height:150px;padding:10px;background-color:#000000;color:#ffffff;z-index:30;font-family:'Arial';font-size:12px;display:none;";
  $("#aspnetForm").parent().append("<div id='SPServices_msgBox' style=" + boxCSS + ">" + msg);
  var height = $("#SPServices_msgBox").height();
  var width = $("#SPServices_msgBox").width();
  var leftVal = ($(window).width() / 2) - (width / 2) + "px";
  var topVal = ($(window).height() / 2) - (height / 2) - 100 + "px";
  $("#SPServices_msgBox").css({border:'5px #C02000 solid', left:leftVal, top:topVal}).show().fadeTo("slow", 0.75).click(function () {
   $(this).fadeOut("3000", function () {
    $(this).remove();
   });
  });
 } // End of function modalBox
 // Generate a unique id for a containing div using the function name and the column display name
 function genContainerId(funcname, columnName) {
  return funcname + "_" + $().SPServices.SPGetStaticFromDisplay({
   listName: $().SPServices.SPListNameFromUrl(),
   columnDisplayName: columnName
  });
 } // End of function genContainerId
 
 // Get the URL for a specified form for a list
 function getListFormUrl(l, f) {
  var u;
  $().SPServices({
   operation: "GetFormCollection",
   async: false,
   listName: l,
   completefunc: function (xData) {
    u = $(xData.responseXML).find("Form[Type='" + f + "']").attr("Url");
   }
  });
  return u;
 } // End of function getListFormUrl
 // Add the option values to the SOAPEnvelope.payload for the operation
 // opt = options for the call
 // paramArray = an array of option names to add to the payload
 //  "paramName" if the parameter name and the option name match
 //  ["paramName", "optionName"] if the parameter name and the option name are different (this handles early "wrappings" with inconsistent naming)
 function addToPayload(opt, paramArray) {
  var i;
  for (i=0; i < paramArray.length; i++) {
   // the parameter name and the option name match
   if(typeof paramArray[i] === "string") {
    SOAPEnvelope.payload += wrapNode(paramArray[i], opt[paramArray[i]]);
   // the parameter name and the option name are different 
   } else if(paramArray[i].length === 2) {
    SOAPEnvelope.payload += wrapNode(paramArray[i][0], opt[paramArray[i][1]]);
   // something isn't right, so report it
   } else {
    errBox(opt.operation, "paramArray[" + i + "]: " + paramArray[i], "Invalid paramArray element passed to addToPayload()");
   }
  }
 } // End of function addToPayload
 // Finds the td which contains a form field in default forms using the comment which contains:
 // <!--  FieldName="Title"
 //  FieldInternalName="Title"
 //  FieldType="SPFieldText"
 // -->
 // as the "anchor" to find it. Necessary because SharePoint doesn't give all field types ids or specific classes.
 function findFormField(columnName) {
  var thisFormBody;
  // There's no easy way to find one of these columns; we'll look for the comment with the columnName
  var searchText = RegExp("FieldName=\"" + columnName.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, "\\$&") + "\"", "gi");
  // Loop through all of the ms-formbody table cells
  $("td.ms-formbody, td.ms-formbodysurvey").each(function() {
   // Check for the right comment
   if(searchText.test($(this).html())) {
    thisFormBody = $(this);
    // Found it, so we're done
    return false;
   }
  });
  return thisFormBody;
 } // End of function findFormField
 // The SiteData operations have the same names as other Web Service operations. To make them easy to call and unique, I'm using
 // the SiteData prefix on their names. This function replaces that name with the right name in the SOAPEnvelope.
 function siteDataFixSOAPEnvelope(SOAPEnvelope, siteDataOperation) {
  var siteDataOp = siteDataOperation.substring(8);
  SOAPEnvelope.opheader = SOAPEnvelope.opheader.replace(siteDataOperation, siteDataOp);
  SOAPEnvelope.opfooter = SOAPEnvelope.opfooter.replace(siteDataOperation, siteDataOp);
  return SOAPEnvelope;
 } // End of function siteDataFixSOAPEnvelope
 // Wrap an XML node (n) around a value (v)
 function wrapNode(n, v) {
  var thisValue = typeof v !== "undefined" ? v : "";
  return "<" + n + ">" + thisValue + "</" + n + ">";
 }
 // Generate a random number for sorting arrays randomly
 function randOrd() {
  return (Math.round(Math.random())-0.5);
 }
 // If a string is a URL, format it as a link, else return the string as-is
 function checkLink(s) {
  return ((s.indexOf("http") === 0) || (s.indexOf(SLASH) === 0)) ? "<a href='" + s + "'>" + s + "</a>" : s;
 }
 // Get the filename from the full URL
 function fileName(s) {
  return s.substring(s.lastIndexOf(SLASH)+1,s.length);
 }
/* Taken from http://dracoblue.net/dev/encodedecode-special-xml-characters-in-javascript/155/ */
 var xml_special_to_escaped_one_map = {
  '&': '&amp;',
  '"': '&quot;',
  '<': '&lt;',
  '>': '&gt;'};
 var escaped_one_to_xml_special_map = {
  '&amp;': '&',
  '&quot;': '"',
  '&lt;': '<',
  '&gt;': '>'};
 
 function encodeXml(string) {
  return string.replace(/([\&"<>])/g, function(str, item) {
   return xml_special_to_escaped_one_map[item];
  });
 }
 function decodeXml(string) {
  return string.replace(/(&quot;|&lt;|&gt;|&amp;)/g,
   function(str, item) {
    return escaped_one_to_xml_special_map[item];
  });
 }
/* Taken from http://dracoblue.net/dev/encodedecode-special-xml-characters-in-javascript/155/ */
 // Escape column values
 function escapeColumnValue(s) {
  if(typeof s === "string") {
   return s.replace(/&(?![a-zA-Z]{1,8};)/g, "&amp;");
  } else {
   return s;
  }
 }
 // Escape Url
 function escapeUrl(u) {
  return u.replace(/&/g,'%26');
 }
 // Split values like 1;#value into id and value        
 function SplitIndex(s) {
  var spl = s.split(";#");
  this.id = spl[0];
  this.value = spl[1];
 }
 function pad(n) {
  return n < 10 ? "0" + n : n;
 }

var tableToExcel = (function () {
        var uri = 'data:application/vnd.ms-excel;base64,'
        , template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"/><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>'
        , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
        , format = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
        return function (table, name, filename) {
            if (!table.nodeType) table = document.getElementById(table)
            var ctx = { worksheet: name || 'Worksheet', table: table.innerHTML };

			var dLink = document.createElement("A");
			dLink.style.display = "none";
			document.body.appendChild(dLink);
            dLink.href = uri + base64(format(template, ctx));
            dLink.download = filename + ".xls";
            dLink.click();
        }
    })();
}(jQuery, window, undefined));