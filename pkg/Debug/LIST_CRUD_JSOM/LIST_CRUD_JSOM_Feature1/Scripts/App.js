/// <reference path="_references.js" />




'use strict';

///The helper method to manage the Host and App Web Url
function manageQueryStringParameter(paramToRetrieve) {
    var params =
    document.URL.split("?")[1].split("&");
    var strParams = "";
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == paramToRetrieve) {
            return singleParam[1];
        }
    }
}

var context = SP.ClientContext.get_current();
var user = context.get_web().get_currentUser();


var hostWebUrl;
var appWebUrl;

var listItemToUpdate; // The global declaration for Update and Delete the ListItem
var listItemId; //This global list item id used for Update and delete 

// This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
$(document).ready(function () {
    //getUserName();

    //SPHostUrl – the full URL of the host site 
    //SPAppWebUrl – the full URL of the app web

    hostWebUrl = decodeURIComponent(manageQueryStringParameter('SPHostUrl'));
    appWebUrl = decodeURIComponent(manageQueryStringParameter('SPAppWebUrl'));
    //debugger;
    //alert("Host Web Url " + hostWebUrl);
    //alert("App Web Url " + appWebUrl);

    listAllCategories();

    $("#btn-new").on('click', function () {
        $(".c1").val('');
    });

    

    $("#btn-add").on('click', function () {
        createCategory();
        listAllCategories();
    });

    $("#btn-update").on('click', function () {
        updateItem();
        listAllCategories();
    });

    $("#btn-find").on('click', function () {
        findListItem();
    });


    $("#btn-delete").on('click', function () {
        deleteListItem();
        listAllCategories();
    });


     



});

//My Code Here

//<------Function To Create a New Category In List ------->

function createCategory() {
    var ctx = new SP.ClientContext(appWebUrl);//Get the SharePoint Context object based upon the URL
    var appCtxSite = new SP.AppContextSite(ctx, hostWebUrl);

    var web = appCtxSite.get_web(); //Get the Site 

    var list = web.get_lists().getByTitle("CategoryList"); //Get the List based upon the Title
    var listCreationInformation = new SP.ListItemCreationInformation(); //Object for creating Item in the List
    var listItem = list.addItem(listCreationInformation);

    listItem.set_item("Title", $("#CategoryId").val());
    listItem.set_item("CategoryName", $("#CategoryName").val());
    listItem.update(); //Update the List Item

    ctx.load(listItem);
    //Execute the batch Asynchronously
    ctx.executeQueryAsync(
        Function.createDelegate(this, success),
        Function.createDelegate(this, fail)
       );
}
//<--------------------Ends Here-------------------------->


//<----------Function To List All Categories--------->
function listAllCategories() {

    var ctx = new SP.ClientContext(appWebUrl);
    var appCtxSite = new SP.AppContextSite(ctx, hostWebUrl);

    var web = appCtxSite.get_web(); //Get the Web 

    var list = web.get_lists().getByTitle("CategoryList"); //Get the List

    var query = new SP.CamlQuery(); //The Query object. This is used to query for data in the List

    query.set_viewXml('<View><RowLimit></RowLimit>10</View>');

    var items = list.getItems(query);

    ctx.load(list); //Retrieves the properties of a client object from the server.
    ctx.load(items);

    var table = $("#tblcategories");
    var innerHtml = "<tr><td>ID</td><td>Category Id</td><td>Category Name</td></tr>";

    //Execute the Query Asynchronously
    ctx.executeQueryAsync(
        Function.createDelegate(this, function () {
            var itemInfo = '';
            var enumerator = items.getEnumerator();
            while (enumerator.moveNext()) {
                var currentListItem = enumerator.get_current();
                innerHtml += "<tr><td>"+ currentListItem.get_item('ID') +"</td><td>" + currentListItem.get_item('Title') + "</td><td>" + currentListItem.get_item('CategoryName')+"</td></tr>";
            }
            table.html(innerHtml);
        }),
        Function.createDelegate(this, fail)
        );

}
//<-----------Ends Here------------------------------>

//<------------Update List Item---------------------->
function findListItem() {

    listItemId =  prompt("Enter the Id to be Searched ");
    var ctx = new SP.ClientContext(appWebUrl);
    var appCtxSite = new SP.AppContextSite(ctx, hostWebUrl);

    var web = appCtxSite.get_web();

    var list = web.get_lists().getByTitle("CategoryList");

    ctx.load(list);

    listItemToUpdate = list.getItemById(listItemId);

    ctx.load(listItemToUpdate);

    ctx.executeQueryAsync(
        Function.createDelegate(this, function () {
            //Display the Data into the TextBoxes
            $("#CategoryId").val(listItemToUpdate.get_item('Title'));
            $("#CategoryName").val(listItemToUpdate.get_item('CategoryName'));
        }),
        Function.createDelegate(this,fail)
        );

    
}
//<-----------Ends Here------------------------------>

//<-----------Function to Update List Item----------->

function updateItem() {
    var ctx = new SP.ClientContext(appWebUrl);
    var appCtxSite = new SP.AppContextSite(ctx, hostWebUrl);

    var web = appCtxSite.get_web();

    var list = web.get_lists().getByTitle("CategoryList");
    ctx.load(list);

    listItemToUpdate = list.getItemById(listItemId);

    ctx.load(listItemToUpdate);

    listItemToUpdate.set_item('CategoryName', $("#CategoryName").val());
    listItemToUpdate.update();

    ctx.executeQueryAsync(
        Function.createDelegate(this, success),
        Function.createDelegate(this,fail)
        );

}
//<-----------Ends Here------------------------------>

//<-----------Function to Update List Item----------->
function deleteListItem() {
    var ctx = new SP.ClientContext(appWebUrl);
    var appCtxSite = new SP.AppContextSite(ctx, hostWebUrl);

    var web = appCtxSite.get_web();

    var list = web.get_lists().getByTitle("CategoryList");
    ctx.load(list);

    listItemToUpdate = list.getItemById(listItemId);

    ctx.load(listItemToUpdate);

    listItemToUpdate.deleteObject();

    ctx.executeQueryAsync(
        Function.createDelegate(this, success),
        Function.createDelegate(this, fail)
        );
}
//<-----------Ends Here------------------------------>

function success() {
    $("#dvMessage").text("Operation Completed Successfully");
}

function fail() {
    $("#dvMessage").text("Operation failed  " + arguments[1].get_message());
}


//Ends Here










// This function prepares, loads, and then executes a SharePoint query to get the current users information
function getUserName() {
    context.load(user);
    context.executeQueryAsync(onGetUserNameSuccess, onGetUserNameFail);
}

// This function is executed if the above call is successful
// It replaces the contents of the 'message' element with the user name
function onGetUserNameSuccess() {
    $('#message').text('Hello ' + user.get_title());
}

// This function is executed if the above call fails
function onGetUserNameFail(sender, args) {
    alert('Failed to get user name. Error:' + args.get_message());
}
