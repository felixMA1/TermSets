'use strict';

//Set up the contoso namespace

var Contoso = window.Contoso || {};

//Set up the Suggestions App module

Contoso.CorporateStructureApp = function () {

    //Variables

    var context = SP.ClientContext.get_current();

    var hostweb;

    var user = context.get_web().get_currentUser();

    var taxonomySession;

    var termStore;

    var groups;

    var corporateGroup;

    var corporateTermSet;

    var corporateGroupGUID = createGuid();//"91B8E283-E5C0-47A8-91CF-EC15A5285F55";

    var corporateTermSetGUID = createGuid();// "15673DB6-04B6-403A-81BD-FD19EB08C7FA";

    var hrTermGUID = createGuid();// "8E2B6697-A9D9-4FDA-B607-BF4362ED072D";

    var salesTermGUID = createGuid();// "8F46D27B-C9F9-4D41-8029-C48992C86B24";

    var technicalTermGUID = createGuid();//"57836ED6-7199-4D6E-B7CA-E99B1DCCA66E";

    var engineeringTermGUID = createGuid();// "6990621D-C692-4DB7-9315-C6A6989CB1C6";

    var softwareTermGUID = createGuid();//"FFBE39F5-BD2A-4E41-9034-E482912D4B73";

    var supportTermGUID = createGuid();//"EB67773A-0168-40ED-82C3-48ED2A70AE0E";

    var metaDataField;
    

    //This function loads the default termstore for the site collection

    var loadTermStore = function () {

        //Open a taxonomy session and get the Managed Metadate term store

        taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);

        termStore = taxonomySession.get_termStores().getByName("Taxonomy_k9+x9EibJQbtsU+uYeEg1w==");

        context.load(taxonomySession);

        context.load(termStore);

        context.executeQueryAsync(function () {

            $("#status-message").text("Term store loaded.");

            checkGroups();

        }, function (sender, args) {

            $("#status-message").text("Error: Term store could not be loaded.");

        });

    };

    //This function lists the taxonomy groups in the term store

    var checkGroups = function () {

        //Get the groups

        var groups = termStore.get_groups();

        context.load(groups);

        context.executeQueryAsync(function () {

            $("#groups-list").children().remove();

            //Loop through the groups

            var groupEnum = groups.getEnumerator();

            while (groupEnum.moveNext()) {

                var currentGroup = groupEnum.get_current();

                var currentGroupID = currentGroup.get_id();

                //Add an element for the current group

                var groupDiv = document.createElement("div");

                groupDiv.appendChild(document.createTextNode(currentGroup.get_name()));

                $("#groups-list").append(groupDiv);

            }

        }, function (sender, args) {

            $("#status-message").text("Error: Groups could not be loaded");

        });

    };

    //This function execute when the user clicks the Create Corporate Term Set button

    //It creates a new group and term set for Contoso divisions and teams

    var createTermSet = function () {

        $("#status-message").text("Creating the group and term set...");

        //Create a dedicated group for the term set to go in

        corporateGroup = termStore.createGroup("Zona geografica", corporateGroupGUID);

        context.load(corporateGroup);

        //Create the term set itself

        corporateTermSet = corporateGroup.createTermSet("Area Norte", corporateTermSetGUID, 1033);

        context.load(corporateTermSet);

        context.executeQueryAsync(function () {

            $("#status-message").text("Group and term set created.");

            //Create the terms in the term set

            createTerms();

        }, function (sender, args) {

            $("#status-message").text("Error: Group and term set creation failed.");

        });

    };

    //This function creates new terms in the term set

    var createTerms = function () {

        //Create and load the terms

        $("#status-message").text("Creating terms...");

        var hrTerm = corporateTermSet.createTerm("Galicia", 1033, hrTermGUID);

        context.load(hrTerm);

        var salesTerm = corporateTermSet.createTerm("Asturias", 1033, salesTermGUID);

        context.load(salesTerm);

        var technicalTerm = corporateTermSet.createTerm("Pais Vasco", 1033, technicalTermGUID);

        context.load(technicalTerm);

        //var engineeringTerm = technicalTerm.createTerm("Engineering", 1033, engineeringTermGUID);

        //context.load(engineeringTerm);

        //var softwareTerm = technicalTerm.createTerm("Software", 1033, softwareTermGUID);

        //context.load(softwareTerm);

        //var supportTerm = technicalTerm.createTerm("Support", 1033, supportTermGUID);

        //context.load(supportTerm);

        context.executeQueryAsync(function () {

            $("#status-message").text("Terms created.");

            //Update the display

            checkGroups();

        }, function (sender, args) {

            $("#status-message").text("Error: Could not create the terms");

        });

    };

    //This executes when the user clicks the Create Corporate Site Columns button

    var createColumns = function () {

        //Start by getting the parent web info object

        $("#status-message").text("Obtaining the parent web...");

        var hostwebinfo = context.get_web().get_parentWeb();

        context.load(hostwebinfo);

        context.executeQueryAsync(function () {

            hostweb = context.get_site().openWebById(hostwebinfo.get_id());

            context.load(hostweb);

            context.executeQueryAsync(function () {

                $("#status-message").text("Parent web loaded.");

                //Now we can add the columns

                addColumns();

            }, function (sender, args) {

                $("#status-message").text("Could not load parent web");

            });

        }, function (sender, args) {

        });

    };

    var addColumns = function () {

        //Create the new fields

        $("#status-message").text("Creating site columns...");

        var webFieldCollection = hostweb.get_fields();

        var noteField = webFieldCollection.addFieldAsXml(xmlNoteField, true, SP.AddFieldOptions.defaultValue);

        context.load(noteField);

        metaDataField = webFieldCollection.addFieldAsXml(xmlMetaDataField, true, SP.AddFieldOptions.defaultValue);

        context.load(metaDataField);

        context.executeQueryAsync(function () {

            $("#status-message").text("Columns added.");

            //Now connect the managed metadata column to the termset

            connectFieldToTermset();

        }, function (sender, args) {

            $("#status-message").text("Error: Could not create the columns.");

        });

    };

    //This function connects the Corporate Unit managed metadata column to

    //the Contoso term set

    var connectFieldToTermset = function () {

        $("#status-message").text("Connecting the columns to the term set");

        //Store the term store ID

        var sspID = termStore.get_id();

        //Get the field, casting as a taxonomy field

        var metaDataTaxonomyField = context.castTo(metaDataField, SP.Taxonomy.TaxonomyField);

        context.load(metaDataTaxonomyField);

        context.executeQueryAsync(function () {

            //Make the connection to the term store and term set

            metaDataTaxonomyField.set_sspId(sspID);

            metaDataTaxonomyField.set_termSetId(corporateTermSetGUID);

            metaDataTaxonomyField.update();

            context.executeQueryAsync(function () {

                $("#status-message").text("Connection made. Operations complete.");

            }, function (sender, args) {

                $("#status-message").text("Error: Could not connect the taxonomy field.");

            });

        }, function (sender, args) {

        });

    };

    //Functions for displaying the current user name

    var getUserName = function () {

        context.load(user);

        context.executeQueryAsync(onGetUserNameSuccess, onGetUserNameFail);

    };

    var onGetUserNameSuccess = function () {

        $('#message').text('Hello ' + user.get_title());

    };

    var onGetUserNameFail = function (sender, args) {

        alert('Failed to get user name. Error:' + args.get_message());

    };

    return {

        create_columns: createColumns,

        create_termset: createTermSet,

        get_username: getUserName,

        load_termstore: loadTermStore

    }

}();

// This code runs when the DOM is ready

function createGuid() {

    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {

        var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);

        return v.toString(16);

    });

}

$(document).ready(function () {

    Contoso.CorporateStructureApp.load_termstore();

});