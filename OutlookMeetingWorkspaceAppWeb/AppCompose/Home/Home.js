/// <reference path="../App.js" />

(function () {
	'use strict';

	var spoRoot = "https://sponlineaccout.sharepoint.com/sites/";
	var provisioningUrl = "/Provision/handler.cshtml?tsname=";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#unLinkTeamsiteUrl').click(UnLinkTeamsiteUrl);
            $('#SetTeamsiteUrl').click(SetUrlToBody);

            $.when(FindTeamsiteUrl()).then(function (data) {
                initPanels(data);
                $('li').hide();

                //for demo only
                fadeItem();

            });

        });
    };


    function initPanels(data) {
        if (data != "") {
            $("#inputUrlPanel").hide();
            $("#linkedUrlPanel").show();

        } else {
            $("#inputUrlPanel").show();
            $("#linkedUrlPanel").hide();

        }

    };

    function FindTeamsiteUrl() {

        var d = jQuery.Deferred();

        Office.context.mailbox.item.body.getAsync("html", function (data) {
            var url = LocateHref(data.value);
            $("#linkedUrl").attr("href", data.value);
            d.resolve(url);

        });
        return d.promise();

    };

    function LocateHref(htmlString) {
        //locate the data-teamsite attribute and then find the href value
        var dataOffset = htmlString.indexOf("x_teamsiteurl");
        if (dataOffset >= 0) {
            htmlString = htmlString.substring(dataOffset);
            var resArray1 = htmlString.match(/href=\'[^']+?\'/);
            if (resArray1 != null) {
                return resArray1[0].substring(6, resArray1[0].length - 1);
            } else {
                var resArray2 = htmlString.match(/href=\"[^"]+?\"/);
                if (resArray2 != null) {
                    return resArray2[0].substring(6, resArray2[0].length - 1);
                }
            }
        }
        return '';
    };

    function SetUrlToBody() {
    	var teamSiteUrl = $("#teamsiteUrl").val();
    	var teamSiteCreate = $("#teamsitecreate")[0].checked;
    	var teamSiteBuiltUrl = teamSiteCreate ? spoRoot + teamSiteUrl : teamSiteUrl;
        var html = "<a class='teamsiteurl' href='" + teamSiteBuiltUrl + "' id='teamsiteurlid'>Link to meeting teamsite</a>";

        Office.context.mailbox.item.body.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, function (result) {
        	console.log("text is now in body");
        	$.when(FindTeamsiteUrl()).then(function (data) {
        		initPanels(data);
        	});

        	// request the message be placed on the queue
        	if (teamSiteCreate) {
        		var requestSiteUrl = provisioningUrl +teamSiteUrl;
	        	$.get(requestSiteUrl);
				console.log("Message handler called at " +teamSiteBuiltUrl);
        	}
        });

    };


    function fadeItem() {
        $('ul li:hidden:first').delay(500).fadeIn(fadeItem);
    };

    function UnLinkTeamsiteUrl() {

        $.when(FindTeamsiteUrl()).then(
                 function (data) {
                     var url = data;

                     var html = "";


                     //set the body again without the url
                     Office.context.mailbox.item.body.setAsync("", { coercionType: Office.CoercionType.Html }, function (result) {
                         console.log("text is now in body");
                         $.when(FindTeamsiteUrl()).then(function (data) {
                             initPanels(data);
                         });
                     });


                     //    console.log("unlink url: " + url);

                     //    Office.context.mailbox.item.body.getAsync("html", function (data) {

                     //        var mailBody = data.value;

                     //        console.log(mailBody);

                     //        var mailDOM = jQuery.parseHTML(mailBody);

                     //        // Gather the parsed HTML's node names

                     //        $.each(mailDOM, function (i, el) {
                     //            console.log(el);
                     //            if (el.getElementsByClassName("x_teamsiteurl").length>0) {
                     //                var itemToRemove = el.getElementsByClassName("x_teamsiteurl")[0];
                     //                html = itemToRemove.parentNode.innerHTML;
                     //                mailBody = mailBody.replace(html, "");
                     //                console.log(mailBody);
                     //            }


                     //        });




                     //    }
                     //);

                 });
    }


})();