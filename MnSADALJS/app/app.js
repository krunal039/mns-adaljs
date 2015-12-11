jQuery(function () {

    //authorization context
    var resource = 'https://kmpdev.sharepoint.com';
    var endpoint = 'https://kmpdev.sharepoint.com/sites/pinterest/_api/web/lists';

    var authContext = new AuthenticationContext({
        instance: 'https://login.microsoftonline.com/',
        tenant: 'kmpdev.onmicrosoft.com',
        clientId: 'dced1fa7-3b26-4131-af32-56575e5a7883',
        postLogoutRedirectUri: window.location.origin,
        cacheLocation: 'localStorage', 
    });

    //sign in and out
    jQuery("#signInLink").click(function () {
        authContext.login();
    });
    jQuery("#signOutLink").click(function () {
        authContext.logOut();
    });

    //save tokens if this is a return from AAD
    authContext.handleWindowCallback();

    var user = authContext.getCachedUser();
    if (user) {  //successfully logged in

        //welcome user
        jQuery("#loginMessage").html("<b>Welcome, </b>" + user.userName);
        jQuery("#signInLink").hide();
        jQuery("#signOutLink").show();

        //call rest endpoint
        authContext.acquireToken(resource, function (error, token) {

            if (error || !token) {
                jQuery("#loginMessage").text('ADAL Error Occurred: ' + error);
                return;
            }

            $.ajax({
                type: 'GET',
                url: endpoint,
                headers: {
                    'Accept': 'application/json',
                    'Authorization': 'Bearer ' + token,
                },
            }).done(function (data) {
                var siteName = jQuery("#loginMessage").html();
				var lists = "<ul>";
				for(var i = 0; i <data.value.length ; i++)
				{ 
					lists += ("<li>" + data.value[i].Title + "</li>");
				}
				lists += "</ul>";
                jQuery("#loginMessage").html(siteName + '<br/> <b>Below is List of lists: </b></br/>' + lists);
            }).fail(function (err) {
                jQuery("#loginMessage").text('Error calling REST endpoint: ' + err.statusText);
            }).always(function () {
            });
        });

    }
    else if (authContext.getLoginError()) { //error logging in
        jQuery("#signInLink").show();
        jQuery("#signOutLink").hide();
        jQuery("#loginMessage").text(authContext.getLoginError());
    }
    else { //not logged in
        jQuery("#signInLink").show();
        jQuery("#signOutLink").hide();
        jQuery("#loginMessage").text("You are not logged in.");
    }

});

