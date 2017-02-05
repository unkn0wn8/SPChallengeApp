(function () {
    'use strict';
    var SPChallengeApp = angular.module('SPChallengeApp', []);
    SPChallengeApp.controller('SPChallengeController', ['$scope', function ($scope) {
        var hostweburl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
        var appweburl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));
        var scriptbase = hostweburl + "/_layouts/15/";
        $scope.documents = [];
        $scope.minViews = -1;
        $scope.maxViews = 1;

        $.getScript(scriptbase + "Search.ClientControls.debug.js", loadRequestExecutor);

        $scope.getDate = function (stringDate) {
            var date = new Date(stringDate);
            return date.toLocaleString();
        }

        $scope.getDelveLink = function (docAuthorOWSUSER) {
            var delveUrl = hostweburl.replace(".sharepoint.com", "-my.sharepoint.com");
            return delveUrl + "/person.aspx?user=" + docAuthorOWSUSER.split('|')[0].trim();
        }

        $scope.getDocIcon = function (docFileType) {
            return scriptbase + "images/ic" + docFileType + ".png"
        }

        $scope.getDocPreview = function (docPreviewUrl){
            if (docPreviewUrl === undefined || docPreviewUrl === null)
                return "../Images/noPreview.png";
            else
                return docPreviewUrl;
        }
        $scope.maxViewsArray = function () {
            return new Array(5);
        }

        $scope.getViews = function (currentViews) {
            var maxStar = 5;
            if (currentViews === undefined || currentViews === null)
                return 1;
            else {
                if (currentViews == $scope.maxViews)
                    return maxStar;
                var ratio = Math.floor((currentViews * maxStar) / $scope.maxViews);
                if (ration < 1){
                    return 1;
                } else {
                    return ratio;
                }
            }
        }

        $scope.calculateViews = function (currentViews) {
            if (currentViews !== undefined || currentViews !== null) {
                if (currentViews > $scope.maxViews)
                    $scope.maxViews = currentViews;
                if (currentViews < $scope.minViews || currentViews === -1) {
                    $scope.minViews = currentViews;
                }
            } else {
                $scope.minViews = 1;
            }
        }

        function loadRequestExecutor() {
            $.getScript(scriptbase + "SP.RequestExecutor.js", getDocuments);
        }


        function getDocuments() {
            var executor = new SP.RequestExecutor(appweburl);
            executor.executeAsync(
                {
                    url: appweburl + "/_api/search/query?Querytext='*'&Properties='IncludeExternalContent:true,GraphQuery:ACTOR(ME\\,OR(action\\:1001\\,action\\:1003))'&SelectProperties=%27AuthorOWSUSER,ViewsLifeTime,OriginalPath,Title,Author,Write,FileType,Created,DocumentPreviewMetadata,SiteId,WebId,DocId,UniqueId,ServerRedirectedPreviewURL%27",
                    method: "GET",
                    headers: { "Accept": "application/json; odata=verbose", "X-RequestDigest": $("#__REQUESTDIGEST").val() },
                    success: successHandler,
                    error: errorHandler
                }
            );
        }
        function getQueryStringParameter(paramToRetrieve) {

            var strParams = "";
            var params = document.URL.split("?")[1].split("&");

            for (var i = 0; i < params.length; i = i + 1) {
                var singleParam = params[i].split("=");

                if (singleParam[0] == paramToRetrieve)
                    return singleParam[1];
            }
        }

        function errorHandler(data, errorCode, errorMessage) {
            alert(errorMessage);
        }

        function successHandler(data) {
            var jsonBody = JSON.parse(data.body);
            var results = jsonBody.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
            for (var i = 0; i < results.length; i++) {
                var doc = {};
                var docProperties = results[i].Cells.results;
                for (var j = 0; j < docProperties.length; j++) {
                    doc[docProperties[j].Key] = docProperties[j].Value;
                }
                $scope.calculateViews(doc.ViewsLifeTime);
                $scope.documents.push(doc);
            }
            $scope.$apply();
        }
    }]);
})();