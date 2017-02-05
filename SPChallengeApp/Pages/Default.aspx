<%-- The following 4 lines are ASP.NET directives needed when using SharePoint components --%>

<%@ Page Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" MasterPageFile="~masterurl/default.master" Language="C#" %>

<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<%-- The markup and script in the following Content element will be placed in the <head> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript" src="../Scripts/jquery-1.9.1.min.js"></script>
    <script type="text/javascript" src="../Scripts/angular.min.js"></script>
    <script type="text/javascript" src="../Scripts/angular-route.min.js"></script>
    <SharePoint:ScriptLink name="sp.js" runat="server" OnDemand="true" LoadAfterUI="true" Localizable="false" />
    <meta name="WebPartPageExpansion" content="full" />

    <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/2.6.1/fabric.min.css" />
    <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/2.6.1/fabric.components.min.css" />
    <!-- Add your CSS styles to the following file -->
    <link rel="Stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Add your JavaScript to the following file -->
    <script type="text/javascript" src="../Scripts/App.js"></script>
</asp:Content>

<%-- The markup in the following Content element will be placed in the TitleArea of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    Page Title
</asp:Content>

<%-- The markup and script in the following Content element will be placed in the <body> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <div ng-app="SPChallengeApp">
        <div ng-controller="SPChallengeController">
            <div class="ms-List ms-List--grid">
                <div class="ms-ListItem ms-ListItem--document" ng-repeat="doc in documents">
                    <div><img src="{{getPreview(doc)}}" /></div>
                    <span class="ms-ListItem-primaryText"><img src="{{getDocIcon(doc)}}" /><a href="{{doc.OriginalPath}}">{{doc.Title}}</a></span>
                    <span class="ms-ListItem-secondaryText"><a href="{{getDelveLink(doc)}}">{{doc.Author}}</a></span>
                    <span class="ms-ListItem-tertiaryText">{{getDate(doc.Created)}}</span>
                    <span>
                        <i ng-repeat="i in maxViewsArray() track by $index" ng-class="{'ms-Icon--star':getViews(doc.ViewsLifeTime) > $index,'ms-Icon--starEmpty':getViews(doc.ViewsLifeTime) <= $index}" class="ms-Icon" aria-hidden="true"></i>
                    </span>
                </div>
            </div>
        </div>
    </div>
</asp:Content>
