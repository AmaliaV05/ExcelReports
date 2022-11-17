<%@ Page Title="Reports" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Reports.aspx.cs" Inherits="CompanyReports.Reports" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <h2><%: Title %></h2> 
    <div class="panel panel-default">
        <div class="panel-heading">
            <h3 class="panel-title">RFM Analysis</h3>
        </div>
        <div class="panel-body">
            <div>
                <asp:Label runat="server" Text="Country:" />
                <asp:DropDownList runat="server" AutoPostBack="true" ID="DisplayCountry">
                    <asp:ListItem Text="Northwest" />
                    <asp:ListItem Text="Northeast" />
                    <asp:ListItem Text="Central" />
                    <asp:ListItem Text="Southwest" />
                    <asp:ListItem Text="Southeast" />
                    <asp:ListItem Text="Canada" />
                    <asp:ListItem Text="France" />
                    <asp:ListItem Text="Germany" />
                    <asp:ListItem Text="Australia" />
                    <asp:ListItem Text="United Kingdom" />
                </asp:DropDownList>
            </div>
            <br />
            <div>
                <asp:Button ID="Button1" runat="server" Text="RFM Analysis Report" OnClick="OnButtonClickedGetRfmAnalysisData" />
            </div>
        </div>
    </div>

    
</asp:Content>
