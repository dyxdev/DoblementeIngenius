<%@ Page Title="Home Page" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="Web._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <div class="alert-danger margin-sm" style="margin: 10px">
        <asp:Label ID="lblError" runat="server" Text="" Visible="false"></asp:Label>
    </div>
    <div class="alert-danger margin-sm" style="margin: 10px">
        <asp:Label ID="Label2" runat="server" Text="Loading data" Visible="false"></asp:Label>
    </div>
   
    <div class="row" style="overflow-x:auto;padding: 20px">
        <asp:DataGrid runat="server" ID="TableGrid" Visible="false" CssClass="table table-striped table-bordered"></asp:DataGrid>
    </div>
    
    <div class="row">
        <div class="col-md-6 col-lg-12 col-12 col-sm-12 " style="margin: 10px">
            <h2>Server </h2>
             <asp:TextBox CssClass="form-control input-group" ID="TextBoxserver" runat="server"></asp:TextBox>
        </div>
        <div class="col-md-6 col-lg-12 col-12 col-sm-12 " style="margin: 10px">
            <h2>Database</h2>
             <asp:TextBox CssClass="form-control input-group" ID="TextBoxDatabase" runat="server"></asp:TextBox>
        </div>
         <div class="col-md-6 col-lg-12 col-12 col-sm-12 " style="margin: 10px">
            <h2>User</h2>
             <asp:TextBox CssClass="form-control input-group" ID="TextBoxUser" runat="server"></asp:TextBox>
        </div>
         <div class="col-md-6 col-lg-12 col-12 col-sm-12 " style="margin: 10px">
            <h2>Password</h2>
             <asp:TextBox CssClass="form-control input-group" ID="TextBoxPassword" runat="server" TextMode="Password"></asp:TextBox>
        </div>
        
         <div class="col-md-6 col-lg-12 col-12 col-sm-12 margin-sm" style="margin: 10px">
             <hr />
             <asp:Button ID="Button1" CssClass="btn btn-btn-primary margin-sm" runat="server" Text="View Data" />
        </div>
       
    </div>

</asp:Content>
