<HEAD>
<TITLE>Debugging Forms and Pages</TITLE>
<STYLE>  {Font-Family="Arial"} </STYLE>
<BASEFONT SIZE=2>
</HEAD>
<BODY>
<H1>Debugging Forms and Pages</H1>

<H3> QueryString Collection </H3>
<% For Each Item in Request.QueryString
      For intLoop = 1 to Request.QueryString(Item).Count %>
        <% = Item & " = " & Request.QueryString(Item)(intLoop) %> <BR>
   <% Next 
   Next %>
<H3> Form Collection </H3>
<% For Each Item in Request.Form
      For intLoop = 1 to Request.Form(Item).Count %>
        <% = Item & " = " & Request.Form(Item)(intLoop) %> <BR>
   <% Next 
   Next %>
<H3> Cookies Collection </H3>
<% For Each Item in Request.Cookies
      If Request.Cookies(Item).HasKeys Then
         'use another For...Each to iterate all keys of dictionary
         For Each ItemKey in Request.Cookies(Item) %>
            Sub Item: <%= Item %> (<%= ItemKey %>) 
                      = <%= Request.Cookies(Item)(ItemKey)%>
      <% Next 
      Else
         'Print out the cookie string as normal %>
         <%= Item %> = <%= Request.Cookies(Item)%> <BR>
   <% End If
   Next %>
<H3> ClientCertificate Collection </H3>
<% For Each Item in Request.ClientCertificate
      For intLoop = 1 to Request.ClientCertificate(Item).Count %>
        <% = Item & " = " & Request.ClientCertificate(Item)(intLoop) %> <BR>
   <% Next 
   Next %>
<H3> ServerVariables Collection </H3>
<% For Each Item in Request.ServerVariables
      For intLoop = 1 to Request.ServerVariables(Item).Count %>
        <% = Item & " = " & Request.ServerVariables(Item)(intLoop) %> <BR>
   <% Next 
   Next %>

</BODY>
</HTML>

