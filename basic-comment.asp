<!--#include file="layouts/header.asp"-->

<h1>Comment</h1>
<%
  set foo = createobject("WScript.Shell")
  myPath = foo.ExpandEnvironmentStrings("%USRVAR%")
  response.write("My variable is: " &myPath)
%>
<!--#include file="layouts/footer.asp"-->
