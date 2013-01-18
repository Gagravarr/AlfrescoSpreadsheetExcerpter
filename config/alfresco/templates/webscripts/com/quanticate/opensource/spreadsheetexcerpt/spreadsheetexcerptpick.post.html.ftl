<html>
<head>
  <title>Spreadsheet Excerpt - Pick Sheets</title>
</head>
<body>
  <h1>Spreadsheet Excerpt - Pick Sheets</h1>

  <form action="excerpt" method="post">
    <input type="hidden" name="file" value="${file}" />

    <div>Select sheets to keep:</div>

    <br />
    <#list sheets as sn>
      <div>
        <input type="checkbox" name="sheets" value="${sn_index}">${sn_index} - ${sn}</input>
      </div>
    </#list>
    <br />

    <div><input type="submit" value="Process File" /></div>
  </form>
</body>
</html>
