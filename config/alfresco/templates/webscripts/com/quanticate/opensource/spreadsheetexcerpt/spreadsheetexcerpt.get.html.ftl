<html>
<head>
  <title>Spreadsheet Excerpt</title>
</head>
<body>
  <h1>Spreadsheet Excerpt</h1>

  <p>Two options are available. The first is to upload a 
   file, pick the sheets, and have them processed. The
   second is to upload the file along with a list of sheet
   indicies, and have it processed in one go.</p>

  <h2>Pick Sheets To Process</h2>
  <form action="excerptpick" method="post" 
        enctype="multipart/form-data">
     <div>Please upload the file, before picking sheets:</div>

     <div><input type="file" name="spreadsheet" /></div>

     <div><input type="submit" value="Upload File" /></div>
  </form>

  <h2>Enter Sheets And Process</h2>
  <form action="excerpt" method="post" 
        enctype="multipart/form-data">
     <div>Please enter the sheet indicies (comma seperated, 
      zero based) and upload a file:</div>

     <div>
      <label for="sheets">Sheet Indicies:</label>
      <input name="sheets" />
     </div>
     <div><input type="file" name="spreadsheet" /></div>

     <div><input type="submit" value="Process File" /></div>
  </form>
</body>
</html>
