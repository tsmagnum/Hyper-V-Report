$header = @"
<style>
/* Outlook-Compatible Colorful Theme */

body {
  background-color: #fef6e4;
  font-size: 13px;
  font-family: "Segoe UI", Arial, sans-serif;
  color: #172b4d;
  margin: 20px;
}

h1, h2 {
  text-align: center;
  color: #ff6b6b;
  margin: 0 0 12px;
}

table {
  width: 100%;
  border-collapse: collapse;
  background-color: #ffffff;
}

th, td {
  text-align: left;
  padding: 12px;
  word-break: break-word;
}

th {
  background-color: #ff6b6b;
  color: #ffffff;
  font-weight: 400;
}

td {
  border-bottom: 1px solid #e0e0e0;
}

/* Zebra striping: apply manually in HTML */
tr.even {
  background-color: #fdf2ff;
}

tr.odd {
  background-color: #ffffff;
}

/* Hover effects not supported in Outlook */

/* Responsive scroll and sticky headers not supported in Outlook */

/* Box-sizing reset (optional for email clients) */
* {
  box-sizing: border-box;
}
</style>
"@