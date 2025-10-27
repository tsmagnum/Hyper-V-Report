$header = @"
<style>
/* Outlook-Compatible Dark Theme */

body {
  background-color: #1e1e2f;
  font-size: 13px;
  font-family: "Segoe UI", Arial, sans-serif;
  color: #e5e7eb;
  margin: 20px;
}

a, a:link, a:focus, a:hover, a:active {
  color: #f64d10;
  text-decoration: underline;
}

a:visited {
  color: #f0f414;
}

h1, h2 {
  text-align: center;
  color: #f3f4f6;
  margin: 0 0 12px;
}

table {
  width: 100%;
  border-collapse: collapse;
  background-color: #2a2a3c;
}

th, td {
  text-align: left;
  padding: 12px;
  word-break: break-word;
}

th {
  background-color: #3b82f6;
  color: #ffffff;
  font-weight: 400;
}

td {
  border-bottom: 1px solid #3f3f4f;
}

/* Zebra striping must be manually applied in HTML rows */
tr.even {
  background-color: #252536;
}

tr.odd {
  background-color: #2a2a3c;
}

/* Hover effects not supported in Outlook, so omitted */

/* Responsive and sticky headers not supported in Outlook */

/* Scrollbars and shadow effects not supported in Outlook */

/* Box-sizing reset (optional for email clients) */
* {
  box-sizing: border-box;
}
</style>
"@