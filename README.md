# Simple Google App Script Client Management System 

A proof of concept that you can use Google App Script with Sheets as a database-like CRM system. 

HOW TO: 

Provide sheet IDs for Client and User sheet. Just create two new Google Sheets and copy the ID from the URL.
For example: 
https://docs.google.com/spreadsheets/d/THIS_IS_YOUR_ID

Give it access to your Drive Account when asked. 
IMPORTANT: leave the first row BLANK or provide row names like "id, name, address, etc" for your data. 

This features address validation via the built in Google Maps Geocoder API. 
Address validation works best for German addresses as this was built for a German company.  
   **UPDATE**: I've reworked the address validation. All types of addresses are supported now. 

Click on 'Add Client' to see it.  

Working demo: https://script.google.com/macros/s/AKfycbzi2OzIZa9Z3J_c77qJIdfc4MzDX_mZCY1V9xCLK1jCVrjtgMGj/exec
