Problem: 
Support for sunsetting order management system, main use case was for 3rd party supplier to fulfill dropship order requests without giving them access to store and sensitive customer information. The supplier is in China, many of the apps they suggested do not meet CCPA or GDPR regulations.

Must haves:
- CSV/bulk order fulfillment
- Partial order fulfillments
- Split order fulfillment

Research:
- Many 3rd party ERP systems were full ERP systems with high monthly expense, with features client did not need
- The ERP they were migrating to did not handle order allocation and split orders
- Shopify JSON does not carry the location information that specifies where the order is allocated to. Not sure why I can see the order's allocated location_id on the admin order but not in the JSON. This seems to be the reason why many apps - if they do not handle the allocation on their side - are not able to handle split orders.

Solution:
- Client already use data import/export app that could perform tasks
- Export needed to be massaged and filtered prior to passing to supplier
- Client short on hands and time to perform this every day

Analysis:
Having worked with macros in MS Excel, I knew that macros could cut down on time to prepare the file for their vendors. I picked Google Sheets because using a browser based spreadsheet program would mean whatever Macro/Script didn't necessitate having specific software or hardware. With Excel, a macro can be made available across different spreadsheets by saving a personal macro workbook in a hidden configuration file within the application's folders. Alternatively, I could give them the script to copy and paste with each file they used on the daily. This is inefficient in a situation where they would have multiple files each day. I found that I was able to publish a script as an app on the Google Workspace Marketplace, which allowed any user of Google Sheets to install the script as an app across all their workbooks.

Solution:
- Build a script with Gemini
- Publish the script as an app

Building a script with Gemini was fairly straightforward, I knew what queries and filters I needed to apply. I refined my ask to ensure that references were based off the spreadsheet headers instead of column names, as VLOOKUP and XLOOKUP use absolute references which may be problematic if someone alters the export to add in additional columns.

Publishing the script was a little more difficult. I needed to configure OAuth access and create an app listing. Fortunately the help documentation on using Google Cloud is pretty accurate. The app is published as "under testing", where I can add specific users to both the Oauth and the app test user list. The link was sent to client for them to install the app.
