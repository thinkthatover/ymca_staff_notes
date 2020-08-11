# Staff Notes Project

## Introduction
I created a Google Sheets program in 2019 to help organize communication between employees and management at the gym I worked at, and practice using some of the JavaScript skills I had learned in Harvards CS50x course on an Excel-adjacent application. The sheets were split into different tabs:

- Facility/General
- Billing
- Missed Hours
- Trainer Notes

as well as a couple of data sheets for populating dropdowns, storing deleted messages and sending automated messages to trainers when a personal training form had been filled. All of the Javascript for the project was saved in the 'code.gs' file found in the "Tools-> Script Editor" dropdown.


## Facilities Page
Primarily though, most of the communication happened on the Facility/General Tab.
![facility page](/images/Facilities.png)

### Usage
Using the page was pretty straightforward. An employee selects their name from the `From` column, selects a `Tag` (another employee, broken equipment, etc.) and explains the issue/note in the `Description` Column. From there, the `Completed` column cell would automatically be set to "No", and add a datestamp to the next column. This was accomplished by creating a trigger (edit->Current Project Triggers in the Script Editor) and calling a created `checkedit` function. When someone responded to the note in the `Response` column, The `checkedit` function would also change the value of the `Completed` column to "Addressed", marking it for cleanup in the weekly triggered `cleansheets` function 

Below is a snippet of some of the logic used in the checkedit function for marking the `Response` column.

```javascript
//mark rows as addressed if response added
if (editCol == respCol && compString !== "Yes"){
    sheet.getRange(editRow, compCol).setValue("Addressed");
}

if (dateCol > 0 && (editCol == tagCol || editCol == nameCol) && editRow !== 1){   //if date header exists, edited row not in header) 
    sheet.getRange(editRow, dateCol).setValue(Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yyyy"));
    sheet.getRange(editRow, compCol).setValue("No");
}
```

As mentioned above, the response column served two purposes. By highlighting the color of each `Completed Cell` (right click cell-> color formatting-> add rule), it provided a visual cue for employees to notice unaddressed notes. The second was marking each row for the `cleansheets` function.

## Keeping Organized: `cleansheets`
By far the most difficult part (besides headaches with permission's to create triggers and send automated emails) of maintaining the application was just keeping each page up to date; removing successfully addressed issues and keeping the number of rows/messages in each tab at a manageable, not-having-to-scroll-for-hours level. Also, we wanted to keep a record of each of the completed entries in case there was some related issue down the line, there would be an electronic paper trail to follow. But manually moving each row to a deleted files sheet was not something I was willing to do.  

Removing(and moving) completed rows was accomplished by the `cleansheets` function. This function would, by a trigger run weekly on Sundays remove all messages in all designated sheets older than 14 days that weren't marked as "Urgent" or "No" to a separate sheet called `Completed Entries`, along with the Sheet and the day the message was added.
