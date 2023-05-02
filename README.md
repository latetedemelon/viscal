# VISCAL - Google Calendar Bi-directional Sync with Google Sheets

VISCAL is a Google Apps Script that enables bi-directional synchronization between Google Sheets and Google Calendar. It allows you to pull data from a Google Calendar to a Google Sheet for editing and then push the edited data back to the calendar.

## Table of Contents

- [Features](#features)
- [Setup](#setup)
- [Usage](#usage)
- [Limitations](#limitations)
- [Contributing](#contributing)
- [Donations](#donations)

## Features

- Pull events from Google Calendar to Google Sheets
- Edit events in Google Sheets
- Push updated events back to Google Calendar
- Add, update, and delete events
- Filter events
- Synchronize specific event fields

## Setup

1. Open a new Google Sheet.
2. Click on `Extensions` > `Apps Script`.
3. Delete the default `Code.gs` file.
4. Click on `File` > `New` > `Script` and name the file `VISCAL`.
5. Copy the entire script from the [VISCAL GitHub repository](https://github.com/latetedemelon/viscal/) and paste it into the `VISCAL.gs` file.
6. Save the script by clicking on the floppy disk icon or pressing `Ctrl + S` (or `Cmd + S` on macOS).
7. Close the Apps Script editor.

## Usage

1. After setting up the script, go back to your Google Sheet and refresh the page via the browser.
2. Click on the `Options` sheet
3. Click on `Google Calendar Sync` in the menu bar.
4. Click on `Refresh` to pull the events from your calendar into the sheet.
5. Edit the events in the `Google Calendar` sheet as needed.
   - Update events by changing the `ACTION` column to `UPDATE`.
   - Delete events by changing the `ACTION` column to `DELETE`.
6. Add new rows for new events.
   - Add new events by changing the `ACTION` column to `ADD`.
6. Click on `Google Calendar Sync` > `Synchronize` to push the changes back to your calendar.

## Options

1. **Calendar ID**: Specifies the calendar to be synced with the Google Sheets document. By default, it is set to 'primary', which corresponds to the user's primary Google Calendar.
2. **Start Date**: Sets the start date for the range of events to be fetched from the Google Calendar. The default value is '2023-01-01'.
3. **End Date**: Sets the end date for the range of events to be fetched from the Google Calendar. The default value is '2023-12-31'.
4. **Search Query**: Allows users to filter events by a specific search term or phrase. Only events containing the specified query will be imported into the Google Sheets document. The default value is an empty string, which means that no filter will be applied.
5. **Filter Variable**: This can be used to implement custom filtering logic based on specific event properties. The `eventMatchesFilter()` function in the script can be modified to implement the desired filter logic. By default, it returns true for all events, meaning no filtering is applied.
6. **Sync Fields**: This setting specifies the event properties to be synced between the Google Calendar and the Google Sheets document. The default value is 'calendarId,eventId,title,start,end,location,description,busyStatus,ColorId'. Users can add or remove properties as needed.

To modify the configuration settings, simply change the value in the second column of the corresponding row in the `Options` sheet. After making the necessary changes, run the `refreshSheet()` function to update the `Google Calendar` sheet based on the new settings.

## Limitations

- This script is designed for use with Google Sheets and Google Calendar. It may not work with other spreadsheet or calendar applications.
- The script only supports bi-directional synchronization between a single calendar and a single sheet.
- The script may not work as expected if there are too many events or if the calendar has a large number of recurring events.
- The script does not support event attendees or reminders.
- The synchronization process may take some time depending on the number of events and the complexity of the data.

## Contributing

Contributions to the VISCAL project are welcome! If you have improvements, bug fixes, or new features you'd like to see added, please submit a Pull Request.

## Donations

If you find VISCAL helpful and would like to support its development, consider making a donation to the project. Every little bit helps!

[Donate via PayPal](https://paypal.me/latetedemelon)

## License

This project is licensed under the MIT License. See the [LICENSE](https://github.com/latetedemelon/viscal/LICENSE) file for details.
