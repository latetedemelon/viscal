[![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://paypal.me/latetedemelon) [![Donate](https://img.shields.io/badge/Donate-Buy%20Me%20a%20Coffee-yellow)](https://buymeacoffee.com/latetedemelon) [![Donate](https://img.shields.io/badge/Donate-Ko--Fi-ff69b4)](https://ko-fi.com/latetedemelon)

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
5. Copy the entire script from the [VISCAL GitHub repository](https://github.com/latetedemelon/viscal/viscal.gs) and paste it into the `VISCAL.gs` file.
6. Save the script by clicking on the floppy disk icon or pressing `Ctrl + S` (or `Cmd + S` on macOS).
7. Close the Apps Script editor.

## Usage

1. After setting up the script, go back to your Google Sheet and refresh the page via the browser.
2. Click on `Google Calendar Sync` in the menu bar.
3. Click on `Refresh` to pull the events from your calendar into the sheet.
4. Authorize the script - This is a one time only item.
5. Click on `Google Calendar Sync` in the menu bar.
6. Click on `Refresh` to pull the events from your calendar into the sheet.
7. Three should now be an `Options` Sheet available.  See below for confuration details.
8. Edit the events in the `Google Calendar` sheet as needed.
   - Update events by changing the `ACTION` column to `UPDATE`.
   - Delete events by changing the `ACTION` column to `DELETE`.
9. Add new rows for new events.
   - Add new events by changing the `ACTION` column to `ADD`.
10. Click on `Google Calendar Sync` > `Synchronize` to push the changes back to your calendar.

## Options

1. **Calendar ID**: Specifies the calendar to be synced with the Google Sheets document. 
2. **Start Date**: Sets the start date for the range of events to be fetched from the Google Calendar. 
3. **End Date**: Sets the end date for the range of events to be fetched from the Google Calendar. 
4. **Search Query**: Allows users to filter events by a specific search term or phrase. Only events containing the specified query will be imported into the Google Sheets document. 
5. **Filter Variable**: This can be used to implement custom filtering logic based on specific event properties. 
6. **Sync Fields**: This setting specifies the event properties to be synced between the Google Calendar and the Google Sheets document. 

To modify the configuration settings, simply change the value in the second column of the corresponding row in the `Options` sheet. After making the necessary changes, `Refresh` under the `Google Calendar Sync` menu to update the `Google Calendar` sheet based on the new settings.

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


[!["Donate via PayPal](<img src="https://github.com/stefan-niedermann/paypal-donate-button/blob/master/paypal-donate-button.png" width="90" height="35">)](https://paypal.me/latetedemelon)

<a href='https://ko-fi.com/latetedemelon' target='_blank'><img height='35' style='border:0px;height:46px;' src='https://az743702.vo.msecnd.net/cdn/kofi3.png?v=0' border='0' alt='Buy Me a Coffee at ko-fi.com' />

[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/yellow_img.png)](https://www.buymeacoffee.com/latetedemelon)

## License

This project is licensed under the MIT License. See the [LICENSE](https://github.com/latetedemelon/viscal/blob/main/LICENSE) file for details.
