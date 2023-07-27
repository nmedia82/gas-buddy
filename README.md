# Google Apps Script - Google Sheets Automation

This repository contains a Google Apps Script for automating tasks in Google Sheets. The script provides the following functions to insert or update data into Google Sheets.

- Add a header row
- Add a new row of data
- Update a row based on a column key
- Remove a row based on a column key and row value.

The purpose of this script is to streamline data management in Google Sheets and facilitate automation for various tasks.

## Getting Started

To use this Google Apps Script, follow the steps below:

1. Open your Google Sheets document.
2. Click on "Extensions" in the top menu and select "Apps Script" to open the Apps Script editor.
3. Copy the code from the provided `code.gs` file in this repository.
4. Paste the copied code into the Apps Script editor.
5. Save the script (File > Save).

## Functions and Examples

### 1. addHeader(...headers)

The `addHeader` function adds a header row to the Google Sheet if it doesn't already exist. It takes multiple header values as arguments.

Example:

```javascript
function addHeaderExample() {
  var message = addHeader('ID', 'Name', 'Email');
  Logger.log(message); // This will log the success message or error message, if any.
}
```

### 2. addRow(...values)
The addRow function adds a new row of data with the specified values to the Google Sheet. The number of values provided must match the number of columns in the header row.

Example:

```javascript
function addRowExample() {
  var message = addRow(1, 'John Doe', 'john.doe@example.com');
  Logger.log(message); // This will log the success message or error message, if any.
}
```

### 3. updateRow(colKey, oldVal, newVal)
The updateRow function updates a row in the Google Sheet based on the given column key, old value, and new value.

Example:

```javascript
function addRowExample() {
  var message = addRow(1, 'John Doe', 'john.doe@example.com');
  Logger.log(message); // This will log the success message or error message, if any.
}
```

### 3. updateRow(colKey, oldVal, newVal)
The updateRow function updates a row in the Google Sheet based on the given column key, old value, and new value.

Example:

```javascript
function updateRowExample() {
  var message = updateRow('Name', 'John Doe', 'Jane Smith');
  Logger.log(message); // This will log the success message or error message, if any.
}
```
### 4. removeRow(colKey, rowVal)
The removeRow function removes a row from the Google Sheet based on the given column key and row value.

Example:

```javascript
function removeRowExample() {
  var message = removeRow('Name', 'Jane Smith');
  Logger.log(message); // This will log the success message or error message, if any.
}
```

# Importance of Automation in Google Sheets

Automation in Google Sheets offers numerous benefits, especially for data management and repetitive tasks. With automation, you can:

- Save time and effort by reducing manual data entry and manipulation.
- Minimize the risk of human errors that can occur during manual data processing.
- Improve data accuracy and consistency by using automated processes.
- Streamline workflows and increase productivity by performing tasks automatically.
- Easily update and maintain data in real-time, ensuring data is up-to-date and relevant.
- Enable better data analysis and decision-making through automated data processing.
- Integrate Google Sheets with other Google Workspace applications and third-party tools for seamless data transfer and synchronization.

For businesses and individuals looking to optimize their data management, automation with Google Apps Script in Google Sheets is a valuable solution. It allows you to create custom functions, triggers, and workflows that cater to your specific needs and enhance the efficiency of data-related tasks.

By utilizing the Google Apps Script provided in this repository, you can harness the power of automation in Google Sheets and enhance your data management capabilities. Whether you are a data analyst, spreadsheet enthusiast, or business professional, automating tasks in Google Sheets will undoubtedly improve your workflow and overall productivity.

Feel free to explore the provided functions and examples to get started with automation in Google Sheets. If you have any questions or need further assistance, don't hesitate to reach out!

# Contact Us
If you need any type of automation using Google Apps Script and Google Sheets, feel free to contact us for custom solutions tailored to your specific requirements. sales@najeebmedia.com

Happy automating!

