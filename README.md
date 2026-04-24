# NH Build Tool

A web tool to parse FreshService ticket data and automatically look up all information needed to build a new hire account.

## Features

- Select a local Excel workbook (.xlsx) containing lookup data
- Data is saved to local storage - no need to re-select on revisit
- Paste FreshService ticket request data to auto-parse job information
- Uses Excel data from the workbook to look up:
  - AD Roles (by Employee Type, Job Title, and Location)
  - CMiC User Settings and Access roles
  - O365 Groups
  - Area Codes
- Dropdown interfaces to audit and override parsed values
- All processing happens locally in the browser - nothing is uploaded

## Usage

1. Select an Excel workbook (.xlsx) containing the lookup sheets
2. Paste FreshService ticket data into the text area
3. Click "Parse Ticket Info" to parse and look up information
4. Review the dropdown values in the "Job Information" section
5. Override any values as needed
6. Copy the looked up information from the "Lookup Results" section

## Lookup Data Sheets

The Excel workbook should contain these sheets:
- AD Per Type - Maps employee type to AD groups
- Job Titles - List of valid job titles
- Location Map - Maps locations to areas
- AD Per Location - Maps location to AD groups, O365 groups, and area codes
- Job Softwares - Maps area + job title to software roles
- CMiC - Maps group + job title to CMiC settings
- CMiC Location Map - Maps location to CMiC group

## Privacy Policy

No sensitive information is stored anywhere other than your browser's local storage. The repository contains only parsing patterns (no private company data). And the website is static and does not process anything on any external servers. Everything happens locally within your web browser.
