# Vessel Schedule Finder Automation Tool

An automated web scraping tool that extracts and compiles vessel schedules from multiple shipping websites into a formatted Excel report. Built to streamline logistics planning and reduce manual data gathering effort.

## Features

- Automated scraping from ONE shipping line and Busan port terminals
- Compiles schedules for Dubai, Jeddah, and related ports
- Extracts key data:
  - Departure/arrival dates
  - Vessel names
  - Service routes
  - Container yard entry dates
- Outputs formatted Excel file with multiple sheets
- Headless browser operation for efficiency

## Technical Implementation

- Python with Selenium WebDriver
- Pandas for data processing
- Excel formatting with openpyxl
- Automated browser interaction handling:
  - Dynamic page loading
  - Dropdown navigation
  - Date range selection
  - Multiple website parsing

## Usage

1. Execute the compiled EXE file
2. Tool automatically scrapes latest schedules
3. Excel file opens automatically upon completion

## Current Limitations

- Updates only when manually executed
- Fixed to specific shipping lines/ports
- Requires stable internet connection
- May need updates if websites change

## Future Improvements

1. Real-time Data Integration: Push notifications for schedule changes
2. Deployment Improvements: Web application interface
