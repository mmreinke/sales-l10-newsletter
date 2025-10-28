# Sales L10 Newsletter

Weekly newsletter automation for the sales team, sending highlights and rock progress before Monday team meetings.

## Overview

This project uses Google Apps Script to automatically send weekly newsletters to the sales team. The newsletters include:
- Team highlights from the previous week
- Rock progress updates

Data is sourced from a Google Sheet and sent via email using Apps Script automation.

## Google Sheet Structure

**Sheet URL:** [Sales L10 Newsletter Data](https://docs.google.com/spreadsheets/d/1-8vXxinAl4LR-dDq5XilQQgytTLNcZIFG4DsD0RId0g/edit?gid=840477440#gid=840477440)

### Tab: "Headlines"
- Column A (starting A2): Highlights
- Column B (starting B2): Owner (Team member)

### Tab: "Rock Progress"
- Column A (starting A2): Rock and owner
- Column B (starting B2): Goal
- Column C (starting C2): Progress

## Apps Script

**Script URL:** [Apps Script Project](https://script.google.com/u/0/home/projects/1TGaSanGzgAor1dSceGFaa-718AOaoR0oFN6fXt1MVZN7Xe_m5PpKOlhV/edit)

## Setup

### Prerequisites
- Node.js and npm installed
- Google account with access to the Apps Script project
- Git installed

### Installation

1. Clone this repository:
```bash
git clone https://github.com/mmreinke/sales-l10-newsletter.git
cd sales-l10-newsletter
```

2. Install clasp globally (if not already installed):
```bash
npm install -g @google/clasp
```

3. Log in to clasp:
```bash
clasp login
```

4. Clone the Apps Script project:
```bash
clasp clone 1TGaSanGzgAor1dSceGFaa-718AOaoR0oFN6fXt1MVZN7Xe_m5PpKOlhV
```

## Development Workflow

### Pull changes from Apps Script
```bash
clasp pull
```

### Push changes to Apps Script
```bash
clasp push
```

### Open the Apps Script editor
```bash
clasp open
```

## Usage

The script is designed to:
1. Read highlights from the "Headlines" tab
2. Read rock progress from the "Rock Progress" tab
3. Format the data into an email
4. Send the newsletter to the sales team

Configure the script to run weekly before Monday meetings using Apps Script triggers.

## Version Control

This project uses Git for version control. Make sure to:
- Commit changes regularly
- Push to GitHub after making updates
- Pull the latest changes before editing

## License

Internal use only
