# Calendar Updater — Figma Plugin

A Figma plugin that populates a 12-month calendar design with dates, holidays, moon phases, sunrise/sunset times, and more.

## Features

- **Federal Holidays** — New Year's Day, MLK Jr. Day, Presidents' Day, Memorial Day, Juneteenth, Independence Day, Labor Day, Indigenous People's Day, Veteran's Day, Thanksgiving, Christmas Day (with observed-day adjustments)
- **Observances** — Groundhog Day, Valentine's Day, St. Patrick's Day, Easter, Earth Day, Cinco de Mayo, Mother's Day, Father's Day, Halloween, Christmas Eve, New Year's Eve, Lunar New Year, Election Day, DST Start/End
- **Full Moons** — Monthly full moons with traditional names, super moon and blue moon labels
- **Sunrise & Sunset** — Times for the four corners of each month's grid, based on zip code
- **Equinoxes & Solstices** — Seasonal astronomical events
- **Birthdays** — Custom birthdays with optional birth year for age display

All event categories can be toggled on/off independently via checkboxes.

## Setup

```bash
cd plugin
npm install
npx tsc
```

Then load the plugin in Figma via **Plugins > Development > Import plugin from manifest** and select `plugin/manifest.json`.

## Usage

1. Open the calendar Figma file
2. Run the plugin
3. Set the year and zip code (for sunrise/sunset location)
4. Toggle event categories on/off
5. Add birthdays in the format `Mon DD [YYYY] Name` (year is optional — include it to show age)
6. Click **Update Calendar**

Settings are saved between sessions.
