# Apps Script schedule updater

This Apps Script turns dictated text from an iPhone Shortcut into schedule codes in the Google Sheet.

## What it understands

Default behavior advances the day in this cycle:

```text
Monday -> Tuesday -> Wednesday -> Thursday -> Monday
```

Examples:

```text
9 orange, 10:30 orange, 11 blue, 1 orange, 3 orange, 4:30 orange
```

```text
Thursday 9 orange, 11 blue, 2 green
```

```text
same day 1:30 purple, 3 orange
```

Color mapping:

```text
orange -> o
blue -> b
green -> g
purple -> p
```

Default behavior clears the old schedule before writing the new one. Say `add` or `append` if you want to keep existing entries and add/update only the times you mention.

## Install

1. Open the Google Sheet.
2. Go to `Extensions -> Apps Script`.
3. Paste `Code.gs` into the Apps Script editor.
4. Save.
5. Click `Deploy -> New deployment`.
6. Choose type `Web app`.
7. Set `Execute as` to `Me`.
8. Set `Who has access` to `Anyone with the link`.
9. Deploy and approve the permissions.
10. Copy the Web app URL.

## iPhone Shortcut

Create a Shortcut with:

1. `Dictate Text`
2. `Get Contents of URL`
   - URL: your Apps Script Web app URL
   - Method: `POST`
   - Request Body: `JSON`
   - JSON body:

```json
{
  "text": "Dictated Text"
}
```

In Shortcuts, use the dictated text variable as the value for `text`.

The Apps Script returns JSON with the parsed day and appointments.
