# Camera Pull

Fetch photos and videos from a Fujifilm camera connected via USB — should work with any Fujifilm model (X-S10, X-T5, X100VI, GFX, etc.).

![Demo](res/demo.gif)

## Features

- Auto-detects the camera (MTP or Mass Storage mode)
- Shows camera name on connect
- Summary table of files by type with sizes and already-copied counts
- Checkbox selection of file types to copy (JPG pre-selected)
- Real-time progress bar with transfer speed and ETA
- Skips files already in the destination

## Requirements

- Windows (uses WPD/MTP API)
- [uv](https://docs.astral.sh/uv/)

## Usage

```
uv run camera_pull.py [destination_folder]
```

Default destination is `D:\camera` — change `DEST_FOLDER` at the top of the script.

## Camera USB modes

Both modes are supported:

| Mode | How to set on camera |
|------|----------------------|
| **MTP/PTP** (default) | MENU → Set Up → Connection Setting → USB Mode → MTP/PTP |
| **Mass Storage** | MENU → Set Up → Connection Setting → USB Mode → USB Mass Storage |
