# Room availability board

A single-file Python web server that displays free/busy status for meeting rooms in Microsoft Exchange Online, using the Microsoft Graph API.

## Requirements

- Python 3.9+
- **Windows only:** `pip install tzdata`

## Azure setup

1. In the Azure portal, open the **"Room information"** App Registration
2. Go to **API permissions → Add a permission → Microsoft Graph → Application permissions**
3. Add `Calendars.ReadBasic.All`
4. Click **Grant admin consent for [your organisation]**

## Configuration

The server reads credentials from environment variables:

| Variable | Description |
|---|---|
| `AZURE_TENANT_ID` | Your Azure tenant ID |
| `AZURE_CLIENT_ID` | App registration client ID |
| `AZURE_CLIENT_SECRET` | App registration client secret value |

## Running

**macOS / Linux**
```bash
export AZURE_TENANT_ID="your-tenant-id"
export AZURE_CLIENT_ID="your-client-id"
export AZURE_CLIENT_SECRET="your-client-secret"
python3 server.py
```

**Windows (PowerShell)**
```powershell
$env:AZURE_TENANT_ID="your-tenant-id"
$env:AZURE_CLIENT_ID="your-client-id"
$env:AZURE_CLIENT_SECRET="your-client-secret"
python3 server.py
```

Then open [http://localhost:8888](http://localhost:8888) in a browser.

## Configuration file

Rooms and theming are configured in `config.json`.

**Rooms** are organised into named groups (rendered as separate sections). Each room needs a display name and its Exchange email address:

```json
{
  "groups": [
    {
      "name": "Floor 3",
      "rooms": [
        { "email": "room.a@example.com", "name": "Room A" }
      ]
    }
  ]
}
```

**Theme** — add a `theme` key to override any of the default colours or font:

```json
{
  "theme": {
    "font_family": "Inter, sans-serif",
    "color_background": "#0b1f1c",
    "color_surface": "#112b26",
    "color_text": "#ffffff",
    "color_text_secondary": "#6b9e94",
    "color_free_bg": "#1a4a3a",
    "color_free_fg": "#3ddc84",
    "color_busy_bg": "#4a1a1a",
    "color_busy_fg": "#ff6b6b",
    "color_error_bg": "#2a2a2a",
    "color_error_fg": "#888888"
  }
}
```

All `theme` keys are optional — any omitted key falls back to the default value shown above.
