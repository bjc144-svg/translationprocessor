# Assets Folder

This folder contains images and assets used in the translation processor.

## Required Files

Place the following image files in this folder:

### 1. **logo.png** (Required)
- **Purpose**: Company logo displayed in document headers and certificate pages
- **Recommended size**: 300x100 pixels or similar aspect ratio
- **Format**: PNG (preferably with transparent background)
- **Usage**: Appears on every page header and on both certificate pages

### 2. **park_signature.png** (Optional)
- **Purpose**: Authorized signature for Park Evaluation Services certification page
- **Recommended size**: 200-300 pixels wide
- **Format**: PNG (preferably with transparent background)
- **Usage**: Appears on the Park certificate page above the signature line

## How to Add Files

1. Save your logo image as `logo.png` in this folder
2. (Optional) Save Park employee signature as `park_signature.png` in this folder
3. Restart the webapp if it's currently running

## File Location

Full path: `webapp/assets/`

Example:
```
webapp/
├── assets/
│   ├── logo.png              ← Place logo here
│   ├── park_signature.png    ← Place Park signature here (optional)
│   └── README.md             ← This file
```

## Notes

- If `logo.png` is not found, the text "Park Evaluation Services" will be displayed instead
- If `park_signature.png` is not found, only a signature line will appear on the Park certificate
- Translator signatures are managed through the Translators page in the webapp
