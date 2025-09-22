# UPC Barcode Encoder for Excel

**📖 Documentation:** [EN English](README.md) | [FR Français](README-FR.md)

![Release](https://img.shields.io/github/v/release/Dolphin2ii/UPC-Barcode-Encoder-Excel-Addin)
![Downloads](https://img.shields.io/github/downloads/Dolphin2ii/UPC-Barcode-Encoder-Excel-Addin/total)
![License](https://img.shields.io/github/license/Dolphin2ii/UPC-Barcode-Encoder-Excel-Addin)

A comprehensive Excel VBA solution for generating UPC-A and EAN-13 barcodes directly in Excel without external dependencies or paid software.

## 🚀 Quick Download

📁 **[Download Latest Release v1.0](https://github.com/Dolphin2ii/UPC-Barcode-Encoder-Excel-Addin/releases/latest)**

Choose your preferred version:
- **UPC Excel Addin.zip** - Complete package with all languages
- **UPC Excel Addin - ENG.zip** - English version only  
- **UPC Excel Addin - FR.zip** - French version only

## 🎯 Project Overview

This project provides free alternatives to paid barcode software (like TEC-IT) for generating UPC-A and EAN-13 barcodes in Excel. The solution is designed for business environments requiring offline, secure barcode generation.

## ✨ Features

- **🔒 Secure & Offline**: No external connections or data transmission
- **📊 Multiple Formats**: Supports UPC-A (11-12 digits) and EAN-13 (13 digits)
- **🔍 Auto-Detection**: Automatically detects code format based on length
- **✅ Check Digit Calculation**: Automatically calculates and validates check digits
- **🎨 Multiple Font Support**: Compatible with Code39 and EAN-13 fonts
- **🌍 International**: Compatible with French Excel and other localized versions
- **🆓 100% Free**: No licensing costs or external dependencies

## 📋 Supported Code Formats

| Input Type | Example | Output Format | Font Required |
|------------|---------|---------------|---------------|
| 11-digit UPC | `82542200004` | `*825422000045*` | Code39 |
| 12-digit UPC | `123456789012` | `*123456789012*` | Code39 |
| 13-digit EAN | `6418029906397` | `*6418029906397*` | Code39 |

## 🚀 Quick Start

### 1. Install Fonts
Extract and install the provided barcode fonts:
- `Libre_Barcode_39.zip` - Open source Code39 font (recommended)

### 2. Download Add-in
Choose your preferred version:
- 🇬🇧 **English users**: Download `UPC Excel Addin - ENG.zip`
- 🇫🇷 **French users**: Download `UPC Excel Addin - FR.zip`
- 🌍 **Bilingual**: Download `UPC Excel Addin.zip` (contains both languages)

### 3. Install Excel Add-in
1. Open Excel
2. Go to **File** → **Options** → **Add-ins**
3. Select **Excel Add-ins** → **Go...**
4. Click **Browse...** and select your downloaded ZIP file (extract first)
5. Check the box to enable the add-in

### 4. Enable Macros (Important!)
🔒 **Security Information**: This add-in contains VBA macros for barcode generation.

**Why macros are needed:**
- ✅ **Local processing only** - No internet connection required
- ✅ **No data transmission** - Everything stays on your computer
- ✅ **Open source code** - All VBA code is visible and auditable
- ✅ **No external dependencies** - Pure Excel VBA functions only

**When Excel shows a security warning:**
1. Click **"Enable Macros"** - This is safe for this add-in
2. If blocked: **File** → **Options** → **Trust Center** → **Trust Center Settings**
3. Select **Macro Settings** → **Enable all macros** (temporarily)
4. Or add the file location to **Trusted Locations** (recommended for permanent use)

**Security guarantee:** This add-in only performs mathematical calculations locally. No network access, no file access outside Excel.

### 5. Use in Excel
```excel
=EncodeUPCOrEAN(A1)    # Recommended - Auto-detects format
=EncodeUPCA(A1)        # UPC-A only (11-12 digits)
=EncodeEAN13(A1)       # EAN-13 only (13 digits)
```

## 📚 Function Reference

### Primary Functions

#### `EncodeUPCOrEAN(code)`
**Recommended for most use cases**
- Auto-detects UPC-A vs EAN-13 based on code length
- Returns Code39 formatted string for all code types
- Example: `=EncodeUPCOrEAN("6418029906397")` → `*6418029906397*`

#### `EncodeUPCA(code)`
**UPC-A specific encoding**
- Handles 11-12 digit UPC codes only
- Automatically calculates check digit for 11-digit codes
- Example: `=EncodeUPCA("82542200004")` → `*825422000045*`

#### `EncodeEAN13(code)`
**EAN-13 specific encoding**
- Handles 12-13 digit EAN codes
- Returns raw format (no asterisks)
- Example: `=EncodeEAN13("6418029906397")` → `6418029906397`

#### `EncodeEAN13AsCode39(code)`
**EAN-13 with Code39 formatting**
- Formats EAN-13 codes for Code39 fonts
- Better scanner compatibility
- Example: `=EncodeEAN13AsCode39("6418029906397")` → `*6418029906397*`

### Legacy Functions
- `EncodeUPCAAsEAN13()` - Converts UPC-A to EAN-13 format
- `EncodeUPCACompatible()` - Code128 compatible encoding

## 🛠️ Installation Guide

### Add-in Installation
1. **Download** the project files
2. **Extract** fonts and install them on your system
3. **Copy the .xlam file** to Excel's default add-ins folder:
   - Path: `C:\Users\[YourUsername]\AppData\Roaming\Microsoft\AddIns\`
   - This ensures the add-in stays in the correct location
4. **Open Excel** and go to File → Options → Add-ins
5. **Browse** to the `.xlam` file in the AddIns folder and install
6. **Start using** the functions in your spreadsheets

### Updating the Add-in
If you need to update the add-in with new features:
1. **Disable the current add-in first**: File → Options → Add-ins → Excel Add-ins → Go → **Uncheck** your UPC add-in → OK
2. **Replace** the old `.xlam` file with the new version in the AddIns folder
3. **Re-enable** the add-in: File → Options → Add-ins → Excel Add-ins → Go → **Check** your UPC add-in → OK
4. **Restart Excel** to ensure changes take effect

## 🔧 Troubleshooting

### Common Issues

**"Excel cannot open two workbooks with the same name"**
- This happens when the add-in is already loaded in Excel
- Disable the add-in first: File → Options → Add-ins → Uncheck your add-in
- Edit the `.xlam` file, then re-enable

**Add-in location is important**
- Always keep the `.xlam` file in Excel's AddIns folder
- If you move the file after installation, Excel will show errors
- Use the folder: `C:\Users\[Username]\AppData\Roaming\Microsoft\AddIns\`

**Codes display but don't scan**
- Ensure you're using Code39 fonts (LibreBarcode39 recommended)
- Verify the cell contains the formula result with asterisks: `*code*`
- Try `EncodeUPCOrEAN()` function for best compatibility

**13-digit codes show error**
- Use `EncodeUPCOrEAN()` instead of `EncodeUPCA()`
- The old `EncodeUPCA()` only supports 11-12 digits

### Check Digit Information

**11-digit codes get extra digit**
- This is **correct behavior**
- The extra digit is the calculated check digit
- UPC-A codes are always 12 digits total
- Example: `82542200004` → `825422000045` (5 is the check digit)

## 📁 Project Structure

```
UPC-Barcode-Encoder-Excel-Addin/
├── Excel_UPC_Barcode_Business_VBA.bas     # Main VBA source code (for reference)
├── Libre_Barcode_39.zip                   # Open source Code39 font (recommended)
├── UPC Excel Addin - ENG.zip              # English version for distribution
├── UPC Excel Addin - FR.zip               # French version for distribution
├── UPC Excel Addin.zip                    # Complete packaged add-in (bilingual)
├── UPC.xlam                               # Excel add-in file
├── LICENSE                                # MIT License (English)
├── LICENSE-FR                             # MIT License (French)
├── README.md                              # Documentation (English)
└── README-FR.md                           # Documentation (French)
```

**For distribution:** Share the `UPC Excel Addin.zip` file which contains everything needed.

## 🧪 Testing

Run the built-in test function to verify everything works:
1. **Open Excel** with the add-in installed
2. **Press** `Alt + F8` to open macros
3. **Run** `TestUPCFunctions`
4. **Check** results in Immediate Window (`Ctrl + G`)

## 🤝 Business Use

This solution was developed to replace paid barcode software in business environments:
- ✅ **Cost Effective**: Eliminates licensing fees
- ✅ **Security Compliant**: No external data transmission
- ✅ **Easy Distribution**: Share add-in file with team members
- ✅ **International Support**: Works with localized Excel versions

## 📈 Version History

- **v2.0** (Current) - Added EAN-13 support, auto-detection, improved compatibility
- **v1.0** - Initial UPC-A support with Code39 fonts

## 🛡️ Security & Compliance

- **No external connections** required
- **No data transmission** outside Excel
- **Uses only standard Excel VBA features**
- **All processing happens locally** on user's computer
- **Compatible with corporate security policies**

## 📞 Support

For issues or questions:
1. Check the troubleshooting section above
2. Run the test function to verify installation
3. Ensure fonts are properly installed
4. Verify you're using `EncodeUPCOrEAN()` for mixed code types

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

Font licenses apply separately:
- LibreBarcode fonts: Open source (SIL Open Font License)

---

**💡 Tip**: For best results, use `EncodeUPCOrEAN()` function with LibreBarcode39 fonts. This combination provides maximum compatibility with all barcode scanners.