# QXport

Export your entire Qlik app. One file.

---

## Why QXport?

QXport lets you export **not just a single sheet - but your entire Qlik app** into one clean Excel file.

It's perfect for:

- **Full App Extraction** - Export charts across multiple sheets into one structured workbook.
- **Dashboard Sharing** - Deliver complete insights to stakeholders outside Qlik.
- **Reporting & Analysis** - Continue working with your data in Excel when needed.
- **Self-Service Simplicity** - No scripting, no setup - just select and export.

---

## Features

- **App Export (NEW)** - Select charts across the entire app and export them in one click.
- **Current Sheet Export** - Export selected charts from the active sheet.
- **Native XLSX export** - generates a real Excel `.xlsx` file.
- Multi-sheet workbook - each chart becomes its own sheet.
- Smart selection UI:
  - Search charts and sheets
  - Select visible / clear visible
  - Expand / collapse sheets
- Built-in progress bar and export status.
- Configurable button text, file name, and max rows per sheet.

---

## Installation

1. Download the latest `QXport.zip` from the **[Releases page](https://github.com/EliGohar85/QXport/releases)**.
2. Open the Qlik Sense Management Console (QMC) and navigate to **Extensions**.
3. Upload `QXport.zip`.
4. In your app, drag the **QXport** visualization onto a sheet.
5. Choose **Current Sheet** or **App Export**, select charts, and click **Export**.

---

## Export formats

- **XLSX (native Excel):** Exports selected chart data into a real `.xlsx` workbook with multiple sheets, auto column widths, wrapped text, filters, and frozen header row.

---

## Extension Settings

- **Button Text** - Customize the export button label.
- **File Name** - Define the exported Excel file name.
- **Max Rows per Sheet** - Control export size per object. Default: `100,000`.

---

## Contributions & Feedback

Do you have a feature idea or have you found a bug?

- Open an issue or submit a PR on **GitHub:** https://github.com/EliGohar85/QXport

---

## Project Links

- **GitHub (source & releases):** https://github.com/EliGohar85/QXport
- **Ko-fi (support):** https://ko-fi.com/eligohar

---

## Support

If QXport saves you time, consider supporting development:

**Ko-fi:** https://ko-fi.com/eligohar

---

## Changelog

### v1.1.0

🚀 **Major Feature - App Export**

- **Export the entire app:** Select charts across multiple sheets and export everything into a single Excel file.
- **App Export mode:** Browse all sheets, search charts, and control selection in a structured UI.
- **Bulk selection tools:** Select visible, clear visible, expand all, collapse all.

📊 **Excel Improvements**

- **Native XLSX export (ExcelJS):** Real `.xlsx` file instead of legacy XML/XLS.
- Multi-sheet workbook (one sheet per chart).
- Auto column widths, wrapped text, auto filters, and frozen header row.

🧠 **UX Enhancements**

- Clear separation between **Current Sheet** and **App Export** modes.
- Improved progress indication and export feedback.
- Better handling of large selections.

🧩 **Compatibility**

- Supports charts inside containers.
- Works across Qlik Sense Enterprise and Qlik Cloud.

### v1.0.0

- Initial release.
- Export selected visualizations from the current Qlik Sense sheet.
- Support for objects inside containers.
- Export to `.xls` format with multiple sheets.
- Built-in progress bar and export status.
- Configurable button text, file name, and max rows per sheet.

---

## License

- Distributed under the MIT License. See `LICENSE` for details.
- This project uses **ExcelJS** (MIT License) for XLSX export.

---

Made with ❤️ by **Eli Gohar**
