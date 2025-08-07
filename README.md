# QXport
Export all charts on your Qlik Sense sheet into a single Excel file — clean, fast, and hassle-free.
# QXport

**Your Dashboard. One File.**  
Export all charts on your Qlik Sense sheet into a single Excel file — clean, fast, and hassle-free.

## 🌟 Features
- Select which visualizations to export
- Supports objects inside containers
- Exports to `.xls` format with multiple sheets
- Built-in progress bar and export status
- No external libraries required

## 📦 Installation

1. **Download the ZIP** of the extension.
2. Go to **Qlik Management Console (QMC)** > *Extensions*.
3. Click **Import**, then upload the ZIP.
4. Refresh your Qlik Sense app and add **QXport** to any sheet.

> ℹ️ Works in both Qlik Sense Enterprise and Qlik Cloud (via dev-hub or side-loading).

## ⚙️ Extension Settings
- **Button Text** – Customize the export button label
- **File Name** – Define the exported Excel file name
- **Max Rows per Sheet** – Control export size per object (default: 100,000)

## 📁 File Structure
- `qxport.js` – Core logic
- `qxport.css` – Stylesheet
- `qxport.qext` – Metadata
- `qxport.wbl` – Extension bundle list
- `preview.png` – Extension icon

## 📸 Preview
![preview](preview.png)

---

## 🛠️ To Do – Planned for Future Versions

- [ ] Support for **CSV** and other output formats
- [ ] Option to **exclude empty data**

> Have feature requests or want to contribute? Open an issue or reach out.

---

## 🔗 Project Page

Explore the code, report bugs, or contribute on GitHub:  
👉 [github.com/EliGohar85/QXport](https://github.com/EliGohar85/QXport)

---

## 📜 License
MIT – see [LICENSE](LICENSE)

---

Made with ❤️ by [Eli Gohar](https://www.linkedin.com/in/eli-gohar/)
