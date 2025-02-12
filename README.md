# Get-ADGroup-Users-Excel

## 📌 Description

This PowerShell script extracts users from an Active Directory group, displaying them in a formatted table and exporting the results to an Excel file (.xlsx). The exported Excel file includes automatically formatted columns, bold headers, a frozen top row, and an auto-filter for easy data manipulation.

## 🚀 Features

- **Prompts the user** to enter an Active Directory group name.
- **Checks if the group exists** before proceeding.
- **Retrieves all users** in the group and extracts:
  - First Name
  - Last Name
  - Account Status (Active/Disabled)
- **Displays results in a formatted PowerShell table**.
- **Automatically saves output to an Excel file (.xlsx) with:**
  - Auto-sized columns
  - Bold headers
  - Frozen top row
  - Auto-filter enabled
- **Does not require Microsoft Excel installed** (uses `ImportExcel` module).

## 📂 Installation

1. **Ensure you have the required modules installed**:
   ```powershell
   Install-Module -Name ImportExcel -Force -Scope CurrentUser
   ```
2. **Clone the repository**:
   ```bash
   git clone https://github.com/Stufo76/Get-ADGroup-Users-Excel.git
   ```
3. **Navigate to the script location**:
   ```bash
   cd Get-ADGroup-Users-Excel
   ```

## 🔧 Usage

1. **Run the script in PowerShell (as Administrator)**:
   ```powershell
   .\Get-ADGroup-Users-Excel.ps1
   ```
2. **Enter the Active Directory group name when prompted**.
3. **The script will display the results and save the file in**:
   ```
   C:\Temp\Users_<GroupName>.xlsx
   ```

## 📜 License

This project is licensed under the **GNU General Public License v3.0 (GPL-3.0)**. You are free to use, modify, and distribute the script under the terms of the license.

## 👤 Author

**Diego Pastore** ("Stufo76")  
📧 Email: [stufo76@gmail.com](mailto:stufo76@gmail.com)  
📌 GitHub: [Stufo76](https://github.com/Stufo76)

## 🤝 Contributing

If you’d like to contribute:

- Fork the repository.
- Create a new branch (`feature-branch`)
- Commit your changes.
- Submit a Pull Request.

## ⚠️ Disclaimer

This script is provided **as-is**, without any warranties. Use it at your own risk!
