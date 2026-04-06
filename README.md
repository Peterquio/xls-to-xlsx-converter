# 📊 XLS to XLSX Converter  
  
A simple and elegant desktop application to convert legacy Excel .xls files into modern .xlsx format.  
Built with Python and a clean graphical interface, this tool ensures reliable conversion by leveraging Microsoft Excel itself — preserving formatting, formulas, and structure.  
  
## 🚀 Features  
✅ Convert .xls → .xlsx easily  
📁 Saves the converted file in the same folder as the original  
🔁 Automatically avoids overwriting existing files  
🖥️ Clean and user-friendly interface  
⚡ Uses Microsoft Excel for maximum compatibility  
📊 Supports large files and multiple sheets  
  
## 🛠️ Requirements  
Windows OS  
Microsoft Excel installed  
Python 3.8+ (only if running from source)  
  
## 📦 Installation (from source)  
pip install pywin32  
python conversor_xls_para_xlsx.py  
  
## 🧱 Build Executable (.exe)  
  
To generate a standalone executable:  
  
pip install pyinstaller pywin32  
pyinstaller --noconfirm --onefile --windowed conversor_xls_para_xlsx.py  
  
The executable will be available at:  
  
dist/conversor_xls_para_xlsx.exe  
  
Optional (with icon):  
pyinstaller --noconfirm --onefile --windowed --icon=icon.ico conversor_xls_para_xlsx.py  
  
If you encounter win32com issues:  
pyinstaller --noconfirm --onefile --windowed ^  
--hidden-import=pythoncom ^  
--hidden-import=pywintypes ^  
--hidden-import=win32com ^  
--hidden-import=win32com.client ^  
conversor_xls_para_xlsx.py  
  
## 🧠 How It Works  
  
This tool automates Microsoft Excel to:  
  
Open the .xls file  
Reprocess the workbook  
Save it as .xlsx  
  
This ensures full compatibility, including:  
  
Formulas  
Formatting  
Multiple sheets  
  
## ⚠️ Notes  
This application requires Microsoft Excel installed  
Conversion time depends on file size and complexity  
Large files (e.g., 100MB+) may take several minutes  
First execution of the .exe may be slower due to antivirus scanning  
  
## 💡 Tips for Better Performance  
  
If your file is very large:  
Open it in Excel manually  
  
Go to:  
Formulas → Calculation Options → Manual  
Save the file before converting  
  
This can significantly reduce conversion time.  
  
## 📥 Download  
  
Download the latest version here:  
👉 https://github.com/Peterquio/xls-to-xlsx-converter/releases  
  
## 👨‍💻 Author  
  
Developed by Diego Morgado  
Automation & Software Development Enthusiast  
