# CIS Benchmark PDF to Excel Converter

## 📋 Overview

Universal Python script that converts **ANY CIS Benchmark PDF** into a professional, section-wise **Excel audit checklist**. Works with Windows, Linux, Cloud platforms, network devices, and more.

## ✨ Features

- ✅ **Universal Compatibility**: Works with any CIS benchmark (Windows Server, Ubuntu, AWS, Azure, FortiGate, Cisco, etc.)
- ✅ **Auto-Detection**: Automatically identifies sections and recommendations
- ✅ **Complete Extraction**: 
  - Description
  - Rationale
  - Impact
  - Audit Steps (both CLI and GUI)
  - Remediation
  - Default Values
  - References
- ✅ **Professional Excel Output**:
  - Multi-tab workbook (one tab per section)
  - Index/overview sheet
  - Color-coded by Level 1/Level 2
  - Professional formatting
  - Ready for audit use

## 🔧 Requirements

```bash
pip install pypdf openpyxl --break-system-packages
```

or

```bash
pip install pypdf openpyxl
```

## 📖 Usage

### Basic Usage

```bash
python cis_pdf_to_excel_converter.py <input_pdf>
```

**Example:**
```bash
python cis_pdf_to_excel_converter.py CIS_Ubuntu_22_04.pdf
# Output: CIS_Ubuntu_22_04_Audit_Checklist.xlsx
```

### Specify Custom Output

```bash
python cis_pdf_to_excel_converter.py <input_pdf> <output_excel>
```

**Example:**
```bash
python cis_pdf_to_excel_converter.py CIS_Windows_Server_2022.pdf Windows_2022_Audit.xlsx
```

### With --output Flag

```bash
python cis_pdf_to_excel_converter.py <input_pdf> --output <output_excel>
```

**Example:**
```bash
python cis_pdf_to_excel_converter.py CIS_FortiGate_7_4.pdf --output FortiGate_Audit.xlsx
```

### Help

```bash
python cis_pdf_to_excel_converter.py --help
```

## 📊 Output Excel Structure

### Generated Workbook Contains:

1. **📑 INDEX Sheet**
   - Benchmark title and version
   - Generation date
   - Section summary table
   - Quick navigation

2. **Section Sheets** (One per section)
   - Section title and number
   - Complete recommendations with:
     - Control number and title
     - Level (1 or 2) - color coded
     - Description + Rationale + Impact
     - Audit steps (CLI & GUI)
     - Remediation procedures
     - Default values
     - References
   - Professional formatting
   - Print-ready

### Excel Columns:

| # | Control Title | Level | Description & Impact | Audit Steps (CLI & GUI) | Remediation | Default Value | References/Status |
|---|---------------|-------|---------------------|------------------------|-------------|---------------|-------------------|

## 🎯 Supported CIS Benchmarks

This script works with **ANY** CIS benchmark including:

### Operating Systems
- ✅ Windows Server (2016, 2019, 2022)
- ✅ Windows 10/11
- ✅ Ubuntu Linux (18.04, 20.04, 22.04, 24.04)
- ✅ Red Hat Enterprise Linux (RHEL 7, 8, 9)
- ✅ CentOS
- ✅ Debian
- ✅ macOS

### Cloud Platforms
- ✅ AWS Foundations Benchmark
- ✅ Microsoft Azure Foundations
- ✅ Google Cloud Platform (GCP)
- ✅ Oracle Cloud Infrastructure (OCI)

### Network & Security Devices
- ✅ Cisco IOS
- ✅ FortiGate/FortiOS
- ✅ Palo Alto Networks
- ✅ Juniper
- ✅ Check Point

### Databases & Applications
- ✅ Microsoft SQL Server
- ✅ MySQL/MariaDB
- ✅ PostgreSQL
- ✅ Oracle Database
- ✅ MongoDB
- ✅ Docker
- ✅ Kubernetes

### Virtualization
- ✅ VMware ESXi
- ✅ Hyper-V
- ✅ Citrix XenServer

## 📝 Examples

### Example 1: Ubuntu Linux

```bash
python cis_pdf_to_excel_converter.py CIS_Ubuntu_Linux_22.04_LTS_Benchmark_v2.0.0.pdf

# Output:
# ✅ CIS_Ubuntu_Linux_22.04_LTS_Benchmark_v2.0.0_Audit_Checklist.xlsx
# - 300+ controls organized into 6 sections
# - Ready for compliance audit
```

### Example 2: Windows Server 2022

```bash
python cis_pdf_to_excel_converter.py CIS_Microsoft_Windows_Server_2022_v3.0.0.pdf WinServer2022_Audit.xlsx

# Output:
# ✅ WinServer2022_Audit.xlsx
# - 400+ controls across 18 sections
# - GPO and Registry audit steps included
```

### Example 3: AWS Foundations

```bash
python cis_pdf_to_excel_converter.py CIS_Amazon_Web_Services_Foundations_Benchmark_v3.0.0.pdf --output AWS_Security_Audit.xlsx

# Output:
# ✅ AWS_Security_Audit.xlsx
# - Cloud security controls
# - AWS CLI and Console audit steps
```

### Example 4: FortiGate Firewall

```bash
python cis_pdf_to_excel_converter.py CIS_FortiGate_7_4_x_Benchmark_v1_0_1.pdf

# Output:
# ✅ CIS_FortiGate_7_4_x_Benchmark_v1_0_1_Audit_Checklist.xlsx
# - 60 security controls
# - 7 sections: Network, System, Firewall, Security, etc.
# - CLI and GUI audit steps
```

## ⚡ Script Workflow

```
1. Read CIS Benchmark PDF
   ↓
2. Auto-detect benchmark title & version
   ↓
3. Find recommendations start page
   ↓
4. Extract all recommendations:
   - Number, Title, Status
   - Description, Rationale, Impact
   - Audit steps (CLI & GUI)
   - Remediation
   - Default values, References
   ↓
5. Organize by section (1, 2, 3, etc.)
   ↓
6. Generate Excel workbook:
   - Index sheet
   - Section sheets (one per section)
   - Professional formatting
   ↓
7. Save Excel file
   ✅ DONE!
```

## 🎨 Customization

### Modify Section Names

Edit the `SECTION_NAMES` dictionary in the script:

```python
SECTION_NAMES = {
    '1': 'Initial Setup',           # Change these to match
    '2': 'System Configuration',    # your benchmark
    '3': 'Network & Services',
    # ... etc
}
```

### Adjust Column Widths

Modify in `create_section_sheet()` method:

```python
ws.column_dimensions['D'].width = 60  # Description column
ws.column_dimensions['E'].width = 70  # Audit column
# ... etc
```

### Change Color Coding

Modify in `_add_recommendation_row()` method:

```python
# Level 1 color
cell.fill = PatternFill(start_color='FFC000', ...)  # Orange

# Level 2 color  
cell.fill = PatternFill(start_color='FFFF00', ...)  # Yellow
```

## 🐛 Troubleshooting

### Issue: "Module not found: pypdf"

**Solution:**
```bash
pip install pypdf openpyxl --break-system-packages
```

### Issue: "No recommendations found"

**Possible causes:**
1. PDF is password-protected → Remove password first
2. PDF is scanned/image-based → Use OCR tool first
3. Non-standard CIS format → Check if PDF follows CIS structure

### Issue: "Missing some recommendations"

**Solution:**
- CIS PDFs can have varied formats
- Check the PDF manually for the missing controls
- Script extracts based on standard patterns
- Some edge cases may need manual addition

## 📄 Output Example

**Generated Excel will look like:**

```
📑 INDEX
┌────────┬────────────────────┬──────────┬─────────────────┐
│Section │ Section Name       │ Controls │ Tab Name        │
├────────┼────────────────────┼──────────┼─────────────────┤
│   1    │ Initial Setup      │    12    │ 1. Initial...   │
│   2    │ System Config      │    35    │ 2. System...    │
│   3    │ Network Services   │    28    │ 3. Network...   │
└────────┴────────────────────┴──────────┴─────────────────┘

1. Initial Setup
┌─────┬────────────────────────┬────────┬───────────────┬──────────┐
│  #  │ Control Title          │ Level  │ Audit Steps   │ Status   │
├─────┼────────────────────────┼────────┼───────────────┼──────────┤
│ 1.1 │ Ensure filesystem...  │Level 1 │ In CLI: ...   │          │
│     │                        │[Orange]│ In GUI: ...   │          │
└─────┴────────────────────────┴────────┴───────────────┴──────────┘
```

## 🔐 Security Note

This script only **reads** PDF files and **generates** Excel files. It does **not**:
- Connect to the internet
- Modify the original PDF
- Send data anywhere
- Execute system commands (except reading PDFs)

## 📞 Support

For issues or questions:
1. Check this README
2. Review the CIS benchmark PDF format
3. Ensure PDF is not password-protected
4. Verify pypdf and openpyxl are installed

## 📜 License

This script is provided as-is for security audit purposes.

## 🎯 Version

**Version:** 1.0  
**Last Updated:** 2026-01-29  
**Compatible with:** Python 3.8+

## ✅ Success Indicators

The script is working correctly if you see:

```
📄 Detected: CIS <Benchmark Name> v<version>
📍 Recommendations start at page X
✅ Extracted XX complete recommendations
📊 Generating Excel workbook...
✅ CONVERSION COMPLETE!
```

## 🚀 Quick Start

```bash
# 1. Install dependencies
pip install pypdf openpyxl --break-system-packages

# 2. Download your CIS benchmark PDF

# 3. Run the script
python cis_pdf_to_excel_converter.py your_benchmark.pdf

# 4. Open the generated Excel file
# Your_Benchmark_Audit_Checklist.xlsx

# 5. Start your audit! ✅
```

---

**Happy Auditing! 🔒**
