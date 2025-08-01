# Exhibit-Linker

## Overview
File hyperlinking is helpful for submitting legal memoranda, investigative reports, and other documents where you want the recipient to have instantaneous access to the files cited in your work product.

Exhibit Linker is a Python script that allows users to select a Word or Excel file to automatically create relative links to exhibits or Bates-stamped documents. The script reads your Word or Excel document, locates exhibit or Bates citations, uses regex to find the cited documents in a user-designed folder, and then creates relatively linked Excel and PDF output files. A Word documents can also be created, but the linking therein will be static only (unlike relative links, static links will only work on the script user's computer).

As long as your exhibits are in the same folder as your PDF or are otherwise in the same relative position (e.g., the parent PDF in one folder and exhibits in a given subfolder), the linking in the output PDF or Excel will work. Even if the PDF and exhibits are moved elsewhere on your hard drive or to another PC, so long as the exhibits travel with it in in the same relative position, the linking will work.

Further, Bates citations will open the operative PDF even if the cited page is mid-document (e.g., if SMITH_005 is found within SMITH_003.pdf, it will link to that file) and, if the output PDF is opened in Chrome, the link will even open to the correct Bates-stamped page. So, in the SMITH_005 example, the link would open a Chrome window to page 3 of that PDF.

To best ensure compatabiltity across non-Acrobat PDF viewers (e.g., if the end user is going to use Chrome to view the output PDF), it is best that the linked documents lack spacing and periods in their filenames. This is because such formatting can confuse Chrome into thinking that the links are to the internet. This script can modify the linked files' names accordingly (if the user chooses) by swapping in underscores (e.g. Ex. 1 Memo.pdf becomes Ex_1_Memo.pdf).

![Screenshot A](./images/Screenshot_A1.png)

## 🔄 How It Works

### Step-by-Step Workflow

```mermaid
graph TD
    A[📄 Select Input Document] --> B[Word/Excel File]
    B --> C[🔍 Scan for Citations]
    C --> D[Exhibit References Found]
    C --> E[Bates Numbers Found]
    D --> F[📁 Search Document Folder]
    E --> F
    F --> G[🎯 Match Files with Citations]
    G --> H[🔗 Create relative Links]
    H --> I[📊 Generate Output Files]
    I --> J[📑 Linked PDF]
    I --> K[📈 Linked Excel]
    I --> L[📝 Static Word Doc]
```


### 🎯 Smart Matching Examples

The script intelligently matches various citation formats:

- **Exhibit References**: `Ex. 1`, `Exhibit A`
- **Bates Numbers**: `SMITH_001`, `CASE_A_123`
- **Page-Specific**: Opens to exact Bates page within multi-page PDFs (if PDF is viewed via Chrome -- otherwise to first page of the relevant PDF).

### 🔧 File Processing 

- **✅ Relative Links**: Work across different computers 
- **📱 Chrome Optimization**: Direct page navigation in Chrome browser
- **🔄 Filename Sanitization**: Replace spaces/periods with underscores for compatibility (optional)
- **📂 Relative Paths**: Maintain links when files are moved together

# Tips

- Ensure that Word and Excel are closed before running the script.

- If the script feels frozen, it is likely just a hiccup.  There are multiple backup methods that it may employ if it hits a hurdle.  That takes time.

- Zip your resulting files if sending to others for easy transport. 


# Quick Start

**Install dependencies**
pip install pywin32 ttkbootstrap pypdf

**Run the application**
python exhibit_linker.py

