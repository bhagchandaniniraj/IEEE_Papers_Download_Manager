**IEEE Paper Download Manager**

An automated desktop application designed to batch-download research papers from IEEE Xplore using exported CSV metadata. Features institutional cookie authentication, dynamic rate limiting, real-time memory monitoring, and automatic folder organization.

**Table of Contents**

*   [Quick Start (Executable)](https://www.google.com/search?q=#quick-start-executable)
    
*   [Key Features](https://www.google.com/search?q=#key-features)
    
*   [How It Works](https://www.google.com/search?q=#how-it-works)
    
*   [Screenshots & Workflow](https://www.google.com/search?q=#screenshots--workflow)
    
*   [Running from Source](https://www.google.com/search?q=#running-from-source)
    
*   [Building the Executable](https://www.google.com/search?q=#building-the-executable)
    
*   [CSV Structure](https://www.google.com/search?q=#csv-structure)
    
*   [Author & Credits](https://www.google.com/search?q=#author--credits)
    
*   [License](https://www.google.com/search?q=#license)
    

**Quick Start (Executable)**

If you don't have Python installed, you can run the standalone Windows executable directly:

1.  Navigate to the dist/ folder in this repository.
    
2.  Download **ieee\_gui.exe**.
    
3.  Double-click **ieee\_gui.exe** to launch the application. _(Note: Accepts Windows Administrator prompt to allow browser process handling and institutional cookie retrieval)._
    

**Key Features**

*   **Institutional Cookie Fetcher**: Automatically extracts required IEEE session cookies (JSESSIONID, ERIGHTS, xpluserinfo, ipCheck) directly from your browser (browser\_cookie3) to inherit institutional download access.
    
*   **Anti-Bot Delay System**: Implements randomized delays between downloads to ensure smooth processing and prevent rate limits.
    
*   **Dynamic File Organization**: Automatically creates subfolders based on your CSV file name and Document Identifier metadata.
    
*   **Full Execution Controls**: Pause, Resume, or Stop download queues at any time.
    
*   **Resource Monitoring**: Tracks application memory and system RAM usage in real time.
    
*   **Duplicate Skipping & Validation**: Detects pre-existing files to skip re-downloads and verifies valid PDF headers (%PDF / %%EOF).
    
*   **File & Folder Actions**: Dedicated buttons to open downloaded PDFs directly or navigate to their folder location in Windows Explorer.
    

**How It Works**

1.  **Authentication**: Retrieves IEEE session cookies from an active browser session where institutional access is active.
    
2.  **Parsing**: Processes the selected CSV file containing PDF links, titles, and category identifiers.
    
3.  **Download Execution**: Resolves arnumber parameters into direct PDF endpoints (stampPDF/getPDF.jsp) and streams content asynchronously.
    

Screenshots & Workflow
----------------------

**1\. Exporting CSV Metadata from IEEE Xplore**

Search for research papers on IEEE Xplore, select **Subscribed Content** or **Open Access**, and click **Export** to save the query metadata as a CSV file.

**2\. Processing Downloads via IEEE Download Manager**

Import the CSV file, initiate the automated queue, and monitor live progress across dedicated tabs (**Downloaded**, **Failed**, **Skipped**).

**Running from Source**

Prerequisites

*   Python **3.8+**
    
*   Windows OS
    

Setup Instructions

1.  Bashgit clone https://github.com/bhagchandaniniraj/IEEE\_Papers\_Download\_Manager.gitcd IEEE\_Papers\_Download\_Manager
    
2.  Bashpip install customtkinter requests browser\_cookie3 psutil
    
3.  Bashpython ieee\_gui.py
    

**Building the Executable**

If you modify the source code and want to recompile the .exe:

1.  pip install pyinstaller
    
2.  pyinstaller --noconfirm --onefile --windowed --uac-admin --icon=icon.ico --collect-all customtkinter --collect-all browser\_cookie3 ieee\_gui.py
    
3.  IEEE\_Papers\_Download\_Manager/└── dist/ └── ieee\_gui.exe
    

**Author & Credits**

*   **Developer**: Niraj Bhagchandani
    
*   **GitHub Repository**: [bhagchandaniniraj/IEEE\_Papers\_Download\_Manager](https://www.google.com/search?q=https://github.com/bhagchandaniniraj/IEEE_Papers_Download_Manager)
    

**License**

This project is licensed under the [MIT License](https://www.google.com/search?q=LICENSE).