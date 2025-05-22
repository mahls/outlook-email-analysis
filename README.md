Outlook Email Analysis Dashboard

This project provides a Streamlit-based web application designed for comprehensive analysis and visualization of personal Outlook email data. It enables users to gain insights into their communication patterns, identify key topics, and perform sentiment analysis.
Features

    Interactive Dashboard: A user-friendly and responsive web interface powered by Streamlit.

    Flexible Data Ingestion: Supports uploading CSV/TSV files or automatically loading from a designated local file.

    Dynamic Filtering:

        Date Range: Analyze emails within specified timeframes.

        Sender: Focus on communications from particular individuals.

        Keyword Search: Locate emails containing specific terms in subjects or bodies.

    Communication Metrics:

        Email Volume Trends: Visualizations of email frequency over daily and monthly periods.

        Activity Heatmap: A heatmap illustrating email activity by hour and day of the week.

        Top Senders: Identification and ranking of the most prolific email senders.

        Outlier Detection: Automated flagging of days with unusually high email volumes.

    Content Analysis:

        Common Keywords: Extraction and visualization of frequently occurring words and bigrams in subject lines.

        Word Cloud: A visual representation of prominent terms within email bodies.

        Named Entity Recognition (NER): Identification of entities (e.g., persons, organizations, locations) within email content.

    Sentiment Analysis: Assessment of the emotional tone of emails, highlighting highly positive or negative communications.

    Behavioral Insights: Dynamic summaries providing interpretive insights into communication patterns and potential anomalies.

Getting Started

Follow these instructions to set up and run the Email Analysis Dashboard.
0. Export Your Emails to CSV

To prepare your email data for analysis, you will need to export it from Outlook into a CSV file. A VBA script can facilitate this process.
Using the Provided VBA Script

The following VBA script can be executed directly within Outlook to export emails from a specified shared mailbox folder to a CSV file.

Sub ExportSharedInboxEmailsToCSV()
    ' === CONFIGURATION VARIABLES ===
    ' IMPORTANT: Fill these variables with your specific details
    Dim sharedMailboxName As String: sharedMailboxName = "Your Shared Mailbox Name" ' e.g., "Sales Team Inbox"
    Dim folderName As String: folderName = "Inbox" ' The specific folder within the shared mailbox (e.g., "Inbox", "Sent Items")
    Dim outputFileName As String: outputFileName = "shared-inbox.csv" ' The name of the CSV file to be created
    Dim outputFolder As String: outputFolder = "C:\Users\YourName\Documents\EmailExports\" ' IMPORTANT: Set your desired output folder path, ending with a backslash.
    Dim bodyTruncateLength As Long: bodyTruncateLength = 1000 ' Truncate email body to this many characters to avoid excessively large cells

    ' === OUTLOOK OBJECTS ===
    Dim ns As Outlook.NameSpace
    Dim inboxFolder As Outlook.Folder
    Dim emailItem As Object ' Use Object for broader compatibility, then check TypeOf
    Dim emailItemCount As Long

    ' === FILE OBJECTS ===
    Dim csvFilePath As String
    Dim csvFile As Integer

    ' === EMAIL FIELDS ===
    Dim i As Long
    Dim emailSubject As String
    Dim emailSender As String
    Dim emailDate As String
    Dim emailBody As String

    ' === SETUP ===
    Set ns = Application.GetNamespace("MAPI")
    On Error Resume Next ' Handle errors gracefully, e.g., if mailbox/folder not found
    Set inboxFolder = ns.Folders(sharedMailboxName).Folders(folderName)
    On Error GoTo 0 ' Turn off error handling

    If inboxFolder Is Nothing Then
        MsgBox "Could not find the folder '" & folderName & "' under mailbox '" & sharedMailboxName & "'. " & _
               "Please check the mailbox name and folder name, and ensure you have access.", vbCritical, "Folder Not Found"
        Exit Sub
    End If

    ' Ensure the output folder exists
    If Dir(outputFolder, vbDirectory) = "" Then
        MkDir outputFolder
    End If

    csvFilePath = outputFolder & outputFileName
    csvFile = FreeFile ' Get the next available file number
    
    ' Open the CSV file for writing. If it exists, it will be overwritten.
    Open csvFilePath For Output As csvFile

    ' === HEADER ROW ===
    ' Use double quotes to enclose each field to handle commas within data
    Print #csvFile, """Subject"",""Sender"",""Date"",""Body"""

    ' === LOOP THROUGH EMAILS ===
    emailItemCount = inboxFolder.Items.Count
    For i = 1 To emailItemCount
        On Error Resume Next ' Handle potential errors with individual items
        Set emailItem = inboxFolder.Items(i)
        On Error GoTo 0

        If Not emailItem Is Nothing Then
            ' Ensure the item is a MailItem (email)
            If TypeOf emailItem Is MailItem Then
                ' Replace commas within data fields to prevent CSV parsing issues
                emailSubject = Replace(emailItem.Subject, """", """""") ' Escape existing double quotes
                emailSubject = Replace(emailSubject, ",", " ") ' Replace commas with spaces

                emailSender = Replace(emailItem.SenderName, """", """""")
                emailSender = Replace(emailSender, ",", " ")

                ' Format date as "DD/MM/YYYY HH:MM:SS AM/PM" to match Python script's expected format
                emailDate = Format(emailItem.ReceivedTime, "DD/MM/YYYY hh:mm:ss AM/PM")

                emailBody = Replace(emailItem.Body, """", """""")
                emailBody = Replace(emailBody, Chr(13), " ") ' Remove carriage returns
                emailBody = Replace(emailBody, Chr(10), " ") ' Remove line feeds
                emailBody = Replace(emailBody, ",", " ") ' Replace commas with spaces
                emailBody = Replace(emailBody, vbCrLf, " ") ' Remove Windows-style newlines
                emailBody = Replace(emailBody, vbCr, " ") ' Remove Mac-style newlines
                emailBody = Replace(emailBody, vbLf, " ") ' Remove Unix-style newlines

                ' Truncate body if too long
                If Len(emailBody) > bodyTruncateLength Then
                    emailBody = Left(emailBody, bodyTruncateLength) & "..."
                End If

                ' Print the email data to the CSV file, enclosing each field in double quotes
                Print #csvFile, """" & emailSubject & """,""" & emailSender & """,""" & emailDate & """,""" & emailBody & """"
            End If
        End If
    Next i

    Close csvFile ' Close the file after writing
    MsgBox "Export complete! Emails saved to: " & csvFilePath, vbInformation, "Export Successful"
End Sub

How to Manually Add and Run the VBA Script in Outlook:

    Enable Developer Tab:

        In Outlook, navigate to File > Options > Customize Ribbon.

        Under "Main Tabs" on the right, select the Developer checkbox.

        Click OK. The Developer tab will now be visible in your Outlook ribbon.

    Open VBA Editor:

        From the Outlook ribbon, click the Developer tab.

        Click Visual Basic (or press Alt + F11).

    Insert a New Module:

        In the VBA editor's "Project Explorer" pane (left side), expand "Microsoft Outlook Objects" or "VBAProject (VBAProject.OTM)".

        Right-click on "Modules" (or "VBAProject (VBAProject.OTM)" if "Modules" is not present) and select Insert > Module.

    Paste the Code:

        Copy the entire VBA code provided above (Sub ExportSharedInboxEmailsToCSV() ... End Sub) and paste it into the newly opened module window.

    Configure Variables:

        Crucially, locate the === CONFIGURATION VARIABLES === section within the VBA code.

        Update sharedMailboxName with the exact name of your shared mailbox (e.g., "Sales Team Inbox").

        Update outputFolder with the full desired path for the CSV file (e.g., "C:\Users\YourName\Documents\EmailExports\"). Ensure the path ends with a backslash \. The script will attempt to create the folder if it doesn't exist.

    Run the Macro:

        Place your cursor anywhere within the Sub ExportSharedInboxEmailsToCSV() code in the VBA editor.

        Press F5 to execute the macro, or go to Run > Run Sub/UserForm.

        A message box will confirm successful export and display the file's location.

Automating with a .bat file (Optional)

For automated execution, create a .bat file (e.g., SharedMailboxExportCSV.bat) with content similar to this:

"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" /m "ExportSharedInboxEmailsToCSV"

    Outlook Path: Adjust the path to OUTLOOK.EXE if your Office installation differs (e.g., Office15 for Office 2013, Program Files (x86) for 32-bit Office).

    Macro Name: Ensure the macro name ("ExportSharedInboxEmailsToCSV") precisely matches your VBA subroutine.

Executing this .bat file will launch Outlook (if not already open) and run the specified VBA macro.
1. Save Your Email Data (Dashboard Input)

Place your exported email data file (e.g., shared-inbox.csv or a .tsv file) in the same directory as your Python scripts. This file will serve as the default input for the dashboard if no file is uploaded via the Streamlit interface.

The file should contain a header row and be delimited (e.g., tab or comma-separated) with the following mandatory columns: Subject, Sender, Date, and Body.

Example format:

Subject	Sender	Date	Body
Meeting Reminder	john.doe@example.com	01/01/2024 10:30:00 AM	Hi Team, just a reminder about our meeting tomorrow...
Project Update	jane.smith@example.com	01/01/2024 02:15:30 PM	The project is progressing well. See attached for details...
...

The dashboard is configured to automatically detect the delimiter (sep=None, engine='python'), but a tab-separated file is recommended for consistency with common email export formats.
2. Install Dependencies

It is highly recommended to use a Python virtual environment to manage project dependencies, ensuring a clean and isolated environment.

# Navigate to your project directory
cd /path/to/your/outlook-email-analysis-project

# Create a virtual environment (if you haven't already)
python -m venv venv

# Activate the virtual environment
# On Windows:
.\venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# --- IMPORTANT FOR WINDOWS USERS ---
# If installation errors occur related to "Microsoft Visual C++ 14.0 or greater is required"
# or "Could not find vswhere.exe", you must install Microsoft Visual C++ Build Tools.
#
# 1. Download "Build Tools for Visual Studio" from Microsoft's website:
#    [https://visualstudio.microsoft.com/downloads/](https://visualstudio.microsoft.com/downloads/) (Look under "Tools for Visual Studio" -> "Other Tools and Frameworks")
# 2. Run the installer.
# 3. Select the "Desktop development with C++" workload.
# 4. Ensure "MSVC v143 - VS 2022 C++ x64/x86 build tools" (or the latest equivalent, e.g., v142 for VS 2019)
#    and "Windows 10 SDK" (or the latest Windows SDK) are selected under "Individual components".
# 5. Complete the installation.
# -----------------------------------

# Execute the provided installation script to install all necessary Python packages
python install_dependencies.py

The install_dependencies.py script will:

    Install core Python libraries: numpy, streamlit, pandas, matplotlib, seaborn, wordcloud, textblob, stanza, plotly, scikit-learn, tqdm, and nltk.

    Download essential NLTK and TextBlob linguistic corpora.

    Download the English language model for Stanza.

3. Run the Dashboard

Once all dependencies are successfully installed and your virtual environment is active, launch the Streamlit application:

streamlit run outlook-email-analysis-dashboard.py

This command will open the Streamlit application in your default web browser, typically at http://localhost:8501.
Dashboard Overview

The interactive dashboard provides various sections for email analysis:

    Sidebar Filters: Utilize the left sidebar to apply filters based on date range, sender, or keywords, dynamically updating the displayed data.

    Overview Metrics: Quick summary statistics, including total emails and unique senders.

    Interactive Charts: Dynamic plots visualizing email frequency, sender distribution, keyword analysis, and sentiment.

    Expandable Sections: Click on section headers to expand and view detailed charts, data tables, and specific insights.

    Dynamic Summary: A dedicated section offering a narrative summary of key findings and potential behavioral patterns derived from the filtered email data.

Dependencies

This project relies on the following Python libraries:

    streamlit: For building interactive web applications.

    pandas: Essential for data manipulation and analysis.

    matplotlib, seaborn, plotly: Comprehensive libraries for data visualization.

    wordcloud: Generates visual word clouds from text data.

    textblob: Provides a simple API for common natural language processing (NLP) tasks, including sentiment analysis.

    stanza: An advanced NLP library for tasks like Named Entity Recognition.

    scikit-learn: Utilized for text vectorization (e.g., CountVectorizer for bigram analysis).

    tqdm: A fast, extensible progress bar for loops (included in install_dependencies.py).

    nltk: A foundational library for natural language processing, used for stopwords and tokenization.

Author

Matthew Holmes
Phone: 0412262967
Notes

    The dashboard script expects the input CSV/TSV file to contain the columns: Subject, Sender, Date, and Body.

    The Date column is specifically parsed expecting the format DD/MM/YYYY HH:MM:SS AM/PM.

    Email bodies undergo basic preprocessing to remove common signatures and polite closings.

    Sentiment analysis generates a polarity score (ranging from -1 for negative to 1 for positive).

    Named Entity Recognition requires the download of an English language model for Stanza, which is handled by the install_dependencies.py script.
