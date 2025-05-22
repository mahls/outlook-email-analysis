Outlook Email Analysis Dashboard

This project provides a Streamlit web application for interactive analysis and visualization of email data, designed to help users understand their communication patterns, identify key topics, and perform sentiment analysis on their exported Outlook emails.
Features

    Interactive Dashboard: A user-friendly web interface built with Streamlit.

    Flexible Data Loading: Supports uploading CSV/TSV files or loading from a default local file.

    Date Filtering: Filter emails by a specific date range.

    Sender Filtering: Select specific senders to analyze their communications.

    Keyword Search: Search for keywords within email subjects and bodies.

    Email Volume Analysis: Visualize email frequency over time (daily, monthly).

    Activity Heatmap: See email activity by hour and day of the week.

    Top Senders: Identify the most active email senders.

    Outlier Detection: Flag days with unusually high email volumes.

    Common Keywords: Analyze frequently used words and bigrams in subject lines.

    Word Cloud: Generate a visual representation of common words in email bodies.

    Named Entity Recognition (NER): Extract entities (like people, organizations, locations) from email content.

    Sentiment Analysis: Determine the emotional tone of emails and identify highly positive/negative communications.

    Dynamic Insights: Provides a summary of key findings and behavioral patterns based on the filtered data.

How to Run

Follow these steps to set up and run the Email Analysis Dashboard.
0. Export Your Emails to CSV

To get your email data in a format suitable for this dashboard, you'll need to export it from Outlook into a CSV file. One way to do this is using a VBA script.
Using the Provided VBA Script

This VBA script can be used directly within Outlook to export emails from a specified shared mailbox folder to a CSV file.

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

    Enable Developer Tab (if not already visible):

        In Outlook, go to File > Options > Customize Ribbon.

        Under "Main Tabs" on the right, check the box next to Developer.

        Click OK. The Developer tab will now appear in your Outlook ribbon.

    Open VBA Editor:

        In Outlook, click on the Developer tab.

        Click Visual Basic (or press Alt + F11).

    Insert a New Module:

        In the VBA editor, in the left-hand "Project Explorer" pane, expand "Microsoft Outlook Objects" or "VBAProject (VBAProject.OTM)".

        Right-click on "Modules" (if it exists, otherwise right-click on "VBAProject (VBAProject.OTM)" and choose Insert > Module).

        Select Insert > Module.

    Paste the Code:

        A new, blank module window will appear. Copy the entire VBA code provided above (Sub ExportSharedInboxEmailsToCSV() ... End Sub) and paste it into this module.

    Configure Variables:

        IMPORTANT: In the VBA code you just pasted, locate the === CONFIGURATION VARIABLES === section.

        Change "Your Shared Mailbox Name" to the exact name of your shared mailbox as it appears in Outlook (e.g., "Sales Team Inbox").

        Change "C:\Users\YourName\Documents\EmailExports\" to the full path where you want the shared-inbox.csv file to be saved (e.g., "C:\Users\MHolmes\Documents\outlook-data\"). Ensure the path ends with a backslash \. The script will attempt to create the folder if it doesn't exist.

    Run the Macro:

        In the VBA editor, place your cursor anywhere within the Sub ExportSharedInboxEmailsToCSV() code.

        Press F5 to run the macro, or go to Run > Run Sub/UserForm.

        A message box will confirm when the export is complete and show the file path.

Automating with a .bat file (Optional)

You can create a .bat file to run the VBA macro without manually opening the VBA editor.
Create a file named SharedMailboxExportCSV.bat (or any other name) with content similar to this:

"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" /m "ExportSharedInboxEmailsToCSV"

    Adjust the Outlook Path: The path to OUTLOOK.EXE might vary depending on your Office installation (e.g., Office15 for Office 2013, Office14 for Office 2010, or Program Files (x86) for 32-bit Office on a 64-bit system).

    Macro Name: Ensure "ExportSharedInboxEmailsToCSV" exactly matches the name of your VBA subroutine.

Running this .bat file will open Outlook (if not already open) and execute the specified macro.
1. Save Your Email Data (Dashboard Input)

Ensure your exported email data is saved in a plain text file named shared-inbox.csv (or .tsv) in the same directory as your Python scripts. This is the default file the dashboard will look for if you don't upload one. The file should have the following header and be tab-separated (or comma-separated if you change the sep parameter in the pd.read_csv call in your dashboard script):

Subject	Sender	Date	Body
Meeting Reminder	john.doe@example.com	01/01/2024 10:30:00 AM	Hi Team, just a reminder about our meeting tomorrow...
Project Update	jane.smith@example.com	01/01/2024 02:15:30 PM	The project is progressing well. See attached for details...
...



The dashboard is configured to use sep=None and engine='python' for automatic delimiter detection, but a tab-separated file is recommended for consistency with common email export formats.
2. Install Dependencies

It's highly recommended to use a Python virtual environment to manage project dependencies.

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
# If you encounter errors related to "Microsoft Visual C++ 14.0 or greater is required"
# or "Could not find vswhere.exe" during package installation, you need to install
# Microsoft Visual C++ Build Tools.
#
# 1. Download the "Build Tools for Visual Studio" from Microsoft's website:
#    https://visualstudio.microsoft.com/downloads/ (Look under "Tools for Visual Studio" -> "Other Tools and Frameworks")
# 2. Run the installer.
# 3. In the installer, select the "Desktop development with C++" workload.
# 4. Ensure "MSVC v143 - VS 2022 C++ x64/x86 build tools" (or the latest equivalent, e.g., v142 for VS 2019)
#    and "Windows 10 SDK" (or the latest Windows SDK) are selected under "Individual components" if not included by default.
# 5. Install the selected components.
# -----------------------------------

# Run the provided installation script to install all necessary Python packages
python install_dependencies.py


The install_dependencies.py script will:

    Install numpy, streamlit, pandas, matplotlib, seaborn, wordcloud, textblob, stanza, plotly, scikit-learn, tqdm, and nltk.

    Download necessary NLTK and TextBlob corpora.

    Download the English model for Stanza.

3. Run the Dashboard

Once all dependencies are successfully installed, you can launch the Streamlit application:

# Ensure your virtual environment is still active
streamlit run outlook-email-analysis-dashboard.py


This command will open the Streamlit application in your default web browser.
Dashboard Overview

Upon running the application, you will see an interactive dashboard.

    Sidebar Filters: Use the filters on the left sidebar to narrow down the email data by date range, sender, or keywords.

    Overview Metrics: See quick statistics like total emails and unique senders.

    Interactive Charts: Explore various plots for email frequency, sender distribution, keyword analysis, and sentiment.

    Expandable Sections: Click on the expanders (+) to reveal detailed charts and data tables for each analysis section.

    Dynamic Summary: A dedicated section at the bottom provides a narrative summary of key findings and potential behavioral insights based on the applied filters.

Dependencies

The project relies on the following Python libraries:

    streamlit: For building the interactive web application.

    pandas: For data manipulation and analysis.

    matplotlib, seaborn, plotly: For data visualization.

    wordcloud: For generating word clouds.

    textblob: For performing sentiment analysis.

    stanza: For advanced natural language processing, specifically Named Entity Recognition.

    scikit-learn: Used for text vectorization (e.g., CountVectorizer for bigrams).

    tqdm: For progress bars (though not explicitly used in the dashboard script, it's a common data science utility).

    nltk: For general natural language processing tasks (like stopwords).

Author

Matthew Holmes
Phone: 0412262967
Notes

    The dashboard script expects the input file to have columns named Subject, Sender, Date, and Body.

    The Date column is expected in the format DD/MM/YYYY HH:MM:SS AM/PM.

    Email bodies are pre-processed to remove common footers and polite endings before analysis.

    Sentiment analysis is performed using TextBlob, which provides a polarity score between -1 (negative) and 1 (positive).

    Named Entity Recognition uses the Stanza library, which requires downloading its English model (handled by install_dependencies.py).