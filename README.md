# Exchange Bulk Deleting Email Script

This PowerShell script provides a powerful and automated way to delete specific emails from multiple user mailboxes in an Exchange on-premises environment. It is designed for administrators who need to perform targeted cleanups, such as removing phishing emails or messages sent in error.

---

##  Mechanism

-   **Bulk Processing:** Reads a CSV file to process multiple deletion tasks at once.
-   **Targeted Deletion:** Deletes emails based on a combination of Sender (`From`), Recipient (`To`), and `Subject`.
-   **Automated Logging:** Automatically creates a detailed transcript log of all its operations, including successes and failures.
-   **Compatibility:** Designed for Exchange on-premises (2010-2019) by attempting to load the correct PowerShell snap-ins.
-   **Error Handling:** Includes validation to ensure the CSV file exists and contains the required information.

## ⚠️ WARNING: Use with Extreme Caution!

This script uses the `Search-Mailbox -DeleteContent` cmdlet, which **permanently deletes data** from user mailboxes. This action is **irreversible**.

-   **There is no "undo" button.**
-   Always double-check your CSV file for accuracy before running the script.
-   It is highly recommended to have a valid, recent backup of your Exchange databases before proceeding.
-   Consider running the search first without the `-DeleteContent` switch to review the items that will be deleted.

## Prerequisites

1.  **Exchange Server On-Premises:** This script is designed for Exchange 2010, 2013, 2016, or 2019. It uses the legacy `Search-Mailbox` cmdlet, which is not available in pure Exchange Online.
2.  **Administrator Permissions:** The account running this script must have the **`Mailbox Import Export`** and **`Mailbox Search`** management roles assigned. Typically, membership in the **"Discovery Management"** and **"Organization Management"** role groups is required.
3.  **PowerShell Execution Policy:** Your server's execution policy must allow for running scripts. You may need to run `Set-ExecutionPolicy RemoteSigned`.

## How to Use

### Step 1: Prepare the CSV File

Create a CSV file that contains the list of emails you want to delete. The script requires this file to have three specific column headers: **`From`**, **`To`**, and **`Subject`**.

-   **`From`**: The email address of the sender.
-   **`To`**: The email address of the mailbox to search within.
-   **`Subject`**: The exact subject of the email to be deleted.

**Example `Delete.csv` file:**

```csv
From,To,Subject
phisher@evil.com,user1@yourcompany.com,"URGENT: Your password has expired"
phisher@evil.com,user2@yourcompany.com,"URGENT: Your password has expired"
internal.user@yourcompany.com,all.staff@yourcompany.com,"Recall: Please disregard previous email"
```

### Step 2: Place the CSV File

By default, the script looks for the CSV file at `C:\Users\msx.support\Desktop\delete(in).csv`.

You can either place your file there or specify a custom path when you run the script.

### Step 3: Run the Script

1.  Open the **Exchange Management Shell (EMS)** as an Administrator.
2.  Navigate to the directory where you saved the `Delete-Emails-Exchange.ps1` script.
3.  Execute the script.

**Default Usage (using the default CSV path):**
```powershell
.\Delete-Emails-Exchange.ps1
```

**Custom CSV Path:**
```powershell
.\Delete-Emails-Exchange.ps1 -CsvPath "D:\Path\To\Your\File.csv"
```

**Custom Log Path:**
```powershell
.\Delete-Emails-Exchange.ps1 -CsvPath "D:\Path\To\Your\File.csv" -LogPath "D:\Logs\DeletionLog.txt"
```

### Step 4: Review the Log File

After the script finishes, a transcript log file will be created. The path to this log will be displayed in the console. Review this file to confirm which operations were successful and to see details of any errors that occurred.

## License

This project is licensed under the MIT License. See the [LICENSE.md](LICENSE.md) file for details.
  
