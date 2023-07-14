# Invoke-PhishingEmailPurge

As phishing attempts grow more frequent, automating the purging process for user mailboxes becomes increasingly valuable. This repository hosts a script I've written to automate the task of blocking and purging phishing attempts from user mailboxes.

The script can be found here: https://github.com/ctejeda/Invoke-PhishingEmailPurge.git

## Prerequisites

The script works with the following technologies:

- Office 365 (Exchange)
- Mimecast (Email Filtering)
- PowerShell

## Installation

1. Clone the repository: 
    ```
    git clone https://github.com/ctejeda/Invoke-PhishingEmailPurge.git
    ```
2. Edit the lines 46, 47, 48, and 49 in the script with your Mimecast information.
3. Save the changes.

## Usage

To purge all user mailboxes from messages sent by a specific sender, use the following command:

```
Invoke-PhishingEmailPurge -PhishingEmail "BadSender@baddomain.com" -Logfile "\some\log\directory"
```

After executing the command, you will be prompted to enter your Office 365 admin credentials.

The first section of the script calls the Mimecast API, adding the malicious sender (`BadSender@baddomain.com`) to Mimecast's block list. This effectively stops further emails from the specified sender.

The script then proceeds to purge the emails from the user mailboxes. If successful, the output will be logged to the specified logfile.

Example of a successful operation would be the deletion of the phishing email from user "Chris Tejeda’s" Office 365 Mailbox.

Please replace the `"BadSender@baddomain.com"` and `"\some\log\directory"` with the actual phishing email sender and the directory where you want to store the log file respectively.

Note: Always ensure to use a valid Office 365 admin account for the process.
