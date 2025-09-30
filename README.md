
# Vendor Expiry Notification System

This project automates vendor expiry notifications and generates daily reports using a Windows Service. It tracks expiry dates and sends email notifications to vendors based on specific conditions, helping with compliance and vendor management.

## Features:

* **Automated Notifications**: Sends email alerts to vendors at:

  * 15 days before expiry
  * 15 days after expiry
  * 60 days after expiry
* **Database Integration**: Uses MySQL to store vendor data, including expiry dates and email addresses.
* **Daily Reports**: Generates CSV reports for vendors whose expiry is more than 20 days away.
* **Logging**: Logs activities and errors in a `log.txt` file for traceability.

## File Structure:

```
Vendor-Expiry-Notification-System/
│
├── VendorExpiryService/
│   └── ServiceImplementationCode.vb (Windows Service code)
├── MySQL/
│   └── vendor_table.sql (SQL script for creating the vendor table)
├── Reports/
│   └── sample_report.csv (Sample CSV report)
├── Logs/
│   └── log.txt (Activity log)
├── README.md (Project description)
```

## Technologies Used:

* **Windows Service**: Handles background task execution.
* **MySQL**: Stores vendor data and expiry details.
* **System.Net.Mail**: Used for sending email notifications via SMTP.
* **StreamWriter**: Generates CSV reports.
* **IIS**: Used for hosting the Windows Service.

## How It Works:

1. **Store Vendor Data**: Vendor expiry information is stored in a MySQL database.
2. **Service Execution**: A Windows service runs daily, checking for upcoming or expired vendor dates. It sends emails based on the expiry conditions.
3. **Report Generation**: Daily CSV reports are created for vendors with expiry dates 20+ days away.
4. **Logging**: All actions are logged in a `log.txt` file to track the service's operations.

## Setup Instructions:

1. Clone the repository:

   ```bash
   git clone https://github.com/kavinarasan-005/Vendor-Expiry-Notification-System.git
   ```
2. Set up the MySQL database using the `vendor_table.sql` script.
3. Configure and run the Windows Service on your machine.
4. Set up SMTP settings for email notifications.

## Outcome:

* Automated the vendor expiry notification process, saving time and reducing manual work.
* Daily CSV reports help keep track of vendors and manage compliance efficiently.
* Activity logs offer transparency and make it easy to troubleshoot if needed.
