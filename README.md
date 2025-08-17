# 🎉 Smart Event Manager

A **command-line Event Management System** built in Python.  
It allows administrators to add, edit, delete, and manage events, attendees, and email reminders.  
Events are stored in a JSON file, and attendees are managed in an Excel (`.xlsx`) file.

---

## ✨ Features

- 👀 **View Today's Events** (for all users)  
- 🔍 **Search Events** by name or type  
- 🔑 **Admin Login** (`password: admin123`)  
- 📝 **Add, Edit, Delete Events** (Admin only)  
- 📅 **View Events by Day**  
- 📑 **List All Events**  
- 👥 **Manage Attendees** (stored in `attendees.xlsx`)  
- 📧 **Send Email Reminders** to all registered attendees for upcoming events

## Notes
- The current version of this project **does not send actual emails** since no direct app password or SMTP credentials are provided.
- Instead, emails are **simulated** and displayed in the console/logs.


---

## ⚙️ Requirements

Install Python dependencies before running the program:

```bash
pip install openpyxl
