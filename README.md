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


# MY EXPERIENCES(CHALLENGES+EXPERIENCE)

Honestly, at first I was pretty overwhelmed when I saw the problem statement. It felt so different from just solving LeetCode problems. I did take some help from AI — especially since I decided to implement the project in Python because file handling (like with Excel) is easier there. I wrote the basic code for the CLI, adding events, searching events, etc. For the trickier parts, like handling time conflicts and sending emails to attendees, I used AI since I had no clue how to do that. But now I’ve almost fully understood what was done in the project. Overall, it was a really nice experience building something like this
