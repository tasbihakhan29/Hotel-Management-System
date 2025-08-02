# 🏨 Hotel Management System (VB6)

A desktop-based Hotel Management System developed during my internship at **Iqra Solutions**. This project is built using **Visual Basic 6.0** and includes a fully functional `.exe` file. It helps hotel staff manage guests, rooms, check-ins, billing, and more, with ease and efficiency.

---

## ✨ Features

### 🔐 Login System
- Secure access with username and password.
- Only authenticated users can access core functionalities.

### 🏠 Home Page
- Redirects to the Check-In form after login.
- Clean, form-based layout for better usability.

### 🧳 Check-In Module
- Collects guest information like name, age, number of children, etc.
- Automatically assigns unique Guest ID.
- Allows selection of:
  - Room Type: Single, Double, or Suite.
  - Facilities: Hairdryer, Spa, AC, etc.

### ➕ Add Guest
- Stores all guest details in database.
- Guest ID is reused across modules like Billing, Check-Out, and Info retrieval.

### 🏨 Add Room
- Adds newly constructed or available rooms to the system.
- Room ID auto-generated.
- Rent is auto-filled based on Room Type (defined in Settings).

### 🧾 Billing System
- Generates bills based on entered Guest ID.
- Calculates:
  - Room charges
  - Services used
  - Total amount payable
- Bill view includes all stay details.

> **Note:** Guest must be billed before performing Check-Out.

### 📤 Check-Out Module
- Accessible only after successful bill generation.
- Uses Guest ID to mark Check-Out.

### 👁️ View Guest Info / Bill / Check-Out
- Enter Guest ID to fetch:
  - Personal details
  - Room info
  - Stay duration
  - Billing history

### 📊 Hotel Status
- Displays:
  - Total checked-in guests
  - No. of adults and children
  - Rooms occupied vs. available
  - Current active services

### ⚙️ Settings Module
- Protected by password.
- Manage:
  - Room types and rates
  - Available services
  - Change admin username/password
  - Toggle service availability

---

## 🧪 Tech Stack

- **Frontend & Backend**: Visual Basic 6.0  
- **Database**: MS Access (.mdb)  
- **Executable**: `.exe` available  
- **OS Support**: Windows OS (32-bit compatibility preferred)

---

## 🧑‍💻 Developed By

**Tasbiha Khan**  
Diploma Student – Second Year  
Computer Engineering Department  
Government polytechnic Yavatmal

Internship Project @ **Iqra Solutions**

---

## 📁 Running the Project

1. Copy the `.exe` file and all dependent files to the specified path mentioned in the `repository_path.txt`.
2. Double-click the `.exe` file to run.
3. Default login credentials (can be changed later via Settings):
   - **Username:** `admin`
   - **Password:** `add`

---

## 📄 License

This project is developed for **educational and learning** purposes during internship. You may use, modify, or extend it for learning, demo, or academic requirements. Not recommended for production deployment.

---

