# 📝 EnrollEdits – Bulk Header/Footer Editor for Word Files (Kotlin + Apache POI)

**EnrollEdits** is a Kotlin-based app that automates the process of editing Microsoft Word documents. Using the Apache POI library, it can replace header and footer text (such as enrollment numbers or IDs) and generate multiple `.docx` files in one go — perfect for institutions or bulk documentation workflows.

---

![EnrollEdits Banner](screenshots/enrolledits-banner.png)

## 📸 Screenshots

| Home Screen | Dark Mode |
|-------------|-----------------|----------------|----------------|
| ![Home](https://github.com/user-attachments/assets/6157f685-8b3c-48e7-b464-6b13cd04801e) | ![Dark Mode](https://github.com/user-attachments/assets/6b990afb-e901-4c1f-974b-8ad75eaa3b58) | 
---

## ✨ Key Features

- 📝 Edit **header** of `.docx` files in bulk
- 📄 Based on a Word **template** with placeholders
- 🔁 Replace content like **enrollment numbers** etc.
- 📦 Generates multiple Word files in one tap
- 📂 Organize output files into folders
- ⚙ Built with Kotlin & Apache POI for performance and flexibility
- 💼 Use-case: schools, colleges, admin departments, event management

---

## ⚙️ How It Works

1. **Template**: You create a Word file with placeholders in the header/footer like: Enrollment No: {{22BECE30157}}
2. 
3. **Generate**: The app reads the list, replaces placeholders in each copy, and outputs:
- `ENR001.docx`
- `ENR002.docx`
- `ENR003.docx`
...in your chosen directory.

---

## 🛠️ Tech Stack

- **Language**: Kotlin
- **UI**: XML-based layout
- **File Handling**: Kotlin IO
- **Target Platform**:  Android
---
