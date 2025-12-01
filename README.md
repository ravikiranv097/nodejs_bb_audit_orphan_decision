---

# # 📘 Orphan Bitbucket Access Audit – Automation Script

This project automates the complete **Orphan Decision Sheet → Bitbucket Access Validation** workflow.
It reads an XLSX file, extracts unique users, checks their Bitbucket project access, captures HTML & PNG evidence, and generates DOCX reports.

This automation replaces manual review work and produces **audit-ready outputs**.

---

# ## 🚀 Features

### ✔ 1. **Parse Orphan Decision Sheet (XLSX)**

* Reads `Orphan_Decision_Sheet_Dummy.xlsx`
* Extracts unique rows based on **User SSO**
* Parses entitlement to extract:

  * **Project Key**
  * **Access Permission**

### ✔ 2. **Bitbucket Access Check**

For each user/project combination:

* Calls Bitbucket REST API
* Extracts permission data
* Categorizes results into:

  * **HAS_ACCESS**
  * **NO_ACCESS**

### ✔ 3. **Evidence Generation**

For every user:

| Status     | HTML Evidence | PNG Evidence |
| ---------- | ------------- | ------------ |
| HAS_ACCESS | ✔ yes         | ✔ yes        |
| NO_ACCESS  | ✔ yes         | ✔ yes        |

Outputs:

```
output_files/
    html/
      has_access/
      no_access/
    png/
      has_access/
      no_access/
```

### ✔ 4. **CSV Outputs**

Two CSV files are generated:

```
orphan_access_results.csv         → users with Access
orphan_no_access_results.csv      → users without Access
```

Each row includes reference paths to HTML and PNG evidence.

### ✔ 5. **DOCX Audit Reports**

Automatically generates:

```
output_files/doc/
    Orphan_Has_Access_Report.docx
    Orphan_No_Access_Report.docx
```

Each DOCX file contains screenshots (PNG) of all users in that category.

---

# ## 📂 Project Structure

```
project-root/
│── orphan_audit.js
│── package.json
│── README.md
│── input_files/
│     └── Orphan_Decision_Sheet_Dummy.xlsx
│
└── output_files/
      ├── orphan_unique_rows.csv
      ├── formatted_orphan_rows.csv
      ├── orphan_access_results.csv
      ├── orphan_no_access_results.csv
      │
      ├── html/
      │     ├── has_access/
      │     └── no_access/
      │
      ├── png/
      │     ├── has_access/
      │     └── no_access/
      │
      └── doc/
            ├── Orphan_Has_Access_Report.docx
            └── Orphan_No_Access_Report.docx
```

---

# ## 🛠 Installation

1. Clone the project

   ```bash
   git clone <repo-url>
   cd project-root
   ```

2. Install dependencies

   ```bash
   npm install
   ```

3. Create `.env` file

   ```
   BB_URL=localhost:7990
   BB_USERNAME=admin
   BB_KEYNAME=YOUR_BITBUCKET_TOKEN
   ```

4. Place the XLSX input inside:

   ```
   input_files/Orphan_Decision_Sheet_Dummy.xlsx
   ```

---

# ## ▶️ Run the script

```bash
node orphan_audit.js
```

After the run completes, results will be inside the `output_files/` directory.

---

# ## 🔗 Bitbucket API Used

This script uses Bitbucket Server (self-hosted) REST API:

```
GET /rest/api/1.0/projects/{projectKey}/permissions/users?filter={username}
```

---

# ## 📦 Dependencies

The project uses:

| Package       | Purpose                         |
| ------------- | ------------------------------- |
| **xlsx**      | Read XLSX input                 |
| **axios**     | Bitbucket REST API calls        |
| **puppeteer** | Generate PNG evidence           |
| **officegen** | Generate DOCX audit reports     |
| **dotenv**    | Environment variable management |

---

# ## 📄 Output Samples

### ✔ Example CSV row

```
user1,account1,PA,Write,HAS_ACCESS,2025-01-01 11:20:00,html/has_access/user1_PA.html,png/has_access/user1_PA.png
```

### ✔ DOCX screenshot example

Each page of the report contains one screenshot for audit purposes.

---

