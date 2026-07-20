# 📊 Excel Report Generator

A full-stack **Next.js** application that automates the generation of standardized Excel reports from a predefined Excel template.

Users can fill in a simple form, upload photographic evidence, and instantly download a fully formatted Excel report. The application processes everything **in memory**, ensuring that no personal data or uploaded files are permanently stored.

---

# ✨ Features

- 📄 Generate Excel reports from a predefined template
- 📷 Drag & Drop image upload
- 🖼️ Automatic photo placement into the "Photographic Record" worksheet
- 📑 Preserve Excel formatting, styles, formulas, merged cells, and print settings
- 📐 Keep A4 layout for PDF export
- 📥 Instant Excel download
- 🔒 No database required
- 🚫 No image persistence
- 📱 Mobile-first responsive interface

---

# 🛠 Tech Stack

## Framework

- Next.js (App Router)
- React
- TypeScript

## UI

- Tailwind CSS
- React Hook Form
- React Dropzone

## Excel Generation

- ExcelJS

## Validation

- Zod

---

# 🏗 Architecture

This project follows a **Backend for Frontend (BFF)** architecture using Next.js.

Instead of maintaining separate frontend and backend applications, Next.js handles both the UI and server-side Excel generation.

```text
Browser
    │
    ▼
Next.js Application
│
├── React UI
│
├── Route Handlers
│
├── Excel Generator Service
│
└── Excel Template
```

---

# 📁 Project Structure

```text
excel-report-generator/

├── public/
│
├── src/
│
│   ├── app/
│   │   ├── page.tsx
│   │   └── api/
│   │       └── reports/
│   │           └── route.ts
│   │
│   ├── components/
│   │   ├── Form/
│   │   ├── Inputs/
│   │   ├── PhotoUpload/
│   │   └── Button/
│   │
│   ├── lib/
│   │   ├── excel-generator.ts
│   │   ├── image-service.ts
│   │   └── template-loader.ts
│   │
│   ├── mapping/
│   │   ├── cells.ts
│   │   └── photos.ts
│   │
│   ├── types/
│   └── utils/
│
├── templates/
│   └── model.xlsx
│
├── package.json
└── README.md
```

---

# ⚙️ Request Flow

```text
User

│

├── Fill form

├── Upload photos

│

▼

POST /api/reports

│

▼

Load model.xlsx

│

▼

Create Workbook (Memory)

│

▼

Fill mapped cells

│

▼

Insert uploaded images

│

▼

Generate XLSX Buffer

│

▼

Return Download

│

▼

Memory Released
```

---

# 📄 Excel Template

The project uses a single immutable template.

```text
/templates/model.xlsx
```

The template is **never modified**.

For every request:

1. Load the template.
2. Create a workbook instance in memory.
3. Populate mapped cells.
4. Insert uploaded photos.
5. Generate a new Excel file.
6. Send the file to the client.
7. Release memory.

The original template remains untouched.

---

# 🗂 Cell Mapping

Business logic is completely separated from the Excel layout.

```ts
export const cells = {
  holderName: 'B5',
  cpf: 'B6',
  city: 'B7',
  inspectionDate: 'H4',
};
```

Photo positioning:

```ts
export const photos = [
  'B3:E15',
  'G3:J15',
  'B18:E30',
  'G18:J30',
  'B33:E45',
  'G33:J45',
];
```

If the report layout changes, only the mapping files need to be updated.

---

# 🔐 Privacy

The application intentionally avoids storing user information.

Nothing is persisted:

- Uploaded images
- CPF
- Holder names
- Generated reports

Everything exists only during the request lifecycle.

---

# 📱 Mobile First

The interface is designed using a mobile-first approach.

Design goals:

- Simple workflow
- Large touch targets
- Responsive layout
- Drag & Drop support
- Accessible form controls

---

# 🚀 Future Improvements

- PDF generation
- QR Code support
- Digital signatures
- Multiple report templates
- Batch report generation
- Template manager
- User authentication
- Report history
- Cloud storage integration

---

# 💻 Getting Started

## Install dependencies

```bash
npm install
```

## Start the development server

```bash
npm run dev
```

Open:

```
http://localhost:3000
```

---

# 📦 Building for Production

```bash
npm run build
npm start
```

---

# 📜 License

This project is licensed under the MIT License.

---

## 💡 Project Philosophy

This project was designed around a simple principle:

> **Keep the Excel template as the single source of truth.**

Instead of recreating spreadsheets through code, the application loads an existing Excel template, injects dynamic data, places uploaded images in predefined positions, and returns a new workbook without ever modifying the original file.

This approach keeps business logic separate from document layout, making the application easier to maintain as report formats evolve.
