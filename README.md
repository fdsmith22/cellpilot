# CellPilot

Advanced data manipulation and automation tool for Google Sheets.

## 🚀 Features

- **Smart Table Parser**: Convert unstructured data into organized columns
- **Advanced Data Cleaning**: Remove duplicates, standardize text, format dates
- **Formula Builder**: Create complex formulas using natural language
- **Industry Templates**: Pre-built templates for various industries
- **ML-Powered Features**: Smart suggestions and pattern recognition
- **Cross-Sheet Operations**: Work with data across multiple sheets
- **Data Pipeline Manager**: Import/export data with transformations

## 📦 Installation

### For Beta Users

1. Sign up at [cellpilot.io](https://www.cellpilot.io)
2. Activate beta access in your dashboard
3. Click "Install CellPilot" 
4. Follow the installation guide

### For Developers

See [docs/PROJECT_STRUCTURE.md](docs/PROJECT_STRUCTURE.md) for development setup.

## 🛠️ Technology Stack

- **Google Apps Script**: Core add-on functionality
- **Next.js 14**: Web application and user dashboard
- **Supabase**: Authentication and database
- **Tailwind CSS**: Styling
- **TypeScript**: Type safety

## 📁 Project Structure

```
cellpilot/
├── apps-script/        # Google Apps Script code
├── site/              # Next.js web application  
├── docs/              # Documentation
└── test-sheet-proxy.js # Testing proxy functions
```

## 🔧 Development

### Apps Script Development

```bash
cd apps-script
npm install

# Push to Google Apps Script
npm run push:library
npm run push:installer
```

### Web Development

```bash
cd site
npm install
npm run dev
```

## 📄 License

Proprietary - All rights reserved

## 🤝 Support

For support, email support@cellpilot.io
