# SpreadsheetApp

A modern web application built with React and Vite for creating and managing spreadsheets.

## Prerequisites

Before you begin, ensure you have the following installed on your system:

- **Node.js** (version 18 or higher) - [Download](https://nodejs.org/)
- **npm** (comes with Node.js) or **yarn**
- **Git** - [Download](https://git-scm.com/)

## Getting Started

### 1. Clone the Repository

```bash
git clone https://github.com/tauhidst07/spreadhsheet.git
cd SpreadsheetApp
```

### 2. Install Dependencies

Install all required project dependencies:

```bash
npm install
```

Or if you prefer yarn:

```bash
yarn install
```

### 3. Run Development Server

Start the development server with hot module replacement (HMR):

```bash
npm run dev
```

The application will be available at `http://localhost:5173`

### 4. Build for Production

Create an optimized production build:

```bash
npm run build
```

The build output will be in the `dist/` directory.

### 5. Preview Production Build

Preview the production build locally:

```bash
npm run preview
```

### 6. Lint Code

Run ESLint to check for code quality issues:

```bash
npm run lint
```

## Project Structure

```
SpreadsheetApp/
├── src/
│   ├── App.jsx           # Main React component
│   ├── App.css           # Application styles
│   ├── main.jsx          # Application entry point
│   ├── index.css         # Global styles
│   ├── assets/           # Static assets (images, icons, etc.)
│   └── engine/           # Core application logic
│       └── core.js       # Engine core functionality
├── public/               # Static files served as-is
├── package.json          # Project dependencies and scripts
├── vite.config.js        # Vite configuration
├── eslint.config.js      # ESLint configuration
├── index.html            # HTML entry point
└── README.md             # This file
```

## Technologies Used

- **React** - A JavaScript library for building user interfaces
- **Vite** - A next-generation frontend build tool
- **ESLint** - JavaScript linting utility
- **CSS** - Styling and layout

## Available Scripts

| Command | Description |
|---------|-------------|
| `npm run dev` | Start the development server with hot reload |
| `npm run build` | Build the application for production |
| `npm run preview` | Preview the production build locally |
| `npm run lint` | Run ESLint to check code quality |

## Development Workflow

1. Create a new branch for your feature:
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. Make your changes and ensure the code passes linting:
   ```bash
   npm run lint
   ```

3. Commit your changes:
   ```bash
   git commit -m "Add description of your changes"
   ```

4. Push to your fork and create a Pull Request

## Browser Support

This application works on all modern browsers that support ES2020+ JavaScript:

- Chrome (latest)
- Firefox (latest)
- Safari (latest)
- Edge (latest)

## Troubleshooting

### Dependencies won't install
- Clear npm cache: `npm cache clean --force`
- Delete `node_modules` and `package-lock.json`, then reinstall: `rm -rf node_modules package-lock.json && npm install`

### Port 5173 already in use
- The dev server will automatically try the next available port
- Or specify a custom port: `npm run dev -- --port 3000`

### Build fails
- Ensure all dependencies are installed: `npm install`
- Clear any build cache: `rm -rf dist`
- Try rebuilding: `npm run build`



# task_AI_native_Office_intern

## My Implementation

### 1. Project Overview

Spreadsheet-like web app built with **React + Vite** and a custom **spreadsheet engine** that manages cell data, parses formulas, tracks dependencies, and recalculates computed values.

### 2. Implemented Features

#### Task 1 – Column Sort & Filter

- **Sorting**: Three states per column—ascending, descending, none. Uses `engine.getCell(row, col).computed` so formula results sort with numbers/strings.
- **View-layer only**: `rowOrder` state controls row display; engine data and formulas stay unchanged.
- **Filtering**: Excel-style dropdown per column with checkboxes of unique computed values. Hides rows in UI; engine data untouched.
- **Combined**: Sort via `rowOrder`, then filter via `passesFilters(rowIndex)`.

#### Task 2 – Multi-Cell Copy & Paste

- **Ctrl+C**: Copies computed value of selected cell (not formula text).
- **Ctrl+V**: Parses tab/newline clipboard; compatible with Excel/Sheets. Multi-row/column paste from selected cell.
- **Internal + external**: Paste preserves formulas when pasted as raw text; `engine.setCell` used per cell.
- **Undo**: Pastes are undoable with Ctrl+Z (each cell goes through engine undo).

#### Task 3 – Local Storage Persistence

- **Auto-save**: State written to `localStorage` when data or styles change; **debounced ~500 ms** (setTimeout) to avoid writes on every keystroke.
- **Auto-restore**: On load, snapshot read and applied (engine + styles); same spreadsheet after refresh.
- **Persisted**: Cell raw values/formulas, cell styles, grid dimensions.
- **Not persisted**: Undo/redo history, selection, editing, filter/sort state.
- **Safety**: try/catch on read/parse; corrupted data ignored, app starts fresh.

### 3. Architecture Overview

- **`App.jsx`**: Toolbar, formula bar, grid; state for selection, editing, `rowOrder`, `filters`, `cellStyles`; handles edit, sort, filter, clipboard, undo/redo, row/column ops.
- **`engine/core.js`**: Sparse `Map` of cells (key = `"A1"` etc.); tokenize → parse (shunting-yard) → evaluate (incl. SUM/AVG/MIN/MAX); dependency graph + topological recalc; public API: `getCell`, `setCell`, insert/delete row/col, undo/redo, `serialize`/`deserialize`.

### 4. Key Design Decisions

- **Sort/filter in view layer**: Row order and visibility only; engine addresses unchanged so formulas stay correct.
- **Formulas in engine**: UI only requests computed values.
- **Debounced persist**: Fewer localStorage writes; data still saved promptly.
- **Styles separate**: In React state + persisted with engine snapshot; engine stays data/formula-only.

### 5. Edge Cases Handled

- Empty cells: Treated as distinct in sort/filter; persistence skips empty cells.
- Mixed number/string: Sort orders by type (number → string → empty → error).
- Formula results: Sort/filter use `.computed` so formulas participate correctly.
- Corrupted localStorage: Caught and ignored; fresh sheet on load.
- Large paste: Rows/cols parsed; cells outside grid bounds skipped.

### 6. Running the Project

```bash
npm install    # or yarn install
npm run dev    # or yarn dev
```

Open `http://localhost:5173` and test sort, filter, clipboard, and reload persistence.

