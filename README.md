# Property Review Email Generator

A small React + Vite web app that reads a property review Excel workbook and generates a standardized recommendation email.

## Run locally

```bash
npm install
npm run dev
```

## Build for production

```bash
npm install
npm run build
```

## Deploy on Vercel

1. Create a new GitHub repository.
2. Upload these files to the repo.
3. In Vercel, import the GitHub repo.
4. Use the default settings:
   - Framework preset: Vite
   - Build command: `npm run build`
   - Output directory: `dist`
5. Deploy.

## Notes

- The parser is currently aligned to the workbook structure from the provided sample.
- If future workbooks shift cell locations or sheet names, update `extractReviewData()` in `src/App.jsx`.
