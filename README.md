# Footer Batch 🚀

**Footer Batch** is a lightning-fast, privacy-first web application that allows you to add custom footers (Name, Class, Roll No, and Page Numbers) to multiple PDF and Word (.docx) documents simultaneously.

## ✨ Why Footer Batch?

- **Zero Latency**: All processing happens directly in your browser. No more waiting for file uploads or server responses.
- **Privacy First**: Your documents **never leave your computer**. Processing is 100% local, keeping your data secure.
- **Batch Processing**: Process up to 10 documents at once and download them as a single ZIP or individual files.
- **Live Preview**: See exactly how your footer will look before you process the documents.

## 🛠️ Tech Stack

- **Frontend**: [React](https://reactjs.org/) + [Vite](https://vitejs.dev/)
- **PDF Manipulation**: [pdf-lib](https://pdf-lib.js.org/)
- **Word Manipulation**: [jszip](https://stuk.github.io/jszip/) (Custom OOXML injection)
- **Deployment**: [Vercel](https://vercel.com/) (Static Hosting)

## 🚀 Getting Started

### Prerequisites

- [Node.js](https://nodejs.org/) (v18 or higher)
- npm or yarn

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/Footer-Batch.git
   cd Footer-Batch
   ```

2. Install dependencies for the frontend:
   ```bash
   cd frontend
   npm install
   ```

3. Run the development server:
   ```bash
   npm run dev
   ```

4. Open `http://localhost:5173` in your browser.

## 📦 Deployment

Since the app is now a fully static site, deployment is simple:

1. Connect your repository to **Vercel**.
2. Set the **Root Directory** to `frontend`.
3. Vercel will automatically detect Vite and deploy your app.

## 📄 License

Distributed under the MIT License. See `LICENSE` for more information (if applicable).

---

Built with ❤️ by [Mayank Rai](https://github.com/mayankrai5225-art)
