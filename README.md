# DataMorph // Universal Data Converter & Cleaner 🔄✨

DataMorph is a super functional, Python-based web application that allows you to upload, clean, optimize, and convert complex data files (.csv, .xlsx, .json) into your desired format in seconds.

Equipped with a custom premium dark mode interface, glassmorphism design effects, and real-time data metrics, it makes modern data preparation, cleaning, and conversion an engaging and effortless experience.

### 🔗 Live Demo
Try the application instantly in your browser: **[https://cinarsdatamorphapp.streamlit.app/](https://cinarsdatamorphapp.streamlit.app/)**

---

## 🚀 Features

### 1. Universal File Support (Input & Output)
* **Input Formats**: Easily upload `.csv`, `.xlsx`, or `.json` files.
* **Output Formats**: Convert and download your cleaned data instantly as `.csv`, `.xlsx`, or `.json`.

### 2. Smart Data Cleaning Options
* **Remove Empty & Missing Rows**: Instantly filters out any rows containing empty fields or missing (NaN/null) values.
* **Standardize Text Fields (Title Case)**: Automatically detects all text (string) columns and formats them into Title Case (e.g., `jOHN sMITH` ➔ `John Smith`). Safely preserves existing missing values without corrupting columns.
* **Clean Duplicate Rows**: Removes identical duplicate records to safeguard your data integrity.

### 3. Premium UI & Advanced UX
* **Dynamic Metric Cards**: Displays total rows, column count, file size, and missing values in real-time inside sleek glassmorphic containers.
* **Real-time Change Analysis**: Dynamically calculates and displays exactly how many rows are filtered out as you select/deselect cleaning options.
* **Side-by-Side Preview Tabs**: Seamlessly compare the "Processed Data" preview and the "Original Raw Data" preview side-by-side using high-fidelity tabs.
* **Custom Styling**: Fully customized theme containing the modern "Outfit" font family, frosted glass cards (`backdrop-filter: blur(12px)`), elegant gradients, and micro-animated widgets.

---

## 🛠️ Tech Stack

* **Frontend & Web Server**: [Streamlit](https://streamlit.io/)
* **Data Processing & Conversion**: [Pandas](https://pandas.pydata.org/)
* **Excel File Support**: [Openpyxl](https://openpyxl.readthedocs.io/)
* **Design & Styling**: HTML5, Custom Vanilla CSS3

---

## 💻 Installation & Running Locally

Follow these simple steps to run the application locally on your computer:

### 1. Clone the Repository
```bash
git clone https://github.com/username/DataMorph.git
cd DataMorph
```

### 2. Install Dependencies
Open your terminal in the project directory and install all required libraries using:
```bash
pip install -r requirements.txt
```

### 3. Run the Application
Start the Streamlit local web server:
```bash
python -m streamlit run app.py
```

The application will launch automatically in your web browser. If it doesn't open, copy and paste the local network address (usually **`http://localhost:8501`**) shown in your terminal.

---

## 📂 Project Structure

```text
DataMorph/
├── .gitignore          # Files to ignore in Git version control
├── README.md           # Project documentation and guide
├── app.py              # Main application source code (UI & Backend)
├── requirements.txt    # Required Python packages list
├── ornek_veri.csv      # Sample CSV file for testing
└── ornek_veri.json     # Sample JSON file for testing
```

---

## 📜 License

This project is licensed under the [MIT License](LICENSE). You are free to use, modify, and distribute it as you wish.
