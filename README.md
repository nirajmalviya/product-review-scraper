

# Product Review Scraper

### Introduction
This project is a Python-based web scraper designed to extract product reviews from an e-commerce website. It uses the `curl_cffi` library for making HTTP requests, `openpyxl` for handling Excel files, and `pydantic` for data validation. The script supports sorting reviews by various criteria such as "Most Recent," "Most Helpful," and more. The reviews are saved in an Excel file, categorized by product.

### Prerequisites

Make sure you have the following installed on your system:

- **Python 3.7+**
- Required Python libraries:
  - `curl_cffi`
  - `openpyxl`
  - `pydantic`

### Installation

1. **Clone the repository:**
   Open your terminal or command prompt and run:
   ```bash
   git clone https://github.com/your-username/product-review-scraper.git
   ```

2. **Navigate to the project directory:**
   ```bash
   cd product-review-scraper
   ```

3. **Create a virtual environment (optional but recommended):**
   ```bash
   python -m venv .venv
   ```

4. **Activate the virtual environment:**
   - On **Windows**:
     ```bash
     .venv\Scripts\activate
     ```
   - On **macOS/Linux**:
     ```bash
     source .venv/bin/activate
     ```

5. **Install the dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

### Usage

1. **Prepare the input Excel file:**
   - Create an Excel file with multiple sheets. Each sheet should contain product details such as URL, Product ID, Product Name, Price, and Category.

2. **Run the scraper:**
   - To start scraping, simply run the following command:
     ```bash
     python scraper.py
     ```

3. **Output:**
   - The reviews will be saved in an Excel file with each review containing the product category, name, price, sort type, review description, and rating.

---

This format should give others clear instructions on how to install, configure, and run your project!
