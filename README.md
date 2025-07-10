# StatementWise: A Smart Statement Parser

A Python tool to parse credit card statement Excel files into clean, structured data, with a focus on preparing expenses for tax purposes. It uses Polars for high-performance processing and Rich for beautiful command-line reporting.

---

## Key Features ✨

* **Extensible Design**: Easily add new parsers for different banks (e.g., Chase, Capital One) by inheriting from the abstract base class.
* **Intelligent Header Detection**: Automatically finds the transaction table within an Excel sheet, even if it's not at the top.
* **High-Performance**: Uses the Polars DataFrame library for fast and memory-efficient data manipulation.
* **Rich CLI Output**: Presents statement summaries and data previews in a clean, easy-to-read format using Rich.
* **Configurable Logging**: Leverages Loguru for straightforward and configurable debug messages.

---

## Setup and Usage 🚀

1.  **Clone the Repository**
    ```sh
    git clone [https://github.com/ivy113/statementwise.git](https://github.com/ivy113/statementwise.git)
    cd statementwise
    ```

2.  **Install Dependencies**
    ```sh
    # Make sure you have a requirements.txt file
    pip install -r requirements.txt
    ```

3.  **Install for Development**
    To make edits to the source code, install the project in "editable" mode. This allows your changes to be reflected immediately without reinstalling.
    ```sh
    pip install -e .
    ```

4.  **Add Your Data**
    Place your private Excel statement files inside the `data/` directory. This directory is intentionally ignored by Git (via `.gitignore`) to keep your financial data secure and off of GitHub.

5.  **Run the Script**
    Execute the main script to parse your files.
    ```sh
    python scripts/preview.py
    ```