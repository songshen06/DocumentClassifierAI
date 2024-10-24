# DocumentClassifierAI
**Document Classifier Script - Usage Guide**

### **Environment Setup**

To use this document classification script, you need to set up the required environment and dependencies as follows:

1. **Clone the Repository**: Ensure you have the script and its associated files cloned to your local machine.

2. **Install Python Dependencies**:
   - You need Python 3.9 or later.
   - Install dependencies using `pip`. Run the following command in your terminal:
     ```bash
     pip install -r requirements.txt
     ```
   - The required Python libraries are:
     - `openai` (for interacting with LLMs)
     - `PyPDF2` and `pdfminer` (for PDF text extraction)
     - `python-pptx` (for extracting text from PowerPoint files)
     - `python-docx` (for extracting text from Word files)
     - `python-dotenv` (for loading environment variables)

3. **Set Up Environment Variables**:
   - Create a `.env` file in the root directory of your project with the following keys:
     ```
     DEEPKEY=<your_deepseek_api_key>
     NVIDIA_key=<your_nvidia_api_key>
     ```
   - Replace `<your_deepseek_api_key>` and `<your_nvidia_api_key>` with valid API keys from DeepSeek and NVIDIA, respectively.

### **Requirements**

- Python 3.9 or later
- Access to the APIs for `ollama`, `deepseek`, and `nvidia_nim` models.
- Required software libraries: `openai`, `PyPDF2`, `pdfminer`, `python-pptx`, `python-docx`, and `dotenv`.
- Ensure that your machine allows running the Ollama model locally, which listens on `http://localhost:11434/v1`.

### **Usage Instructions**

You can use this script to classify documents based on content using different LLM providers. The script can be run via the command line with customizable arguments.

1. **Command Line Arguments**

   - `-FT` or `--file_type` (default: `all`): Specify the type of files to process. Options are `pdf`, `docx`, `pptx`, or `all`.
   - `-OD` or `--output_dir` (default: `.`): Specify the directory where the output file will be saved. Defaults to the current directory.
   - `-OF` or `--output_filename` (default: `classification_results.csv`): Specify the output CSV file name.
   - `-P` or `--provider` (default: `ollama`): Specify the LLM provider for classification. Options are `ollama`, `deepseek`, or `nvidia_nim`.

2. **Example Usage**

   To classify all supported document types and save the output in the current directory:
   ```bash
   python DocumentClassifierAI.py -FT all -OD . -OF classification_results.csv -P ollama
   ```

   To classify only DOCX files and save the output to a specific folder:
   ```bash
   python DocumentClassifierAI.py -FT docx -OD /path/to/output -OF docx_classification.csv -P deepseek
   ```

3. **How the Script Works**

   - **Document Selection**: The script reads documents from the specified directory (`/Users/shens/Downloads` by default).
   - **Document Extraction**: Depending on the file type (`pdf`, `docx`, or `pptx`), text is extracted using the respective library (`PyPDF2`, `python-docx`, or `python-pptx`). Text extraction is limited to 1000 characters per document to manage token limitations in LLMs.
   - **Summarization**: The Ollama model is used to summarize the document content before classification.
   - **Classification**: The chosen LLM provider (`ollama`, `deepseek`, or `nvidia_nim`) classifies the summarized document into one of four categories: `Omniverse`, `vGPU`, `NVAIE`, or `Uncategorized`.
   - **Save Results**: The classification results are saved in a CSV file, with columns for the document name and category.

### **Customizing Document Classification**

If you want to classify your documents into custom categories, you need to modify the **prompt** used for classification and update the product keywords accordingly.

1. **Modify Product Keywords**:
   - In the script, locate the `product_keywords` dictionary in the `main()` function. You can add, remove, or update the categories and associated keywords to suit your needs.
   - Example:
     ```python
     product_keywords = {
         "Finance": ["investment", "stocks", "banking", "financial planning"],
         "Healthcare": ["medical", "healthcare", "hospital", "pharmaceutical"],
         "Technology": ["AI", "machine learning", "software", "cloud computing"]
     }
     ```

2. **Update the Prompt for Classification**:
   - Locate the `create_completion` method inside the `classify_documents()` function. Modify the system message to reflect your new categories.
   - Example:
     ```python
     response = ai_wrapper.create_completion(
         messages=[
             {"role": "system", "content": (
                 "You are an assistant that classifies documents into the following categories: {}. "
                 "Please only return the category name without any additional information. The categories are: Finance, Healthcare, Technology, Uncategorized.\n"
                 "For example, if the document discusses investment or stocks, return 'Finance'. If the document discusses medical or healthcare topics, return 'Healthcare'.\n"
                 "Make sure to only return one of the following words: Finance, Healthcare, Technology, Uncategorized. No other text should be included in the response."
             ).format(list(product_keywords.keys()))},
             {"role": "user", "content": summary}
         ],
         stream=False
     )
     ```

### **Notes**

- The script ignores temporary files that start with `$` (often generated by office applications on macOS).
- Ensure that the Ollama server is running locally (`http://localhost:11434/v1`) if using the Ollama model for classification.
- The output directory must exist; otherwise, an error message will be displayed.
- The output filename must have a `.csv` extension.

