# Adobe Invoice Info extraction

 <p>
    This project aims to extract invoice details from PDF files using the Adobe PDF Extract API. The extracted information includes customer and business details, as well as product and tax information.
  </p>
  <h2>Problem Statement</h2>
  <p>
    The problem at hand is to extract relevant information from invoices provided in PDF format. The required details include the customer's address, business information, product details, and tax information. By automating this extraction process, it becomes easier to analyze and manage large volumes of invoices efficiently.
  </p>
  <h2>Solution Overview</h2>
  <p>
    The solution utilizes the Adobe PDF Extract API to extract both text and table information from the provided PDF files. The extracted data is then processed and divided into three major parts: business information, customer information, and product details.
  </p>
  <h3>Business Information Extraction</h3>
  <p>
    The API generates a JSON file containing the information extracted from the PDF file. This JSON file is parsed line by line and compared against a predefined template provided in the code. If a match is found, the corresponding parts such as the business name, location, zip code, and invoice number are selected and stored.
  </p>
  <h3>Customer Information Extraction</h3>
  <p>
    The next section of the PDF file contains the customer information. The JSON file containing this information is sequentially read from the starting point of the customer details. Each entry in the JSON file is processed one by one, updating the customer's name, email ID, address, and other relevant details. In cases where the JSON file does not match the line-by-line template, the code attempts to find an Excel file containing a table with the customer information. If found, the information from this table is extracted.
  </p>
  <h3>Product Details Extraction</h3>
  <p>
    The code searches for an Excel file containing a table with the product details, including the name, rate, and quantity. This information is extracted from the Excel file and stored in a list for further processing.
  </p>
  <h2>Data Integration and Output</h2>
  <p>
    Once the customer and business information, as well as the product details, have been extracted, the two sets of information are combined. The combined data is then saved in CSV format, which satisfies the required output format.
  </p>
  <h2>Code Structure</h2>
  <p>
    The code is structured in an object-oriented manner, utilizing classes and functions to organize the workflow and provide flexibility for future enhancements. The main parts used in the code include:
    <ul>
      <li>PDFProcessor: This class encapsulates the functionality related to PDF extraction using the Adobe PDF Extract API. It handles the authentication process and provides methods to extract both text and table information from PDF files. This class object needs to be created to </li>
      <li>extract_info_from_pdf: This method is responsible for calling the pdf extraction API to extract information and save the extracted files in the form of json and excel files. These are initially saved as compressed files and later extracted.</li>
      <li>process: This method encapsulates the other small methods for extraction of the information from the files. We can pass parameters such as save_csv, delete_files_after_processing and save_pdf_file_name from this method.</li>
    </ul>
  </p>

![image](https://github.com/Pranay-Pandey/Adobe-Invoice-Info-extraction/assets/79053599/3a2fc3bd-3882-444c-9b52-254f82e45ce4)

