from processing import *

# Initialize PDFProcessor object
pdf_processor = PDFProcessor(working_directory='Extractions')

# Process the first PDF file and create a pandas DataFrame
pdf_file_path = 'TestDataSet/output80.pdf'
df = pdf_processor.process(pdf_file_path, save_csv=False,
                           delete_files_after_processing=True,
                           save_pdf_file_name=False)

# Process remaining PDF files and concatenate their extracted data to the first DataFrame
for i in range(1, 100):
    pdf_file_path = fr'InvoicesData/TestDataSet/output{i}.pdf'
    df_temp = pdf_processor.process(pdf_file_path, save_csv=False)
    df = pd.concat([df, df_temp], axis=0)

# Save the final DataFrame as a CSV file in the working directory
csv_file_path = r'Extractions/extracted_data.csv'
df.to_csv(csv_file_path, index=False)