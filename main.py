from processing import *

# running the main function for a file
working_directory = r'Procedure\Extractions'
pdf_processor = PDFProcessor(working_directory)

pdf_file_loc = r'InvoicesData\TestDataSet\output0.pdf'
df1 = pdf_processor.process(pdf_file_loc, save_csv=False, 
                            delete_files_after_processing=True, save_pdf_file_name=False)

for i in range(1, 100):
    pdf_file_loc = r'InvoicesData\TestDataSet\output' + str(i) + '.pdf'
    df2 = pdf_processor.process(pdf_file_loc, False)
    df1 = pd.concat([df2, df1], axis=0)
    print(i)

df1.to_csv(r'Procedure\Extractions\extracted_data.csv', index=False)