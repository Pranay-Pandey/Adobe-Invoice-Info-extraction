from processing import *

# running the main function for a file
working_directory = r'TestDataSet'
pdf_processor = PDFProcessor(working_directory)

pdf_file_loc = r'TestDataSet\output0.pdf'
df1 = pdf_processor.process(pdf_file_loc, save_csv=False, 
                            delete_files_after_processing=True, save_pdf_file_name=False)

for i in range(1, 100):
    pdf_file_loc = r'TestDataSet\output' + str(i) + '.pdf'
    df2 = pdf_processor.process(pdf_file_loc, False)
    df1 = pd.concat([df2, df1], axis=0)

df1.to_csv(r'extracted_data.csv', index=False)