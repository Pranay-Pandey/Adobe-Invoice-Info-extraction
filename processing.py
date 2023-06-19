import json
import re 
import os
import zipfile
import pandas as pd
from extract_pdf import extract_text_and_table_info

class PDFProcessor:
    def __init__(self, working_directory):
        self.working_directory = working_directory # The directory where the pdfs are located
        self.json_location = os.path.join(working_directory, 'structuredData.json') # The location of extracted json data
        self.excel_directory = os.path.join(working_directory, 'tables')    # The directory where the extracted tables are located
        self.output_csv_file_name = os.path.join(working_directory, 'output.csv')   
        self.output_file_name = os.path.join(working_directory, 'output.zip')

    def delete_contents(self, folder_path):
        # Function to delete the files and folders that were extracted from the pdf
        for item in os.listdir(folder_path):
            if ('.csv' in item):
                continue
            item_path = os.path.join(folder_path, item)
            if os.path.isfile(item_path):
                os.remove(item_path)
            elif os.path.isdir(item_path):
                self.delete_contents(item_path) 
                os.rmdir(item_path)

    def extract_info_from_pdf(self, loc_of_pdf):
        # Extracts the text and table information from the pdf and stores it as json and xlsx files
        extract_text_and_table_info(loc_of_pdf, self.output_file_name)

        with zipfile.ZipFile(self.output_file_name, 'r') as zip_ref:
                zip_ref.extractall(self.working_directory)


    def merge(self, dict1, dict2):
        # Merges two dictionaries
        return (dict2.update(dict1))


    def check_for_date(self, text, extracted_data):
        # Checks for date in the text and stores it in the json
        for segments in text:
            try:
                dt = pd.to_datetime(segments, dayfirst=True, errors="coerce")
                if pd.notnull(dt):  # if the date is valid
                    extracted_data['Invoice__IssueDate'] = segments
                    break
            except ValueError:
                pass

    def fragmented_text(self, data):
        # Splits the text into segments
        segments = re.split(',| ', data)
        return [part for part in segments if part!='']


    def set_Header_attributes(self, extracted_data, index, data):
        """
        This function is responsible for setting the information about the business and the invoice in a JSON file.
        When called, this function scans the data from the provided index and sequentially reads the JSON file,
        extracting and storing the required information according to a predefined template based on a PDF document.
        The extracted information is processed accordingly and stored in the JSON file.

        Args:
            index (int): The index to start scanning the data from.
            new_json (dict): The JSON file to store the extracted information in.
            data (list): The list of dictionaries containing the information extracted from the PDF document.

        Returns:
            index (int): The index till which the data has been scanned.
        """

        while('Text' not in data[index]):
            index = index+1
        # Extracting the business information
        if data[index+1]['Text'].replace(" ", "").isnumeric():
            text = self.fragmented_text(data[index]['Text'])
            extracted_data['Bussiness__Country'] = text[-2]+", "+text[-1]
            extracted_data['Bussiness__City'] = text[-3]
            extracted_data['Bussiness__StreetAddress'] = " ".join(text[:-3])
            index = index+1
            extracted_data['Bussiness__Zipcode'] = data[index]['Text']
        else:
            text = self.fragmented_text(data[index]['Text'])
            extracted_data['Bussiness__StreetAddress'] = " ".join(text[:-1])
            extracted_data['Bussiness__City'] = text[-1]
            index = index+1
            extracted_data['Bussiness__Country'] = data[index]['Text']
            index = index+1
            if 'Text' in data[index]:
                extracted_data['Bussiness__Zipcode'] = data[index]['Text']
        index = index+1
        return index

    def get_personal_info_from_table(self, extracted_data):
        """
        Locates and reads an Excel file containing the customer's personal information.

        This function is responsible for finding the Excel file that contains the personal information of the customer.
        Once the file is located, it reads the Excel file and extracts the customer's personal information and other details.
        """
        main_file = None
        for file in os.listdir(self.excel_directory):
            if '.xlsx' in file:
                excel_file = pd.read_excel(os.path.join(self.excel_directory, file))
                if (len(excel_file.columns)==3):
                    # if the excel file has 3 columns, it is the file containing the customer's personal information
                    main_file = pd.read_excel(os.path.join(self.excel_directory, file))
                    break
        if main_file is None:
            return
        
        # Removing the extra characters from the data
        for i in main_file:
            main_file[i] = main_file[i].astype(str).str.replace("_x000D_", "")

        if (len(main_file[main_file.columns[0]])!=5):
            # if the excel file has only one column, then all information 
            # needs to be extracted from the combined text of all the cells
            complete_text = ""
            for i in range(len(main_file[main_file.columns[0]])):
                complete_text += main_file[main_file.columns[0]][i]

            complete_text = complete_text.replace("_x000D_", "")
            texts = self.fragmented_text(complete_text)

            # Extracting the customer's personal information
            extracted_data['Customer__Name'] = " ".join(texts[:2])
            index = 2
            extracted_data['Customer__Email'] = texts[index]
            if '.com' not in texts[index]:
                index = index+1
                extracted_data['Customer__Email'] = extracted_data['Customer__Email'][:-1] +texts[index]
            extracted_data['Customer__PhoneNumber'] = texts[index+1]
            extracted_data['Customer__Address__line1'] = " ".join(texts[index+2:-2])
            extracted_data['Customer__Address__line2'] = " ".join(texts[-2:])

        else:   
            # if the excel file has 5 columns, then all fields are seperated
            for i in range(len(main_file[main_file.columns[0]])):
                if i==0:
                    extracted_data['Customer__Name'] = main_file[main_file.columns[0]][i]
                if i==1:
                    extracted_data['Customer__Email'] = main_file[main_file.columns[0]][i]
                if i==2:
                    extracted_data['Customer__PhoneNumber'] = main_file[main_file.columns[0]][i]
                if i==3:
                    extracted_data['Customer__Address__line1'] = main_file[main_file.columns[0]][i]
                if i==4:
                    extracted_data['Customer__Address__line2'] = main_file[main_file.columns[0]][i]
        
        extracted_data['Invoice__Description'] = ''
        for text in main_file[main_file.columns[1]]:
            extracted_data['Invoice__Description'] += str(text)

        text = self.fragmented_text(main_file[main_file.columns[2]][0])
        if '$' not in text[-1]:
            extracted_data['Invoice__Due_Date'] = text[-1]
        else:
            extracted_data['Invoice__Due_Date'] = text[-2]

        return  

            
    def start_personal_info(self, index, json_object, data):
        """
        Extracts customer information from a PDF document or an Excel file.

        This function is responsible for extracting the information about the customer from the provided PDF document.
        When called, it scans the data from the specified index and sequentially reads the PDF document,
        extracting and storing the required information based on a predefined template specific to the PDF format.
        If the information does not adhere to the predefined template, the function falls back to extracting the
        information from an Excel file using the get_personal_info_from_table function.

        Args:
            index (int): The index to start scanning the data from.
            json_object (object): The JSON object to store the extracted information.
            data (object): The data source containing the customer information.

        Returns:
            index (int): The index till which the data has been scanned.
        """ 
        try:
            # if first element contains the email address then the customer information
            # is not segregated and needs to be extracted from the combined text
            if '@' in data['elements'][index]['Text']:
                texts = self.fragmented_text(data['elements'][index]['Text'])
                json_object['Customer__PhoneNumber'] = texts[-1]
                json_object['Customer__Email'] = texts[-2]
                if ".com" not in texts[-2]:
                    json_object['Customer__Email'] = texts[-3]+texts[-2]
                    json_object['Customer__Name'] = " ".join(texts[:-3])
                else:
                    json_object['Customer__Name'] = " ".join(texts[:-2])
                index += 1

                texts = self.fragmented_text(data['elements'][index]['Text'])
                json_object['Customer__Address__line1'] = " ".join(texts[:-1])
                json_object['Customer__Address__line2'] = texts[-1]
                index += 1

            else:

                # extract the customer information from the predefined template
                if 'Text' in data['elements'][index]:
                    json_object['Customer__Name'] = data['elements'][index]['Text']

                index += 1
                if 'Text' in data['elements'][index]:
                    json_object['Customer__Email'] = data['elements'][index]['Text']

                if 'Text' not in data['elements'][index] or '.com' not in json_object['Customer__Email']:
                    index += 1
                    if 'Text' in data['elements'][index]:
                        json_object['Customer__Email'] = json_object['Customer__Email'][:-1] + \
                        data['elements'][index]['Text']

                index += 1
                if 'Text' in data['elements'][index]:
                    json_object['Customer__PhoneNumber'] = data['elements'][index]['Text']
                
                index += 1
                if 'Text' in data['elements'][index]:
                    json_object['Customer__Address__line1'] = data['elements'][index]['Text']
                    index += 1
                
                if 'Text' in data['elements'][index]:
                    json_object['Customer__Address__line2'] = data['elements'][index]['Text']
                    index += 1
            
            while 'Text' not in data['elements'][index] or 'DETAILS' in data['elements'][index]['Text'] or \
            'PAYMENT' in data['elements'][index]['Text']:
                index += 1

            json_object['Invoice__Description'] = ""
            if ('Text' in data['elements'][index-1] and 'DETAILS' in data['elements'][index-1]['Text']):
                texts = self.fragmented_text(data['elements'][index]['Text'])
                if (len(texts)>5):
                    index -= 1

            while 'Text' not in data['elements'][index] or not (('PAYMENT' in data['elements'][index]['Text']) or
                                                            ('Due date' in data['elements'][index]['Text'])):
                if 'Text' in data['elements'][index]:
                    json_object['Invoice__Description'] += data['elements'][index]['Text'] + " "
                index += 1

            if 'DETAILS' in json_object['Invoice__Description']:
                json_object['Invoice__Description'] = json_object['Invoice__Description'][7:]
            
            while 'Text' not in data['elements'][index] or 'PAYMENT' in data['elements'][index]['Text']:
                index += 1
            
            if 'Text' in data['elements'][index]:
                json_object['Invoice__Due_Date'] = data['elements'][index]['Text'].split(" ")[2]
                index += 1

        except:
            # if the information is not in the predefined template, 
            # extract the information from the table
            self.get_personal_info_from_table(json_object)

        return index

    def initialize_json(self, dict_object):
        # initialize the json object with empty strings
        attribute_list = [
            "Bussiness__Name",
            "Bussiness__StreetAddress",
            "Bussiness__City",
            "Bussiness__Country",
            "Bussiness__Zipcode",
            "Invoice__Number",
            "Invoice__IssueDate",
            "Bussiness__Description",
            "Customer__Name",
            "Customer__Email",
            "Customer__PhoneNumber",
            "Customer__Address__line1",
            "Customer__Address__line2",
            "Invoice__Description",
            "Invoice__Due_Date",
            "Invoice__Tax"
        ]   
        for attribute in attribute_list:
            dict_object[attribute] = ""


    def parse_json(self, json_location):
        """
        Extracts business and customer information from a JSON file.

        This function initializes the JSON object and calls the necessary functions to extract
        business and customer information from the data extracted from a PDF. The extraction process
        is performed based on the order of appearance of various attributes or flags in the PDF.

        Args:
            json_location (str): The location of the data source containing the information from the PDF.

        Returns:
            new_json (object): The JSON object containing the extracted information.
        """
        extracted_data = {}
        self.initialize_json(extracted_data)
        with open(json_location) as f:

            data = json.load(f)
            extracted_data['Bussiness__Name'] = data['elements'][0]['Text']
            # the first element of the JSON file contains the business name
            index_scanned = self.set_Header_attributes(extracted_data, 1, data['elements'])
        
            for i in range(index_scanned, len(data['elements'])):
                block = data['elements'][i]
                loc = block['Path']
                if 'Text' not in block:
                    continue

                if "Invoice" in block['Text']:
                    extracted_data['Invoice__Number'] = block['Text'].split(" ")[1]
                    if block['Text'].split(" ")[1]=='' or block['Text'].split(" ")[1]==' ':
                        extracted_data['Invoice__Number'] = data['elements'][i+1]['Text']

                self.check_for_date(self.fragmented_text(block['Text']), extracted_data)
                if loc == '//Document/Sect/Title' or extracted_data['Bussiness__Name'] in block['Text']:
                    extracted_data['Bussiness__Description'] = data['elements'][i+1]['Text']
                    break

            index_scanned = i
            table_started_flag = False

            for i in range(0, len(data['elements'])):

                block = data['elements'][i]
                loc = block['Path']

                if 'Text' not in block:
                    continue

                if 'BILL TO' in block['Text']:
                    table_started_flag = True
                    continue

                if table_started_flag:
                    if 'BILL TO' in block['Text'] or 'DETAILS' in block['Text'] or 'PAYMENT' in block['Text']:
                        continue

                    i = self.start_personal_info(i, extracted_data, data)
                    table_started_flag = False
                    break

            tax_text_start_flag = False
            for i in range(len(data['elements'])-1, 0, -1):

                if (tax_text_start_flag):
                    while('Text' not in data['elements'][i]):
                        i -= 1
                    text = self.fragmented_text(str(data['elements'][i]['Text']))
                    extracted_data['Invoice__Tax'] = text[-1]
                    break

                if 'Text' in data['elements'][i] and 'Total Due' in data['elements'][i]['Text']:
                    tax_text_start_flag = True

        return extracted_data
    
    def get_content_details(self):
        """
        Extracts content details from an Excel file.

        This function is responsible for extracting content details, such as the name of the product,
        quantity, and rate of the product mentioned in the invoice PDF file. The content details are
        stored in one of the Excel files saved in the Excel directory. This function searches for the
        relevant Excel file and returns a list of JSON objects containing the content details.

        Returns:
            list_of_json_objects (list): The list of JSON objects containing the content details.
        """
        names=['Invoice__BillDetails__Name', 'Invoice__BillDetails__Quantity', 'Invoice__BillDetails__Rate', 'amount']
                
        main_file = None
        for file in os.listdir(self.excel_directory):

            if '.xlsx' in file:
                excel_file = pd.read_excel(os.path.join(self.excel_directory, file), header=None)
                if (len(excel_file.columns)==4 and 'Subtotal' not in excel_file[0].values[0]):
                    # the table should have 4 columns 
                    main_file = pd.read_excel(os.path.join(self.excel_directory, file), header=None, 
                                            names=names)
                    
        if (main_file is None):
            print("Error! No excel file found")
            exit()

        list_of_items = []
        for i in main_file:
            main_file[i] = main_file[i].astype(str).str.replace("_x000D_", "")

        for i in range(len(main_file)):
            new_item = {}
            for name in names:
                if name!='amount':
                    new_item[name] = main_file[name][i]
            list_of_items.append(new_item)

        return list_of_items

    def process(self, loc_of_pdf, save_csv = True, delete_files_after_processing = True, 
                save_pdf_file_name = False):
         
         """
            Process the PDF file and extract information.

            This is the main function of the class that encapsulates all the other functions. It is called
            by the user to extract information from a PDF file. The function calls the necessary functions
            to extract information and returns a Pandas DataFrame containing the information.

            Args:
                loc_of_pdf (str): The location of the PDF file.
                save_csv (bool, optional): Flag indicating whether to save the extracted information as a CSV file. 
                    Defaults to True.
                delete_files_after_processing (bool, optional): Flag indicating whether to delete the files containing the
                raw information extracted from the Pdf file. Defaults to True.
                save_pdf_file_name (bool, optional): Flag indicating whether to save the PDF file name as part of 
                    the extracted information. Defaults to False.

            Returns:
                pandas.DataFrame: A DataFrame containing the extracted information.
            """
         self.extract_info_from_pdf(loc_of_pdf)
         extracted_info = self.parse_json(self.json_location)
         # extracted_info now contains the information extracted from the PDF file

         if save_pdf_file_name:
             extracted_info['pdf_file_name'] = os.path.basename(loc_of_pdf).split('/')[-1]
         
         # get the content details from the excel file
         list_of_json_objects = self.get_content_details()
                 
         for json_object in list_of_json_objects:
             self.merge(extracted_info, json_object)
                     
         df = pd.DataFrame(list_of_json_objects)
         
         if save_csv:            
             df.to_csv( self.output_csv_file_name, index=False)
         
         if (delete_files_after_processing):
             self.delete_contents(self.working_directory)

         return df