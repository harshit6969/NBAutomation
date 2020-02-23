import sys
import pandas as pd
import numpy as np
headers = {
    "Area": ["Area", "Type", "Cluster"],
    "Apartment": ["Block", "Floor", "Flat", "Intercom"],
    "FlatOwner": ["Block", "Flat", "BHK", "Sq Feet Area", "Owner Name", "PARKING", "Owner Phone", "Owner Email", "Accomodation Type", "Tenant Name", "Tenant Phone", "Tenant Email", "Resident Type"]
}
def levenshtein_ratio(s, t):
    """ levenshtein_ratio:
        Calculates levenshtein distance between two strings.
        The function computes the
        levenshtein distance ratio of similarity between two strings
        For all i and j, distance[i,j] will contain the Levenshtein
        distance between the first i characters of s and the
        first j characters of t
    """
    # Initialize matrix of zeros
    rows = len(s)+1
    cols = len(t)+1
    distance = np.zeros((rows,cols),dtype = int)

    # Populate matrix of zeros with the indeces of each character of both strings
    for i in range(1, rows):
        for k in range(1,cols):
            distance[i][0] = i
            distance[0][k] = k

    # Iterate over the matrix to compute the cost of deletions,insertions and/or substitutions    
    for col in range(1, cols):
        for row in range(1, rows):
            if s[row-1] == t[col-1]:
                cost = 0 # If the characters are the same in the two strings in a given position [i,j] then the cost is 0
            else:
                # In order to align the results with those of the Python Levenshtein package, if we choose to calculate the ratio
                # the cost of a substitution is 2. If we calculate just distance, then the cost of a substitution is 1.
                cost = 2
            distance[row][col] = min(distance[row-1][col] + 1,      # Cost of deletions
                                 distance[row][col-1] + 1,          # Cost of insertions
                                 distance[row-1][col-1] + cost)     # Cost of substitutions
    # Computation of the Levenshtein Distance Ratio
    return ((len(s)+len(t)) - distance[row][col]) / (len(s)+len(t))

def get_header(value, sheetname):
    header = value
    distance = 0
    for fixed_header in headers[sheetname]:
        lev = levenshtein_ratio(value.upper(), fixed_header.upper())
        if(lev > distance):
            distance = lev
            header = fixed_header
    return header

def get_headers(current_headers, sheetname):
    output_headers = []
    for header in current_headers:
        output_headers.append(get_header(header, sheetname))
    return output_headers
    
if __name__=="__main__":
    try:
        filename = sys.argv[1]
        filepath = "./{0}.xlsx".format(filename)
        f = pd.ExcelFile(filepath)
        excel = pd.read_excel(f, sheet_name=None)
        sheet_names = f.sheet_names
        area_sheet_name = sheet_names[0]
        apartment_sheet_name = sheet_names[1]
        resident_sheet_name = sheet_names[2]
        area_sheet = excel[area_sheet_name]
        apartment_sheet = excel[apartment_sheet_name]
        resident_sheet = excel[resident_sheet_name]

        area_sheet.columns = get_headers(area_sheet.columns.tolist(), area_sheet_name)
        apartment_sheet.columns = get_headers(apartment_sheet.columns.tolist(), apartment_sheet_name)
        resident_sheet.columns = get_headers(resident_sheet.columns.tolist(), resident_sheet_name)
        area = area_sheet[area_sheet['Type'].str.contains('BLOCK')]
        blocks = area["Area"].tolist()
        apartments = apartment_sheet[apartment_sheet.Block.isin(blocks)]
        errors = []
        # apartments_to_skip = apartment_sheet[~apartment_sheet.Block.isin(blocks)]["Block"].tolist()
        for index, row in apartments.iterrows():
            flats = row["Flat"].split(",")
            intercoms = row["Intercom"].split(",")
            if len(flats) is not len(intercoms):
                errors.append({"Type": "ApartmentSheetError", "Message": "Flat - Intercom length does not match at row {0}.".format(index+2)})
            for flat in flats:
                if flat.isnumeric():
                    resident_record = resident_sheet[(resident_sheet.Block == row["Block"]) & (resident_sheet.Flat == int(flat))]
                    if not resident_record.Flat.count():
                        errors.append({"Type": "FlatOwnerError", "Message": "Flat not found for block - {0} and flat - {1}.".format(row["Block"], flat)})
        if errors:
            for index in range(len(errors)): 
                print("{0}  -   {1} : {2}".format((index+1), errors[index]["Type"], errors[index]["Message"]))
        else:
            area_sheet["Cluster"] = "NULL"
            resident_sheet["Accomodation Type"] = "VACANT"
            resident_sheet["BHK"] = "NA"
            out_file = filename + " - " + area_sheet_name + ".csv"
            area_sheet.to_csv(out_file, index=False)
            print("File {0} created successfully".format(out_file))
            out_file = filename + " - " + apartment_sheet_name + ".csv"
            apartment_sheet.to_csv(out_file, index=False)
            print("File {0} created successfully".format(out_file))
            out_file = filename + " - " + resident_sheet_name + ".csv"
            resident_sheet.to_csv(out_file, index=False)
            print("File {0} created successfully".format(out_file))
    except IndexError:
        print("Filename is mandatory")
    except FileNotFoundError as error:
        print("File {0}.xlsx does not exist".format(filename))
