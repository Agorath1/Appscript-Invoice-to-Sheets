import fitz  # PyMuPDF
import pandas as pd
import re

# Open the PDF file
pdf_path = "document.pdf"
pdf_document = fitz.open(pdf_path)

# Initialize a list to store extracted data
columns = ["STORE", "DELV DT", "INVOICE#", "PAGE", "DEPT", "QTY ORD", "QTY SHP", "ITEM #", "DESCRIPTION",
           "PRODUCT UPC", "AWGSELL", "FREIGHT", "TOTAL ALLOW", "NET COST", "PACK", "UNT COST", "EXT NT COST",
           "TOTAL WEIGHT", "PB"]
data = []

line = []
for i in range(len(columns)):
    line.append("")

# Extract text from each page and process
for page_num in range(len(pdf_document)):
    page = pdf_document.load_page(page_num)
    text = page.get_text("text")
    text = re.sub(r'\s+', ' ', text).split()

    counter = -1
    while counter < len(text) - 1:
        counter += 1

        if text[counter] == "STORE":
            line[columns.index("STORE")] = int(text[counter + 1])
            counter += 2
        if text[counter] == "DEPT":
            line[columns.index("DEPT")] = int(text[counter + 1][0:2])
            counter += 2
        if text[counter] == "DT:":
            line[columns.index("DELV DT")] = text[counter + 1]
            counter += 2
        if text[counter] == "INVOICE#":
            line[columns.index("INVOICE#")] = int(text[counter + 1])
            counter += 2
        if text[counter] == "PAGE" and (not text[counter + 1].isalpha()):
            line[columns.index("PAGE")] = int(text[counter + 1])
            counter += 2



        try:
            if (text[counter].isalnum() and
                    text[counter + 1].isalnum() and
                    text[counter + 2][1] == "-" and
                    text[counter + 2][0].isalnum()):

                for i in range(5, len(line)):
                    line[i] = ""

                if text[counter - 1] == "PB":
                    line[columns.index("PB")] = text[counter - 1]

                # Quantity Order
                line[columns.index("QTY ORD")] = int(text[counter])
                # Quantity Shipped
                line[columns.index("QTY SHP")] = int(text[counter + 1])
                # Item Code
                line[columns.index("ITEM #")] = int(text[counter + 2].split("-")[1])

                if line[columns.index("ITEM #")] == 350334:
                    counter = counter

                counter += 3
                i = 0
                try:
                    while text[counter + i][0] != "0" or len(text[counter + i]) != 15:
                        i += 1
                except IndexError:
                    pass
                line[columns.index("DESCRIPTION")] = " ".join(text[counter:(counter + i)])
                counter += i

                line[columns.index("PRODUCT UPC")] = int(text[counter])
                counter += 1

                # Checks for not shipped.
                i = 0
                while True:
                    try:
                        float(text[counter + i])
                        break
                    except ValueError:
                        i += 1
                if i != 0:

                    line[columns.index("AWGSELL")] = " ".join(text[counter:counter + i])
                    counter += i

                else:

                    line[columns.index("AWGSELL")] = float(text[counter])
                    counter += 1

                    line[columns.index("TOTAL ALLOW")] = float(text[counter])
                    counter += 1

                    line[columns.index("NET COST")] = float(text[counter])
                    counter += 2

                    try:
                        if str(int(text[counter])) == text[counter]:
                            line[columns.index("PACK")] = int(text[counter])
                            counter += 1
                    except ValueError:
                        pass

                    try:
                        float(text[counter + 1])
                        line[columns.index("UNT COST")] = float(text[counter])
                        counter += 1
                    except ValueError:
                        pass

                    line[columns.index("EXT NT COST")] = float(text[counter])
                    counter += 1

                    if text[counter] == "PB":
                        line[columns.index("PB")] = text[counter]
                        counter += 1

                    while True:
                        try:
                            float(text[counter])
                            line[columns.index("FREIGHT")] = float(text[counter])
                            break
                        except ValueError:
                            counter += 1
                    counter += 1

                    if "WEIGHT:" in text[counter:counter + 20]:
                        while text[counter] != "WEIGHT:":
                            counter += 1
                        line[columns.index("TOTAL WEIGHT")] = float(text[counter + 1].replace(",", ""))
                        counter += 2

                counter -= 1
                data.append(list(line))
        except IndexError:
            pass

# Convert list to DataFrame
df = pd.DataFrame(data, columns=columns)

# Save DataFrame to Excel file
excel_path = "extracted_data.xlsx"
df.to_excel(excel_path, index=False)
