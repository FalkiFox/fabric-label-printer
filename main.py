from jinja2 import Template
import pandas as pd
import qrcode
import platform
import os

# Script parameters
FILE_NAME = "Y:\Hobbies\Furry\Fursuit_Making\Fur_Database.xlsx"
SHEET_NAME = "Fabric Database"
TEMPLATE_STR = """
<html>
  <body>
    {% for row in rows %}
    <div style="margin-bottom: 50px">
        <table border="1">
            <tr>
              <th>QR Code</th>
              <th style="width: 100px">Type / Color</th>
              <th style="width: 150px">Name</th>
            </tr>
            <tr>
                <td rowspan="2"><img src="{{ row.index }}.png" width="100" height="100"></td>
                <td>{{ row.Type }}</td>
                <td rowspan="2">{{ row.Name }}</td>
            </tr>
            <tr>
                <td>{{ row.Color }}</td>
            </tr>
        </table>
    </div>
    {% endfor %}
  </body>
  <style>
    table, th, td {
      border: 1px solid black;
      border-collapse: collapse;
      text-align: center
    }
  </style>
</html>
"""

def clear():
    """Clears the console for windows or linux systems"""
    currentPlatform = platform.system()
    command = "cls" if currentPlatform == "Windows" else "clear"
    os.system(command)


def getUserInput(options, selection):
    """Shows user selection and returns the users input"""
    clear()

    print("=================================================== Current Stock ===================================================")
    print(f"{options}")
    print("================================================= Current Selection =================================================")
    print(f"{selection}")
    print("=====================================================================================================================")

    userInput = input("Please enter the index of the fabric you want to add to your selection. (Type 'Done' to finish selecting): ")
    return userInput


if __name__ == "__main__":
    # Get excel file as DataFrame object
    try:
        xl_file = pd.ExcelFile(FILE_NAME)
        fabrics = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME, index_col=0)
        fabrics = fabrics.reset_index(drop=True)  # Reset index column
    except (FileNotFoundError, ValueError) as error:
        print(error)
        quit()

    isUserAdding = True
    addedFabrics = pd.DataFrame()

    while isUserAdding:
        userInput = getUserInput(fabrics, addedFabrics)

        if userInput == "Done":
            isUserAdding = False
        else:
            try:
                # Drop the row from the original DataFrame and re-add the popped row to the second DataFrame
                userInput = int(userInput)
                poppedRow = fabrics.iloc[[userInput]]
                fabrics = fabrics.drop(index=userInput).reset_index(drop=True)
                addedFabrics = pd.concat([addedFabrics, poppedRow], ignore_index=True)
            except (IndexError, ValueError):  
                print("Please enter a valid index!")

    # Generate QR codes
    for index, row in addedFabrics.iterrows():
        img = qrcode.make(row['Link'])
        img.save(f"{index}.png")

    # Create a Jinja2 template object
    template = Template(TEMPLATE_STR)

    # Convert the DataFrame rows to dictionaries so that we can iterate through them
    rows = addedFabrics.to_dict(orient='records')
    for i, row in enumerate(rows):
        row['index'] = i

    # Render the template with the rows from the DataFrame
    rendered_html = template.render(rows=rows)

    # Save the rendered HTML to a file
    with open("output.html", "w") as file:
        file.write(rendered_html)

    print("Generated HTML file!")