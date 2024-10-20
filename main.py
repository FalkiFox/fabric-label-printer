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
              <th style="width: 100px">Type</th>
              <th style="width: 200px">Name</th>
              <th style="width: 150px">Color</th>
            </tr>
            <tr>
                <td><img src="{{ row.index }}.png" width="100" height="100"></td>
                <td>{{ row.Type }}</td>
                <td>{{ row.Name }}</td>
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
    currentPlatform = platform.system()
    command = "cls" if currentPlatform == "Windows" else "clear"
    os.system(command)


if __name__ == "__main__":
    # Get excel file as dataframe
    try:
        xl_file = pd.ExcelFile(FILE_NAME)
        fabrics = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME, index_col=0)
        fabrics = fabrics.reset_index(drop=True)  # Reset index column
    except (FileNotFoundError, ValueError) as e:
        print(e)
        quit()

    isUserAdding = True
    addedFabrics = pd.DataFrame()

    while isUserAdding:
        clear()
        print("=================================================== Current Stock ===================================================")
        print(fabrics)
        print("================================================= Current Selection =================================================")
        print(addedFabrics)
        print("=====================================================================================================================")
        index = input("Please enter the index of the fabric you want to add to your selection. (Type 'Done' to finish selecting): ")

        if index == "Done":
            isUserAdding = False    
        else:
            try:
                # Drop the row from the original DataFrame and re-add the popped row to the second DataFrame
                index = int(index)

                poppedRow = fabrics.iloc[[index]]
                fabrics = fabrics.drop(index=index).reset_index(drop=True)
                addedFabrics = pd.concat([addedFabrics, poppedRow], ignore_index=True)
            except (IndexError, ValueError):  
                print("Please enter a valid index!")

    # Generate the QR codes
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