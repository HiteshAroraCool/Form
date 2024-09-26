from flask import Flask, render_template, request
import pandas as pd
import openpyxl
import os

app = Flask(__name__)

excel_file = 'data.xlsx'

@app.route('/', methods=['GET'])
def Form():
    return render_template('index.html')

@app.route("/form", methods=['POST'])
def receive_data():
    if request.method == 'POST':
        try:
            # Data Received form fields
            name = str(request.form['Name'])
            age = int(request.form['Age'])
            gender = str(request.form['Gender'])
            interest = str(request.form['Interest'])
            data = {
                'Name': [name],
                'Age': [age],
                'Gender': [gender],
                'Interest': [interest]
            }

            df_new = pd.DataFrame(data)

            # Check if the file exists
            if os.path.exists(excel_file):
                # Read the existing file
                df_existing = pd.read_excel(excel_file)

                # Append the new data
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)

                # Write the updated data back to the Excel file
                df_combined.to_excel(excel_file, index=False)
            else:
                # Write the new data if the file does not exist
                df_new.to_excel(excel_file, index=False)
                
            # new page to show received information
            return render_template('form-submitted.html', name=name, age=age, gender=gender, interest=interest)
                
        except Exception as e:
            return render_template('index.html', text=f"Error Occurred: {str(e)}")
    else:
        return render_template('index.html')
    
if __name__ == "__main__":
    app.run(debug=True)