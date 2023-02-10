from flask import Flask, request, render_template, make_response,send_file
import pandas as pd
import json
import io
import os
from pandas import json_normalize

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        file = request.files['file']
        df = pd.read_excel(file)
        dataframes = []
        for i in range(df.shape[0]):
            json_string = df.iloc[i, 0] 
            json_df = json_normalize(json.loads(json_string))
            dataframes.append(json_df)
        result = pd.concat(dataframes, ignore_index=True)

        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        result.to_excel(r'converted_file.xlsx', index=False)

        writer.save()
        output.name = "converted_file.xlsx"
        output.seek(0)
        
        response = make_response(output.getvalue())
        response.headers["Content-Disposition"] = "attachment; filename=converted_file.xlsx"
        response.mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return render_template("upload.html", download_message=True, filename='converted_file.xlsx')


    else:
        return render_template('upload.html')
print(os.path.abspath('converted_file.xlsx'))
@app.route('/download_file')
def download_file():
    filename = 'converted_file.xlsx'
    file_path = os.path.join('/home/mobm93', filename)

    response = make_response(send_file(file_path))
    response.headers["Content-Disposition"] = "attachment; filename=converted_file.xlsx"
    response.mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return response


if __name__ == '__main__':
    app.run(debug=True)
