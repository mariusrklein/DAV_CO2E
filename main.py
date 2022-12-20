import io
from flask import Flask, abort, render_template, request, send_file, session
from PyPDF2 import PdfFileReader
import pandas as pd

app = Flask(__name__)
app.secret_key = '''eb ef 28 c1 fd 42 00 3b a8 7e 91 d9 eb 0f 19 88 
c3 7a 33 04 6f a7 0c 9c db 52 7a 3f bf 66 2f 3b 
42 82 48 93 98 3e 22 5b ca 91 7b d9 5c e1 b6 d6 
aa 87 92 53 f7 46 f9 69 7c 72 4f 2c 65 31 22 ab 
45 76 2c 43 68 d3 93 d4 60 c1 6a ab e7 b3 af 1e 
58 7c 78 7d 07 92 64 4c a4 f7 08 16 4f 2f 22 53 
2b 4e c0 08'''

@app.route('/')
def main():
    return render_template("index.html")
  
  
@app.route('/upload', methods=['POST'])
def upload():
    if request.method == 'POST':
  
        res_dict = {}
        # Get the list of files from webpage
        files = request.files.getlist("file")

        out_string = '''<h1>Files Uploaded Successfully.!</h1>
        <a href="/" id="zurueck" class="btn btn-outline-info">Zurück</a>
        <a href="get_csv" id="download" class="btn btn-outline-info">Download</a>'''
# {{url_for('get_csv', filename= 'some_csv')}}
        for i, file in enumerate(files):
            pdf_reader = PdfFileReader(file)
            dictionary = pdf_reader.getFormTextFields() # returns a python dictionary
            res_dict[str(i) + "_" + file.filename] = dictionary

        res_df = pd.DataFrame.from_dict(res_dict)
        session["df"] = res_df.to_csv(index=True, header=True, sep=";")

        out_string = out_string + res_df.to_html()

        return out_string

@app.route("/get_csv", methods=['GET', 'POST'])
def get_csv():
    
    if "df" in session:
        csv = session["df"]
    else:
        abort(404)

    
    
    # Create a string buffer
    buf_str = io.StringIO(csv)

    # Create a bytes buffer from the string buffer
    buf_byt = io.BytesIO(buf_str.read().encode("utf-8"))
    
    # Return the CSV data as an attachment
    return send_file(buf_byt,
                     mimetype="text/csv",
                     as_attachment=True,
                     attachment_filename="data.csv")

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8080, debug=True)