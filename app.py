import os
import pandas as pd
from flask import Flask, request, render_template, redirect, url_for, send_from_directory
from werkzeug.utils import secure_filename
from zipfile import ZipFile


ALLOWED_EXTENSIONS = set(['csv', 'xls', 'xlsx', 'zip'])

# checks filename is valid and if if it's an allowed extension
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# extract zip file and convert excel files to csv
def process_zip(zipfile, archive_name):
    with ZipFile(zipfile, 'r') as zip: 
    
        # extract all the files to ./unpacked 
        zip.extractall('unpacked') 

        # now that our zip is extracted, we need to convert the excel files to csv
        convert_to_csv(archive_name)

# converts folder of excel files to csv
def convert_to_csv(archive_name):
    # get path of unpacked folder
    archive_path = f'unpacked/{archive_name}'

    # loop through files in unpacked folder
    for file in os.scandir(archive_path):

        # if file is an excel file
        if file.path.endswith(('.xlsx', '.xls')):
            # convert excel file to csv
            excel_file = pd.read_excel(file.path)
            excel_file.to_csv(f'csv/{file.name}.csv', index = None, header=True)

        # delete file
        os.remove(file.path)
    
    os.rmdir(archive_path) 

def merge_csvs():
    # get path of csv folder
    csv_path = 'csv'

    # create empty dataframe
    df = pd.DataFrame()

    # loop through files in csv folder
    for file in os.scandir(csv_path):
        # if file is a csv file
        if file.path.endswith('.csv'):
            # read csv file
            csv_file = pd.read_csv(file.path)
            # append csv file to dataframe
            df = pd.concat([df, csv_file], ignore_index=True)

        # delete file
        os.remove(file.path)

    # convert dataframe to csv
    df.to_csv('output/merged.csv', index = None, header=True)

def convert_csv_to_excel():
    # get path of csv folder
    csv_path = 'output/merged.csv'
    # read csv file
    csv_file = pd.read_csv(csv_path)
    # convert csv file to excel
    csv_file.to_excel('output/output.xlsx', index = None, header=True)
    # delete csv file 
    os.remove(csv_path)

def create_app():
    app = Flask(__name__)

    @app.route('/', methods=['GET', 'POST'])
    def index():
        if request.method == 'POST':
            try:
                # get file from html form
                file = request.files['file']

                # if file exists and its an allowed file extension
                if file and allowed_file(file.filename):
                    # create secure filename to prevent malicious files
                    filename = secure_filename(file.filename)

                    # assign save path to input folder
                    save_location = os.path.join('input', filename)

                    # save file to input folder
                    file.save(save_location)

                    folder_name = file.filename.rsplit( ".", 1 )[ 0 ]

                    # takes in zip file and creates folder of csv files
                    process_zip(save_location, folder_name)

                    # delete zip file
                    os.remove(save_location)

                    merge_csvs()

                    convert_csv_to_excel()
                    
                    return redirect(url_for('download'))
            except:
                return 'something went wrong :('

        return render_template('index.html')

    @app.route('/download')
    def download():
        return render_template('download.html', files=os.listdir('output'))

    @app.route('/download/<filename>')
    def download_file(filename):
        return send_from_directory('output', filename)

    return app