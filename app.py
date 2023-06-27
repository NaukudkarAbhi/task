from flask import Flask, render_template, request, send_file
import os
import pandas as pd
import sqlite3

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = 'uploads'
DATABASE = 'database.db'

# Create uploads folder if it doesn't exist
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


# Route to the home page
@app.route('/')
def index():
    return render_template('index.html')


# Route to handle the file upload
@app.route('/upload', methods=['POST'])
def upload():
    # Get the uploaded file
    file = request.files['file']
    
    # Save the file to the uploads folder
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)
    
    # Insert data from the Excel file into the database
    insert_data_into_database(file_path)
    
    return 'File uploaded successfully!'


# Function to insert data from the Excel file into the database
def insert_data_into_database(file_path):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path)
    
    # Connect to the database
    conn = sqlite3.connect(DATABASE)
    
    # Insert the DataFrame into the database table
    df.to_sql('orders', conn, if_exists='append', index=False)
    
    # Close the database connection
    conn.close()


# Route to handle the file download
@app.route('/download')
def download():
    # Connect to the database
    conn = sqlite3.connect(DATABASE)
    
    # Read data from the database table into a pandas DataFrame
    df = pd.read_sql_query('SELECT * FROM orders', conn)
    
    # Create a new Excel file
    output_file = 'database_details.xlsx'
    df.to_excel(output_file, index=False)
    
    # Close the database connection
    conn.close()
    
    # Send the file for download
    return send_file(output_file, as_attachment=True)


if __name__ == '__main__':
    app.run()


