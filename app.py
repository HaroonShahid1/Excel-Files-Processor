from flask import Flask, render_template, request, send_file, session
from flask_session import Session
import pandas as pd
from io import BytesIO
import os

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Change this to a random secret key
app.config['SESSION_TYPE'] = 'filesystem'
Session(app)

def compare_quantities(df1, df2):
    if 'Quantity' not in df1 or 'Quantity' not in df2 or 'Pcode' not in df1 or 'Pcode' not in df2:
        return pd.DataFrame(), "Both files must have 'Quantity' and 'Pcode' columns."

    merged_df = pd.merge(df1, df2, on='Pcode', suffixes=('_old_software', '_new_software'))
    result_df = merged_df[merged_df['Quantity_old_software'] != merged_df['Quantity_new_software']]
    return result_df, ""

@app.route('/third', methods=['GET', 'POST'])
def third_page():
    if request.method == 'POST':
        if '_old_software' in request.files and '_new_software' in request.files:
            file1 = request.files['_old_software']
            file2 = request.files['_old_software']
            
            if file1.filename != '' and file2.filename != '':
                df1 = pd.read_excel(file1, engine='openpyxl')
                df2 = pd.read_excel(file2, engine='openpyxl')
                
                session['df1'] = df1.to_json()  # Store DataFrame as JSON string in session
                session['df2'] = df2.to_json()

        if 'compare' in request.form:
            if 'df1' in session and 'df2' in session:
                df1 = pd.read_json(session['df1'])
                df2 = pd.read_json(session['df2'])
                
                result_df, error_message = compare_quantities(df1, df2)
                
                if error_message:
                    return render_template('third_page.html', message=error_message)

                session['result_df'] = result_df.to_json()
                result_data_html = result_df.to_html(classes='data', header="true", index=False)
                return render_template('third_page.html', data1=df1.to_html(), data2=df2.to_html(), result_data=result_data_html)
            else:
                return render_template('third_page.html', message='Please upload both files.')

        if 'save_file' in request.form:
            if 'result_df' in session:
                result_df = pd.read_json(session['result_df'])
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)
                output.seek(0)
                return send_file(output, as_attachment=True, download_name='comparison_result.xlsx')
            else:
                return render_template('third_page.html', message='No comparison result to save.')

    return render_template('third_page.html', data1=None, data2=None, result_data=None)


def process_second_dataframe(df, action):
    if action == 'rename_columns':
        rename_map = {
            'Cost': 'Cost Price',
            'Total Cost': 'Amount',
            'SKU': 'Pcode'
        }
        df.rename(columns=rename_map, inplace=True)
    return df

@app.route('/second', methods=['GET', 'POST'])
def second_page():
    if request.method == 'POST':
        action = request.form['action']

        if 'second_file' in request.files:
            file = request.files['second_file']
            if file.filename != '':
                df = pd.read_excel(file, engine='openpyxl')
                session['second_dataframe'] = df.to_json()

        if 'second_dataframe' in session:
            df = pd.read_json(session['second_dataframe'])

            if action != 'save_file':
                df = process_second_dataframe(df, action)
                session['second_dataframe'] = df.to_json()
                processed_data_html = df.to_html(classes='data', header="true", index=False)
                return render_template('second_page.html', data=processed_data_html)

            # Save and send file
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            output.seek(0)
            return send_file(output, as_attachment=True, download_name='modified_new_software.xlsx')

        return render_template('second_page.html', message='No file uploaded or selected')

    return render_template('second_page.html', data=None)


def process_dataframe(df, action):
    if action == 'remove_nulls':
        df.dropna(inplace=True)
    elif action == 'combine_columns':
        df['Product Name'] = df['Item Name'] + ' ' + df['Potency'] + ' ' + df['Pack Size']
    elif action == 'drop_columns':
        df.drop(['Item Name', 'Potency', 'Pack Size'], axis=1, inplace=True)
    return df

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        action = request.form['action']

        if 'file' in request.files:
            file = request.files['file']
            if file.filename != '':
                df = pd.read_excel(file, engine='openpyxl')
                session['dataframe'] = df.to_json()  # Store DataFrame in session

        if 'dataframe' in session:
            df = pd.read_json(session['dataframe'])

            if action != 'save_file':
                df = process_dataframe(df, action)
                session['dataframe'] = df.to_json()  # Update DataFrame in session
                processed_data_html = df.to_html(classes='data', header="true", index=False)
                return render_template('index.html', data=processed_data_html)

             # Save and send file
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            output.seek(0)
            return send_file(output, as_attachment=True, download_name='modified_old_software.xlsx')


        return render_template('index.html', message='No file uploaded or selected')

    return render_template('index.html', data=None)

    

if __name__ == '__main__':
    app.run(debug=True)
