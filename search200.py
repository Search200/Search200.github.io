from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__, template_folder='.')

#Define the DataFrame globally
df = None

# Read the Excel workbook into a pandas DataFrame
excel_file_path = './documents/Current200s.xlsx'
try:
    df = pd.read_excel(excel_file_path)
except Exception as e:
    print(f"Error reading Excel file: {e}")


@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')


@app.route('/search', methods=['POST'])
def search():
    query = request.form.get('search_query')

    print(f"Query: {query}")
    print(f"DataFrame Values: {df.iloc[:, 0].astype(str).values}")

    # Check if the query is in the DataFrame index
    if not pd.isna(df.iloc[:, 0]).any() and query in df.iloc[:, 0].astype(str).values:
        # Find the index of the query in the first column
        index_of_query = df.index[df.iloc[:, 0].astype(str) == query].tolist()[0]

        # Get the value in the adjacent cell (next column)
        adjacent_value = df.iloc[index_of_query, 1]

        return f"   Your ID number was found! Provide this number: {adjacent_value} to the authorities at the University of Ghana Business School Academic Affairs Office for your ID card. Thank you!."
    else:
        return "   Sorry, Your ID number was not found in the database. This means that your ID card is not available, Kindly see the Academic Affairs Office at Business School for more info. Thank you!"


if __name__ == '__main__':
    app.run(debug=True)