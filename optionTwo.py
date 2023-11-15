import pandas as pd
import plotly.express as px
from flask import Flask, send_file

# Load your transaction data into a Pandas DataFrame
data = pd.read_csv('transactions_within_saccos.csv')

# Convert the 'date_time' column to datetime format
data['transactionDate'] = pd.to_datetime(data['transactionDate'])

# Extract hour from the 'date_time' column
data['hour'] = data['transactionDate'].dt.hour

# Group data by 'sacco' and 'hour', and count the number of transactions
grouped_data = data.groupby(['Sacco', 'hour'])['transactionAmount'].count().reset_index()

# Create an interactive line chart using Plotly Express
fig = px.line(grouped_data, x='hour', y='transactionAmount', color='Sacco',
              labels={'transactionAmount': 'Number of Transactions', 'hour': 'Hour of the Day'},
              title='Number of Transactions by Hour for Different Financial Groups',
              hover_data=['Sacco'])

# Create a Flask web application
app = Flask(__name__)

# Route for the plot
@app.route('/')
def plot():
    return fig.to_html(full_html=False)

# Route for downloading the data as an Excel file
@app.route('/download_excel')
def download_excel():
    # Save the grouped data as an Excel file
    grouped_data.to_excel('transactions_data.xlsx', index=False)
    # Send the Excel file as a download attachment
    return send_file('transactions_data.xlsx', as_attachment=True, after_write=True)

if __name__ == '__main__':
    app.run(debug=True)
