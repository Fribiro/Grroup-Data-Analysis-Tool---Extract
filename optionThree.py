import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from flask import Flask, render_template_string

# Load your transaction data into a Pandas DataFrame
data = pd.read_csv('transactions_within_saccos_2024-01-24.csv')

# Convert the 'date_time' column to datetime format
data['transactionDate'] = pd.to_datetime(data['transactionDate'])

# Extract hour from the 'date_time' column
data['hour'] = data['transactionDate'].dt.hour

# Group data by 'Sacco' ie 'tenantName' and 'hour', and count the number of transactions
grouped_data = data.groupby(['tenantName', 'hour'])['transactionAmount'].count().reset_index()

# Create an interactive line chart using Plotly Express
fig = px.line(grouped_data, x='hour', y='transactionAmount', color='tenantName',
              labels={'transactionAmount': 'Number of Transactions', 'hour': 'Hour of the Day'},
              title='Number of Transactions by Hour for Different Financial Groups',
              hover_data=['tenantName'])

# Add a download button for Excel format
fig.add_trace(go.Scatter(
    x=[0],
    y=[0.5],
    mode='text',
    text=['Download Excel'],
    showlegend=False,
    textposition='top right',
    textfont=dict(color='blue'),
    hoverinfo='none',
    xaxis='x1',
    yaxis='y1',
))

# Configure the button to trigger a download
# fig.update_layout(
#     updatemenus=[
#         dict(
#             type='buttons',  # or 'dropdown'
#             x=1.05,
#             y=0.9,
#             buttons=[dict(label='Download Excel', method='relayout', args=['./transactions_within_saccos.csv', 'download_trans_analysis.xlsx'])]
#         )
#     ]
# )
# Create buttons for each Sacco to toggle visibility
button_list = []
sacco_names = grouped_data['tenantName'].unique()

buttons_dropdown = [dict(label='Toggle All', method='update', args=[{'visible': [True] * len(sacco_names)}])]
buttons_dropdown += [dict(label=selected_sacco, method='update', args=[{'visible': [True if sacco == selected_sacco else 'legendonly' for sacco in sacco_names]}]) for selected_sacco in sacco_names]

# Add buttons_dropdown to updatemenus list
fig.update_layout(updatemenus=[
    dict(type='dropdown', direction='down', showactive=True, x=1.0, y=1.07, buttons=buttons_dropdown),
])
# fig.show()

# # Save the data to an Excel file
# grouped_data.to_excel('download.xlsx', index=False)
# Flask App
app = Flask(__name__)

@app.route('/')
def plot():
    return render_template_string(fig.to_html())

if __name__ == '__main__':
    app.run(port=6600)  # Change the port number as needed
