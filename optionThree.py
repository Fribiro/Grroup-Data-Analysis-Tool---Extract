import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Load your transaction data into a Pandas DataFrame
data = pd.read_csv('transactions_within_saccos.csv')

# Convert the 'date_time' column to datetime format
data['transactionDate'] = pd.to_datetime(data['transactionDate'])

# Extract hour from the 'date_time' column
data['hour'] = data['transactionDate'].dt.hour

# Group data by 'Sacco' and 'hour', and count the number of transactions
grouped_data = data.groupby(['Sacco', 'hour'])['transactionAmount'].count().reset_index()

# Create an interactive line chart using Plotly Express
fig = px.line(grouped_data, x='hour', y='transactionAmount', color='Sacco',
              labels={'transactionAmount': 'Number of Transactions', 'hour': 'Hour of the Day'},
              title='Number of Transactions by Hour for Different Financial Groups',
              hover_data=['Sacco'])

# Add a download button for Excel format
fig.add_trace(go.Scatter(
    x=[0],
    y=[0],
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
fig.update_layout(
    updatemenus=[
        dict(
            type='buttons',  # or 'dropdown'
            x=1.05,
            y=0.9,
            buttons=[dict(label='Download Excel', method='relayout', args=['template', 'download.xlsx'])]
        )
    ]
)

# Show the interactive plot
fig.show()

# Save the data to an Excel file
grouped_data.to_excel('download.xlsx', index=False)
