import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px

data = pd.read_csv('./transactions_within_saccos.csv')

print(data.head())

print("CSV file loaded successfully")

data['transactionDate'] = pd.to_datetime(data['transactionDate'])

data['hour'] = data['transactionDate'].dt.hour

grouped_data = data.groupby(['Sacco', 'hour'])['transactionAmount'].count().reset_index()

# for group, group_data in grouped_data.groupby('Sacco'):
#     plt.plot(group_data['hour'], group_data['transactionAmount'], label=f'Sacco {group}')
# Create an interactive line chart using Plotly Express
# fig = px.line(grouped_data, x='hour', y='transactionAmount', color='Sacco',
#               labels={'transactionAmount': 'Number of Transactions', 'hour': 'Hour of the Day'},
#               title='Number of Transactions by Hour for Different Financial Groups',
#               hover_data=['Sacco'])
    
# Create an Excel writer
excel_writer = pd.ExcelWriter('transactions_data.xlsx', engine='openpyxl')

# Write the data to the Excel file
grouped_data.to_excel(excel_writer, index=False)

# Close the Excel writer
excel_writer.close()

# Create an interactive line chart using Plotly Express
fig = px.line(grouped_data, x='hour', y='transactionAmount', color='Sacco',
              labels={'transactionAmount': 'Number of Transactions', 'hour': 'Hour of the Day'},
              title='Number of Transactions by Hour for Different Financial Groups',
              hover_data=['Sacco'])

# Add a download button for Excel format
fig.update_layout(updatemenus=[dict(type='buttons', showactive=False, buttons=[
    dict(label='Download Excel', method='download', args=[{
        'filename': 'transactions_data.xlsx',  # Specify the correct file name with extension
        'title': 'Transaction Data',
        'label': 'Download Excel',
        'args': [{'visible': [True, False, False]}],
    }])
])])



# Customize the plot
plt.xlabel('Hour of the Day')
plt.ylabel('Total Transaction Amount')
plt.title('Transaction Activity by Hour')
plt.legend()
plt.grid(True)
plt.xticks(range(24))  # Assuming transactions cover 24 hours
plt.tight_layout()

print("Plotting completed")

# Display the plot
#plt.show()
fig.show()