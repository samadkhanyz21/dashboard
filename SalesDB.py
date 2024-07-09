import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
import plotly.graph_objects as go
import warnings

warnings.filterwarnings('ignore')

#####################################################################
#                        Feature Engineering
#####################################################################

# Specify the path to the Excel file
file_path = "C:/Users/samad/Downloads/TEST TRACKER.xlsx"

# Load the specific sheet into a DataFrame
test_sheet_df = pd.read_excel(
    file_path,
    sheet_name='Test Sheet',
    engine='openpyxl'
)

# Data preprocessing
test_sheet_df = test_sheet_df.rename(
    columns={'Lead Channel ': 'Lead Channel',
             'Date Lead Received ': 'Date Lead Received'}
)

imp_df = test_sheet_df[[
    'Depot', 'Month', 'Lead Source', 'Brand Source',
    'Lead Category', 'Customer Name', 'Device Type',
    'Email', 'Status', 'Assigned to'
]].copy()

imp_df['Status'] = imp_df['Status'].fillna('Not Quoted')
imp_df['Email'] = imp_df['Email'].fillna('noEmail@example.com')
imp_df['Device Type'] = imp_df['Device Type'].fillna('No Specified Device')
imp_df.dropna(subset='Customer Name', inplace=True)
imp_df['Sales Rep'] = imp_df['Assigned to']
imp_df.drop(['Assigned to'], axis=1, inplace=True)

for col in imp_df.columns:
    imp_df[col] = imp_df[col].astype('category')

#####################################################################
#                     Streamlit App Page Setup
#####################################################################

# Set page config
st.set_page_config(page_title="Sales Dashboard", page_icon=":bar_chart:", layout="wide")

main_logo = 'images/logo.png'
st.logo(main_logo, link='https://oneworldrental.com/')

# Title
st.title("OWR Sales Dashboard")

# Set background color to white
st.markdown(
    """
    <style>
    body {
        background-color: #FFFFFF;
    }
    </style>
    """,
    unsafe_allow_html=True
)
#####################################################################
#                              Graphs
#####################################################################


# Filter categories below 3% in Pie and Donut charts
def filter_below_3_percent(df, column):
    counts = df[column].value_counts(normalize=True)
    filtered_values = counts[counts >= 0.03].index
    return df[df[column].isin(filtered_values)]


# Plotly Donut Chart
filtered_lead_source_df = filter_below_3_percent(imp_df, 'Lead Source')
fig_donut = px.pie(filtered_lead_source_df, names='Lead Source', hole=0.3, title='Lead Source Distribution')
fig_donut.update_traces(textposition='inside', textinfo='percent+label')

# Plotly Pie Chart
filtered_brand_source_df = filter_below_3_percent(imp_df, 'Brand Source')
fig_pie = px.pie(filtered_brand_source_df, names='Brand Source', title='Brand Source Distribution')

# Plotly Bar Chart
fig_bar = px.bar(imp_df, x='Depot', color='Lead Category', title='Lead Category by Depot')

# Filter the DataFrame for January, February, and March only
filtered_df = imp_df[imp_df['Month'].isin(['January', 'February', 'March'])]

# Plotly Histogram for January, February, and March
fig_hist = px.histogram(filtered_df, x='Month', color='Status', title='Monthly Lead Status Distribution')

# Plotly Stacked Bar Chart
fig_stacked_bar = px.bar(imp_df, x='Depot', y='Lead Source', color='Brand Source', title='Stacked Lead Source by Depot')

# Plotly Sales Rep Pie Chart
filtered_sales_rep_df = filter_below_3_percent(imp_df, 'Sales Rep')
fig_sales_rep_pie = px.pie(filtered_sales_rep_df, names='Sales Rep', title='Sales Representatives Distribution for Booking Quotes')

# Adding the Sales Rep Pie Chart
st.plotly_chart(fig_sales_rep_pie, use_container_width=True)

# Streamlit Layout
col1, col2 = st.columns(2)
with col1:
    st.plotly_chart(fig_donut, use_container_width=True)
    st.plotly_chart(fig_bar, use_container_width=True)

with col2:
    st.plotly_chart(fig_pie, use_container_width=True)
    st.plotly_chart(fig_hist, use_container_width=True)

# Single column for stacked bar chart
st.plotly_chart(fig_stacked_bar, use_container_width=True)