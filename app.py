import pandas as pd
import streamlit as st
import xlsxwriter
from io import BytesIO

# Data cleaning function
def clean_data_cgk(df):

        if df['DATE'].dtype == 'O':  # 'O' is the code for object type
        # Clean the date column
            df['DATE'] = df['DATE'].astype(str)
            df['Year'] = df['DATE'].apply(lambda x: x[:4] if '00:00:00' in x else x[-4:])
            df['Month'] = df.apply(lambda row: row['DATE'][8:10] if '00:00:00' in row['DATE'] else row['DATE'][3:5], axis=1)
            df['Day'] = df.apply(lambda row: row['DATE'][5:7] if '00:00:00' in row['DATE'] else row['DATE'][:2], axis=1)
            df['Date_final'] = pd.to_datetime(df[['Year', 'Month', 'Day']], errors='coerce')
            df['Date_final'] = pd.to_datetime(df['Date_final']).dt.strftime('%Y-%m-%d')
            # Convert 'Date_final' to datetime with error handling
            df['Date_final'] = pd.to_datetime(df['Date_final'], errors='coerce')
        else:
            df['Date_final'] = df['DATE']
   
     # Add total pax
        df['PAX'] = df['F.A']+df['F.C']+df['F.I']+df['B.A']+df['B.C']+df['B.I']+df['E.A']+df['E.C']+df['E.I']

        # Reorder desired column that will shown in Excel
        column_orders = ['No', 'AIRLINES', 'FLIGHT NO', 'Date_final', 'AC REG', 'AC TYPE', 'STRETCH',
                     'F.A', 'F.I', 'F.C', 'B.A', 'B.I', 'B.C',
                    'E.A', 'E.I', 'E.C', 'PAX', 'CARGO', 'MAIL', 'ULD', 'Kgs',
                     'Pcs', 'STA', 'ATA', 'STD', 'ATD']
        df = df[column_orders]


        # Add departure/arrival label
        df['arr/dep'] = df['STA'].apply(lambda x: "Departure" if x == "**" else "Arrival")

        # Sort by arr/dep and date
        df = df.sort_values(by=['arr/dep', 'Date_final'])


        # Create function to return AC Code
        def get_ac_code(airline):
            if airline == 'QZ - AIR ASIA INDONESIA (DOMESTIC)':
                return 'QZ-DOM'
            elif airline == 'QZ - AIR ASIA INDONESIA (INTERNATIONAL)':
                return 'QZ-INT'
            elif airline == 'CX - CATHAY PACIFIC FREIGHTER':
                return 'CX-FREIGHTER'
            elif airline == '2Y - MY INDO (DOM) - PREMIER':
                return '2Y-DOM'
            elif airline == '2Y - MY INDO (INTL) - PREMIER':
                return '2Y-INT'
            else:
                return airline[:2]  # Default: Take the first two characters

        # Trigger the function
        df['AC CODE'] = df['AIRLINES'].apply(get_ac_code)
        return df

def clean_data_dps(df):
    if df['DATE'].dtype == 'O':  # 'O' is the code for object type
        # Clean the date column
            df['DATE'] = df['DATE'].astype(str)
            df['Year'] = df['DATE'].apply(lambda x: x[:4] if '00:00:00' in x else x[-4:])
            df['Month'] = df.apply(lambda row: row['DATE'][8:10] if '00:00:00' in row['DATE'] else row['DATE'][3:5], axis=1)
            df['Day'] = df.apply(lambda row: row['DATE'][5:7] if '00:00:00' in row['DATE'] else row['DATE'][:2], axis=1)
            df['Date_final'] = pd.to_datetime(df[['Year', 'Month', 'Day']], errors='coerce')
            df['Date_final'] = pd.to_datetime(df['Date_final']).dt.strftime('%Y-%m-%d')
            # Convert 'Date_final' to datetime with error handling
            df['Date_final'] = pd.to_datetime(df['Date_final'], errors='coerce')
    else:
            df['Date_final'] = pd.to_datetime(df['DATE'])
   

    # Add total pax
    df['PAX'] = df['F.A']+df['F.C']+df['F.I']+df['B.A']+df['B.C']+df['B.I']+df['E.A']+df['E.C']+df['E.I']

    # Reorder desired column that will shown in Excel
    column_orders = ['No', 'AIRLINES', 'FLIGHT NO', 'Date_final', 'AC REG', 'AC TYPE', 'STRETCH',
                     'F.A', 'F.I', 'F.C', 'B.A', 'B.I', 'B.C',
                    'E.A', 'E.I', 'E.C', 'PAX', 'CARGO', 'MAIL', 'ULD', 'Kgs',
                     'Pcs', 'STA', 'ATA', 'STD', 'ATD']
    df = df[column_orders]


    # Add departure/arrival label
    df['arr/dep'] = df['STA'].apply(lambda x: "Departure" if x == "**" else "Arrival")

    # Sort by arr/dep and date
    df = df.sort_values(by=['arr/dep', 'Date_final'])

    # Create function to get AC Code based on AC Type
    def get_ac_code(airline, type):
        print(f"Checking airline: {airline}")
        if 'CX - CATHAY PACIFIC' in airline and "A333" in type:
            print("Matched CX-WB condition")
            return "CX-WB"
        if 'CX - CATHAY PACIFIC' in airline and "A321" in type:
            print("Matched CX-NB condition")
            return "CX-NB"
        if 'SQ - SINGAPORE AIRLINES' in airline and "B787-1000" in type:
            print("Matched SQ-WB condition")
            return "SQ-WB"
        if 'SQ - SINGAPORE AIRLINES' in airline and "B738MAX" in type:
            print("MAtched SQ-NB condition")
            return "SQ-NB"
        if  airline == 'QZ - AIR ASIA INDONESIA (DOMESTIC)':
            return "QZ-DOM " 
        if airline == 'QZ - AIR ASIA INDONESIA (INTERNATIONAL)':
            return "QZ-INT"
        else:
            print("Did not match get AC CODE condition")
            return airline[:2]

    # Trigger the function to get AC Code based on AC Type
    df['AC CODE'] = [get_ac_code(x, y) for x, y in zip(df['AIRLINES'], df['AC TYPE'])]

         
    return df

def clean_data_hlp(df):
    
        #Clean date column
    if df['DATE'].dtype == 'O':  # 'O' is the code for object type
        # Clean the date column
            df['DATE'] = df['DATE'].astype(str)
            df['Year'] = df['DATE'].apply(lambda x: x[:4] if '00:00:00' in x else x[-4:])
            df['Month'] = df.apply(lambda row: row['DATE'][8:10] if '00:00:00' in row['DATE'] else row['DATE'][3:5], axis=1)
            df['Day'] = df.apply(lambda row: row['DATE'][5:7] if '00:00:00' in row['DATE'] else row['DATE'][:2], axis=1)
            df['Date_final'] = pd.to_datetime(df[['Year', 'Month', 'Day']], errors='coerce')
            df['Date_final'] = pd.to_datetime(df['Date_final']).dt.strftime('%Y-%m-%d')
            # Convert 'Date_final' to datetime with error handling
            df['Date_final'] = pd.to_datetime(df['Date_final'], errors='coerce')
    else:
            df['Date_final'] = pd.to_datetime(df['DATE'])
    
    # Add total pax
    df['PAX'] = df['F.A']+df['F.C']+df['F.I']+df['B.A']+df['B.C']+df['B.I']+df['E.A']+df['E.C']+df['E.I']

    # Reorder desired column that will shown in Excel
    column_orders = ['No', 'AIRLINES', 'FLIGHT NO', 'Date_final', 'AC REG', 'AC TYPE', 'STRETCH',
                     'F.A', 'F.I', 'F.C', 'B.A', 'B.I', 'B.C',
                    'E.A', 'E.I', 'E.C', 'PAX', 'CARGO', 'MAIL', 'ULD', 'Kgs',
                     'Pcs', 'STA', 'ATA', 'STD', 'ATD']
    df = df[column_orders]


    # Add departure/arrival label
    df['arr/dep'] = df['STA'].apply(lambda x: "Departure" if x == "**" else "Arrival")

    # Sort by arr/dep and date
    df = df.sort_values(by=['arr/dep', 'Date_final'])


    # Get AC Code
    df['AC CODE'] = df['AIRLINES'].str[:2]

    return df

def clean_data_kjt(df):
    
    #clean date column
    if df['DATE'].dtype == 'O':  # 'O' is the code for object type
        # Clean the date column
            df['DATE'] = df['DATE'].astype(str)
            df['Year'] = df['DATE'].apply(lambda x: x[:4] if '00:00:00' in x else x[-4:])
            df['Month'] = df.apply(lambda row: row['DATE'][8:10] if '00:00:00' in row['DATE'] else row['DATE'][3:5], axis=1)
            df['Day'] = df.apply(lambda row: row['DATE'][5:7] if '00:00:00' in row['DATE'] else row['DATE'][:2], axis=1)
            df['Date_final'] = pd.to_datetime(df[['Year', 'Month', 'Day']], errors='coerce')
            df['Date_final'] = pd.to_datetime(df['Date_final']).dt.strftime('%Y-%m-%d')
            # Convert 'Date_final' to datetime with error handling
            df['Date_final'] = pd.to_datetime(df['Date_final'], errors='coerce')
    else:
            df['Date_final'] = df['DATE']


    # Add total pax
    df['PAX'] = df['F.A']+df['F.C']+df['F.I']+df['B.A']+df['B.C']+df['B.I']+df['E.A']+df['E.C']+df['E.I']

    # Reorder desired column that will shown in Excel
    column_orders = ['No', 'AIRLINES', 'FLIGHT NO', 'Date_final', 'AC REG', 'AC TYPE', 'STRETCH',
                     'F.A', 'F.I', 'F.C', 'B.A', 'B.I', 'B.C',
                    'E.A', 'E.I', 'E.C', 'PAX', 'CARGO', 'MAIL', 'ULD', 'Kgs',
                     'Pcs', 'STA', 'ATA', 'STD', 'ATD']
    df = df[column_orders]


    # Add departure/arrival label
    df['arr/dep'] = df['STA'].apply(lambda x: "Departure" if x == "**" else "Arrival")

    # Sort by arr/dep and date
    df = df.sort_values(by=['arr/dep', 'Date_final'])


    # Get AC Code
    df['AC CODE'] = df['AIRLINES'].str[:2]

    return df

def clean_data_kno(df):
    
        #clean date column
    if df['DATE'].dtype == 'O':  # 'O' is the code for object type
        # Clean the date column
            df['DATE'] = df['DATE'].astype(str)
            df['Year'] = df['DATE'].apply(lambda x: x[:4] if '00:00:00' in x else x[-4:])
            df['Month'] = df.apply(lambda row: row['DATE'][8:10] if '00:00:00' in row['DATE'] else row['DATE'][3:5], axis=1)
            df['Day'] = df.apply(lambda row: row['DATE'][5:7] if '00:00:00' in row['DATE'] else row['DATE'][:2], axis=1)
            df['Date_final'] = pd.to_datetime(df[['Year', 'Month', 'Day']], errors='coerce')
            df['Date_final'] = pd.to_datetime(df['Date_final']).dt.strftime('%Y-%m-%d')
            # Convert 'Date_final' to datetime with error handling
            df['Date_final'] = pd.to_datetime(df['Date_final'], errors='coerce')
    else:
            df['Date_final'] = df['DATE']




    # Add total pax
    df['PAX'] = df['F.A']+df['F.C']+df['F.I']+df['B.A']+df['B.C']+df['B.I']+df['E.A']+df['E.C']+df['E.I']

    # Reorder desired column that will shown in Excel
    column_orders = ['No', 'AIRLINES', 'FLIGHT NO', 'Date_final', 'AC REG', 'AC TYPE', 'STRETCH',
                     'F.A', 'F.I', 'F.C', 'B.A', 'B.I', 'B.C',
                    'E.A', 'E.I', 'E.C', 'PAX', 'CARGO', 'MAIL', 'ULD', 'Kgs',
                     'Pcs', 'STA', 'ATA', 'STD', 'ATD']
    df = df[column_orders]


    # Add departure/arrival label
    df['arr/dep'] = df['STA'].apply(lambda x: "Departure" if x == "**" else "Arrival")

    # Sort by arr/dep and date
    df = df.sort_values(by=['arr/dep', 'Date_final'])


    # Create function to return AC Code
    def get_ac_code(airline):
        if airline == 'QZ - AIR ASIA INDONESIA (DOMESTIC)':
            return 'QZ-DOM'
        elif airline == 'QZ - AIR ASIA INDONESIA (INTERNATIONAL)':
            return 'QZ-INT'
        elif airline == 'CX - CATHAY PACIFIC FREIGHTER':
            return 'CX-FREIGHTER'
        elif airline == '2Y - MY INDO (DOM) - PREMIER':
            return '2Y-DOM'
        elif airline == '2Y - MY INDO (INTL) - PREMIER':
            return '2Y-INT'
        else:
            return airline[:2]  # Default: Take the first two characters

    # Trigger the function
    df['AC CODE'] = df['AIRLINES'].apply(get_ac_code)

    return df

def clean_data_sub(df):
    
        #clean date column
    if df['DATE'].dtype == 'O':  # 'O' is the code for object type
        # Clean the date column
            df['DATE'] = df['DATE'].astype(str)
            df['Year'] = df['DATE'].apply(lambda x: x[:4] if '00:00:00' in x else x[-4:])
            df['Month'] = df.apply(lambda row: row['DATE'][8:10] if '00:00:00' in row['DATE'] else row['DATE'][3:5], axis=1)
            df['Day'] = df.apply(lambda row: row['DATE'][5:7] if '00:00:00' in row['DATE'] else row['DATE'][:2], axis=1)
            df['Date_final'] = pd.to_datetime(df[['Year', 'Month', 'Day']], errors='coerce')
            df['Date_final'] = pd.to_datetime(df['Date_final']).dt.strftime('%Y-%m-%d')
            # Convert 'Date_final' to datetime with error handling
            df['Date_final'] = pd.to_datetime(df['Date_final'], errors='coerce')
    else:
            df['Date_final'] = df['DATE']

    # Add total pax
    df['PAX'] = df['F.A']+df['F.C']+df['F.I']+df['B.A']+df['B.C']+df['B.I']+df['E.A']+df['E.C']+df['E.I']

    # Reorder desired column that will shown in Excel
    column_orders = ['No', 'AIRLINES', 'FLIGHT NO', 'Date_final', 'AC REG', 'AC TYPE', 'STRETCH',
                     'F.A', 'F.I', 'F.C', 'B.A', 'B.I', 'B.C',
                    'E.A', 'E.I', 'E.C', 'PAX', 'CARGO', 'MAIL', 'ULD', 'Kgs',
                     'Pcs', 'STA', 'ATA', 'STD', 'ATD']
    df = df[column_orders]


    # Add departure/arrival label
    df['arr/dep'] = df['STA'].apply(lambda x: "Departure" if x == "**" else "Arrival")

    # Sort by arr/dep and date
    df = df.sort_values(by=['arr/dep', 'Date_final'])


    def get_ac_code(airline, type):
        print(f"Checking airline: {airline}")
        if 'CX - CATHAY PACIFIC' in airline and "A333" in type:
            print("Matched CX-WB condition")
            return "CX-WB"
        if 'CX - CATHAY PACIFIC' in airline and "A321" in type:
            print("Matched CX-NB condition")
            return "CX-NB"
        if 'SQ - SINGAPORE AIRLINES' in airline and "A359" in type:
            print("Matched SQ-WB condition")
            return "SQ-WB"
        if 'SQ - SINGAPORE AIRLINES' in airline and "B738MAX" in type:
            print("MAtched SQ-NB condition")
            return "SQ-NB"
        if '2Y - MY INDO (INTL) - PREMIER' in airline:
            return "2Y-INT-NB"
        if '2Y - MY INDO (DOM) - PREMIER' in airline:
            return "2Y-DOM-NB"
        else:
            print("Did not match CX-WB condition")
            return airline[:2]


    # Trigger the function
    df['AC CODE'] = [get_ac_code(x, y) for x, y in zip(df['AIRLINES'], df['AC TYPE'])]

    return df

    

def clean_and_export_excel(df, unique_code):
    # Use BytesIO to store the Excel file in memory
    excel_buffer = BytesIO()

    # Create an Excel writer
    excel_writer = pd.ExcelWriter(excel_buffer, engine='xlsxwriter')

    # Put master data into excel
    df.to_excel(excel_writer, sheet_name='JABS', index=False)

    

    # Write each dataframe to a different sheet and create a simplified pivot table
    for airline in unique_code:
        df_airline = df[df['AC CODE'] == airline]


        # Write the original dataframe to the sheet
        df_airline.to_excel(excel_writer, sheet_name=airline, index=False)

        # Check if 'Date_final' is datetime-like
        if pd.api.types.is_datetime64_any_dtype(df_airline['Date_final']):
            # Create a simplified pivot table
            pivot_table = pd.pivot_table(df_airline, values='Date_final', index='arr/dep', columns=df_airline['Date_final'].dt.day, aggfunc='count', fill_value=0)
            pivot_table.to_excel(excel_writer, sheet_name=airline, startrow=2, startcol=35, index=True, header=True)
        else:
            print(f"Warning: 'Date_final' column in sheet {airline} is not recognized as datetime-like values.")
            print(f"Unique values in 'Date_final' column for sheet {airline}: {df_airline['Date_final'].unique()}")

    # Save the Excel file
    excel_writer.close()

    ## Seek to the beginning of the BytesIO stream
    excel_buffer.seek(0)

    return excel_buffer

def main():
    st.title("JABS DATA CLEANING AND PROCESSING")

    # File uploader
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"])


    

    if uploaded_file is not None:
        st.sidebar.info("File uploaded successfully!")

        # Importing the data
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"An error occurred: {e}")
            return

        # Choose data cleaning option
        data_cleaning_option = st.selectbox("Select Station", ["CGK", "DPS", "HLP", "KJT", "KNO", "SUB"])

        # Perform selected data cleaning
        if data_cleaning_option == "CGK":
            df_cleaned = clean_data_cgk(df)
        elif data_cleaning_option == "DPS":
            df_cleaned = clean_data_dps(df)
        elif data_cleaning_option == "HLP":
            df_cleaned = clean_data_hlp(df)
        elif data_cleaning_option == "KJT":
            df_cleaned = clean_data_kjt(df)
        elif data_cleaning_option == "KNO":
            df_cleaned = clean_data_kno(df)
        elif data_cleaning_option == "SUB":
            df_cleaned = clean_data_sub(df)
        else:
            st.warning("Please select a valid data cleaning option.")
            return

        # Create list for unique AC Code --> To be a name of sheet
        unique_code = df_cleaned['AC CODE'].unique()

        # Display some information about the cleaned data
        st.write("### Cleaned Data Preview:")
        st.write(df_cleaned.head())

        # Perform data cleaning and Excel export on button click
        if st.button("Clean Data and Export Excel"):
            output_file = clean_and_export_excel(df_cleaned, unique_code)
            st.sidebar.success(f"Data cleaned and Excel exported successfully! [Download Link](./{output_file})")

            # Add download button
            st.download_button(
                label="Download Excel File",
                data=output_file,
                file_name="cleaned_data.xlsx",  # Update file name for consistency
                key="Download_button"
            )

if __name__ == "__main__":
    main()