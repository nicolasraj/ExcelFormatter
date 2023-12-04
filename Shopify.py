from __future__ import unicode_literals
import streamlit as st
import pandas as pd
import base64
import os
from io import BytesIO
from phonenumbers import COUNTRY_CODE_TO_REGION_CODE


#title of microservice
#st.image('russell.gif')
#st.set_option('deprecation.showfileUploaderEncoding', False)
st.title('Shopify Shipping Template Export')

st.sidebar.title("Shopify Template")
#st.sidebar.info("You need to select shipping service")

#ship_list = ['Aramex', 'DHL', 'Skynet']
#address = st.sidebar.selectbox("Select Shipping Service", ship_list)



lookup = pd.read_excel('lookup.xlsx', sheet_name=['Country Code','Malaysian Postcode'], engine='openpyxl')
c_list = lookup['Country Code']

aramex_list = c_list[c_list['Service'] == 'ARAMEX']['Country Code'].tolist()
dhlex_list = c_list[c_list['Service'] == 'DHLex']['Country Code'].tolist()
skynet_list = c_list[c_list['Service'] == 'SKYNET']['Country Code'].tolist()



lookup_fill = lookup['Malaysian Postcode']
lookup_fill['Postcode2'].fillna(lookup_fill['Postcode'], inplace=True)
lookup_fill['Postcode2'] = lookup_fill['Postcode2'].astype(int)



range1 = lookup_fill['Postcode'].tolist()
range2 = lookup_fill['Postcode2'].tolist()

def get_calling_code(iso):
  for code, isos in COUNTRY_CODE_TO_REGION_CODE.items():
    if iso.upper() in isos:
        return code
  return None

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.close()
    processed_data = output.getvalue()
    return processed_data


def get_table_download_link(df):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    val = to_excel(df)
    b64 = base64.b64encode(val)  # val looks like b'...'
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="extract.xlsx">Download XLXS</a>' # decode b'abc' => abc



def get_table_download_link_csv(df):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()  # some strings <-> bytes conversions necessary here
    href = f'<a href="data:file/csv;base64,{b64}" download="extract.csv">Download csv file</a>'
    return href


user_file = st.sidebar.file_uploader("Upload your Shopify File:")
#user_file = io.TextIOWrapper(uploaded_file)
if user_file is not None:
    try:
        with st.spinner("Uploading your Shopify File..."):
            df = pd.read_excel(user_file, dtype=str)
            #GroupBy
            unq = df.groupby('Name').first()
            cnt = df.groupby('Name').count()[['Customer: Email']]
            cnt = cnt.rename(columns={'Customer: Email':'Count'})

            df = pd.merge(unq,cnt, left_on=unq.index, right_on=cnt.index)
            df = df.rename(columns={'key_0':'Name'})

            df = df.fillna('')


            df['Assign']="ARAMEX"
            df.loc[(df['Shipping: Country Code'].isin(dhlex_list)),'Assign']="DHLex"
            df.loc[(df['Shipping: Country Code'].isin(aramex_list)),'Assign']="ARAMEX"

            df.loc[(df['Shipping: Country Code']=='MY'),'Assign']="SKYNET"

            for i in range(len(df['Shipping: Zip'])):
                if df['Shipping: Zip'][i].isdigit():
                    for j in range(len(range1)):
                        if range1[j] <= int(df['Shipping: Zip'][i]) <= range2[j]:
                            df.loc[(df['Shipping: Country Code']=='MY'),'Assign']="DHL"
            df['Shipping: Zip'] = df['Shipping: Zip'].astype(str)



            st.success('Done!')
        st.subheader("Your Shopify File")
        st.write(df)
    except Exception as e:
        st.error(
            f"Sorry, there was a problem processing your Shopify file./n {e}"
        )
        user_file = None


    data = df

    #Aramex
    aramex = data[data['Assign'] == 'ARAMEX']
    if len(aramex) != 0:
        aramex['fullname'] = aramex['Shipping: First Name']+ ' ' + aramex['Shipping: Last Name']
        aramex['fulladdress'] = aramex['Shipping: Address 1']+ ', ' + aramex['Shipping: Address 2']+ ', ' +aramex['Shipping: City']+ ', ' + aramex['Shipping: Province']
        aramex['Descgoods'] = 'Merch'
        aramex['Pieces'] = aramex['Count']
        aramex['Weight_kg'] = aramex['Weight Total'].astype(float)/1000
        aramex['Shipping: Phone'] = aramex['Shipping: Phone'].astype(str)
        aramex['Shipping: Phone'] = aramex['Shipping: Phone'].str.replace("'",'')


        aramex_reorder = aramex[['fullname', 'fullname', 'fulladdress',
                                'Shipping: Country', 'Shipping: Zip', 'Shipping: Phone',
                                'Customer: Email', 'Transaction: Amount', 'Currency',
                                'Descgoods','Pieces', 'Weight_kg',
                                'Name']]


        aramex_reorder.columns = ['Name', 'Contact', 'Address',
                                  'Country', 'Zipcode', 'Phone Number',
                                  'Email', 'Customs Value', 'Currency',
                                  'Description of Goods', 'Pieces', 'Weight',
                                  'Comments']


        #aramex_reorder.to_excel('aramex.xlsx', index=False)
        st.subheader("Your Aramex File")
        st.write(aramex_reorder)
        st.markdown(get_table_download_link(aramex_reorder), unsafe_allow_html=True)


    #DHL
    dhl = data[data['Assign'] == 'DHL']
    if len(dhl) != 0:

        #Split # from ordernumber(name) column
        tempdata = dhl['Name'].str.split("#",expand=True)
        dhl['Name_no']=tempdata[1]

        dhl['fullname'] = dhl['Shipping: First Name']+ ' ' + dhl['Shipping: Last Name']
        dhl['shipdesc'] = 'Merch'
        dhl['PickupNo']='5275477549'
        dhl['ShipServiceCode']='PDO'
        dhl['Currency Code'] = 'MYR'
        dhl['Total Declared Value'] = dhl['Transaction: Amount'].astype(float) * 4.27
        dhl['Total Declared Value'] = dhl['Total Declared Value'].round(2)
        dhl['Weight Total'] = dhl['Weight Total'].astype(float)
        dhl['Weight Total'] = dhl['Weight Total'].astype(int)
        dhl['Na_ad3']=''
        dhl['Na_isinsured']=''
        dhl['Na_insurance']=''
        dhl['Na_iscod']=''
        dhl['Na_codvalue']=''
        dhl['Na_service1']=''
        dhl['Na_ismult']=''
        dhl['Na_delopt']=''
        dhl['Na_pieceid']=''
        dhl['Na_piecedesc']=''
        dhl['Na_pieceweight']=''
        dhl['Na_piececod']=''
        dhl['Na_pieceinsurance']=''
        dhl['Na_piecebilref1']=''
        dhl['Na_piecebilref2']='Merch Item'
        dhl['Shipping: Phone'] = dhl['Shipping: Phone'].astype(str)
        dhl['Shipping: Phone'] = dhl['Shipping: Phone'].str.replace("'",'')



        dhl_reorder = dhl[['PickupNo', 'Name_no', 'ShipServiceCode',
                            'fullname', 'Shipping: Address 1', 'Shipping: Address 2',
                            'Na_ad3', 'Shipping: City', 'Shipping: Province',
                            'Shipping: Zip', 'Shipping: Country Code', 'Shipping: Phone',
                            'Weight Total', 'Currency Code', 'Total Declared Value',
                            'Na_isinsured', 'Na_insurance', 'Na_iscod',
                            'Na_codvalue', 'shipdesc', 'Na_service1',
                            'Na_ismult', 'Na_delopt', 'Na_pieceid',
                            'Na_piecedesc', 'Na_pieceweight', 'Na_piececod',
                            'Na_pieceinsurance', 'Na_piecebilref1', 'Na_piecebilref2']]


        dhl_reorder.columns =['Pick-up Account Number', 'Shipment Order ID', 'Shipping Service Code',
                                'Consignee Name', 'Address Line 1', 'Address Line 2',
                                'Address Line 3', 'City', 'State',
                                'Postal Code', 'Destination Country Code', 'Phone Number',
                                'Shipment Weight (g)', 'Currency Code', 'Total Declared Value',
                                'Is Insured', 'Insurance', 'Is COD',
                                'Cash on Delivery Value', 'Shipment Description', 'Service1',
                                'IsMult', 'Delivery Option', 'PieceID',
                                'Piece Description', 'Piece Weight', 'Piece COD',
                                'Piece Insurance', 'Piece Billing Reference 1', 'Piece Billing Reference 2']


        st.subheader("Your DHL eCom File")
        st.write(dhl_reorder)
        st.markdown(get_table_download_link(dhl_reorder), unsafe_allow_html=True)

    #SKYNET
    skynet = data[data['Assign'] == 'SKYNET']
    if len(skynet) != 0:

        skynet['Weight Total'] =skynet['Weight Total'].astype(float)/1000
        skynet['fullname'] = skynet['Shipping: First Name']+ ' ' + skynet['Shipping: Last Name']
        skynet['Shipping: Phone'] = skynet['Shipping: Phone'].astype(str)
        skynet['Shipping: Phone'] = skynet['Shipping: Phone'].str.replace("'",'')
        skynet['FullAddress'] = skynet['Shipping: Address 1'] + ', ' + skynet['Shipping: Address 2'] + ', ' + skynet['Shipping: Province'] + ', ' + skynet['Shipping: City'] + ', ' + skynet['Shipping: Zip']

        skynet_reorder = skynet[['fullname', 'Shipping: Phone', 'Count', 'Weight Total',
                                'Shipping: Country', 'FullAddress', 'Name']]

        skynet_reorder.columns = ['CompanyName', 'PhoneNo', 'NoOfPackage', 'Weight',
                                  'Country', 'Address1', 'ProductCode']


        st.subheader("Your SKYNET File")
        st.write(skynet_reorder)
        st.markdown(get_table_download_link(skynet_reorder), unsafe_allow_html=True)



    #SKYNET
    dhlex = data[data['Assign'] == 'DHLex']
    if len(dhlex) != 0:
        dhlex['﻿Name (Ship FROM) (Required)'] = 'Kingdomcity KL'
        dhlex['Company (Ship FROM) (Required)'] = 'Kingdomcity KL'
        dhlex['Address 1 (Ship FROM) (Required)'] = 'A4, EV-T-08, Third Floor,'
        dhlex['Address 2 (Ship FROM)'] = 'Evolve Concept Mall D, 2-7,'
        dhlex['Address 3 (Ship FROM)'] = 'Jalan PJU 1/1, Ara Damansara'
        dhlex['ZIP/Postal Code (Ship FROM)'] = '47301'
        dhlex['City (Ship FROM) (Required)'] = 'Petaling Jaya'
        dhlex['Country Code (Ship FROM) (Required)'] = 'MY'
        dhlex['Email Address (Ship FROM) (Required)'] = 'cs.kl@kingdomcity.com'
        dhlex['Phone Country Code (Ship FROM)'] = '60'
        dhlex['Phone Number (Ship FROM) (Required)'] = '327790525'
        dhlex['Name (Ship TO) (Required)'] = dhlex['Shipping: First Name']+ ' ' + dhlex['Shipping: Last Name']
        dhlex['Company (Ship TO) (Required)'] = dhlex['Shipping: First Name']+ ' ' + dhlex['Shipping: Last Name']
        dhlex['Address 1 (Ship TO) (Required)'] = dhlex['Shipping: Address 1']
        dhlex['Address 2 (Ship TO)'] = dhlex['Shipping: Address 2']
        dhlex['Address 3 (Ship TO)'] = ''
        dhlex['ZIP/Postal Code (Ship TO)'] = dhlex['Shipping: Zip']
        dhlex['City (Ship TO) (Required)'] = dhlex['Shipping: City']
        dhlex['Suburb (Ship TO)'] = ''
        dhlex['State/Province (Ship TO)'] = dhlex['Shipping: Province']
        dhlex['State/Province Code (Ship TO)'] = ''
        dhlex['Country Code (Ship TO) (Required)'] = dhlex['Shipping: Country Code']
        dhlex['Email Address (Ship TO)'] = dhlex['Customer: Email']
        dhlex['Phone Country Code (Ship TO)'] = dhlex['Shipping: Country Code'].apply(get_calling_code)
        dhlex['Phone Number (Ship TO) (Required)'] =  dhlex['Shipping: Phone'].astype(str)
        dhlex['Phone Number (Ship TO) (Required)'] = dhlex['Phone Number (Ship TO) (Required)'].str.replace("'",'')
        dhlex['Product Code (Global)'] = 'P'
        dhlex['Product Code (Local)'] = 'P'
        dhlex['Shipment Type (Required)'] = 'P'
        dhlex['Product Code (3 Letter) (Required)'] = 'WPX'
        dhlex['Total Shipment Pieces (Required)'] = dhlex['Count']
        dhlex['Total Weight (Required)'] = dhlex['Weight Total'].astype(float)/1000
        dhlex['Summary of Contents (Required)'] = 'Merch'
        dhlex['Shipment Reference'] = dhlex['Name']
        dhlex['Declared Value (Required)'] = dhlex['Transaction: Amount']
        dhlex['Declared Value Currency (Required)'] = 'USD'
        dhlex['Account Number (Shipper) (Required)'] = '550267931'
        dhlex['Account Number (Duty/Tax)'] = ''
        dhlex['Trade Term'] = 'DAP'



        dhlex_reorder = dhlex[['﻿Name (Ship FROM) (Required)', 'Company (Ship FROM) (Required)', 'Address 1 (Ship FROM) (Required)',
                                'Address 2 (Ship FROM)', 'Address 3 (Ship FROM)', 'ZIP/Postal Code (Ship FROM)',
                                'City (Ship FROM) (Required)', 'Country Code (Ship FROM) (Required)', 'Email Address (Ship FROM) (Required)',
                                'Phone Country Code (Ship FROM)', 'Phone Number (Ship FROM) (Required)', 'Name (Ship TO) (Required)',
                                'Company (Ship TO) (Required)', 'Address 1 (Ship TO) (Required)', 'Address 2 (Ship TO)',
                                'Address 3 (Ship TO)', 'ZIP/Postal Code (Ship TO)', 'City (Ship TO) (Required)',
                                'Suburb (Ship TO)', 'State/Province (Ship TO)', 'State/Province Code (Ship TO)',
                                'Country Code (Ship TO) (Required)', 'Email Address (Ship TO)', 'Phone Country Code (Ship TO)',
                                'Phone Number (Ship TO) (Required)', 'Product Code (Global)', 'Product Code (Local)',
                                'Shipment Type (Required)', 'Product Code (3 Letter) (Required)', 'Total Shipment Pieces (Required)',
                                'Total Weight (Required)', 'Summary of Contents (Required)', 'Shipment Reference',
                                'Declared Value (Required)', 'Declared Value Currency (Required)', 'Account Number (Shipper) (Required)',
                                'Account Number (Duty/Tax)', 'Trade Term']]

        dhlex_reorder.columns = ['﻿Name (Ship FROM) (Required)', 'Company (Ship FROM) (Required)', 'Address 1 (Ship FROM) (Required)',
                                'Address 2 (Ship FROM)', 'Address 3 (Ship FROM)', 'ZIP/Postal Code (Ship FROM)',
                                'City (Ship FROM) (Required)', 'Country Code (Ship FROM) (Required)', 'Email Address (Ship FROM) (Required)',
                                'Phone Country Code (Ship FROM)', 'Phone Number (Ship FROM) (Required)', 'Name (Ship TO) (Required)',
                                'Company (Ship TO) (Required)', 'Address 1 (Ship TO) (Required)', 'Address 2 (Ship TO)',
                                'Address 3 (Ship TO)', 'ZIP/Postal Code (Ship TO)', 'City (Ship TO) (Required)',
                                'Suburb (Ship TO)', 'State/Province (Ship TO)', 'State/Province Code (Ship TO)',
                                'Country Code (Ship TO) (Required)', 'Email Address (Ship TO)', 'Phone Country Code (Ship TO)',
                                'Phone Number (Ship TO) (Required)', 'Product Code (Global)', 'Product Code (Local)',
                                'Shipment Type (Required)', 'Product Code (3 Letter) (Required)', 'Total Shipment Pieces (Required)',
                                'Total Weight (Required)', 'Summary of Contents (Required)', 'Shipment Reference',
                                'Declared Value (Required)', 'Declared Value Currency (Required)', 'Account Number (Shipper) (Required)',
                                'Account Number (Duty/Tax)', 'Trade Term']


        st.subheader("Your DHL Express File")
        st.write(dhlex_reorder)
        st.markdown(get_table_download_link_csv(dhlex_reorder), unsafe_allow_html=True)
