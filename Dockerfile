FROM python:3.7
EXPOSE 8501
WORKDIR /app
RUN pip3 install pandas
RUN pip3 install phonenumbers
RUN pip3 install streamlit
RUN pip3 install xlrd
RUN pip3 install openpyxl
RUN pip3 install xlsxwriter
COPY . .
CMD streamlit run Shopify.py --server.fileWatcherType none --server.port $PORT
