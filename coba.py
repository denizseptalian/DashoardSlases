import streamlit as st
import pandas as pd
import plotly.express as px

# Page settings
st.set_page_config(
    page_title="Dasbor Analitik",
    page_icon=":bar_chart:",
    layout="wide"
)

st.title(":robot_face: Dasboard Analitik Penjualan Benih dan Pupuk Gelar Rasa 2024")
st.markdown("Team Maba Muda")

# Default data path
DEFAULT_DATA_PATH = 'sales.xlsx'  # Change to your default data file path

# File upload section
st.header("Jika Ingin Mengunkan Daa Baru harap sesuaikan dengan Tamplate berikut : https://s.id/TamplateData")
uploaded_file = st.file_uploader(
    label="Unggah file Excel Anda", type="xlsx"
)

# Data loading function
@st.cache
def load_data(path: str):
    return pd.read_excel(path, sheet_name=None)

# Load data: use uploaded file if available, otherwise use default data
if uploaded_file:
    data = load_data(uploaded_file)
    st.success("Data baru berhasil dimuat!")
else:
    data = load_data(DEFAULT_DATA_PATH)
    st.info("Menggunakan file data default. Unggah file baru untuk memperbarui.")

# Ensure data has all required sheets
required_sheets = {'SalesData', 'Product', 'Customer', 'SalesRep'}
if not required_sheets.issubset(data.keys()):
    st.error("File Excel harus memiliki sheet: SalesData, Product, Customer, dan SalesRep.")
else:
    # Extract data from each sheet
    sales_data = data['SalesData']
    product_data = data['Product']
    customer_data = data['Customer']
    salesrep_data = data['SalesRep']

    # Merge data based on common identifier columns
    merged_data = sales_data.merge(product_data, on='ProductID').merge(customer_data, on='CustomerID').merge(salesrep_data, on='SalesRepID')

    # Filter by year or all data
    st.sidebar.header("Filter")
    filter_option = st.sidebar.radio("Pilih Opsi Data", ["Data Seluruhnya", "Filter Berdasarkan Tahun"])

    if filter_option == "Filter Berdasarkan Tahun":
        if 'Year' in merged_data.columns:
            available_years = merged_data['Year'].unique()
            selected_year = st.sidebar.selectbox("Pilih Tahun", sorted(available_years))

            # Filter data by selected year
            filtered_data = merged_data[merged_data['Year'] == selected_year]
        else:
            st.warning("Kolom 'Year' tidak tersedia dalam data.")
            filtered_data = merged_data
    else:
        filtered_data = merged_data

    # Preview filtered data
    with st.expander("Pratinjau Data Berdasarkan Pilihan Filter"):
        st.dataframe(filtered_data)

    # Visualization functions
    def plot_sales_per_product():
        total_sales_per_product = filtered_data.groupby(['Year', 'Product Name'])['Sales Amount (in US$)'].sum().reset_index()
        fig = px.bar(total_sales_per_product, x='Product Name', y='Sales Amount (in US$)', color='Year', title="Penjualan Tahunan per Produk")
        st.plotly_chart(fig, use_container_width=True)

    def plot_avg_sales_per_customer():
        avg_sales_per_customer = filtered_data.groupby(['CustomerType', 'Size', 'Subsidized'])['Sales Amount (in US$)'].mean().reset_index()
        fig = px.bar(avg_sales_per_customer, x='CustomerType', y='Sales Amount (in US$)', color='Size', barmode='group', title="Rata-Rata Penjualan Berdasarkan Karakteristik Pelanggan")
        st.plotly_chart(fig, use_container_width=True)

    def plot_top_sales_reps():
        # Group data to find the top 3 sales reps by sales amount
        top_sales_reps = (
            filtered_data
            .groupby(['Firstnames', 'Gender', 'Tenure', 'Certified Crop Adviser (CCA)', 'Certified Professional Agronomist (CPA)'])['Sales Amount (in US$)']
            .sum()
            .nlargest(3)
            .reset_index()
        )

        # Create composite labels for the legend, including Tenure, CCA, and CPA
        top_sales_reps['Rep Legend'] = (
            'Gender: ' + top_sales_reps['Gender'] + ', ' +
            'Tenure: ' + top_sales_reps['Tenure'].astype(str) + ', ' +
            'CCA: ' + top_sales_reps['Certified Crop Adviser (CCA)'].astype(str) + ', ' +
            'CPA: ' + top_sales_reps['Certified Professional Agronomist (CPA)'].astype(str)
        )

        # Plot with separate x-axis names and legend
        fig = px.bar(
            top_sales_reps,
            x='Firstnames',
            y='Sales Amount (in US$)',
            color='Rep Legend',
            title="3 Perwakilan Penjualan Teratas Berdasarkan Total Penjualan",
            labels={'Firstnames': 'Sales Rep', 'Rep Legend': 'Details (Gender, Tenure, Certified Crop Adviser (CCA), Certified Professional Agronomist (CPA))'}
        )

        # Display the chart in Streamlit
        st.plotly_chart(fig, use_container_width=True)

    def plot_sales_by_customer_type():
        product_sales = filtered_data.groupby(['Product Name', 'CustomerType'])['Sales Amount (in US$)'].sum().reset_index()
        fig = px.bar(product_sales, x='Product Name', y='Sales Amount (in US$)', color='CustomerType', title="Total Penjualan Setiap Produk Berdasarkan Jenis Pelanggan")
        st.plotly_chart(fig, use_container_width=True)

    def plot_monthly_sales_proportion():
        if filtered_data['Month'].dtype == 'object':
            filtered_data['Month'] = pd.to_datetime(filtered_data['Month'], format='%b').dt.month

        monthly_sales = filtered_data.groupby(['Month', 'CustomerType'])['Sales Amount (in US$)'].sum().reset_index()

        fig = px.line(
            monthly_sales,
            x='Month',
            y='Sales Amount (in US$)',
            color='CustomerType',
            title="Monthly Sales Trend by Customer Type",
            labels={'Month': 'Month', 'Sales Amount (in US$)': 'Sales Amount (in USD)', 'CustomerType': 'Customer Type'}
        )

        fig.update_xaxes(type='category', tickvals=list(range(1, 13)), ticktext=['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])
        st.plotly_chart(fig, use_container_width=True)

    def plot_top_reps_performance():
        top_sales_reps = filtered_data.groupby(['Firstnames', 'Gender'])['Sales Amount (in US$)'].sum().nlargest(3).reset_index()
        top_reps_ids = top_sales_reps[['Firstnames', 'Gender']].apply(lambda x: ' '.join(x), axis=1).tolist()
        filtered_data['FullName'] = filtered_data['Firstnames'] + ' ' + filtered_data['Gender']
        top_reps_data = filtered_data[filtered_data['FullName'].isin(top_reps_ids)]
        rep_sales = top_reps_data.groupby(['Year', 'FullName'])['Sales Amount (in US$)'].sum().reset_index()
        fig = px.line(rep_sales, x='Year', y='Sales Amount (in US$)', color='FullName', title="Kinerja Perwakilan Penjualan Teratas Selama Beberapa Tahun")
        st.plotly_chart(fig, use_container_width=True)

    # Streamlit layout
    top_left_column, top_right_column = st.columns((2, 1))
    bottom_left_column, bottom_right_column = st.columns(2)

    with top_left_column:
        st.subheader("Analisis Deskriptif")
        plot_sales_per_product()

    with top_right_column:
        st.subheader("Rata-Rata Penjualan per Karakteristik Pelanggan")
        plot_avg_sales_per_customer()

    with bottom_left_column:
        st.subheader("Penjualan Berdasarkan Jenis Pelanggan")
        plot_sales_by_customer_type()

    # Additional Data Visualizations with Tabs
    st.subheader("Visualisasi Data Tambahan")
    tab1, tab2, tab3 = st.tabs(["Proporsi Penjualan Bulanan", "Perwakilan Penjualan Teratas", "Kinerja Perwakilan Penjualan Teratas Selama Waktu"])

    with tab1:
        plot_monthly_sales_proportion()

    with tab2:
        plot_top_sales_reps()

    with tab3:
        plot_top_reps_performance()

    # Analysis and Reporting Section
    st.header("Analisis dan Pelaporan")
    analysis = """
    Dari analisis ini, kami mengamati:
    1. **Penjualan Tahunan per Produk**: Total penjualan per produk menunjukkan tren yang menyoroti produk-produk utama dengan permintaan yang konsisten.
    2. **Tren Jenis Pelanggan**: Pelanggan besar cenderung melakukan pembelian dalam jumlah besar, dengan jenis pelanggan petani mendominasi pembelian benih.
    3. **Perwakilan Teratas**: 3 perwakilan penjualan teratas menunjukkan kinerja yang konsisten, dengan peningkatan yang signifikan di beberapa tahun.
    4. **Tren Bulanan**: Tren musiman dalam penjualan bulanan menunjukkan penjualan yang lebih tinggi di bulan-bulan awal untuk pupuk, kemungkinan karena musim tanam.

    Secara keseluruhan, berfokus pada segmen pelanggan dengan permintaan tinggi selama musim tanam dapat mengoptimalkan strategi. Meningkatkan program sertifikasi mungkin juga dapat memberikan manfaat lebih lanjut pada kinerja perwakilan penjualan.
    """
    st.write(analysis)
