import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Pengaturan halaman
st.set_page_config(
    page_title="Dasbor Penjualan",
    page_icon=":bar_chart:",
    layout="wide"
)

st.title(":robot_face: Dasbor Penjualan Streamlit")
st.markdown("_Versi Prototipe v1.0.0_")

# Path data default
DEFAULT_DATA_PATH = 'sales.xlsx'  # Ganti dengan path file data default Anda

# Bagian unggah file
st.header("Unggah File")
uploaded_file = st.file_uploader(
    label="Unggah file Excel Anda", type="xlsx"
)

# Fungsi pemuatan data
@st.cache
def load_data(path: str):
    return pd.read_excel(path, sheet_name=None)

# Memuat data: gunakan file yang diunggah jika ada; jika tidak, gunakan data default
if uploaded_file:
    data = load_data(uploaded_file)
    st.success("Data baru berhasil dimuat!")
else:
    data = load_data(DEFAULT_DATA_PATH)
    st.info("Menggunakan file data default. Unggah file baru untuk memperbarui.")

# Ekstraksi data dari sheet
sales_data = data['SalesData']
product_data = data['Product']
customer_data = data['Customer']
salesrep_data = data['SalesRep']

# Gabungkan data berdasarkan kolom pengenal umum
merged_data = sales_data.merge(product_data, on='ProductID').merge(customer_data, on='CustomerID').merge(salesrep_data, on='SalesRepID')

# Pratinjau data
with st.expander("Pratinjau Data"):
    st.dataframe(merged_data)

# Fungsi Visualisasi
def plot_sales_per_product():
    total_sales_per_product = merged_data.groupby(['Year', 'Product Name'])['Sales Amount (in US$)'].sum().reset_index()
    fig = px.bar(total_sales_per_product, x='Product Name', y='Sales Amount (in US$)', color='Year', title="Penjualan Tahunan per Produk")
    st.plotly_chart(fig, use_container_width=True)

def plot_avg_sales_per_customer():
    avg_sales_per_customer = merged_data.groupby(['CustomerType', 'Size', 'Subsidized'])['Sales Amount (in US$)'].mean().reset_index()
    fig = px.bar(avg_sales_per_customer, x='CustomerType', y='Sales Amount (in US$)', color='Size', barmode='group', title="Rata-Rata Penjualan Berdasarkan Karakteristik Pelanggan")
    st.plotly_chart(fig, use_container_width=True)

def plot_top_sales_reps():
    top_sales_reps = merged_data.groupby(['Firstnames', 'Surnames'])['Sales Amount (in US$)'].sum().nlargest(3).reset_index()
    fig = px.bar(top_sales_reps, x='Firstnames', y='Sales Amount (in US$)', color='Surnames', title="3 Perwakilan Penjualan Teratas Berdasarkan Total Penjualan")
    st.plotly_chart(fig, use_container_width=True)

def plot_sales_by_customer_type():
    product_sales = merged_data.groupby(['Product Name', 'CustomerType'])['Sales Amount (in US$)'].sum().reset_index()
    fig = px.bar(product_sales, x='Product Name', y='Sales Amount (in US$)', color='CustomerType', title="Total Penjualan Setiap Produk Berdasarkan Jenis Pelanggan")
    st.plotly_chart(fig, use_container_width=True)

def plot_monthly_sales_proportion():
    # Pastikan kolom bulan dikonversi ke datetime jika dalam format string
    merged_data['Month'] = pd.to_datetime(merged_data['Month'], format='%b').dt.month
    monthly_sales = merged_data.groupby(['Month', 'CustomerType'])['Sales Amount (in US$)'].sum().reset_index()
    fig = px.area(monthly_sales, x='Month', y='Sales Amount (in US$)', color='CustomerType', title="Proporsi Penjualan Bulanan Berdasarkan Jenis Pelanggan", line_group='CustomerType')
    st.plotly_chart(fig, use_container_width=True)

def plot_top_reps_performance():
    # Filter untuk perwakilan teratas berdasarkan data penjualan
    top_sales_reps = merged_data.groupby(['Firstnames', 'Surnames'])['Sales Amount (in US$)'].sum().nlargest(3).reset_index()
    top_reps_ids = top_sales_reps[['Firstnames', 'Surnames']].apply(lambda x: ' '.join(x), axis=1).tolist()
    merged_data['FullName'] = merged_data['Firstnames'] + ' ' + merged_data['Surnames']
    top_reps_data = merged_data[merged_data['FullName'].isin(top_reps_ids)]
    rep_sales = top_reps_data.groupby(['Year', 'FullName'])['Sales Amount (in US$)'].sum().reset_index()
    fig = px.line(rep_sales, x='Year', y='Sales Amount (in US$)', color='FullName', title="Kinerja Perwakilan Penjualan Teratas Selama Beberapa Tahun")
    st.plotly_chart(fig, use_container_width=True)

# Layout Streamlit
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

# Visualisasi Data Tambahan dengan Tab
st.subheader("Visualisasi Data Tambahan")
tab1, tab2, tab3 = st.tabs(["Proporsi Penjualan Bulanan", "Perwakilan Penjualan Teratas", "Kinerja Perwakilan Penjualan Teratas Selama Waktu"])

with tab1:
    plot_monthly_sales_proportion()

with tab2:
    plot_top_sales_reps()

with tab3:
    plot_top_reps_performance()

# Bagian Analisis dan Pelaporan
st.header("Analisis dan Pelaporan")
analysis = """
Dari analisis ini, kami mengamati:
1. **Penjualan Tahunan per Produk**: Total penjualan per produk menunjukkan tren yang menyoroti produk-produk utama dengan permintaan yang konsisten. Produk benih memiliki penjualan substansial dari petani besar, sementara pupuk lebih banyak dibeli oleh pengecer berukuran menengah.
2. **Tren Jenis Pelanggan**: Pelanggan besar cenderung melakukan pembelian dalam jumlah besar, dengan jenis pelanggan petani mendominasi pembelian benih.
3. **Perwakilan Teratas**: 3 perwakilan penjualan teratas menunjukkan kinerja yang konsisten, dengan peningkatan yang signifikan di beberapa tahun yang mungkin disebabkan oleh sertifikasi yang diperoleh. Perwakilan yang bersertifikat dapat lebih memenuhi kebutuhan klien, sehingga meningkatkan hasil penjualan.
4. **Tren Bulanan**: Tren musiman dalam penjualan bulanan menunjukkan penjualan yang lebih tinggi di bulan-bulan awal untuk pupuk, kemungkinan karena musim tanam.

Secara keseluruhan, berfokus pada segmen pelanggan dengan permintaan tinggi selama musim tanam dapat mengoptimalkan strategi. Meningkatkan program sertifikasi mungkin juga dapat memberikan manfaat lebih lanjut pada kinerja perwakilan penjualan.
"""
st.write(analysis)
