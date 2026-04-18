import streamlit as st
import pandas as pd
import pyodbc
import time
import io
from datetime import datetime

# KONFIGURASI KONEKSI
SERVER = '127.0.0.1'
DATABASE = 'CP'
USERNAME = 'sa'
PASSWORD = 'myPass123'

def get_connection():
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        f"UID={USERNAME};"
        f"PWD={PASSWORD};"
        f"Connection Timeout=30;"
    )
    return pyodbc.connect(conn_str)

def update_itemname2(itemid, new_itemname2):
    try:
        conn = get_connection()
        cursor = conn.cursor()
        query = "UPDATE taGoods SET ItemName2 = ? WHERE ItemID = ?"
        cursor.execute(query, (new_itemname2, str(itemid)))
        conn.commit()
        affected = cursor.rowcount
        cursor.close()
        conn.close()
        if affected == 0:
            return False, f"Item ID '{itemid}' tidak ditemukan!"
        return True, f"Berhasil update {affected} data!"
    except Exception as e:
        return False, f"Error: {e}"

def fetch_all_data():
    """Ambil SEMUA data tanpa batasan"""
    try:
        conn = get_connection()
        # HAPUS TOP jika ada, ambil semua data
        query = "SELECT ItemID, ItemName2 FROM taGoods ORDER BY ItemID"
        df = pd.read_sql(query, conn)
        conn.close()
        return df
    except Exception as e:
        st.error(f"Error: {e}")
        return pd.DataFrame()

def get_total_count():
    """Hitung total data di tabel"""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM taGoods")
        count = cursor.fetchone()[0]
        cursor.close()
        conn.close()
        return count
    except Exception as e:
        return 0

def create_template_excel():
    template_df = pd.DataFrame({
        'ItemID': ['contoh_id_1', 'contoh_id_2', 'contoh_id_3'],
        'ItemName2': ['4双本体中心', '测试汉字', '日本語テスト']
    })
    return template_df

def bulk_update(excel_file):
    try:
        df = pd.read_excel(excel_file)
        required_cols = ['ItemID', 'ItemName2']
        if not all(col in df.columns for col in required_cols):
            return False, 0, 0, f"File harus memiliki kolom: {', '.join(required_cols)}"
        
        conn = get_connection()
        cursor = conn.cursor()
        success_count = 0
        fail_list = []
        
        for _, row in df.iterrows():
            try:
                cursor.execute(
                    "UPDATE taGoods SET ItemName2 = ? WHERE ItemID = ?",
                    (str(row['ItemName2']), str(row['ItemID']))
                )
                if cursor.rowcount > 0:
                    success_count += 1
                else:
                    fail_list.append({'ItemID': row['ItemID'], 'Status': 'ID tidak ditemukan'})
            except Exception as e:
                fail_list.append({'ItemID': row['ItemID'], 'Status': str(e)})
        
        conn.commit()
        cursor.close()
        conn.close()
        return True, success_count, len(df), fail_list
    except Exception as e:
        return False, 0, 0, str(e)

def main():
    st.set_page_config(page_title="Update Data taGoods", page_icon="📝", layout="wide")
    st.title("📝 Update Data taGoods")
    
    # Test koneksi dan hitung total data
    try:
        with st.spinner("🔌 Menguji koneksi database..."):
            conn = get_connection()
            conn.close()
        st.success("✅ Koneksi database berhasil!")
        
        # Tampilkan total data
        total_data = get_total_count()
        st.info(f"📊 Total data di database: **{total_data} item**")
        
    except Exception as e:
        st.error(f"❌ Gagal koneksi: {e}")
        st.stop()
    
    # Tombol refresh
    col_btn1, col_btn2 = st.columns([1, 5])
    with col_btn1:
        if st.button("🔄 Refresh Data", use_container_width=True):
            st.rerun()
    
    st.divider()
    
    # Tampilkan SEMUA data
    st.subheader(f"📋 Data Saat Ini (Total: {get_total_count()} item)")
    
    with st.spinner("📊 Memuat semua data dari database..."):
        df = fetch_all_data()
    
    if not df.empty:
        # Tampilkan jumlah data yang berhasil dimuat
        st.caption(f"✅ Berhasil memuat {len(df)} data")
        
        # Gunakan height agar bisa scroll melihat semua data
        st.dataframe(df, use_container_width=True, height=500)
        
        # Opsi download data sebagai CSV
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Data sebagai CSV",
            data=csv,
            file_name=f"taGoods_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
        )
    else:
        st.warning("⚠️ Tidak ada data ditemukan")
    
    # Tab
    tab1, tab2, tab3 = st.tabs(["✏️ Update Single", "📊 Bulk Update Excel", "🔍 Cari & Update"])
    
    # TAB 1: Update Single
    with tab1:
        st.subheader("Update Single Data")
        
        col1, col2 = st.columns(2)
        with col1:
            itemid = st.text_input("Item ID", placeholder="Contoh: 01B022BK", key="single_id")
        with col2:
            new_value = st.text_input("ItemName2 (Baru)", placeholder="Contoh: 4双本体中心", key="single_value")
        
        if st.button("🚀 Update Data", type="primary", key="single_btn"):
            if itemid and new_value:
                with st.spinner(f"💾 Menyimpan data..."):
                    time.sleep(0.5)
                    success, message = update_itemname2(itemid, new_value)
                    if success:
                        st.success(message)
                        st.balloons()
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(message)
            else:
                st.warning("Harap isi Item ID dan nilai baru!")
    
    # TAB 2: Bulk Update
    with tab2:
        st.subheader("Bulk Update via Excel")
        
        # Download template
        template_df = create_template_excel()
        template_buffer = io.BytesIO()
        with pd.ExcelWriter(template_buffer, engine='openpyxl') as writer:
            template_df.to_excel(writer, index=False, sheet_name='Template')
        template_buffer.seek(0)
        
        st.download_button(
            label="📥 Download Template Excel",
            data=template_buffer,
            file_name=f"template_update_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("**Kolom yang diperlukan:** `ItemID` dan `ItemName2`")
        
        uploaded_file = st.file_uploader("Pilih file Excel", type=['xlsx', 'xls'], key="bulk_file")
        
        if uploaded_file:
            preview_df = pd.read_excel(uploaded_file)
            st.write("**Preview data:**")
            st.dataframe(preview_df.head(10), use_container_width=True)
            
            if st.button("🚀 Jalankan Bulk Update", type="primary", key="bulk_btn"):
                with st.spinner("📊 Memproses bulk update..."):
                    success, success_count, total, fail_list = bulk_update(uploaded_file)
                    
                    if success:
                        st.success(f"✅ Berhasil update {success_count} dari {total} data!")
                        if success_count > 0:
                            st.balloons()
                        if fail_list and len(fail_list) > 0:
                            st.warning(f"⚠️ {len(fail_list)} data gagal:")
                            st.dataframe(pd.DataFrame(fail_list))
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(f"❌ Error: {fail_list}")
    
    # TAB 3: Cari & Update
    with tab3:
        st.subheader("Cari Data")
        
        search_id = st.text_input("Masukkan Item ID", placeholder="Contoh: 01B022BK", key="search_id")
        
        if search_id:
            with st.spinner("🔍 Mencari..."):
                try:
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute("SELECT ItemID, ItemName2 FROM taGoods WHERE ItemID = ?", (search_id,))
                    result = cursor.fetchone()
                    conn.close()
                except Exception as e:
                    st.error(f"Error: {e}")
                    result = None
            
            if result:
                st.success(f"✅ Data ditemukan!")
                
                col_show1, col_show2 = st.columns(2)
                with col_show1:
                    st.info(f"**Item ID:** {result[0]}")
                with col_show2:
                    st.info(f"**Current Value:** {result[1] if result[1] else '(kosong)'}")
                
                new_value = st.text_input("Nilai baru", value=result[1] if result[1] else "", key="search_value")
                
                if st.button("Update", type="primary", key="search_btn"):
                    if new_value:
                        with st.spinner("💾 Mengupdate..."):
                            success, message = update_itemname2(search_id, new_value)
                            if success:
                                st.success(message)
                                st.balloons()
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.error(message)
                    else:
                        st.warning("Harap isi nilai baru!")
            else:
                st.error(f"❌ Item ID '{search_id}' tidak ditemukan!")
    
    # Footer
    st.divider()
    st.caption(f"🟢 Server: {SERVER} | Database: {DATABASE} | Total data: {get_total_count()} item")

if __name__ == "__main__":
    main()