import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(
    page_title="Facebook Post Tracker",
    page_icon="📊",
    layout="wide"
)

# ====================== SESSION STATE ======================
if "profiles" not in st.session_state:
    st.session_state.profiles = []

if "df_posts" not in st.session_state:
    st.session_state.df_posts = None

# ====================== SIDEBAR ======================
st.sidebar.title("📊 Facebook Post Tracker")
st.sidebar.markdown("---")

menu = st.sidebar.radio(
    "Menu",
    ["1. Quản lý hồ sơ", "2. Nhập dữ liệu", "3. Thống kê & Cảnh báo", "4. Xuất báo cáo"]
)

st.sidebar.markdown("---")
st.sidebar.info(
    "Công cụ chỉ xử lý dữ liệu bạn đã thu thập hợp lệ.\n"
    "Không tự truy cập Facebook."
)

# ====================== TAB 1: QUẢN LÝ HỒ SƠ ======================
if menu == "1. Quản lý hồ sơ":
    st.header("1. Danh sách hồ sơ cần theo dõi")

    col1, col2 = st.columns([3, 1])
    with col1:
        new_profile = st.text_input("Thêm hồ sơ mới", placeholder="Nhập tên hồ sơ Facebook...")
    with col2:
        st.write("")
        st.write("")
        if st.button("Thêm", use_container_width=True):
            if new_profile.strip() and new_profile.strip() not in st.session_state.profiles:
                st.session_state.profiles.append(new_profile.strip())
                st.success(f"Đã thêm: {new_profile.strip()}")
            elif new_profile.strip() in st.session_state.profiles:
                st.warning("Hồ sơ này đã có trong danh sách")

    st.subheader("Danh sách hiện tại")
    if st.session_state.profiles:
        for i, profile in enumerate(st.session_state.profiles):
            col_a, col_b = st.columns([5, 1])
            with col_a:
                st.write(f"{i+1}. {profile}")
            with col_b:
                if st.button("Xóa", key=f"del_{i}"):
                    st.session_state.profiles.pop(i)
                    st.rerun()
    else:
        st.info("Chưa có hồ sơ nào. Hãy thêm ở trên.")

# ====================== TAB 2: NHẬP DỮ LIỆU ======================
elif menu == "2. Nhập dữ liệu":
    st.header("2. Nhập dữ liệu bài đăng")

    st.markdown("""
    **Định dạng file khuyến nghị (Excel hoặc CSV):**
    | profile_name | author | content | post_time | url | post_type | topic |
    |--------------|--------|---------|-----------|-----|-----------|-------|
    """)

    uploaded_file = st.file_uploader("Upload file Excel (.xlsx) hoặc CSV", type=["xlsx", "csv"])

    if uploaded_file:
        try:
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            # Chuẩn hóa tên cột
            df.columns = [c.strip().lower() for c in df.columns]

            required_cols = ["profile_name", "author", "content", "post_time"]
            missing = [c for c in required_cols if c not in df.columns]
            if missing:
                st.error(f"Thiếu cột bắt buộc: {', '.join(missing)}")
            else:
                # Xử lý thời gian
                df["post_time"] = pd.to_datetime(df["post_time"], dayfirst=True, errors="coerce")
                df["month"] = df["post_time"].dt.to_period("M").astype(str)

                # Các cột tùy chọn
                if "url" not in df.columns:
                    df["url"] = ""
                if "post_type" not in df.columns:
                    df["post_type"] = ""
                if "topic" not in df.columns:
                    df["topic"] = "Chưa phân loại"

                st.session_state.df_posts = df
                st.success(f"Đã tải thành công {len(df)} bài đăng!")
                st.dataframe(df.head(10), use_container_width=True)

        except Exception as e:
            st.error(f"Lỗi đọc file: {e}")

# ====================== TAB 3: THỐNG KÊ & CẢNH BÁO ======================
elif menu == "3. Thống kê & Cảnh báo":
    st.header("3. Thống kê & Cảnh báo")

    if st.session_state.df_posts is None:
        st.warning("Chưa có dữ liệu. Hãy vào tab 'Nhập dữ liệu' trước.")
    else:
        df = st.session_state.df_posts.copy()

        # --- Thống kê ---
        st.subheader("📈 Thống kê tổng quan")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Tổng số bài", len(df))
        col2.metric("Số hồ sơ", df["profile_name"].nunique())
        col3.metric("Số chủ đề", df["topic"].nunique())
        col4.metric("Bài thiếu URL", df["url"].isna().sum() + (df["url"] == "").sum())

        st.markdown("#### Đếm theo hồ sơ")
        st.dataframe(
            df.groupby("profile_name").size().reset_index(name="Số bài").sort_values("Số bài", ascending=False),
            use_container_width=True
        )

        st.markdown("#### Đếm theo tháng")
        st.dataframe(
            df.groupby("month").size().reset_index(name="Số bài").sort_values("month"),
            use_container_width=True
        )

        st.markdown("#### Đếm theo chủ đề")
        st.dataframe(
            df.groupby("topic").size().reset_index(name="Số bài").sort_values("Số bài", ascending=False),
            use_container_width=True
        )

        # --- Cảnh báo ---
        st.subheader("⚠️ Cảnh báo")

        # Bài thiếu URL
        missing_url = df[(df["url"].isna()) | (df["url"] == "")]
        st.markdown(f"**Bài thiếu URL:** {len(missing_url)} bài")
        if len(missing_url) > 0:
            st.dataframe(missing_url[["profile_name", "author", "post_time", "content"]], use_container_width=True)

        # Bài cần xác minh (thiếu nội dung hoặc thời gian)
        need_verify = df[df["content"].isna() | (df["content"] == "") | df["post_time"].isna()]
        st.markdown(f"**Bài cần xác minh:** {len(need_verify)} bài")
        if len(need_verify) > 0:
            st.dataframe(need_verify, use_container_width=True)

        # Bài trùng (dựa trên nội dung giống nhau)
        df["content_clean"] = df["content"].astype(str).str.strip().str.lower()
        duplicates = df[df.duplicated(subset=["content_clean"], keep=False)].sort_values("content_clean")
        st.markdown(f"**Bài trùng nội dung:** {len(duplicates)} bài")
        if len(duplicates) > 0:
            st.dataframe(duplicates[["profile_name", "author", "post_time", "content"]], use_container_width=True)

# ====================== TAB 4: XUẤT BÁO CÁO ======================
elif menu == "4. Xuất báo cáo":
    st.header("4. Xuất báo cáo Excel")

    if st.session_state.df_posts is None:
        st.warning("Chưa có dữ liệu. Hãy vào tab 'Nhập dữ liệu' trước.")
    else:
        df = st.session_state.df_posts.copy()

        if st.button("Tạo file Excel báo cáo", type="primary"):
            output = io.BytesIO()

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                # Sheet 1: Tổng hợp
                summary_profile = df.groupby("profile_name").size().reset_index(name="Số bài")
                summary_month = df.groupby("month").size().reset_index(name="Số bài")
                summary_topic = df.groupby("topic").size().reset_index(name="Số bài")

                summary_profile.to_excel(writer, sheet_name="Tổng hợp", index=False, startrow=0)
                summary_month.to_excel(writer, sheet_name="Tổng hợp", index=False, startrow=len(summary_profile) + 3)
                summary_topic.to_excel(writer, sheet_name="Tổng hợp", index=False, startrow=len(summary_profile) + len(summary_month) + 6)

                # Sheet 2: Chi tiết
                detail_cols = ["profile_name", "author", "post_time", "post_type", "topic", "url", "content"]
                detail_cols = [c for c in detail_cols if c in df.columns]
                df[detail_cols].sort_values("post_time", ascending=False).to_excel(
                    writer, sheet_name="Chi tiết bài đăng", index=False
                )

            output.seek(0)
            st.download_button(
                label="📥 Tải file Excel",
                data=output,
                file_name=f"bao_cao_facebook_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Đã tạo file thành công!")
