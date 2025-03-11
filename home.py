import streamlit as st


def main():
    st.subheader("欢迎使用我的小工具")
    st.write("")
    cols = st.columns(3)
    with cols[0]:
        st.write("")
        st.page_link("pages/1-提取图片.py", label="提取图片", icon="1️⃣")
    with cols[1]:
        st.write("")
        st.page_link("pages/2-表格工具.py", label="表格工具", icon="2️⃣")
    with cols[2]:
        st.write("")
        st.page_link("pages/3-节约打印.py", label="节约打印", icon="3️⃣")


if __name__ == "__main__":
    main()
