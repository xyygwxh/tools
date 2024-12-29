import streamlit as st
import pandas as pd
import zipfile
import io


def main():
    st.subheader("提取员工证件照")
    photos_to_download = []
    employee_ids = []
    num = 0

    st.write("")
    with st.expander("使用方法："):
        st.write("本工具用于提取员工证件照，你需要：")
        st.write("1. 按照下表的格式上传需要提取的员工身份证号文件")
        st.write("2. 上传全体员工照片，注意照片的命名格式为“身份证号.jpg”")
        st.write("3. 点击提取照片按钮，系统会将要提取的照片打包")
        st.write("4. 点击下载照片按钮，系统会将照片打包下载")

        example_df = pd.DataFrame(
            {
                "身份证号": ["123456789012345678"],
                "姓名": ["张三"],
            }
        )
        st.dataframe(example_df, hide_index=True)

    employee_file = st.file_uploader(
        "上传员工身份证号文件", type=["xlsx", "xls", "csv"]
    )

    if employee_file is not None:
        employee_df = pd.read_excel(employee_file)
        try:
            employee_df = employee_df[["身份证号", "姓名"]]
            employee_ids = employee_df["身份证号"].tolist()
            total_num = len(employee_ids)
        except KeyError:
            st.error("请确保上传文件中有身份证号和姓名两列")
            st.stop()

    st.write("")
    employee_photos = st.file_uploader(
        "上传全体员工照片（以“身份证号.jpg”命名）",
        type=["jpg", "png", "jpeg"],
        accept_multiple_files=True,
    )

    if employee_photos is None:
        st.info("请上传照片")
    if st.button("提取照片"):
        for photo in employee_photos:
            if photo.name[:18] in employee_ids:
                num += 1
                photos_to_download.append(photo)
                employee_ids.remove(photo.name[:18])

        st.success(f"照片提取完毕，需提取照片数：{total_num} ，已提取照片数：{num}")
        if len(employee_ids) != 0:
            employee_unfound = employee_df[employee_df["身份证号"].isin(employee_ids)]
            st.info(f"未找到照片数：{len(employee_unfound)}")
            st.error("以下的人员照片未找到：")
            st.dataframe(employee_unfound)

        # 打包并下载 photos_to_download 中的文件
        def zip_files(files):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for file in files:
                    zipf.writestr(file.name, file.getbuffer().tobytes())
            return zip_buffer.getvalue()

        st.download_button(
            label="下载照片",
            data=zip_files(photos_to_download),
            file_name="照片.zip",
            mime="application/zip",
        )


if __name__ == "__main__":
    main()
