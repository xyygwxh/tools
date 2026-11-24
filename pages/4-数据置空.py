# pyright: reportMissingTypeStubs=false
# 项目未安装 pandas 类型存根包；如需完整类型检查，可另行安装 pandas-stubs
import io
from typing import cast

import streamlit as st
import pandas as pd


DEFAULT_SCORE_COLS = ["总分", "语文", "数学", "英语", "物历", "化政", "生地"]
REQUIRED_COLS = ["班级类别", "班级"]


def main():
    _ = st.title("数据置空")
    st.write("本工具将用于将Excel表格中的特定数据按要求置空")

    # 通过交互控件接收置空规则：班级类别及每班置空人数
    _ = st.subheader("置空规则")
    col1, col2 = st.columns(2, gap="large")
    with col1:
        category1 = st.text_input("班级类别 1", value="宏志班").strip()
        count1 = st.number_input("每班置空人数", min_value=1, value=3, key="count1")
    with col2:
        category2 = st.text_input("班级类别 2", value="励志班").strip()
        count2 = st.number_input("每班置空人数", min_value=1, value=5, key="count2")

    # 忽略留空的类别；同名类别只保留一次，避免重复规则导致超额置空
    rule_map = {c: n for c, n in ((category1, count1), (category2, count2)) if c}
    rules = list(rule_map.items())

    # 读取Excel表格
    file = st.file_uploader("请上传成绩表", type=["xlsx"])
    if file is not None:
        df = pd.read_excel(file)  # pyright: ignore[reportUnknownMemberType]

        # 校验分组所需列，避免运行时 KeyError
        missing = [c for c in REQUIRED_COLS if c not in df.columns]
        if missing:
            _ = st.error(f"表格缺少必需列：{', '.join(missing)}")
            st.stop()

        # 学科列用复选框选择，默认勾选常用科目；选项来自实际列名，避免列名不一致报错
        st.write("选择要置空的成绩列")
        candidate_cols = [c for c in list(df.columns) if c not in REQUIRED_COLS]
        score_cols: list[str] = []
        checkbox_cols = st.columns(4)
        for i, col in enumerate(candidate_cols):
            with checkbox_cols[i % len(checkbox_cols)]:
                if st.checkbox(
                    str(col),
                    value=col in DEFAULT_SCORE_COLS,
                    key=f"score_{col}",
                ):
                    score_cols.append(col)

        if st.button("置空数据"):
            if not rules:
                _ = st.error("请至少填写一个班级类别")
                st.stop()
            if not score_cols:
                _ = st.error("请至少选择一列要置空的成绩列")
                st.stop()

            result = df.copy()

            # 一次 groupby 完成所有学科排名；并列时按出现顺序只取指定人数
            for category, n in rules:
                is_target = result["班级类别"] == category
                subset = result.loc[is_target, score_cols]
                ranks = subset.groupby(result.loc[is_target, "班级"])[score_cols].rank(
                    method="first", ascending=True
                )
                result.loc[is_target, score_cols] = subset.mask(ranks.le(n))

            buffer = io.BytesIO()
            result.to_excel(buffer, index=False, engine="openpyxl")  # pyright: ignore[reportUnknownMemberType]
            _ = st.success("处理完成，请下载结果表格")
            _ = st.download_button(
                label="下载置空后的表格",
                data=buffer.getvalue(),
                file_name=cast(str, file.name),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()
