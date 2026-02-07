from typing import Any, Dict, Optional

import numpy as np
import pandas as pd
import streamlit as st


class ExcelAnalyzer:
    def __init__(self):
        self.dataframes: Dict[str, Dict[str, pd.DataFrame]] = {}
        self.file_info: Dict[str, Dict] = {}

    def load_excel_file(self, file_data, filename: str) -> bool:
        """Load Excel file and extract all sheets"""
        try:
            excel_file = pd.ExcelFile(file_data)
            sheets = {}

            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(file_data, sheet_name=sheet_name)
                sheets[sheet_name] = df

            self.dataframes[filename] = sheets
            self.file_info[filename] = {
                "sheets": list(sheets.keys()),
                "total_sheets": len(sheets),
                "total_rows": sum(len(df) for df in sheets.values()),
                "total_columns": sum(len(df.columns) for df in sheets.values()),
            }

            return True
        except Exception as e:
            st.error(f"Error loading {filename}: {str(e)}")
            return False

    def get_file_summary(self, filename: str) -> str:
        """Get summary information about a loaded file"""
        if filename not in self.file_info:
            return "File not found"

        info = self.file_info[filename]
        summary = f"""
**File: {filename}**
- Total Sheets: {info["total_sheets"]}
- Sheet Names: {", ".join(str(s) for s in info["sheets"])}
- Total Rows (across all sheets): {info["total_rows"]}
- Total Columns (across all sheets): {info["total_columns"]}
"""
        return summary

    def get_sheet_info(self, filename: str, sheet_name: str) -> str:
        """Get detailed information about a specific sheet"""
        if (
            filename not in self.dataframes
            or sheet_name not in self.dataframes[filename]
        ):
            return "Sheet not found"

        df = self.dataframes[filename][sheet_name]

        info = f"""
**Sheet: {sheet_name} (from {filename})**
- Shape: {df.shape[0]} rows × {df.shape[1]} columns
- Columns: {", ".join(str(c) for c in df.columns.tolist())}
- Data Types: {dict(df.dtypes)}
- Memory Usage: {df.memory_usage(deep=True).sum() / 1024**2:.2f} MB
"""

        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            info += f"\n\n**Numeric Columns Summary:**\n"
            for col in numeric_cols:
                info += f"- {col}: min={df[col].min():.2f}, max={df[col].max():.2f}, mean={df[col].mean():.2f}\n"

        return info

    # ── Comparison Methods ──────────────────────────────────────────────

    def compare_columns(
        self,
        file1: str,
        sheet1: str,
        file2: str,
        sheet2: str,
    ) -> Dict[str, Any]:
        """Compare column structure between two sheets."""
        df1 = self.dataframes[file1][sheet1]
        df2 = self.dataframes[file2][sheet2]

        cols1 = set(str(c) for c in df1.columns)
        cols2 = set(str(c) for c in df2.columns)

        return {
            "only_in_first": sorted(cols1 - cols2),
            "only_in_second": sorted(cols2 - cols1),
            "common": sorted(cols1 & cols2),
            "first_columns": [str(c) for c in df1.columns],
            "second_columns": [str(c) for c in df2.columns],
        }

    def compare_shape(
        self,
        file1: str,
        sheet1: str,
        file2: str,
        sheet2: str,
    ) -> Dict[str, Any]:
        """Compare shape / size of two sheets."""
        df1 = self.dataframes[file1][sheet1]
        df2 = self.dataframes[file2][sheet2]

        return {
            "first_shape": df1.shape,
            "second_shape": df2.shape,
            "row_diff": df1.shape[0] - df2.shape[0],
            "col_diff": df1.shape[1] - df2.shape[1],
        }

    def compare_cell_diff(
        self,
        file1: str,
        sheet1: str,
        file2: str,
        sheet2: str,
        key_column: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Deep cell-level comparison.

        If key_column is provided, rows are matched on that column.
        Otherwise rows are compared positionally (by index).
        """
        df1 = self.dataframes[file1][sheet1].copy()
        df2 = self.dataframes[file2][sheet2].copy()

        # Stringify column names for safety
        df1.columns = [str(c) for c in df1.columns]
        df2.columns = [str(c) for c in df2.columns]

        common_cols = sorted(set(df1.columns) & set(df2.columns))
        if not common_cols:
            return {"error": "No common columns to compare."}

        # ── Key-based comparison ────────────────────────────────────
        if key_column and key_column in common_cols:
            df1_keyed = df1.set_index(key_column, drop=False).sort_index()
            df2_keyed = df2.set_index(key_column, drop=False).sort_index()

            keys1 = set(df1_keyed.index)
            keys2 = set(df2_keyed.index)

            only_in_first_keys = sorted(keys1 - keys2, key=str)
            only_in_second_keys = sorted(keys2 - keys1, key=str)
            common_keys = sorted(keys1 & keys2, key=str)

            rows_only_first = df1_keyed.loc[
                df1_keyed.index.isin(only_in_first_keys)
            ].reset_index(drop=True)
            rows_only_second = df2_keyed.loc[
                df2_keyed.index.isin(only_in_second_keys)
            ].reset_index(drop=True)

            # Cell-level changes on common keys & common columns
            compare_cols = [c for c in common_cols if c != key_column]
            changes = []
            for key in common_keys:
                row1 = df1_keyed.loc[[key]]
                row2 = df2_keyed.loc[[key]]
                # Take first occurrence if duplicate keys
                r1 = row1.iloc[0]
                r2 = row2.iloc[0]
                for col in compare_cols:
                    v1 = r1.get(col)
                    v2 = r2.get(col)
                    # Treat NaN == NaN as equal
                    both_nan = (
                        pd.isna(v1) and pd.isna(v2)
                        if not isinstance(v1, str) and not isinstance(v2, str)
                        else False
                    )
                    if not both_nan and str(v1) != str(v2):
                        changes.append(
                            {
                                key_column: key,
                                "Column": col,
                                "File 1 Value": v1,
                                "File 2 Value": v2,
                            }
                        )

            return {
                "mode": "key",
                "key_column": key_column,
                "rows_only_in_first": rows_only_first,
                "rows_only_in_second": rows_only_second,
                "cell_changes": pd.DataFrame(changes) if changes else pd.DataFrame(),
                "total_changes": len(changes),
                "common_keys_count": len(common_keys),
                "only_first_count": len(only_in_first_keys),
                "only_second_count": len(only_in_second_keys),
            }

        # ── Positional (index-based) comparison ────────────────────
        else:
            min_rows = min(len(df1), len(df2))
            df1_cmp = df1[common_cols].iloc[:min_rows].reset_index(drop=True)
            df2_cmp = df2[common_cols].iloc[:min_rows].reset_index(drop=True)

            # Cast both to string to compare safely
            diff_mask = df1_cmp.astype(str) != df2_cmp.astype(str)

            changes = []
            for row_idx in range(min_rows):
                for col in common_cols:
                    if diff_mask.at[row_idx, col]:
                        changes.append(
                            {
                                "Row": row_idx,
                                "Column": col,
                                "File 1 Value": df1_cmp.at[row_idx, col],
                                "File 2 Value": df2_cmp.at[row_idx, col],
                            }
                        )

            extra_first = (
                df1[common_cols].iloc[min_rows:].reset_index(drop=True)
                if len(df1) > min_rows
                else pd.DataFrame()
            )
            extra_second = (
                df2[common_cols].iloc[min_rows:].reset_index(drop=True)
                if len(df2) > min_rows
                else pd.DataFrame()
            )

            return {
                "mode": "positional",
                "rows_compared": min_rows,
                "cell_changes": pd.DataFrame(changes) if changes else pd.DataFrame(),
                "total_changes": len(changes),
                "extra_rows_in_first": extra_first,
                "extra_rows_in_second": extra_second,
            }

    def compare_stats(
        self,
        file1: str,
        sheet1: str,
        file2: str,
        sheet2: str,
    ) -> Optional[pd.DataFrame]:
        """Compare descriptive stats for common numeric columns."""
        df1 = self.dataframes[file1][sheet1].copy()
        df2 = self.dataframes[file2][sheet2].copy()
        df1.columns = [str(c) for c in df1.columns]
        df2.columns = [str(c) for c in df2.columns]

        num1 = set(df1.select_dtypes(include=[np.number]).columns)
        num2 = set(df2.select_dtypes(include=[np.number]).columns)
        common_num = sorted(num1 & num2)

        if not common_num:
            return None

        rows = []
        for col in common_num:
            rows.append(
                {
                    "Column": col,
                    "File1 Mean": df1[col].mean(),
                    "File2 Mean": df2[col].mean(),
                    "Mean Diff": df1[col].mean() - df2[col].mean(),
                    "File1 Sum": df1[col].sum(),
                    "File2 Sum": df2[col].sum(),
                    "Sum Diff": df1[col].sum() - df2[col].sum(),
                    "File1 Min": df1[col].min(),
                    "File2 Min": df2[col].min(),
                    "File1 Max": df1[col].max(),
                    "File2 Max": df2[col].max(),
                }
            )
        return pd.DataFrame(rows)

    # ── Q&A Methods ─────────────────────────────────────────────────

    def answer_question(self, question: str) -> str:
        """Answer questions about the loaded data."""
        question_lower = question.lower()

        if not self.dataframes:
            return "No Excel files loaded. Please upload files first."

        try:
            if "how many" in question_lower and (
                "rows" in question_lower or "records" in question_lower
            ):
                return self._count_rows_question()
            elif "what columns" in question_lower or "column names" in question_lower:
                return self._list_columns_question()
            elif "summary" in question_lower or "describe" in question_lower:
                return self._summary_question()
            elif (
                "compare" in question_lower
                or "diff" in question_lower
                or "difference" in question_lower
            ):
                return self._compare_question_text()
            elif "average" in question_lower or "mean" in question_lower:
                return self._average_question()
            elif "maximum" in question_lower or "max" in question_lower:
                return self._max_question()
            elif "minimum" in question_lower or "min" in question_lower:
                return self._min_question()
            elif (
                "null" in question_lower
                or "missing" in question_lower
                or "empty" in question_lower
            ):
                return self._nulls_question()
            elif "duplicate" in question_lower:
                return self._duplicates_question()
            else:
                return self._general_question_help()

        except Exception as e:
            return f"Error processing question: {str(e)}"

    def _count_rows_question(self) -> str:
        result = "📊 **Row Counts:**\n\n"
        for filename, sheets in self.dataframes.items():
            result += f"**{filename}:**\n"
            for sheet_name, df in sheets.items():
                result += f"  - {sheet_name}: {len(df)} rows\n"
            result += "\n"
        return result

    def _list_columns_question(self) -> str:
        result = "📋 **Column Information:**\n\n"
        for filename, sheets in self.dataframes.items():
            result += f"**{filename}:**\n"
            for sheet_name, df in sheets.items():
                result += f"  - **{sheet_name}**: {', '.join(str(c) for c in df.columns.tolist())}\n"
            result += "\n"
        return result

    def _summary_question(self) -> str:
        result = "📈 **Data Summary:**\n\n"
        for filename, sheets in self.dataframes.items():
            result += f"**{filename}:**\n"
            for sheet_name, df in sheets.items():
                numeric_df = df.select_dtypes(include=[np.number])
                if len(numeric_df.columns) > 0:
                    result += f"  **{sheet_name}** (Numeric columns):\n"
                    stats = numeric_df.describe()
                    result += f"```\n{stats.to_string()}\n```\n\n"
                else:
                    result += f"  **{sheet_name}**: No numeric columns\n\n"
        return result

    def _compare_question_text(self) -> str:
        if len(self.dataframes) < 2:
            return "⚠️ Need at least 2 files to compare. Please upload more files.\n\nUse the **🔄 Compare Files** tab above for a detailed comparison."

        result = "🔄 **File Comparison Overview:**\n\n"
        filenames = list(self.dataframes.keys())

        for filename in filenames:
            info = self.file_info[filename]
            result += f"**{filename}:**\n"
            result += f"  - Sheets: {info['total_sheets']} — {', '.join(str(s) for s in info['sheets'])}\n"
            result += f"  - Total Rows: {info['total_rows']}\n"
            result += f"  - Total Columns: {info['total_columns']}\n\n"

        # Common sheets
        all_sheet_sets = [set(info["sheets"]) for info in self.file_info.values()]
        common_sheets = set.intersection(*all_sheet_sets) if all_sheet_sets else set()

        if common_sheets:
            result += f"**Common Sheets:** {', '.join(str(s) for s in sorted(common_sheets))}\n\n"
            for sheet in sorted(common_sheets):
                result += f"  📄 **{sheet}:**\n"
                for fn in filenames:
                    df = self.dataframes[fn][sheet]
                    result += (
                        f"    - {fn}: {df.shape[0]} rows × {df.shape[1]} columns\n"
                    )
                result += "\n"
        else:
            result += "⚠️ No common sheet names found between files.\n\n"

        only_sheets = {}
        for fn in filenames:
            unique_to_file = set(self.file_info[fn]["sheets"]) - common_sheets
            if unique_to_file:
                only_sheets[fn] = sorted(unique_to_file)
        if only_sheets:
            result += "**Sheets unique to one file:**\n"
            for fn, sheets in only_sheets.items():
                result += f"  - {fn}: {', '.join(str(s) for s in sheets)}\n"

        result += "\n👉 **For a detailed cell-level diff, use the 🔄 Compare Files tab above.**"
        return result

    def _average_question(self) -> str:
        result = "📊 **Average Values:**\n\n"
        for filename, sheets in self.dataframes.items():
            result += f"**{filename}:**\n"
            for sheet_name, df in sheets.items():
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    result += f"  **{sheet_name}:**\n"
                    for col in numeric_cols:
                        result += f"    - {col}: {df[col].mean():.2f}\n"
                result += "\n"
        return result

    def _max_question(self) -> str:
        result = "📈 **Maximum Values:**\n\n"
        for filename, sheets in self.dataframes.items():
            result += f"**{filename}:**\n"
            for sheet_name, df in sheets.items():
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    result += f"  **{sheet_name}:**\n"
                    for col in numeric_cols:
                        result += f"    - {col}: {df[col].max()}\n"
                result += "\n"
        return result

    def _min_question(self) -> str:
        result = "📉 **Minimum Values:**\n\n"
        for filename, sheets in self.dataframes.items():
            result += f"**{filename}:**\n"
            for sheet_name, df in sheets.items():
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    result += f"  **{sheet_name}:**\n"
                    for col in numeric_cols:
                        result += f"    - {col}: {df[col].min()}\n"
                result += "\n"
        return result

    def _nulls_question(self) -> str:
        result = "🕳️ **Null / Missing Values:**\n\n"
        for filename, sheets in self.dataframes.items():
            result += f"**{filename}:**\n"
            for sheet_name, df in sheets.items():
                nulls = df.isnull().sum()
                has_nulls = nulls[nulls > 0]
                if len(has_nulls) > 0:
                    result += f"  **{sheet_name}:**\n"
                    for col, count in has_nulls.items():
                        pct = (count / len(df)) * 100
                        result += f"    - {col}: {count} missing ({pct:.1f}%)\n"
                else:
                    result += f"  **{sheet_name}**: No missing values ✅\n"
                result += "\n"
        return result

    def _duplicates_question(self) -> str:
        result = "🔁 **Duplicate Rows:**\n\n"
        for filename, sheets in self.dataframes.items():
            result += f"**{filename}:**\n"
            for sheet_name, df in sheets.items():
                dup_count = df.duplicated().sum()
                result += f"  - **{sheet_name}**: {dup_count} duplicate rows"
                if dup_count > 0:
                    result += " ⚠️"
                else:
                    result += " ✅"
                result += "\n"
            result += "\n"
        return result

    def _general_question_help(self) -> str:
        return """
🤔 **I can help you with questions like:**

📊 **Data Overview:**
- "How many rows are in each sheet?"
- "What columns are available?"
- "Give me a summary of the data"

🔢 **Statistics:**
- "What's the average of the numeric columns?"
- "What's the maximum value?"
- "Show me the minimum values"

🕳️ **Data Quality:**
- "Are there any null or missing values?"
- "Are there any duplicate rows?"

🔄 **Comparisons:**
- "Compare the two files"
- "What's the difference between the files?"

👉 For detailed cell-level diffs, use the **🔄 Compare Files** tab above.
"""

    def get_dataframe(self, filename: str, sheet_name: str) -> Optional[pd.DataFrame]:
        """Get a specific dataframe for direct manipulation"""
        if filename in self.dataframes and sheet_name in self.dataframes[filename]:
            return self.dataframes[filename][sheet_name]
        return None


# ═══════════════════════════════════════════════════════════════════════
# Streamlit UI
# ═══════════════════════════════════════════════════════════════════════


def render_sidebar(analyzer: ExcelAnalyzer):
    """Sidebar: file upload."""
    with st.sidebar:
        st.header("📁 Upload Excel Files")

        uploaded_files = st.file_uploader(
            "Choose Excel files",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="Upload one or more Excel files to analyze",
        )

        if uploaded_files:
            for uploaded_file in uploaded_files:
                if uploaded_file.name not in analyzer.dataframes:
                    with st.spinner(f"Loading {uploaded_file.name}..."):
                        success = analyzer.load_excel_file(
                            uploaded_file, uploaded_file.name
                        )
                        if success:
                            st.success(f"✅ Loaded {uploaded_file.name}")

        if analyzer.dataframes:
            st.divider()
            st.markdown("**Loaded files:**")
            for fname, info in analyzer.file_info.items():
                st.markdown(
                    f"- **{fname}** — {info['total_sheets']} sheet(s), "
                    f"{info['total_rows']} rows"
                )


def render_explore_tab(analyzer: ExcelAnalyzer):
    """Tab 1: Explore individual files."""
    if not analyzer.dataframes:
        st.info("👆 Upload Excel files using the sidebar to get started!")
        return

    for filename in analyzer.dataframes.keys():
        with st.expander(f"📋 {filename}", expanded=False):
            st.markdown(analyzer.get_file_summary(filename))
            for sheet_name in analyzer.dataframes[filename].keys():
                st.subheader(f"Sheet: {sheet_name}")
                st.markdown(analyzer.get_sheet_info(filename, sheet_name))
                df = analyzer.get_dataframe(filename, sheet_name)
                if df is not None:
                    st.dataframe(df.head(20), use_container_width=True)


def render_compare_tab(analyzer: ExcelAnalyzer):
    """Tab 2: Deep comparison between two sheets (possibly from different files)."""
    filenames = list(analyzer.dataframes.keys())

    if len(filenames) < 2:
        st.warning("Upload at least **2 Excel files** to use the comparison tool.")
        return

    st.markdown("Pick the two sheets you want to compare:")

    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("##### 📄 First")
        file1 = st.selectbox("File", filenames, key="cmp_file1")
        sheets1 = list(analyzer.dataframes[file1].keys())
        sheet1 = st.selectbox("Sheet", sheets1, key="cmp_sheet1")

    with col_right:
        st.markdown("##### 📄 Second")
        file2 = st.selectbox(
            "File", filenames, index=min(1, len(filenames) - 1), key="cmp_file2"
        )
        sheets2 = list(analyzer.dataframes[file2].keys())
        sheet2 = st.selectbox("Sheet", sheets2, key="cmp_sheet2")

    # Key column (optional)
    df1 = analyzer.dataframes[file1][sheet1]
    df2 = analyzer.dataframes[file2][sheet2]
    cols1 = [str(c) for c in df1.columns]
    cols2 = [str(c) for c in df2.columns]
    common_cols = sorted(set(cols1) & set(cols2))

    key_column = None
    if common_cols:
        use_key = st.checkbox(
            "Match rows by a key column (e.g. ID) instead of row position",
            value=False,
        )
        if use_key:
            key_column = st.selectbox("Key column", common_cols, key="cmp_key")

    if st.button("🔍 Run Comparison", type="primary"):
        _run_comparison(analyzer, file1, sheet1, file2, sheet2, key_column)


def _run_comparison(
    analyzer: ExcelAnalyzer,
    file1: str,
    sheet1: str,
    file2: str,
    sheet2: str,
    key_column: Optional[str],
):
    """Execute and render a full comparison."""
    label1 = f"{file1} → {sheet1}"
    label2 = f"{file2} → {sheet2}"

    # ── 1. Shape ────────────────────────────────────────────────────
    st.subheader("📐 Shape Comparison")
    shape = analyzer.compare_shape(file1, sheet1, file2, sheet2)
    c1, c2, c3 = st.columns(3)
    c1.metric(
        label1, f"{shape['first_shape'][0]} rows × {shape['first_shape'][1]} cols"
    )
    c2.metric(
        label2, f"{shape['second_shape'][0]} rows × {shape['second_shape'][1]} cols"
    )
    c3.metric(
        "Difference", f"{abs(shape['row_diff'])} rows, {abs(shape['col_diff'])} cols"
    )

    # ── 2. Column diff ──────────────────────────────────────────────
    st.subheader("📋 Column Comparison")
    col_cmp = analyzer.compare_columns(file1, sheet1, file2, sheet2)

    if col_cmp["only_in_first"]:
        st.warning(
            f"**Columns only in {label1}:** {', '.join(col_cmp['only_in_first'])}"
        )
    if col_cmp["only_in_second"]:
        st.warning(
            f"**Columns only in {label2}:** {', '.join(col_cmp['only_in_second'])}"
        )
    if col_cmp["common"]:
        st.success(
            f"**Common columns ({len(col_cmp['common'])}):** {', '.join(col_cmp['common'])}"
        )
    if not col_cmp["only_in_first"] and not col_cmp["only_in_second"]:
        st.success("✅ Both sheets have identical column structures.")

    # ── 3. Stats comparison ─────────────────────────────────────────
    stats_df = analyzer.compare_stats(file1, sheet1, file2, sheet2)
    if stats_df is not None and len(stats_df) > 0:
        st.subheader("📊 Numeric Stats Comparison")
        st.dataframe(stats_df, use_container_width=True)

    # ── 4. Cell-level diff ──────────────────────────────────────────
    st.subheader("🔬 Cell-Level Differences")

    if not col_cmp["common"]:
        st.error("No common columns — cannot compare cells.")
        return

    with st.spinner("Computing cell-level differences..."):
        result = analyzer.compare_cell_diff(file1, sheet1, file2, sheet2, key_column)

    if "error" in result:
        st.error(result["error"])
        return

    if result["mode"] == "key":
        st.markdown(f"**Matched on key column:** `{result['key_column']}`")

        m1, m2, m3 = st.columns(3)
        m1.metric("Common Keys", result["common_keys_count"])
        m2.metric(f"Only in {label1}", result["only_first_count"])
        m3.metric(f"Only in {label2}", result["only_second_count"])

        if result["total_changes"] > 0:
            st.warning(
                f"⚠️ **{result['total_changes']} cell value(s) differ** across common keys."
            )
            st.dataframe(result["cell_changes"], use_container_width=True)

            # Download button for changes
            csv = result["cell_changes"].to_csv(index=False)
            st.download_button(
                "⬇️ Download changes as CSV",
                csv,
                "cell_changes.csv",
                "text/csv",
            )
        else:
            st.success("✅ All cell values match for common keys and common columns!")

        if len(result["rows_only_in_first"]) > 0:
            with st.expander(
                f"Rows only in {label1} ({len(result['rows_only_in_first'])})"
            ):
                st.dataframe(result["rows_only_in_first"], use_container_width=True)

        if len(result["rows_only_in_second"]) > 0:
            with st.expander(
                f"Rows only in {label2} ({len(result['rows_only_in_second'])})"
            ):
                st.dataframe(result["rows_only_in_second"], use_container_width=True)

    else:
        # Positional mode
        st.markdown(
            f"**Compared positionally** (row by row) — {result['rows_compared']} rows compared."
        )

        if result["total_changes"] > 0:
            st.warning(f"⚠️ **{result['total_changes']} cell value(s) differ.**")
            st.dataframe(result["cell_changes"], use_container_width=True)

            csv = result["cell_changes"].to_csv(index=False)
            st.download_button(
                "⬇️ Download changes as CSV",
                csv,
                "cell_changes.csv",
                "text/csv",
            )
        else:
            st.success("✅ All compared cells are identical!")

        if len(result["extra_rows_in_first"]) > 0:
            with st.expander(
                f"Extra rows in {label1} ({len(result['extra_rows_in_first'])})"
            ):
                st.dataframe(result["extra_rows_in_first"], use_container_width=True)

        if len(result["extra_rows_in_second"]) > 0:
            with st.expander(
                f"Extra rows in {label2} ({len(result['extra_rows_in_second'])})"
            ):
                st.dataframe(result["extra_rows_in_second"], use_container_width=True)


def render_questions_tab(analyzer: ExcelAnalyzer):
    """Tab 3: Ask questions about your data."""
    if not analyzer.dataframes:
        st.info("👆 Upload Excel files using the sidebar to get started!")
        return

    question = st.text_input(
        "What would you like to know?",
        placeholder="e.g., How many rows are in each sheet? What's the average? Compare the files.",
        help="Ask natural language questions about your Excel data",
    )

    if st.button("🔍 Get Answer", type="primary"):
        if question:
            with st.spinner("Analyzing..."):
                answer = analyzer.answer_question(question)
                st.markdown("### 💡 Answer:")
                st.markdown(answer)
        else:
            st.warning("Please enter a question first!")

    # Quick question buttons
    st.subheader("🚀 Quick Questions")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        if st.button("📊 Row Counts"):
            st.markdown(analyzer.answer_question("How many rows in each sheet?"))

    with col2:
        if st.button("📋 Column Names"):
            st.markdown(analyzer.answer_question("What columns are available?"))

    with col3:
        if st.button("📈 Summary Stats"):
            st.markdown(analyzer.answer_question("Give me a summary of the data"))

    with col4:
        if st.button("🕳️ Missing Values"):
            st.markdown(analyzer.answer_question("Are there any null values?"))

    col5, col6, col7 = st.columns(3)

    with col5:
        if st.button("📊 Averages"):
            st.markdown(analyzer.answer_question("What's the average?"))

    with col6:
        if st.button("🔁 Duplicates"):
            st.markdown(analyzer.answer_question("Are there any duplicates?"))

    with col7:
        if st.button("🔄 Compare Files"):
            st.markdown(analyzer.answer_question("Compare the two files"))


# ═══════════════════════════════════════════════════════════════════════
# Main
# ═══════════════════════════════════════════════════════════════════════


def main():
    st.set_page_config(page_title="Excel Analyzer", page_icon="🔍", layout="wide")

    st.title("🔍 Excel File Analyzer")
    st.markdown("Upload Excel files, explore data, compare sheets, and ask questions!")

    # Initialize analyzer in session state
    if "analyzer" not in st.session_state:
        st.session_state.analyzer = ExcelAnalyzer()

    analyzer = st.session_state.analyzer

    # Render sidebar (file upload)
    render_sidebar(analyzer)

    # Tabs
    tab_explore, tab_compare, tab_questions = st.tabs(
        ["📊 Explore Data", "🔄 Compare Files", "❓ Ask Questions"]
    )

    with tab_explore:
        render_explore_tab(analyzer)

    with tab_compare:
        render_compare_tab(analyzer)

    with tab_questions:
        render_questions_tab(analyzer)


if __name__ == "__main__":
    main()
