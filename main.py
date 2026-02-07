import io
from pathlib import Path
from typing import Any, Dict, List

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
            # Read all sheets from the Excel file
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
        📊 **File: {filename}**
        - Total Sheets: {info["total_sheets"]}
        - Sheet Names: {", ".join(info["sheets"])}
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
        📈 **Sheet: {sheet_name} (from {filename})**
        - Shape: {df.shape[0]} rows × {df.shape[1]} columns
        - Columns: {", ".join(df.columns.tolist())}
        - Data Types: {dict(df.dtypes)}
        - Memory Usage: {df.memory_usage(deep=True).sum() / 1024**2:.2f} MB
        """

        # Add basic statistics for numeric columns
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        if len(numeric_cols) > 0:
            info += f"\n\n**Numeric Columns Summary:**\n"
            for col in numeric_cols:
                info += f"- {col}: min={df[col].min():.2f}, max={df[col].max():.2f}, mean={df[col].mean():.2f}\n"

        return info

    def answer_question(self, question: str) -> str:
        """Answer questions about the loaded data using natural language processing"""
        question_lower = question.lower()

        if not self.dataframes:
            return "No Excel files loaded. Please upload files first."

        try:
            # Simple question matching - you can enhance this with more sophisticated NLP
            if "how many" in question_lower and (
                "rows" in question_lower or "records" in question_lower
            ):
                return self._count_rows_question(question)
            elif "what columns" in question_lower or "column names" in question_lower:
                return self._list_columns_question(question)
            elif "summary" in question_lower or "describe" in question_lower:
                return self._summary_question(question)
            elif "compare" in question_lower:
                return self._compare_question(question)
            elif "filter" in question_lower or "where" in question_lower:
                return self._filter_question(question)
            elif "average" in question_lower or "mean" in question_lower:
                return self._average_question(question)
            elif "maximum" in question_lower or "max" in question_lower:
                return self._max_question(question)
            elif "minimum" in question_lower or "min" in question_lower:
                return self._min_question(question)
            else:
                return self._general_question_help()

        except Exception as e:
            return f"Error processing question: {str(e)}"

    def _count_rows_question(self, question: str) -> str:
        """Answer questions about row counts"""
        result = "📊 **Row Counts:**\n\n"
        for filename, sheets in self.dataframes.items():
            result += f"**{filename}:**\n"
            for sheet_name, df in sheets.items():
                result += f"  - {sheet_name}: {len(df)} rows\n"
            result += "\n"
        return result

    def _list_columns_question(self, question: str) -> str:
        """Answer questions about columns"""
        result = "📋 **Column Information:**\n\n"
        for filename, sheets in self.dataframes.items():
            result += f"**{filename}:**\n"
            for sheet_name, df in sheets.items():
                result += f"  - **{sheet_name}**: {', '.join(df.columns.tolist())}\n"
            result += "\n"
        return result

    def _summary_question(self, question: str) -> str:
        """Provide summary statistics"""
        result = "📈 **Data Summary:**\n\n"
        for filename, sheets in self.dataframes.items():
            result += f"**{filename}:**\n"
            for sheet_name, df in sheets.items():
                numeric_df = df.select_dtypes(include=[np.number])
                if len(numeric_df.columns) > 0:
                    result += f"  **{sheet_name}** (Numeric columns):\n"
                    stats = numeric_df.describe()
                    result += f"```\n{stats.to_string()}\n```\n\n"
        return result

    def _compare_question(self, question: str) -> str:
        """Compare data between files or sheets"""
        if len(self.dataframes) < 2:
            return "Need at least 2 files to compare. Please upload more files."

        result = "🔄 **Comparison:**\n\n"
        filenames = list(self.dataframes.keys())

        for i, filename in enumerate(filenames):
            info = self.file_info[filename]
            result += f"**{filename}:**\n"
            result += f"  - Sheets: {info['total_sheets']}\n"
            result += f"  - Total Rows: {info['total_rows']}\n"
            result += f"  - Total Columns: {info['total_columns']}\n\n"

        return result

    def _filter_question(self, question: str) -> str:
        """Help with filtering data"""
        return """
        🔍 **Filtering Help:**

        I can help you filter data! Here are some examples:
        - "Show me rows where column X > 100"
        - "Filter data where status = 'Active'"
        - "Find records with missing values"

        Please specify:
        1. Which file and sheet
        2. The filtering condition
        3. Which columns to display
        """

    def _average_question(self, question: str) -> str:
        """Calculate averages for numeric columns"""
        result = "📊 **Average Values:**\n\n"
        for filename, sheets in self.dataframes.items():
            result += f"**{filename}:**\n"
            for sheet_name, df in sheets.items():
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    result += f"  **{sheet_name}:**\n"
                    for col in numeric_cols:
                        avg_val = df[col].mean()
                        result += f"    - {col}: {avg_val:.2f}\n"
                result += "\n"
        return result

    def _max_question(self, question: str) -> str:
        """Find maximum values for numeric columns"""
        result = "📈 **Maximum Values:**\n\n"
        for filename, sheets in self.dataframes.items():
            result += f"**{filename}:**\n"
            for sheet_name, df in sheets.items():
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    result += f"  **{sheet_name}:**\n"
                    for col in numeric_cols:
                        max_val = df[col].max()
                        result += f"    - {col}: {max_val}\n"
                result += "\n"
        return result

    def _min_question(self, question: str) -> str:
        """Find minimum values for numeric columns"""
        result = "📉 **Minimum Values:**\n\n"
        for filename, sheets in self.dataframes.items():
            result += f"**{filename}:**\n"
            for sheet_name, df in sheets.items():
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    result += f"  **{sheet_name}:**\n"
                    for col in numeric_cols:
                        min_val = df[col].min()
                        result += f"    - {col}: {min_val}\n"
                result += "\n"
        return result

    def _general_question_help(self) -> str:
        """Provide help for asking questions"""
        return """
        🤔 **I can help you with questions like:**

        📊 **Data Overview:**
        - "How many rows are in each sheet?"
        - "What columns are available?"
        - "Give me a summary of the data"

        🔢 **Statistics:**
        - "What's the average of column X?"
        - "What's the maximum value in column Y?"
        - "Show me the minimum values"

        🔄 **Comparisons:**
        - "Compare the two files"
        - "What's different between sheet A and sheet B?"

        Try asking a more specific question, and I'll do my best to help!
        """

    def get_dataframe(self, filename: str, sheet_name: str) -> pd.DataFrame:
        """Get a specific dataframe for direct manipulation"""
        if filename in self.dataframes and sheet_name in self.dataframes[filename]:
            return self.dataframes[filename][sheet_name]
        return None


# Streamlit UI
def main():
    st.set_page_config(page_title="Excel Analyzer", layout="wide")

    st.title("🔍 Excel File Analyzer")
    st.markdown("Upload Excel files and ask questions about your data!")

    # Initialize analyzer
    if "analyzer" not in st.session_state:
        st.session_state.analyzer = ExcelAnalyzer()

    analyzer = st.session_state.analyzer

    # Sidebar for file upload
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

    # Main content area
    if analyzer.dataframes:
        # Display file summaries
        st.header("📊 Loaded Files Summary")

        for filename in analyzer.dataframes.keys():
            with st.expander(f"📋 {filename}", expanded=False):
                st.markdown(analyzer.get_file_summary(filename))

                # Sheet details
                for sheet_name in analyzer.dataframes[filename].keys():
                    st.subheader(f"Sheet: {sheet_name}")
                    st.markdown(analyzer.get_sheet_info(filename, sheet_name))

                    # Show first few rows
                    df = analyzer.get_dataframe(filename, sheet_name)
                    st.dataframe(df.head(), use_container_width=True)

        # Question interface
        st.header("❓ Ask Questions About Your Data")

        question = st.text_input(
            "What would you like to know?",
            placeholder="e.g., How many rows are in each sheet? What's the average of column X?",
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

        # Pre-made question buttons
        st.subheader("🚀 Quick Questions")
        col1, col2, col3 = st.columns(3)

        with col1:
            if st.button("📊 Row Counts"):
                st.markdown("### 💡 Answer:")
                st.markdown(analyzer.answer_question("How many rows in each sheet?"))

        with col2:
            if st.button("📋 Column Names"):
                st.markdown("### 💡 Answer:")
                st.markdown(analyzer.answer_question("What columns are available?"))

        with col3:
            if st.button("📈 Summary Stats"):
                st.markdown("### 💡 Answer:")
                st.markdown(analyzer.answer_question("Give me a summary of the data"))

    else:
        st.info("👆 Please upload Excel files using the sidebar to get started!")
        st.markdown("""
        ### 🚀 How to use this tool:

        1. **Upload Files**: Use the sidebar to upload one or more Excel files
        2. **Explore Data**: View summaries and preview your data
        3. **Ask Questions**: Use natural language to query your data

        ### 💡 Example questions you can ask:
        - "How many rows are in each sheet?"
        - "What columns are available in the data?"
        - "What's the average of the sales column?"
        - "Compare the two files I uploaded"
        - "Show me a summary of the numeric data"
        """)


if __name__ == "__main__":
    # For running with Streamlit: streamlit run main.py
    main()
