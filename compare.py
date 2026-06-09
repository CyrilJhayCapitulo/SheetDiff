import pandas as pd


def load_excel(file_path):
    """
    Load Excel file.
    """
    return pd.read_excel(file_path)


def compare_excel(file_a, file_b, key_column=None):
    """
    Compare two Excel files using a key column.

    Returns:
        differences_df
        aligned_df_a
        aligned_df_b
    """

    df_a = pd.read_excel(file_a)
    df_b = pd.read_excel(file_b)

    # Auto-select first column if none provided
    if key_column is None:
        key_column = df_a.columns[0]

    if key_column not in df_a.columns:
        raise Exception(
            f"'{key_column}' not found in File A"
        )

    if key_column not in df_b.columns:
        raise Exception(
            f"'{key_column}' not found in File B"
        )

    df_a = df_a.fillna("")
    df_b = df_b.fillna("")

    # Keep original copies for aligned display
    display_df_a = df_a.copy()
    display_df_b = df_b.copy()

    # Set index for comparison
    df_a = df_a.set_index(key_column)
    df_b = df_b.set_index(key_column)

    differences = []

    all_keys = sorted(
        set(df_a.index).union(set(df_b.index)),
        key=str
    )

    # ==================================================
    # Compare Data
    # ==================================================

    for key in all_keys:

        # ------------------------------------------
        # Added Row
        # ------------------------------------------

        if key not in df_a.index:

            differences.append(
                {
                    "Status": "ADDED",
                    "Key": key,
                    "Column": "",
                    "Old Value": "",
                    "New Value": "Entire Row Added",
                }
            )

            continue

        # ------------------------------------------
        # Deleted Row
        # ------------------------------------------

        if key not in df_b.index:

            differences.append(
                {
                    "Status": "DELETED",
                    "Key": key,
                    "Column": "",
                    "Old Value": "Entire Row Deleted",
                    "New Value": "",
                }
            )

            continue

        row_a = df_a.loc[key]
        row_b = df_b.loc[key]

        common_columns = [
            col
            for col in df_a.columns
            if col in df_b.columns
        ]

        for column in common_columns:

            value_a = row_a[column]
            value_b = row_b[column]

            if str(value_a) != str(value_b):

                differences.append(
                    {
                        "Status": "MODIFIED",
                        "Key": key,
                        "Column": column,
                        "Old Value": value_a,
                        "New Value": value_b,
                    }
                )

    # ==================================================
    # Build Aligned DataFrames
    # ==================================================

    aligned_rows_a = []
    aligned_rows_b = []

    display_df_a = display_df_a.set_index(key_column)
    display_df_b = display_df_b.set_index(key_column)

    aligned_columns = [key_column] + list(df_a.columns)

    blank_row = [""] * len(aligned_columns)

    for key in all_keys:

        # Left Side
        if key in display_df_a.index:

            row_data = (
                [key]
                + display_df_a.loc[key].tolist()
            )

            aligned_rows_a.append(row_data)

        else:

            aligned_rows_a.append(blank_row.copy())

        # Right Side
        if key in display_df_b.index:

            row_data = (
                [key]
                + display_df_b.loc[key].tolist()
            )

            aligned_rows_b.append(row_data)

        else:

            aligned_rows_b.append(blank_row.copy())

    aligned_df_a = pd.DataFrame(
        aligned_rows_a,
        columns=aligned_columns
    )

    aligned_df_b = pd.DataFrame(
        aligned_rows_b,
        columns=aligned_columns
    )

    return (
        pd.DataFrame(differences),
        aligned_df_a,
        aligned_df_b
    )