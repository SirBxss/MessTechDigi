import pandas as pd


class DataLoader:
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = None

    def load_excel(self):
        self.data = pd.read_excel(self.file_path)

    def save_excel(self, output_path):
        self.data.to_excel(output_path, index=False)


class DataCleaner:
    @staticmethod
    def unmerge_vertical(column):
        filled = column.ffill()
        mask = column.notna()
        for i in range(1, len(mask)):
            if mask.iloc[i]:
                filled.iloc[i] = column.iloc[i]
            elif not mask.iloc[i - 1]:  # it was NaN before
                filled.iloc[i] = column.iloc[i]
        return filled

    @staticmethod
    def format_date(column, format_string='%d.%m.%Y'):
        """Convert and format date column."""
        return pd.to_datetime(column).dt.strftime(format_string)


class DataManager:
    def __init__(self, file_path):
        self.loader = DataLoader(file_path)
        self.loader.load_excel()
        self.data = self.loader.data

    def clean_data(self):
        """Apply cleaning functions to data columns."""
        id_column = self.find_id_column()
        birthdate_column = self.find_birth_date_column()
        if id_column:
            self.data[id_column] = DataCleaner.unmerge_vertical(self.data[id_column])
        if 'Last Name' in self.data.columns:
            self.data['Last Name'] = DataCleaner.unmerge_vertical(self.data['Last Name'])
        if birthdate_column:
            self.data['Date of birth'] = DataCleaner.format_date(self.data['Date of birth'])

    def fill_first_column(self, method='ffill'):
        """Fill empty cells in the first column using the specified method."""
        first_column_name = self.data.columns[0]
        self.data[first_column_name] = self.data[first_column_name].fillna(method=method)

    def find_id_column(self):
        """Dynamically find the column that likely represents an ID."""
        id_variants = ['ID', 'identification number', 'Identity number']
        for column in self.data.columns:
            if column.lower().replace(" ", "") in (variant.lower().replace(" ", "") for variant in id_variants):
                return column
        return None

    def find_birth_date_column(self):
        """Dynamically find the column that likely represents a Birth date."""
        id_variants = ['date of birth', 'Birthdate', 'BD', 'Birth date']
        for column in self.data.columns:
            if column.lower().replace(" ", "") in (variant.lower().replace(" ", "") for variant in id_variants):
                return column
        return None

    def search_by_id(self, search_id):
        """Search for a record by ID."""
        id_column = self.find_id_column()
        if id_column:
            filtered_data = self.data[self.data[id_column] == search_id]
            return filtered_data
        return pd.DataFrame()  # Return an empty DataFrame if no suitable ID column is found

    # def display_data(self):
    #     """Display the first 20 rows of the DataFrame."""
    #     print(self.data.head(20))

    def save_cleaned_data(self, output_path):
        """Save cleaned data to an Excel file."""
        self.loader.data = self.data
        self.loader.save_excel(output_path)


if __name__ == "__main__":
    data_manager = DataManager('Sample_list.xlsx')
    data_manager.clean_data()
    data_manager.fill_first_column()
    search_id = input("Enter the ID to search for: ")
    # found_data = data_manager.search_by_id('50Bb061cB30B461')
    found_data = data_manager.search_by_id(search_id)
    if not found_data.empty:
        print(found_data.to_string(index=True))
    else:
        print("No data found for the given ID.")
    # data_manager.display_data()
    data_manager.save_cleaned_data('Cleaned_Sample_list.xlsx')
