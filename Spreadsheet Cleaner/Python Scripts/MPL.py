from typing import Text
import pandas as pd
import numpy as np
import os
import sys
import time
from UliPlot.XLSX import auto_adjust_xlsx_column_width
from fuzzywuzzy import process
from pandas.io.formats.style_render import Subset


class MPL:

    # Properties
    script_location = os.path.realpath(__file__)
    references_folder = script_location[:-
                                        (len(os.path.basename(sys.argv[0])))] + "References"
    filesinfolder = os.listdir(references_folder)

    # variables for files that will be needed for script. (Filepaths gathered later on)
    # if script and other files needed won't be together, please add locations for files and remove search functions below for commodity/programs files
    commodities_reference = ""  # location reference
    programs_reference = ""  # location reference
    commodity_file = ""  # reading commodity file
    programs_file = ""  # reading programs file

    uncleaned_mpl = ""  # location reference
    # only the filename which will be modified for new clean file name default
    just_filename_uncleaned_mpl = ""
    mpl = ""  # dataframe of unclean MPL
    reduced = ""  # dataframe after duplicates removed from MPL
    cleaned_mpl = ""  # dataframe of cleaned MPL
    cleaned_mpl_with_changes = ""  # dataframe with records of things that changed
    name_of_new_cleaned_mpl = ""
    step_count = 0
    file_validity = 0
    program_col = 0
    parts_col = 0
    commodity_col = 0
    continue_regardless = ""
    programs_reference_last_update = ""

    # Columns names that MPL file should have
    mpl_cols = ['Program*', 'Site Group', 'Site Building', 'Part*',
                'Part Description', 'Procurement Commodity', 'Ops Finance Commodity*',
                'Cost', 'Supplier Description', 'Supplier Code', 'GSM DRI Name', 'GSM DRI Email',
                'Part Tier', 'FG PO FLAG']

    # To ask for excel file
    def openfile(self, filepath):
        self.uncleaned_mpl = filepath

        print(f'✅ File path: {filePath}')
        self.mpl = pd.read_excel(self.uncleaned_mpl)

        self.check_unclean_file()

        if (filepath is None) or (self.file_validity == 0):
            print("======================")
            print(
                "❌ Error: Provided file was invalid, please select a valid MPL file to openfile()")
            print(f'❌ File path: {filepath}')
            print(f'❌ File validity: {self.file_validity}')
            print("======================")
        else:
            self.just_filename_uncleaned_mpl = self.uncleaned_mpl.rpartition(
                '/')[-1]
            self.name_of_new_cleaned_mpl = os.path.splitext(
                self.just_filename_uncleaned_mpl)[0] + "_CLEANED"
            self.step_count += 1

    # function used to check file validity

    def check_unclean_file(self):
        getting_columns = self.mpl.columns.tolist()

        if getting_columns == self.mpl_cols:
            self.program_col = 0
            self.parts_col = 3
            self.commodity_col = 6
            self.file_validity = 1
        elif len(getting_columns) < 7:
            self.file_validity = 0
        else:
            while (self.parts_col == "") or (self.parts_col == "") or (self.parts_col == ""):

                for col in getting_columns:
                    if self.program_col == "":
                        if col.lower() == 'Program*':
                            self.program_col = getting_columns.index(col)
                        elif (process.extractOne('Program*', getting_columns)[1] >= 90):
                            self.program_col = getting_columns.index(
                                process.extractOne('Program*', getting_columns)[0])

                    if self.parts_col == "":
                        if col.lower() == 'Part*':
                            self.parts_col = getting_columns.index(col)
                        elif (process.extractOne('Part*', getting_columns)[1] >= 90):
                            self.parts_col = getting_columns.index(
                                process.extractOne('Part*', getting_columns)[0])

                    if self.commodity_col == "":
                        if col.lower() == 'Ops Finance Commodity*':
                            self.commodity_col = getting_columns.index(col)
                        elif (process.extractOne('Ops Finance Commodity*', getting_columns)[1] >= 90):
                            self.commodity_col = getting_columns.index(process.extractOne(
                                'Ops Finance Commodity*', getting_columns)[0])

                if (self.program_col == "") or (self.parts_col == "") or (self.commodity_col == ""):
                    self.get_confirmation()

                    print("I got to this point")
                    print(self.continue_regardless)
                    print("This is after supposedly printing the value from pop up")
                    if self.continue_regardless.lower() in ['yes', 'yea', 'y']:
                        print('Perfect, lets get started....\n')
                        # time.sleep(3) # Sleep for 3 seconds
                        self.program_col = 0
                        self.parts_col = 3
                        self.commodity_col = 6
                        self.file_validity = 1
                    else:
                        print(self.continue_regardless)
                        print("Unable to locate necessary columns")
                        self.file_validity = 0
                        break

    # removing duplicate rows before the clean up and providing reference of quantity removed
    def remove_duplicates(self):
        # show count of rows prior to removal of APN-Program duplicates
        print('Total rows in file:', self.mpl.shape[0])

        if (self.mpl.duplicated(subset=[(self.mpl.columns[self.program_col]), (self.mpl.columns[self.parts_col])]).sum()) > 0:
            # notifying of duplicates that will be removed
            print('---> Removing', self.mpl.duplicated(subset=[
                (self.mpl.columns[self.program_col]), (self.mpl.columns[self.parts_col])]).sum(), 'duplicates\n')

            # remove duplicates based on several columns
            self.reduced = self.mpl.drop_duplicates(
                subset=[(self.mpl.columns[self.program_col]), (self.mpl.columns[self.parts_col])], keep="last")
            print('\nTotal rows after removing duplicates:',
                  self.reduced.shape[0])
        else:
            print("\nNo duplicates rows in file")
            self.reduced = self.mpl

    # removing duplicate rows after clean up was done and providing reference of quantity removed
    def remove_duplicates_after_cleanup(self):

        # show count of rows prior to removal of APN-Program duplicates
        print('Total rows in file after clean:',
              self.cleaned_mpl_with_changes.shape[0])
        y = self.cleaned_mpl_with_changes.shape[0]
        if (self.cleaned_mpl_with_changes.duplicated(subset=["Program*", "Part*"]).sum()) > 0:
            # notifying of duplicates that will be removed
            print('---> Removing', (self.cleaned_mpl_with_changes.duplicated(
                subset=["Program*", "Part*"]).sum()), 'duplicates\n')
            x = self.cleaned_mpl_with_changes.duplicated(
                subset=["Program*", "Part*"]).sum()
            # remove duplicates based on several columns
            self.cleaned_mpl_with_changes = self.cleaned_mpl_with_changes.drop_duplicates(
                subset=["Program*", "Part*"], keep="last")
            print('\nTotal rows after removing duplicates:',
                  self.cleaned_mpl_with_changes.shape[0])
        else:
            print("\nNo duplicates rows in file")

    # function to return programs that were changed
    def returning_changed_programs(self, word):
        programs_list = self.programs_file["Program"].unique().tolist()

        if process.extractOne(word, programs_list)[1] == 100:
            if process.extractOne(word, programs_list)[0] == word:
                return np.nan  # sending blank since it matched exactly
            else:
                return word
        elif process.extractOne(word, programs_list)[1] >= 95:
            return word
        elif process.extractOne("O_"+word, programs_list)[1] >= 95:
            return word
        elif len(word[2:]) > 1:
            if process.extractOne(word[2:], programs_list)[1] >= 95:
                return word
            else:
                return np.nan
        else:
            return np.nan

    # function to clean up programs
    def comparing_programs(self, word):
        programs_list = self.programs_file["Program"].unique().tolist()

        if process.extractOne(word, programs_list)[1] == 100:
            return process.extractOne(word, programs_list)[0]
        elif process.extractOne(word, programs_list)[1] >= 95:
            return process.extractOne(word, programs_list)[0]
        elif process.extractOne("O_"+word, programs_list)[1] >= 95:
            return process.extractOne(word, programs_list)[0]
        elif len(word[2:]) > 1:
            if process.extractOne(word[2:], programs_list)[1] >= 95:
                return process.extractOne(word, programs_list)[0]
            else:
                return word
        else:
            return word

    # function to return programs that were changed
    def returning_changed_commodities(self, word):
        commodities_list = self.commodity_file["Commodity"].unique().tolist()
        sub_commodities_list = self.commodity_file["Sub-Commodity"].unique(
        ).tolist()

        if process.extractOne(word, commodities_list)[1] == 100:
            if process.extractOne(word, commodities_list)[0] == word:
                return np.nan  # sending blank since it matched exactly
            else:
                return word
        elif process.extractOne(word, commodities_list)[1] >= 95:
            return word
        else:
            if word not in ["Uncategorized", "AppleCare Only", ""]:
                if process.extractOne(word, sub_commodities_list)[1] >= 89:
                    match = process.extractOne(word, sub_commodities_list)[0]
                    return word
                else:
                    return np.nan
            else:
                return np.nan

    # function to clean up commodities
    def comparing_commodities(self, word):
        commodities_list = self.commodity_file["Commodity"].unique().tolist()
        sub_commodities_list = self.commodity_file["Sub-Commodity"].unique(
        ).tolist()

        if process.extractOne(word, self.commodities_list)[1] == 100:
            return process.extractOne(word, self.commodities_list)[0]
        elif process.extractOne(word, self.commodities_list)[1] >= 95:
            return process.extractOne(word, self.commodities_list)[0]

        else:
            if word not in ["Uncategorized", "AppleCare Only", ""]:
                if process.extractOne(word, self.sub_commodities_list)[1] >= 89:
                    match = process.extractOne(
                        word, self.sub_commodities_list)[0]
                    return self.commodity_file['Commodity'][self.commodity_file['Sub-Commodity'][self.commodity_file['Sub-Commodity'] == match].index.tolist()[0]]
                else:
                    return word
            else:
                return word

    # main function to process unclean MPL file
    def process_mpl(self):
        # global cleaned_mpl_with_changes
        # global cleaned_mpl

        self.remove_duplicates()

        data_new = []

        try:
            for index, row in self.reduced.iterrows():
                print(index)
                dict_data = {
                    # work in progress to only get values which have a change
                    "Original Program if Changed": self.returning_changed_programs(row[self.reduced.columns.tolist()[self.program_col]]),
                    "Program*": self.comparing_programs(row[self.reduced.columns.tolist()[self.program_col]]),
                    "Site Group": np.nan,
                    "Site Building": np.nan,
                    "Part*": row[self.reduced.columns.tolist()[self.parts_col]],
                    "Part Description": np.nan,
                    "Procurement Commodity": np.nan,
                    # work in progress to only get values which have a change
                    "Original Commodity if Changed": self.returning_changed_commodities(row[self.reduced.columns.tolist()[self.commodity_col]]),
                    "Ops Finance Commodity*": self.comparing_commodities(row[self.reduced.columns.tolist()[self.commodity_col]]),
                    "Cost": np.nan,
                    "Supplier Description": np.nan,
                    "Supplier Code": np.nan,
                    "GSM DRI Name": np.nan,
                    "GSM DRI Email": np.nan,
                    "Part Tier": np.nan,
                    "FG PO FLAG": np.nan
                }

                data_new.append(dict_data)

            self.cleaned_mpl_with_changes = pd.DataFrame(data=data_new)

            self.remove_duplicates_after_cleanup()

            self.cleaned_mpl = self.cleaned_mpl_with_changes.drop(
                columns=["Original Program if Changed", "Original Commodity if Changed"])

        except:
            print("An exception occurred")

    # To save excel file
    def savefile(self):
        newfilename = f'{os.path.splitext(self.just_filename_uncleaned_mpl)[0]}_Cleaned{os.path.splitext(self.just_filename_uncleaned_mpl)[1]}'
        filepath2 = self.filedialog.asksaveasfilename(
            title="Where to save cleaned MPL", defaultextension=".xlsx", initialfile=newfilename)

        with pd.ExcelWriter(filepath2) as writer:
            self.to_excel(writer, sheet_name="Sheet1", index=0)
            auto_adjust_xlsx_column_width(
                self.cleaned_mpl, writer, sheet_name="Sheet1", index=0)  # adjusting column sizing

        if filepath2 is None:
            return
        else:
            print(filepath2)

    # To save excel file with references to original data which changed
    def savefile_w_references(self):
        newfilename = f'{os.path.splitext(self.just_filename_uncleaned_mpl)[0]}_Cleaned_w_Changes{os.path.splitext(self.just_filename_uncleaned_mpl)[1]}'
        filepath2 = self.filedialog.asksaveasfilename(
            title="Where to save cleaned MPL", defaultextension=".xlsx", initialfile=newfilename)

        with pd.ExcelWriter(filepath2) as writer:
            self.cleaned_mpl_with_changes.to_excel(
                writer, sheet_name="Sheet1", index=0)
            auto_adjust_xlsx_column_width(
                self.cleaned_mpl_with_changes, writer, sheet_name="Sheet1", index=0)  # adjusting column sizing

        if filepath2 is None:
            return
        else:
            print(filepath2)

# Order to run:
# openfile()
# process_mpl()
# savefile()
