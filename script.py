
import selenium_tools as st
import pandas as pd
import time

def init_excel():
    while True:
        try:
            open("data_source.xlsx", "r")
            dataframe = pd.read_excel("data_source.xlsx", sheet_name=None)
        except IOError:
            print("Error: Excel file does not appear to exist")
            continue
        break

    return dataframe

def init_driver():
    while True:
        try:
            chromeDriverPath = "chromedriver.exe"
            driver, driver_info = st.get_driver(chromeDriverPath, headless=False)
        except Exception:
            print("Chromedriver error, please check version compatibility")
            continue
        break

    return driver, driver_info

def exec_actions(df, driver, wait_time, sleep_time):

    final_df = df['config'][9:].reset_index(drop=True)
    final_df = final_df[1:]
    final_df.columns = ['no','action','xpath','text','clear_text','enter','dropdown_index','source_column']
    final_df.drop(['no'], axis=1, inplace=True)
    final_df.dropna(how='all', inplace=True)

    index = 0

    excel_source_df = df['data']

    for i in range(len(excel_source_df.index)):

        for index, row in final_df.iterrows():

            action = row['action']

            if(action == 'click_button'):
                st.click_button(driver, wait_time, row['xpath'])

            if(action == 'input_text'):
                temp_clear = row['clear_text'].lower() in ("Y")
                temp_enter = row['enter'].lower() in ("Y")

                st.input_text(driver, wait_time, row['xpath'], row["text"], temp_clear, temp_enter)

            if(action == 'click_dropdown'):
                st.click_dropdown_item(driver, row['xpath'], row['dropdown_index'])

            if(action == 'excel_data'):
                temp_clear = row['clear_text'].lower() in ("Y")
                temp_enter = row['enter'].lower() in ("Y")

                st.input_text(driver, wait_time, row['xpath'], excel_source_df.iloc[[i]][row['source_column']].to_string(header=False, index=False), temp_clear, temp_enter)

            if(action == 'excel_button'):
                st.click_button(driver, wait_time, excel_source_df.iloc[[i]][row['source_column']].to_string(header=False, index=False))

        time.sleep(sleep_time)

    return

dataframe = init_excel()
driver, driver_info = init_driver()

# Get row 7 column 1 (start 0)
target_url = dataframe["config"].iloc[7,1]

driver.get(target_url)

exec_actions(dataframe, driver, wait_time=100, sleep_time=5)

input('Press any key to exit!')
st.driver_quit(driver)