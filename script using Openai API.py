# Imports essentials
import pandas as pd
import openai

# Set up your OpenAI API key
openai.api_key = 'sk-******'

# Specify the input data file path and read it using the 'pd.read' function from pandas library
input_file_path = 'input_data.xlsx'
df_input = pd.read_excel(input_file_path)

# Specify static output columns for storing the cleaned output data
output_columns = ['Department Name', 'CGAC', 'Pop Street Address', 'Pop State', 'Pop Zip', 'Country', 'Award Number',
                  'Award Date', 'Award Amount ($)', 'Awardee Name', 'City', 'State', 'Zipcode', 'Country']


# Create a DataFrame with static output columns
df_output = pd.DataFrame(columns=output_columns)

df_output['Department Name'] = df_input['Department/Ind.Agency'] + ' ' + df_input['Sub-Tier']

# call the GPT-3
def standardize_text(text):
    prompt = f"search for the following text and Standardize the following text and give cleaned name for it: {text}"
    response = openai.Completion.create(
        engine="text-davinci-002",  # Choose an appropriate engine
        prompt=prompt,
        max_tokens=50  # Adjust as needed for a concise result
    )
    return response.choices[0].text.strip()

# These lines from 24-31 makes a call to the OpenAI GPT-3 API using the openai.Completion.create method. It sends the
# prompt to the GPT-3 model, specifying the engine (text-davinci-002) and setting the maximum number of tokens in the
# response (max_tokens=50).

df_output['Department Name'].apply(standardize_text)

# Entering the data from input file (input_data) to the output file (output_data) after cleaning and skiping unwanted
# data
df_output['CGAC'] = df_input['CGAC']
df_output['AAC Code'] = df_input['AAC Code']
df_output['Pop Street Address'] = df_input['PopStreetAddress'] + ', ' + df_input['PopCity']
df_output['Pop State'] = df_input['PopState']
df_output['Pop Zip'] = df_input['PopZip']
df_output['Country'] = df_input['PopCountry']

df_output['Award Number'] = df_input['AwardNumber']
df_output['Award Date'] = df_input['AwardDate']
df_output['Award Amount ($)'] = df_input['Award$']
df_output['Awardee Name'] = df_input['Awardee']
df_output['City'] = df_input['City']
df_output['State'] = df_input['State']
df_output['Zipcode'] = df_input['ZipCode']
df_output['Country'] = df_input['CountryCode']

# Write to output Excel sheet
output_file_path = 'output.xlsx'
df_output.to_excel(output_file_path, index=False)

# Validating if the output file has been successfully created and the data is claned and stored in the output file (
# output_data)
if os.path.exists(output_file_path):
    print("Test case passed: Data successfully written to output_data.xlsx")
else:
    print("Test case failed: Data not written to output_data.xlsx")