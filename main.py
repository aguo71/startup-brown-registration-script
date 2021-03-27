import pandas as pd

# replace with your filename
df = pd.read_excel("/Users/ashleychang/PycharmProjects/registration-startup/first_version_of_startup_registration.xlsx")
df = df[['Name', 'School Email', 'Please rank your *Workshop I* preferences below!']]
df['workshop list'] = df.apply(lambda x: x['Please rank your *Workshop I* preferences below!'].split(","))
print(df)

