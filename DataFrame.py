import os

import pandas as pd

file_path = os.path.abspath("data.xlsx")
data = pd.read_excel(file_path)
file_path = os.path.abspath("State.xlsx")
state = pd.read_excel(file_path)
file_path = os.path.abspath("Agent.xlsx")
agent = pd.read_excel(file_path)
file_path = os.path.abspath("HireDate.xlsx")
hire_date = pd.read_excel(file_path)
# Question 1

drug_company = "drug_company"

# Calculate the length of each value in the column drug_company and store it in a new column
data['value_length'] = data[drug_company].astype(str).apply(len)

# Calculate the average value length
average_length = data['value_length'].mean()

print("Question One")
print("The average size (in letters) of a drug company name is", average_length)
print()

# Question 2

# Create an empty list of the drug companies that have a comma in their name
list_of_drug_companies_with_commas = []

# Iterate over each row in the column drug_company
for value in data["drug_company"]:
    #If it has a comma in its name add it to the list
    if ',' in str(value):
        list_of_drug_companies_with_commas.append(value)

print("Question Two")
print("There are", len(list_of_drug_companies_with_commas), "drug company that have commas in their name")
print("and they are", list_of_drug_companies_with_commas)
print()

# Question 3

# Create a new column by concatenating first_name and last_name
data['full_name'] = data['first_name'] + ' ' + data['last_name']

# first convert the hire date column to be a datetime format, then we can reformat it
data['hire_date'] = pd.to_datetime(data['hire_date'], dayfirst=True).dt.strftime('%Y-%m-%d')

filtered_data = data.loc[data['hire_date'] > '2022-12-31']
# Filter the rows based on the condition
print("Question Three")
print("There were", len(filtered_data), "people hired this year")
print("and they were:")
print(filtered_data['full_name'].to_string(index=False))
print()

# Question 4
for index, value in enumerate(data['full_name']):
    first_and_last = str(value)
    first_and_last = first_and_last.split()
    first = first_and_last[0]
    capital_first = first[0].upper() + first[1:]
    last = first_and_last[1]
    capital_last = last[0].upper() + last[1:]
    first_and_last[0] = capital_first
    first_and_last[1] = capital_last
    data.at[index, 'full_name'] = " ".join(first_and_last)
print("Question Four")
print("Here are the names now capitalized:")
print(data['full_name'].to_string(index=False))
print()

# Question 5

# Which state had the most hires?

newDF = state["drug_company"].str.split(": ")

state["drug_company"] = newDF.str[1]

tier = newDF.str[0]

state["tier"] = tier.str[13:]

merged_df = pd.merge(data, state, on="drug_company")

merged_df.to_excel('merged.xlsx', index=False)

number_per_state = merged_df["state"].value_counts()

print("The State with the most hires was", number_per_state.idxmax(), "and it had", number_per_state[0])
print()

# Question 6

# Which agent had the most hires?

agent["state"] = agent["state"].str[1:]

merged_df = pd.merge(merged_df, agent, on="state")

merged_df.to_excel('merged.xlsx', index=False)

number_per_agent = merged_df["agent_name"].value_counts()

print("The Agent with the most hires was", number_per_agent.idxmax() ,"and he had", number_per_agent[0])
print()

# Question 7

# Which agent had the most hires this month?


filtered_may = merged_df[merged_df['hire_date'] > '2023-04-30']

number_per_agent_this_month = filtered_may["agent_name"].value_counts()

print("The Agent with the most hires in May was", number_per_agent_this_month.idxmax() ,"and she had", number_per_agent_this_month[0])
print()

# Question 8

# Which year saw the highest number of hires?

merged_df['year_hired'] = merged_df['hire_date'].str[0:4]

year_with_the_most_hires = merged_df['year_hired'].value_counts().idxmax()

print("The year with the most hires was", year_with_the_most_hires)
print()

# Questions for the Week

# Question 1

# Agents get a bonus based on which tier drug company they hire from. For Tier A companies, they get a $5 bonus,
# if Tier B, $10, Tier C, $20, Tier D, $50. How much is each agent getting in bonuses for this year?

# Function to calculate the bonus amount for each agent
merged_df["Bonus"] = 0

# add the bonus for each drug_company hired
for index, row in merged_df.iterrows():
    agent = row["agent_name"]
    tier = row["tier"]
    if tier == "TIER A":
        merged_df.loc[index, 'Bonus'] = 5
    elif tier == "TIER B":
        merged_df.loc[index, 'Bonus'] = 10
    elif tier == "TIER C":
        merged_df.loc[index, 'Bonus'] = 20
    elif tier == "TIER D":
        merged_df.loc[index, 'Bonus'] = 50

def calculate_bonus(data_frame, agent_name):
    agent_bonus = 0

    for index, row in data_frame.iterrows():
        if row["agent_name"] == agent_name:
            tier = row["tier"]

            if tier == "TIER A":
                agent_bonus += 5
            elif tier == "TIER B":
                agent_bonus += 10
            elif tier == "TIER C":
                agent_bonus += 20
            elif tier == "TIER D":
                agent_bonus += 50

    return agent_bonus

jonathan_bonuses = merged_df[(merged_df['agent_name'] == 'Jonathan Braiden') & (merged_df['year_hired'] == '2023')]

Jonathan_Braiden_Bonus = calculate_bonus(jonathan_bonuses, "Jonathan Braiden")

print("Jonathan Braiden's Bonus was", Jonathan_Braiden_Bonus, "dollars")
print()

zehava_bonuses = merged_df[(merged_df['agent_name'] == 'Zehava David') & (merged_df['year_hired'] == '2023')]

Zehava_David_Bonus = calculate_bonus(zehava_bonuses, "Zehava David")

print("Zehava David's Bonus was", Zehava_David_Bonus, "dollars")
print()

murray_bonuses = merged_df[(merged_df['agent_name'] == 'Murray Goldman') & (merged_df['year_hired'] == '2023')]

Murray_Goldman_Bonus = calculate_bonus(murray_bonuses, "Murray Goldman")

print("Murray Goldman's Bonus was", Murray_Goldman_Bonus , "dollars")
print()

merged_df.to_excel('merged.xlsx', index=False)

# Question 2

# Agents get an extra bonus if they make mutiple hires in one week.
# The bonus increases depending on how many hires were made that week.
# Thus, if 2 hires were made, they get a $20 bonus.
# If 3 hires, then $30, etc. How much is each agent getting in bonuses this year?
# (Please add this calculation to the previous one).

merged_df["week"] = pd.to_datetime(merged_df['hire_date']).dt.isocalendar().week

df = merged_df[merged_df['year_hired'] == '2023']

jonathan_bonuses = df[df['agent_name'] == 'Jonathan Braiden']

# Iterate over the name counts
for week, count in jonathan_bonuses["week"].value_counts().items():
    if count > 1:
        Jonathan_Braiden_Bonus += (count - 1) * 10

print("Jonathan's bonus was", Jonathan_Braiden_Bonus, "dollars after factoring in bonus for multiple hires in one week")
print()

zehava_bonuses = df[df['agent_name'] == 'Zehava David']

# Iterate over the name counts
for week, count in zehava_bonuses["week"].value_counts().items():
    if count > 1:
        Zehava_David_Bonus += (count - 1) * 10

print("Zehava's bonus was", Zehava_David_Bonus, "dollars after factoring in bonus for multiple hires in one week")
print()

murray_bonuses = df[df['agent_name'] == 'Murray Goldman']

# Iterate over the name counts
for week, count in murray_bonuses["week"].value_counts().items():
    if count > 1:
        Murray_Goldman_Bonus += (count - 1) * 10

print("Murray's bonus was", Murray_Goldman_Bonus, "dollars after factoring in bonus for multiple hires in one week")
print()

merged_df.to_excel('merged.xlsx', index=False)

# Question 3

# After much debate, the company is agreeing to affect these bonus rules back from 2015.
# Please recalculate the bonuses.

merged_df["year"] = pd.to_datetime(merged_df['hire_date']).dt.isocalendar().year

# filter for week and year

merged_df['week_year'] = merged_df['week'].astype(str) + '-' + merged_df['year'].astype(str)

df = merged_df[merged_df['year_hired'] >= '2015']

jonathan_bonuses = df[df['agent_name'] == 'Jonathan Braiden']

Jonathan_Braiden_Bonus = calculate_bonus(jonathan_bonuses, "Jonathan Braiden")

# Iterate over the name counts
for week, count in jonathan_bonuses["week_year"].value_counts().items():
    if count > 1:
        Jonathan_Braiden_Bonus += (count - 1) * 10

print("Jonathan's bonus was", Jonathan_Braiden_Bonus, "dollars from 2015")
print()

df = merged_df[merged_df['year_hired'] >= '2015']

zehava_bonuses = df[df['agent_name'] == 'Zehava David']

Zehava_David_Bonus = calculate_bonus(zehava_bonuses, "Zehava David")

# Iterate over the name counts
for week, count in zehava_bonuses["week_year"].value_counts().items():
    if count > 1:
        Zehava_David_Bonus += (count - 1) * 10

print("Zehava's bonus was", Zehava_David_Bonus, "dollars from 2015")
print()

df = merged_df[merged_df['year_hired'] >= '2015']

murray_bonuses = df[df['agent_name'] == 'Murray Goldman']

Murray_Goldman_Bonus = calculate_bonus(murray_bonuses, "Murray Goldman")

# Iterate over the name counts
for week, count in murray_bonuses["week_year"].value_counts().items():
    if count > 1:
        Murray_Goldman_Bonus += (count - 1) * 10

print("Murray's bonus was", Murray_Goldman_Bonus, "dollars from 2015")
print()


merged_df.to_excel('merged.xlsx', index=False)

# Question 4

# Which state has the most Tier D drug companies?

df = merged_df[merged_df['tier'] == 'TIER D']

state_with_the_most_TIER_D = df['state'].value_counts()

print("The state with the most TIER D drug companies is", state_with_the_most_TIER_D.idxmax())
print()

# Question 5

# The HR department is suggesting that candidates are less likely to be hired
# when first contacted on a Sunday. Is that true?

# Convert 'first_contact_date' column from string to datetime format
hire_date['first_contact_date'] = pd.to_datetime(hire_date['first_contact_date'])


# Extract the day of the week
hire_date['day_of_week_contacted'] = hire_date['first_contact_date'].dt.day_name()

hire_date.to_excel('HireDate.xlsx', index=False)

merged_df = pd.merge(merged_df, hire_date, on=["last_name", "first_name"])

print("This is true since the 'first contact day' with the least amount of hires is",
      merged_df['day_of_week_contacted'].value_counts().idxmin())
print()

merged_df.to_excel('merged.xlsx', index=False)

# Question 6

# The HR department found out it's annoying for people to be contacted on Shabbos.
# They therefore want to impose the following rule: no bonuses are given
# for any candidates first contacted on a Shabbos. Please recalculate the bonuses.

# Filter out rows where 'day_of_week_contacted' is not 'Saturday'
df = merged_df[merged_df['day_of_week_contacted'] != 'Saturday']

df = df[df['year_hired'] >= '2015']

jonathan_bonuses = df[df['agent_name'] == 'Jonathan Braiden']

Jonathan_Braiden_Bonus = calculate_bonus(jonathan_bonuses, "Jonathan Braiden")

# Iterate over the name counts
for week, count in jonathan_bonuses["week_year"].value_counts().items():
    if count > 1:
        Jonathan_Braiden_Bonus += (count - 1) * 10

print("Jonathan's bonus was", Jonathan_Braiden_Bonus, "dollars from 2015")
print()

df = df[df['year_hired'] >= '2015']

zehava_bonuses = df[df['agent_name'] == 'Zehava David']

Zehava_David_Bonus = calculate_bonus(zehava_bonuses, "Zehava David")

# Iterate over the name counts
for week, count in zehava_bonuses["week_year"].value_counts().items():
    if count > 1:
        Zehava_David_Bonus += (count - 1) * 10

print("Zehava's bonus was", Zehava_David_Bonus, "dollars from 2015")
print()

df = df[df['year_hired'] >= '2015']

murray_bonuses = df[df['agent_name'] == 'Murray Goldman']

Murray_Goldman_Bonus = calculate_bonus(murray_bonuses, "Murray Goldman")

# Iterate over the name counts
for week, count in murray_bonuses["week_year"].value_counts().items():
    if count > 1:
        Murray_Goldman_Bonus += (count - 1) * 10

print("Murray's bonus was", Murray_Goldman_Bonus, "dollars from 2015")
print()

merged_df.to_excel('merged.xlsx', index=False)

# Question 7

# The HR department wants to know
# if winter months see more hires than summer months. Is there a significant difference?

# Extract the month hired
merged_df["month_hired"] = pd.to_datetime(merged_df['hire_date']).dt.month

# Filter out the summer months
summer_months = merged_df[merged_df['month_hired'].isin([6, 7, 8])]

# Filter out the winter months
winter_months = merged_df[merged_df['month_hired'].isin([1, 2, 11])]

print("There are more hires in summer months than winter months since the summer months had",
      summer_months['month_hired'].sum(), "hires and the winter months had",
      winter_months['month_hired'].sum(), "hires making there a", summer_months['month_hired'].sum() -
      winter_months['month_hired'].sum(), "difference in hires depending on the season")
print()

merged_df.to_excel('merged.xlsx', index=False)
