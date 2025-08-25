# Expiry use Database Update

This document was created on 7/28/2025 in a effort to document the best processes in updating CEDACS Expiry use database to a new at risk date of 2030.



## General Notes

This section will be used for quick notes. There is no date as it is constantly overwritten.

17.625 % of properties are 100% section 8, and have a expiry date after 2030. (282 total)

858, or 45.61 % of all properties received some form of section 8.

576 properties only have a subset of units covered under section 8.

201 properties have a override date that is past the forecast date.







## Pulling additonal fields (8/20/25)

In order to make significant further progress in sectioning off a significant portion of projects, I have to pull additional information out of the database than what is available in the forecast report currently. These fields are  

| MF_Assistance_and_Sec8_Contracts_archive                     | Properties                                                   | DHCD_archive             |
| ------------------------------------------------------------ | ------------------------------------------------------------ | ------------------------ |
| tracs_effective_date                                         | ProjectID_lending                                            | CBH Awarded              |
| - Will need owner_email_text from the MF_Properties_with_Assistance_and_Sec8_Contracts_archive | Comment: Preservation Status, Type or other new Affordability Te | FCF Awarded (Both types) |
|                                                              | Mortgage status/ TC and other restrictions                   | HIF Awarded              |
|                                                              | Project/AKA                                                  | Date <Loan name> Closed  |

Also will require pulling the ExpUse Property ID field from all of the tables above in order to data match.

## Data Syncing (8/18/25) - (8/19/25)

Due to having many different sheets full of sub querys, when a sub query is updated it needs to be reflect back on the main sheet. That being said, they are cases when the same property is represented in many sub querys, how should that be synced?

The solution I have implemented in a *Bi-Directional Sync* It works by:

When a user initiates a sync, the Jameer notes, Roberta notes, and 1st cut status column are all copied into the **Main Forecast** sheet. However, if two subqueries contain the same property but with differing notes, the notes will be combined and placed into the corresponding notes section in the main forecast. 

For example 

| All TC Properties  | All Overrides            | Main Forecast After Sync                  |
| ------------------ | ------------------------ | ----------------------------------------- |
| "Look at registry" | "Registry book ___ says" | Look at registry \| Registry book __ says |
| "Refi in 2024"     |                          | Refi in 2024                              |
|                    | Owned by non-profit      | Owned by non-profit                       |

After the main forecast has been synced, the newly synced notes from the main forecast are copied back into the subqueries in order to keep everything the same. 

The code for this process can be found in **BiDirectionalSync.py**

> Currently this process can only be done with the execution of a python script. In the future, Roberta will like a to initiate the syncing process herself. This can via a application exe, but the details of which will need to be discussed further in order to comply with CEDAC security protocol.

## Stat taking (8/12/25)



**PRAC Stats**

- 109 out of 159 PRAC properties have <= 20 units (68%). 50 Properties have more than 20 units
___
**202 Stats**
- 14 out of 28 202 properties have <= 20 units (50%) Leaving another 14 with more than 20 units

- Average of number of units of properties that are not preserved is 72

- Median number of uinits of properties that are not preserved is 50

- They are a total of 136 preserved units

- Total of 1084 unpreserved units
___
**TC Stats**
-  Property ID 592,17934,547,358,242 All Have section 8 expiry dates before 2030, but are TC projects made 30 years before forecast date
-  87 TC projects have section 8
___
**Active Mortgage Final Endorsement Notes**

- What should be done about partially S8 Properties on this list?

___
Next, It would be helpful if in the 1st cut status the reason for being marked as preserved by me was noted. I can go back and change that. They are 3 reasons why a property would be marked as preserved for now.
1. 100% S.8 Expiring after the forecast date
2. <= 20 unit PRAC property
3. <= 20 unit 202 property

An Additional question I have is why does prop ID 237 repeat is TC year awarded forecast 5 times? Seems weird for the tree to do that.

  

## Post-Roberta Meeting Sub-sectioning (8/11/25)

There is a lot of work to be done in order to look get a better look at all of the programs involved in the ask risk report. This includes

- [x] Create spreadsheet with all properties that are 100% section 8 expiring after 2030. Mark all as PRESERVED

- [x] Create spreadsheet with all S8 Properties that expire *Before* 2030. 
  - If Hud Contract Program Group code is PRAC and <= 20 units mark as PERSERVED
  
- [x] Create spreadsheet with all the listings that are 
  - Expiring before 2030
  - 100% S.8
  - PRAC
  - In order to get a visualize what the world of PRAC properties looks like currently.
- [x] Back in the S.8 Properties that are expiring before 2030 spreadsheet, if the program type group code is 202/8 and <= 20 units then mark as PERSERVED
- [x] Also a generate a spreadsheet with all 100% S.8, 202/8 propeties that are expiring before 2030
- [x] Generate spreadsheet of all TC year awarded 0 to 30 years before forecast date properties.
- [x] Generate spreadsheet for all projects where Forecast Preserved By Program is baseed on HUD active mortage with final endorsment date 0 to 15 years before forecast date
- [x] Generate spreadsheet of all projects that have a override past the forecast date
- [x] Highlight ones that have a S.8 end date later the forecast date

This is all involving the verification of projects that are not at risk. Come up with any useful statistics and verifications along the way.  It was also reveled that TRACs is a form of section 8, therefore meaning that the HUD_TRACS_CONTRACT field is a hint that a properties receives section 8 assistance. 

## Additional Sub-sectioning (8/8/25)

So far, I have discovered that ~ 17.5% of properies are fully section 8 and have contracts that expire after 2030.  (About 282)

I also know they are about ~220, 230 section 202 properties in MA that will most likely renew themselves. A good course of action would be extracting this exact number of 202's and setting them aside. A descion must be made, If a 202 is set to expire before 2030, should it be looked at or can it be set aside? (Yes)

After obtaining that information I should be able to see how many 202 props are 100% section 8 to see how many are done in total





## Post Jim Meeting (8/4/25)

- [x] Roberta requested a list of all properties in the expiring use report that have a section 8 end date that is later than the section 8 override date. 
- [ ] Run the "CompleteS8" script discussed in the 7/28 pattern matching log in on the new forecast that Jim sent over.

In the aftermath of the meeting with jim, it was revealed that the "Check date" column that Stephan had imported was not a part of the decision tree code. In order to regain the columns Stephan would need to run that script again. However since he is on vacation it makes things a little more complicated for now.






## Meeting with Jim (7/30/25)

Column T has a star where column S and (?) differ

Has the TC date where it says TC awarded 340 years ago been updated to the currant forecast year (2030).

- Cut the records that different from the 2027 at risk report
- Sometimes properties will go through refinancing where they are some units lost or added.
- TC awarded 0 to 30 year before are all not at risk b/ by law they are required  total more than 30 years. 
- Anywhere where it says prac of 202 and the project is less than 20 units, than we can assume that those properties are good.  

- Whenever ForcastPreservedByProgram is blank, then it should be assumed that the units of that property are at risk, 
- If there ais a active HUD contract, and a override present on the same property, the date that is the furthest away takes precedence.

descion tree is written in SQL And Visual basic

As of this moment meeting, I presented this information to the audience of Dilla, Roberta, and Jim.

> 17.625 % of properties are 100% section 8, and have a expiry date after 2030. (282 total)
>
> 858, or 45.61 % of all properties received some form of section 8.
>
> 576 properties only have a subset of units covered under section 8.
>
> 201 properties have a override date that is past the forecast date.

## Pattern matching (7/28/25)



The first step of the updating process is to take subsets of properties based on a common factor in a effort to rule them out.

For example, if 100% of units are under section 8, and the S8_ExpDate is after 2030, that property can be safely consider as not at risk and marked off.  The code for which I checked for this is listed below. The follow script updated about 282 properties out of 1882. This means they are 1600 properties left. In other words this script eliminated17.625% of properties.

```python
import pandas as pd
from openpyxl import load_workbook

# This script is intended to take in the expiry use database forecast through 2030, and run a quick check for all properties that have
# are 100% section 8 units and have a S8_ExpDate after 2030 (the forecast date). The forecast date can easily be changed in the
# variable bellow

# Change this to the appropritae forecast date
forecast_date = "12/31/2030"

df = pd.read_excel("advance Preserve Through date to 2030 v3 draft.xlsx", na_values=["", " ", "  ", "NA", "N/A", "null"])
df.columns = df.columns.str.strip()




def check_section_8_expiry(df, forecast_date):
    # Convert the forecast date to a datetime object
    forecast_date = pd.to_datetime(forecast_date)
    # Ensure the S8_ExpDate column is in datetime format
    df['S8_ExpDate'] = pd.to_datetime(df['S8_ExpDate'], errors='coerce')
    df['TotalUnits'] = pd.to_numeric(df['TotalUnits'], errors='coerce').fillna(0).astype(int)
    df['S8_PBAUnits'] = pd.to_numeric(df['S8_PBAUnits'], errors='coerce')
    # Filter for properties that are 100% Section 8 units and have an expiry date after the forecast date
    section_8_expiry = df[(df['S8_ExpDate'] > forecast_date) & (df['S8_PBAUnits'] == df['TotalUnits'])]

    # Select relevant columns to return
    section_8_expiry = section_8_expiry['PropertyID']

    return section_8_expiry

def find_all_s8_properties(df):
    df['S8_PBAUnits'] = pd.to_numeric(df['S8_PBAUnits'], errors='coerce')
    s8_count = df['S8_PBAUnits'].notna().sum()
    return s8_count


matching_ids = check_section_8_expiry(df, forecast_date)

note = '100% of units are S.8, and S8 expires after forecast date'
status = 'done'
df.loc[df['PropertyID'].isin(matching_ids), 'Jameer notes'] = note
df.loc[df['PropertyID'].isin(matching_ids), 'Status after 1st cut (done, small, <>)'] = status


# Load the workbook and active sheet with openpyxl (preserves formatting)
workbook = load_workbook("advance Preserve Through date to 2030 v3 draft.xlsx")
sheet = workbook.active

# Get header row to find column indexes
header = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
prop_id_idx = header.index("PropertyID")
note_idx = header.index("Jameer notes")
status_idx = header.index("Status after 1st cut (done, small, <>)")

matching_counter = 0
# Iterate through rows and update notes/status for matching IDs
for row in sheet.iter_rows(min_row=2):  # assuming headers are in row 1
    property_id = row[prop_id_idx].value
    if property_id in matching_ids.values:
        matching_counter += 1
        row[note_idx].value = note
        row[status_idx].value = status
    
# Save to a new file (original formatting preserved)
workbook.save("S8_Forecast_Check_Formatted.xlsx")
print(f"\n {matching_counter} properties updated in the Excel file.")
print(f" Total properties with non-null S8_PBAUnits: {find_all_s8_properties(df)}")


```

Next, There is a field inside of the preservation sheet called "ForecastPreservedByProgram". This column represents the reason why the descion tree marked a property down as Not at RIsk.



