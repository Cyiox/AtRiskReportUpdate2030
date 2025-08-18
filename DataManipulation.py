import pandas as pd
from pandas import DataFrame, Series
from pandas._libs.tslibs.timestamps import Timestamp
from openpyxl import load_workbook
import re
from typing import Optional
import matplotlib.pyplot as plt



# This file contains code for running various checks for the CEDAC At-Risk Forecast report. It was built for the 2030 report, but contains functions that can be used for other years.
# Written by Jameer Gomez-Santos, jgomez-santos@cedac.org, 2025



def S8_Expiring_Pre_Forecast(df: DataFrame, forecast_date: str) -> DataFrame:
    '''Returns a Dataframe of all properties from the inputed dataframe that consist of 100% S8 units and have an S8 expiry date after the forcast date.

    Args:
        df (DataFrame): Dataframe all the forecast data. Must have the columns 'S8_ExpDate', 'S8_PBAUnits', and 'TotalUnits'.
        forecast_date (str): The cutoff date for the forecast in MM/DD/YYYY format.

    Returns:
        DataFrame: Subset of input Dataframe containing properties with 100% Section 8 units and an S8 expiry date after the forecast date.
    '''
    # Convert the forecast date to a datetime object
    cutOff_DT:Timestamp = pd.to_datetime(forecast_date)
    # Ensure the S8_ExpDate column is in datetime format
    df['S8_ExpDate'] = pd.to_datetime(df['S8_ExpDate'], errors='coerce')
    df['TotalUnits'] = pd.to_numeric(df['TotalUnits'], errors='coerce').fillna(0).astype(int)
    df['S8_PBAUnits'] = pd.to_numeric(df['S8_PBAUnits'], errors='coerce')
    # Filter for properties that are 100% Section 8 units and have an expiry date after the forecast date
    section_8_expiry:DataFrame = df[(df['S8_ExpDate'] > cutOff_DT) & (df['S8_PBAUnits'] == df['TotalUnits'])]
    return section_8_expiry

def S8_Expiring_Post_Forecast(df: DataFrame, forecast_date: str) -> DataFrame:
    '''Returns a Dataframe of all properties from the inputed dataframe that consist of 100% S8 units and have an S8 expiry date after the forcast date.

    Args:
        df (DataFrame): Dataframe all the forecast data. Must have the columns 'S8_ExpDate', 'S8_PBAUnits', and 'TotalUnits'.
        forecast_date (str): The cutoff date for the forecast in MM/DD/YYYY format.

    Returns:
        DataFrame: Subset of input Dataframe containing properties with 100% Section 8 units and an S8 expiry date after the forecast date.
    '''
    # Convert the forecast date to a datetime object
    cutOff_DT:Timestamp = pd.to_datetime(forecast_date)
    # Ensure the S8_ExpDate column is in datetime format
    df['S8_ExpDate'] = pd.to_datetime(df['S8_ExpDate'], errors='coerce')
    df['TotalUnits'] = pd.to_numeric(df['TotalUnits'], errors='coerce').fillna(0).astype(int)
    df['S8_PBAUnits'] = pd.to_numeric(df['S8_PBAUnits'], errors='coerce')
    # Filter for properties that are 100% Section 8 units and have an expiry date after the forecast date
    section_8_expiry:DataFrame = df[(df['S8_ExpDate'] < cutOff_DT) & (df['S8_PBAUnits'] == df['TotalUnits'])]
    return section_8_expiry

def get_all_s8_properties(df: DataFrame) -> DataFrame:
    """
    Returns a DataFrame of all properties that have at least one S8_PBAUnit.

    Args:
        df (DataFrame): The input DataFrame containing property data, including the 'S8_PBAUnits' column.

    Returns:
        DataFrame: Subset of the input DataFrame where S8_PBAUnits is not null and greater than 0.
    """
    df['S8_PBAUnits'] = pd.to_numeric(df['S8_PBAUnits'], errors='coerce')
    s8_properties = df[df['S8_PBAUnits'].notna() & (df['S8_PBAUnits'] > 0)]
    return s8_properties

def find_all_overrides(df: DataFrame, forecast_date: Optional[str] = None) -> DataFrame:
    """
    Returns a DataFrame of all properties with an override date.
    If a forecast_date is provided, only properties with override dates after that date are included.

    Args:
        df (DataFrame): Input DataFrame containing property data.
        forecast_date (Optional[str]): Optional cutoff date (MM/DD/YYYY) for filtering override dates.

    Returns:
        DataFrame: Subset of DataFrame with 'PropertyID' and 'OverrideDate'.
    """
    pattern = r"At Risk Date Override:\s*(\d{1,2}/\d{1,2}/\d{4})"
    df['OverrideDate'] = df['ForecastPreservedByProgram'].astype(str).str.extract(pattern)
    df['OverrideDate'] = pd.to_datetime(df['OverrideDate'], errors='coerce')

    if forecast_date:
        cutoff_dt = pd.to_datetime(forecast_date)
        overrides = df[df['OverrideDate'] > cutoff_dt]
    else:
        overrides = df[df['OverrideDate'].notna()]
    
    return overrides

def extract_propertyIDs(df: DataFrame) -> Series:
    """
    Extracts unique PropertyIDs from the DataFrame.

    Args:
        df (DataFrame): The input DataFrame containing property data.

    Returns:
        Series: A Series of unique PropertyIDs. A series is an individual column of a Dataframe.
    """
    return df['PropertyID']



def get_larger_date(date1: str, date2: str) -> Optional[str]:
    """
    Returns the later of two dates in MM/DD/YYYY format.
    If both dates are invalid or missing, returns None.
    """
    dateTime1 = pd.to_datetime(date1, errors='coerce')
    dateTime2 = pd.to_datetime(date2, errors='coerce')

    if pd.isna(dateTime1) and pd.isna(dateTime2):
        return None
    if pd.isna(dateTime1):
        return dateTime2.strftime('%m/%d/%Y')  # type: ignore
    if pd.isna(dateTime2):
        return dateTime1.strftime('%m/%d/%Y')

    return max(dateTime1, dateTime2).strftime('%m/%d/%Y')

def S8_expDate_vs_overrideDate(S8Propertys: DataFrame, overridenProperties: DataFrame) -> DataFrame:
    """
    Returns a DataFrame of properties where the Section 8 expiration date is later than the override date.

    Args:
        S8Propertys (DataFrame): DataFrame with at least 'PropertyID' and 'S8_ExpDate' columns.
        overridenProperties (DataFrame): DataFrame with 'PropertyID' and 'OverrideDate'.

    Returns:
        DataFrame: Subset of merged DataFrame where S8_ExpDate > OverrideDate.
    """
    # Merge on PropertyID
    merged: DataFrame = pd.merge(S8Propertys, overridenProperties, on="PropertyID", how="inner")

    # Ensure both date columns are in datetime format
    merged['S8_ExpDate'] = pd.to_datetime(merged['S8_ExpDate'], errors='coerce')
    merged['OverrideDate'] = pd.to_datetime(merged['OverrideDate'], errors='coerce')

    # Filter where Section 8 Expiration Date is later than the override date
    result: DataFrame = merged[merged['S8_ExpDate'] > merged['OverrideDate']]

    return result

def extract_Column_With_IDs(df: DataFrame, columnName: str) -> DataFrame:
    '''Extracts single given column along with another column containing its 'PropertyID' from a DataFrame object. Assumes existance of "PropertyID" column.

    Args:
        df (DataFrame): DataFrame to extract a coulmn from.
        columnName (str): Exact name of the coulumn to extract.

    Raises:
        ValueError: Returns a ValueError if the coulumn is not found in the DataFrane.

    Returns:
        Series: A singlar column of the dataframe.
    '''
    if columnName not in df.columns:
        raise ValueError(f"Column '{columnName}' not found in DataFrame.")
    else:
         return df[['PropertyID', columnName]].copy()

def get_Expired_S8(df: DataFrame, forecast_date: str) -> DataFrame:
    """
    Returns full rows of all properties where the S8_ExpDate is earlier than the forecast date (i.e., already expired).

    Args:
        df (DataFrame): The input DataFrame containing property data.
        forecast_date (str): The forecast cutoff date in MM/DD/YYYY format.

    Returns:
        DataFrame: Subset of original DataFrame with all columns, for expired S8 properties.
    """
    # Convert columns to proper types
    df['S8_ExpDate'] = pd.to_datetime(df['S8_ExpDate'], errors='coerce')
    
    # Convert forecast_date to datetime
    forecast_dt = pd.to_datetime(forecast_date, errors='coerce')
    
    # Filter full DataFrame
    expired_df = df[df['S8_ExpDate'] < forecast_dt]
    
    return expired_df

def get_all_PRAC(df: DataFrame, forecast_date: str, numOfUnits: Optional[int] = None) -> DataFrame:
    """Returns a DataFrame of all PRAC properties that expire before the forecast date, optionally filtered being greater than a number of units.

    Args:
        df (DataFrame): Dataframe containing the atRisk forecast data.
        numOfUnits (Optional[int]): Optional parameter to that will only return PRAC properties with a number of units greater or equal to this number when provided.
        forecast_date (str): Forecast date of at risk forecast in MM/DD/YYYY format.

    Returns:
        DataFrame: A DataFrame containing all PRAC properties that expire before the forecast date. Optionally filtered by being greater than number of units.
    """    
    # Getting all S8 Properties 
    RiskS8 = S8_Expiring_Post_Forecast(df,forecast_date = forecast_date)
    PRACRiskS8 = RiskS8[RiskS8['ForecastPreservedByProgram'].str.contains("PRAC", na=False, case=False)]
    if numOfUnits is not None:
        PRACRiskS8 = PRACRiskS8[PRACRiskS8['TotalUnits'] >= numOfUnits]
        return PRACRiskS8
    else:
        return PRACRiskS8


def get_All_By_Forecast(df: DataFrame, forecast_date: str, forecastFilter: str, numOfUnits: Optional[int] = None) -> DataFrame:
    """Returns a DataFrame of all properties based on a matching a string with the 'ForecastPreservedByProgram' column, that expire before the forecast date, optionally filtered by number of units.

    Args:
        df (DataFrame): Dataframe containing the atRisk forecast data.
        forecastFilter (str): Filter to apply to the 'ForecastPreservedByProgram', Filteed based on if that string is contained in the row of that column.
        forecast_date (str): Forecast date of at risk forecast in MM/DD/YYYY format.
        numOfUnits (Optional[int]): Optional parameter to that will only return PRAC properties with a number of units greater or equal to this number when provided.

    Returns:
        DataFrame: A DataFrame containing all PRAC properties that expire before the forecast date. Optionally filtered by being greater than number of units.
    """    
    # Getting all S8 Properties 
    RiskS8 = S8_Expiring_Post_Forecast(df,forecast_date= forecast_date)
    PRACRiskS8 = RiskS8[RiskS8['ForecastPreservedByProgram'].str.contains(forecastFilter, na=False, case=False)]
    if numOfUnits is not None:
        PRACRiskS8 = PRACRiskS8[PRACRiskS8['TotalUnits'] >= numOfUnits]
        return PRACRiskS8
    else:
        return PRACRiskS8
    
def Preserve_Via_Units(df: DataFrame, numOfUnits: int, statusColumnName: str) -> DataFrame:
    '''Returns a Dataframe of all the properties in a dataframe under a given amount of units, then sets status column to PRESERVED.

    Args:
        df (DataFrame): Dataframe containing properties including a 'TotalUnits' column.
        numOfUnits (int): Property must have less than or equal to this number to be preserved.
        statusColumnName (str): Name of the column used to track the status of the property.

    Returns:
        DataFrame: returns a DataFrame of all properties with less than or equal to the given number of units, with the status column set to PRESERVED.
    '''    
    df = df.copy()
    df.loc[df['TotalUnits'] <= numOfUnits, statusColumnName] = 'PRESERVED'
    return df

def create_pie_chart(df: DataFrame, columnName: str, topN: Optional[int] = None) -> None:
    """
    Plots a pie chart of the top N most frequent values in a column.
    If top_n is not specified, all values are shown with no 'Other' grouping.

    Args:
        df (DataFrame): Input DataFrame.
        columnName (str): Name of the column to analyze.
        top_n (Optional[int]): Number of top values to display. If None, show all.
    """
    # Clean and normalize the column
    clean_series: Series = df[columnName].astype(str).str.strip().str.title()
    value_counts = clean_series.value_counts()

    # Determine if we need to group into "Other"
    if topN is not None and topN < len(value_counts):
        top_values = value_counts.head(topN)
        other_total = value_counts.iloc[topN:].sum()
        final_counts = pd.concat([top_values, pd.Series({'Other': other_total})])
    else:
        final_counts = value_counts

    # Plot the pie chart
    plt.figure(figsize=(8, 8))
    plt.pie(final_counts, labels=final_counts.index, autopct='%1.1f%%',
            startangle=90, textprops={'fontsize': 9})
    title = f"Most Overriden Citys" if topN else f"All {columnName.title()}s"
    plt.title(title, fontsize=14)
    plt.axis('equal')  # Keep pie chart circular
    plt.tight_layout()
    plt.show()
    
    

def plot_date_cutoff_bar(
    df: DataFrame,
    date_column: str,
    cutoff_date: str,
    title: str = "Date Distribution"
) -> None:
    """
    Plots a bar chart showing how many entries in a date column are:
    - Before a cutoff date
    - After a cutoff date
    - Missing (null)

    Args:
        df (DataFrame): Input DataFrame.
        date_column (str): Name of the column containing dates.
        cutoff_date (str): The cutoff date in 'MM/DD/YYYY' or ISO format.
        title (str): Title of the chart.
    """
    # Convert column to datetime, coerce errors to NaT
    dates = pd.to_datetime(df[date_column], errors="coerce")
    cutoff = pd.to_datetime(cutoff_date)

    # Count categories
    before = (dates < cutoff).sum()
    after = (dates > cutoff).sum()
    missing = dates.isna().sum()

    counts = {
        f"Before {cutoff.strftime('%m/%d/%Y')}": before,
        f"After {cutoff.strftime('%m/%d/%Y')}": after,
        "Not S.8": missing
    }

    # Plot bar chart
    plt.figure(figsize=(7, 5))
    bars = plt.bar(counts.keys(), counts.values(), color=["#6baed6", "#74c476", "#fdd0a2"])
    plt.title(title, fontsize=14)
    plt.ylabel("Number of Records")
    plt.xticks(rotation=15)
    plt.tight_layout()

    # Add labels on bars
    for bar in bars:
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2.0, height, f"{int(height)}", ha='center', va='bottom', fontsize=10)

    plt.show()


    
if __name__ == "__main__":
    print("Script Started")
    df: DataFrame = pd.read_excel("2030 Forecast E-U-R Draft -Master.xlsx", na_values=["", " ", "  ", "NA", "N/A", "null"])
    df.columns = df.columns.str.strip()
    print("Dataframe Loaded.")

    overriden = find_all_overrides(df)
    print("Saving All292 to Excel sheet...")
    print("Saved successfully!")
    print(overriden)

    newDf = pd.read_excel("2030 Forecast E-U-R Draft -Master.xlsx", sheet_name="All Active Mortage Properties")
    plot_date_cutoff_bar(
        df=newDf,
        date_column="S8_ExpDate",
        cutoff_date=forecast_date,
        title="S.8 Expiration Distribution"
    )

























