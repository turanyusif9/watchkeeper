from pdf2image import convert_from_path
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pdfminer.high_level import extract_text
import re


def get_ganttdata(image, days):
    """
    Extracts gantt data (worker schedules) from an image.

    Args:
        image (PIL.Image.Image): The input image.
        days (int): The number of days to extract gantt data for.

    Returns:
        pandas.DataFrame: A DataFrame 48x(number of days) containing the gantt data, 0 meaning rest, 1 meaning work hours.
    """
    # Convert the image to RGB
    image = image.convert('RGB')

    # Get the color value of the pixel at a certain position (x, y)
    y = 348
    all_x = [288,303,327,343,366,383,406,422,445,461,484,500,523,540,563,579,602,619,643,658,682,698,721,736,760,776,798,814,839,853,877,894,917,932,957,972,996,1012,1036,1051,1074,1090,1114,1130,1153,1170,1193,1208]

    gantts = []

    for i in range(days):  
        gantt = []
        for x in all_x:
            r0, g0, b0 = image.getpixel((x, y))

            if r0 != 255 and g0 != 255 and b0 != 255:
                gantt.append(1)
            else:
                gantt.append(0)
        y += 26
        gantts.append(gantt)

    gantts = pd.DataFrame(gantts)
    return gantts


def visualize(gantt):
    """
    Visualizes the given Gantt data (worker schedules).

    Parameters:
    gantt (numpy.ndarray): The Gantt data to be visualized.
    """
    
    # VISUALIZATION
    
    plt.figure(figsize=(7, 7))  # Increase the size of the figure
    ax = sns.heatmap(gantt, cmap='Blues', cbar=False, linewidths=.5)  # Use the Blues color map, remove the color bar, and add lines between the cells

    # Get the current y-tick labels
    yticks = ax.get_yticks()
    xticks = ax.get_xticks()

    # Add 1 to each y-tick label
    ax.set_yticklabels([int(ytick + 1) for ytick in yticks])
    ax.set_xticklabels([int(xtick/2) for xtick in xticks])
    
    plt.xlabel('Hours')  # Add an x-label
    plt.ylabel('Days')  # Add a y-label
    plt.title('Schedule')  # Add a title
    plt.show()
    return


def export(gantts):
    """
    Export the given Gantt outputs to an Excel file.
    Each Gantt output will be written to a separate sheet.

    Args:
        gantts (list): A list of Gantt charts represented as 2D arrays.
    """
    writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
    for i in range(len(gantts)):
        # Convert the 2D array to a DataFrame
        df = pd.DataFrame(gantts)

        # Write the DataFrame to a new sheet in the Excel file
        df.to_excel(writer, sheet_name='Page' + str(i + 1), index=False)

    writer.save()
    return


def read_pdf(file):
    """
    Reads a PDF file, extracts information, and returns the extracted data.

    Args:
        file (str): The path to the PDF file.

    Returns:
        tuple: A tuple containing the following elements:
            - gantts (list): A list of Gantt data extracted from the PDF.
            - header (pandas.DataFrame): A DataFrame containing header information extracted from the PDF.
            This includes the vessel, seafarer, position, period, start day, end day, and page.
            - anyhours (dict): A dictionary containing information about hours of rest in any 24 hours and any 7 days.
            - hoursworkedinaday (list): A list of lists containing the number of hours worked in each day.

    """
    # Convert the PDF to an image
    images = convert_from_path(file)

    # Extract text from the PDF
    text = extract_text(file)

    split_text = text.split('RECORD OF HOURS OF REST')[1:]

    header = {'Vessel': [], 'Seafarer': [], 'Position': [], 'Period': [], 'StartDay': [], 'EndDay': [], 'Page': []}
    anyhours = {'Hours of rest in any 24h': [], 'Hours of rest in any 7d': []}

    gantts = []

    for i in range(len(split_text)):
        vessel = re.search(r'Vessel:\n\n(.*?)\n', split_text[i])
        seafarer = re.search(r'Seafarer \(Full Name\):\n\n(.*?)\n', split_text[i])
        position = re.search(r'Position \(Rank\):\n\n(.*?)\n', split_text[i])
        period = re.search(r'\n(.*)\n\nPeriods', split_text[i])
        startday = re.search(r'Date\n(\d\d)/', split_text[i])
        endday = re.search(r'\n(\d\d)/\d\d/\d{4}\n\n', split_text[i])
        page = re.search(r'Page *(.*?) ', split_text[i])

        header['Vessel'].append(vessel.group(1))
        header['Seafarer'].append(seafarer.group(1))
        header['Position'].append(position.group(1))
        header['Period'].append(period.group(1))
        header['StartDay'].append(int(startday.group(1)))
        header['EndDay'].append(int(endday.group(1)))
        header['Page'].append(page.group(1))

        hoursofrest24 = re.search(r'in any 24h\n([\s\S]+?)\n\n', split_text[i])
        hoursofrest7 = re.search(r'in any 7d\n([\s\S]+?)\n\n', split_text[i])

        anyhours['Hours of rest in any 24h'].append(hoursofrest24.group(1).split('\n'))
        anyhours['Hours of rest in any 7d'].append(hoursofrest7.group(1).split('\n'))

        for key in anyhours.keys():
            for i in range(len(anyhours[key])):
                anyhours[key][i] = [float(x) if x != 'N/A' else 'N/A' for x in anyhours[key][i]]
        
        gantt = get_ganttdata(images[i], header['EndDay'][i] - header['StartDay'][i] + 1)
        gantts.append(gantt)
    
    hoursworkedinaday = []
    for gantt in gantts:
        hoursworked = []
        gantt = gantt.T
        
        for days in range(len(gantt.T)):
            hoursworked.append(gantt[days][gantt[days] == 1].count()/2)

        hoursworkedinaday.append(hoursworked)


    header = pd.DataFrame(header)
    return gantts, header, anyhours, hoursworkedinaday


def report_average_hours(header, hoursworkedinaday, month):
    """
    Generate an Excel report of the average hours worked by position in a given month.

    Args:
        header (dict): A dictionary containing header information.
        hoursworkedinaday (list): A list of hours worked in a day.
        month (str): The month for which the report is generated.

    Returns:
        tuple: A tuple containing the positions and their corresponding average hours worked.
    """
    hoursworkedinaday = pd.DataFrame(hoursworkedinaday)
    means = hoursworkedinaday[header['Period'] == month].T.mean()
    positions = header['Position'][header['Period'] == month]

    plt.figure(figsize=(20, 7))  # Increase the size of the figure
    # plt.scatter(positions, means, marker='o', linestyle='None', )  # Plot the means
    plt.bar(positions, means)  # Plot the means
    plt.xlabel('Position')  # Add an x-label
    plt.ylabel('Average Hours Worked')  # Add a y-label
    plt.title(f'Average Hours Worked by Position in {month} for Ship {header["Vessel"][0]}')  # Add a title
    plt.show()

    # make excel file with positions and means
    writer = pd.ExcelWriter(f'{header["Vessel"][0]} {month} average hours by positions.xlsx', engine='xlsxwriter')
    df = pd.DataFrame({'Position': positions, 'Average Hours Worked': means})
    df.to_excel(writer, sheet_name=f'{header["Vessel"][0]} {month}', index=False)
    writer.save()

    return positions, means


def report_overtime_monthly(header, hoursworkedinaday, limit):
    """
    Generate an Excel report of overtime hours worked by months.

    Args:
        header (dict): A dictionary containing header information.
        hoursworkedinaday (list): A list of hours worked per day.
        limit (int): The maximum number of hours considered as regular working hours.

    Returns:
        tuple: A tuple containing the months, overtime hours, and total hours worked.
    """
    months = [i for i in range (1, 13)]
    overtimes = []
    totalhoursworked = []
    hoursworkedinaday = pd.DataFrame(hoursworkedinaday)
    for month in months:
        hoursworked = hoursworkedinaday[header['Period'] == month].T
        totalhoursworked.append(hoursworked.sum().sum())
        overtime = hoursworked[hoursworked > limit] - limit
        overtimes.append(overtime.sum().sum())

    # plot overtimes
    plt.figure(figsize=(20, 7))  # Increase the size of the figure
    plt.bar(months, overtimes)  # Plot the means
    plt.xlabel('Month')  # Add an x-label
    plt.ylabel('Total Overtime')  # Add a y-label
    plt.title('Total Overtime by Month')  # Add a title
    plt.show()

    # make excel file with months and overtimes
    writer = pd.ExcelWriter(f'{header["Vessel"][0]} overtime by month over {limit} hours.xlsx', engine='xlsxwriter')
    df = pd.DataFrame({'Month': months, 'Overtime': overtimes, 'Total Hours Worked': totalhoursworked})
    df.to_excel(writer, sheet_name=f'{header["Vessel"][0]}', index=False)
    writer.save()

    return months, overtimes, totalhoursworked


def report_overtime_bypositions(header, hoursworkedinaday, limit):
    """
    Generate an Excel report of total overtime worked by positions.

    Args:
        header (pandas.DataFrame): DataFrame containing header information.
        hoursworkedinaday (pandas.DataFrame): DataFrame containing hours worked per day.
        limit (int): The maximum number of hours considered as regular working hours.

    Returns:
        tuple: A tuple containing the positions, overtimes, and total hours worked.

    """
    positions = header['Position'].unique()
    totalhoursworked = []
    overtimes = []
    hoursworkedinaday = pd.DataFrame(hoursworkedinaday)
    
    for position in positions:
        hoursworked = hoursworkedinaday[header['Position'] == position].T
        totalhoursworked.append(hoursworked.sum().sum())
        overtime = hoursworked[hoursworked > limit] - limit
        overtimes.append(overtime.sum().sum())


    # plot overtimes
    plt.figure(figsize=(20, 7))  # Increase the size of the figure
    plt.bar(positions, overtimes)  # Plot the means
    plt.xlabel('Position')  # Add an x-label
    plt.ylabel('Total Overtime')  # Add a y-label
    plt.title('Total Overtime by Positions')  # Add a title
    plt.show()

    # make excel file with months and overtimes
    writer = pd.ExcelWriter(f'{header["Vessel"][0]} overtime by positions over {limit} hours.xlsx', engine='xlsxwriter')
    df = pd.DataFrame({'Position': positions, 'Overtime': overtimes, 'Total Hours Worked': totalhoursworked})
    df.to_excel(writer, sheet_name=f'{header["Vessel"][0]}', index=False)
    writer.save()

    return positions, overtimes, totalhoursworked


def report_overtime_bypositions_monthly(header, hoursworkedinaday, limit):
    """
    Generate an Excel report of overtime by positions and months.

    Args:
        header (pandas.DataFrame): DataFrame containing header information.
        This includes the vessel, seafarer, position, period, start day, end day, and page.
        hoursworkedinaday (pandas.DataFrame): DataFrame containing hours worked per day.
        limit (int): The limit of hours for overtime calculation.

    Returns:
        tuple: A tuple containing two DataFrames - monthly_overtimes and monthly_totalhoursworked.
            - monthly_overtimes: DataFrame with overtime hours by positions and months.
            - monthly_totalhoursworked: DataFrame with total hours worked by positions and months.
    """

    months = [i for i in range (1, 13)]
    positions = header['Position'].unique()
    
    hoursworkedinaday = pd.DataFrame(hoursworkedinaday)

    monthly_totalhoursworked = []
    monthly_overtimes = []

    for month in months:
        totalhoursworked = []
        overtimes = []
        for position in positions:
            hoursworked = hoursworkedinaday[(header['Period'] == month) & (header['Position'] == position)]
            totalhoursworked.append(hoursworked.sum().sum())
            overtime = hoursworked[hoursworked >= limit] - limit
            overtimes.append(overtime.sum().sum())
        monthly_totalhoursworked.append(totalhoursworked)
        monthly_overtimes.append(overtimes)

    monthly_overtimes = pd.DataFrame(monthly_overtimes).T
    monthly_totalhoursworked = pd.DataFrame(monthly_totalhoursworked).T

    monthly_overtimes.columns = months
    monthly_overtimes.index = positions

    monthly_totalhoursworked.columns = months
    monthly_totalhoursworked.index = positions

    # make excel file with months and overtimes
    writer = pd.ExcelWriter(f'{header["Vessel"][0]} overtime by positions and months over {limit} hours.xlsx', engine='xlsxwriter')
    monthly_overtimes.to_excel(writer, sheet_name=f'Overtimes')
    monthly_totalhoursworked.to_excel(writer, sheet_name=f'Total Hours Worked')
    writer.save()

    return monthly_overtimes, monthly_totalhoursworked


def clean_anyhours(anyhours):  
    """
    Cleans the 'Hours of rest' values in the 'anyhours' dictionary of 'N/A' values 
    and sets them as 24 and 168 respectively meaning no work hours.

    Args:
        anyhours (dict): A dictionary containing the 'Hours of rest' values.
    """

    # replace 'N/A' with 24 in anyhours['Hours of rest in any 24h']
    for i in range(len(anyhours['Hours of rest in any 24h'])):
        for j in range(len(anyhours['Hours of rest in any 24h'][i])):
            if anyhours['Hours of rest in any 24h'][i][j] == 'N/A':
                anyhours['Hours of rest in any 24h'][i][j] = 24

    # replace 'N/A' with 77 in anyhours['Hours of rest in any 7d']
    for i in range(len(anyhours['Hours of rest in any 7d'])):
        for j in range(len(anyhours['Hours of rest in any 7d'][i])):
            if anyhours['Hours of rest in any 7d'][i][j] == 'N/A':
                anyhours['Hours of rest in any 7d'][i][j] = 168


def plot_violations(anyhours, header, dayorweek, limit):
    """
    Plot the total violations by month.

    Parameters:
    anyhours (DataFrame): The data containing the hours information.
    header (dict): The header information.
    dayorweek (str): The column name for the day or week information,
    either 'Hours of rest in any 24h' or 'Hours of rest in any 7d'.
    limit (int): The limit for violations.

    Returns:
    list: The list of violations for each month.
    """
    months = [i for i in range (1, 13)]

    # for month in months:
    violations = []
    anyhours = pd.DataFrame(anyhours)
    for month in months:
        hoursrested = pd.Series(anyhours[dayorweek][header['Period'] == month])
        violation = 0
        for page in hoursrested:
            page = pd.Series(page)
            violation += page[page < limit].count()
        violations.append(violation)

    # plot violations
    plt.figure(figsize=(20, 7))  # Increase the size of the figure
    plt.bar(months, violations)  # Plot the means
    plt.xlabel('Month')  # Add an x-label
    plt.ylabel('Total Violations')  # Add a y-label
    plt.title('Total Violations by Month')  # Add a title
    plt.show()

    return violations


def report_violations(anyhours, header):
    """
    Generates an Excel report of violations based on the given data.

    Args:
        anyhours (list): List of hours of rest data.
        header (dict): Dictionary containing header information.

    Returns:
        tuple: A tuple containing the following:
            - violationsinday (list): List of violations in any 24 hours.
            - violationsinweek (list): List of violations in any 7 days.
            - totaldaysworked_list (list): List of total days worked.
    """

    clean_anyhours(anyhours)
    months = [i for i in range (1, 13)]
    violationsinday = plot_violations(anyhours, header, 'Hours of rest in any 24h', 10)
    violationsinweek = plot_violations(anyhours, header, 'Hours of rest in any 7d', 77)

    totaldaysworked_list = []
    for month in months:
        endday = header['EndDay'][header['Period'] == month]
        startday = header['StartDay'][header['Period'] == month]
        totaldaysworked = (endday - startday + 1).sum()
        totaldaysworked_list.append(totaldaysworked)

    # make excel file with months and violations
    writer = pd.ExcelWriter(f"{header['Vessel'][0]} violations by month.xlsx", engine='xlsxwriter')
    df = pd.DataFrame({'Month': months, 'Violations in any 24h': violationsinday, 'Violations in any 7d': violationsinweek, 'Total Days Worked': totaldaysworked_list})
    df.to_excel(writer, sheet_name=f"{header['Vessel'][0]}", index=False)
    writer.save()
    return violationsinday, violationsinweek, totaldaysworked_list


def mean_std_per_seafarer(header, hoursworkedinaday):
    """
    Calculate the mean, standard deviation, and number of days worked for each seafarer and generate an Excel report.

    Args:
        header (pandas.DataFrame): DataFrame containing the header information.
        hoursworkedinaday (list): List of lists containing the hours worked by each seafarer.

    Returns:
        tuple: A tuple containing three lists - means, stds, and numberofdays.
            - means (list): List of mean values for each seafarer.
            - stds (list): List of standard deviation values for each seafarer.
            - numberofdays (list): List of the number of days worked for each seafarer.
    """

    hoursworkedinaday = pd.DataFrame(hoursworkedinaday)
    means = []
    stds = []
    numberofdays = []
    for seafarer in header['Seafarer'].unique():
        # merge the listof lists hoursworkedinaday[header['Seafarer'] == seafarer] into a single list
        merged_list = [item for sublist in hoursworkedinaday[header['Seafarer'] == seafarer].values for item in sublist]
        merged_list = pd.Series(merged_list)

        means.append(merged_list.mean())
        stds.append(merged_list.std())

        numberofdays.append(merged_list.count())

    # create excel file with seafarer, mean, and std
    writer = pd.ExcelWriter(f'{header["Vessel"][0]} mean and std of each seafarer.xlsx', engine='xlsxwriter')
    df = pd.DataFrame({'Seafarer': header['Seafarer'].unique(), 'Mean': means, 'Std': stds, 'Number of Days': numberofdays})
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()
    return means, stds, numberofdays
