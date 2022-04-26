# Import os and system modules
import os
import sys
# Import csv module
import csv
# Import datetime module
import datetime

# Get the execution path
path = os.path.dirname(os.path.realpath(__file__))
# Assign the csv file name without the extension
csv_file = "sample3.csv"
# Add the current timestamp to the csv_file name
csv_file = csv_file + "_" + datetime.datetime.now().strftime("%Y%m%d%H%M%S") + ".csv"

# Join the csv_file to the path
csv_file = os.path.join(path, csv_file)

# Create a function to create the csv file
def create_csv(csv_file):
    # Open the csv file in write mode
    with open(csv_file, 'w', newline='') as csvfile:
        # Create a csv writer object
        csvwriter = csv.writer(csvfile)
        # Write the header row
        csvwriter.writerow(['Name', 'Age', 'Job', 'City'])
        # Create random data to the file:
        # Write the data rows
        csvwriter.writerow(['John', '25', 'Engineer', 'New York'])
        csvwriter.writerow(['Mary', '29', 'Doctor', 'New York'])
        csvwriter.writerow(['Dave', '25', 'QA', 'New York'])
        csvwriter.writerow(['Mary', '27', 'Doctor', 'New York'])
        csvwriter.writerow(['Willis', '25', 'DEV', 'New York'])
        csvwriter.writerow(['Mary', '24', 'Doctor', 'Ho Chi Minh'])

# Call the function to create the csv file
create_csv(csv_file)
# Show the created csv file
print("Created csv file: " + csv_file)
