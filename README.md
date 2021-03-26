# sw_report
A program that takes two PCScale-generated reports and converts them to a csv file that can be used to auto-populate a tedious fillable pdf form we are required to submit monthly.

As of right now, it's ugly, it's delicate, and it's extremely verbose... but it works. This is my first useful piece of code - it saves me a humble 90 minutes per month. As I learn more, I intend to revisit this program to improve the efficiency and readability of the code.

### What does it do?
Each month, we are required to fill out and submit the Solid Waste Facility Monthly Disposal and Materials Recovery Report to the New Jersey Department of Environmental Protection. This is essentially a fillable pdf form detailing how much of each type of solid waste we received from each municipality of each County in our area, as well as our total inbound waste and a few other figures.

This information is all readily available through our billing & data software, PCScale. My original process involved running and printing two reports from PCScale, then plugging the information into the fillable pdf form manually. The process took about 90 minutes, on average.

This program parses those same two reports (saved as csv files) and converts the data into a new csv file formatted so that it can be imported directly into the blank form in Foxit Reader (a free pdf reader with some form functionality). It's designed with my coworkers in mind, so the dialogue box really holds the user's hand through the process of running the reports, pointing the program to the right folder, importing the csv to the form, etc.

Overall, the 90-minutes required to fill out the report have been reduced to under five minutes.

### How does it work?
1. Program opens in terminal and begins dialogue by explaining how it works and how to run the reports it needs. Asks the user to continue once reports are ready.
2. User is asked for a few pieces of information, including report locations, preferred location of finished csv file, form header info, etc. This information is saved in a dictionary for later use.
3. Program takes information from the two source csv files and sorts it based on waste type, municipality, and county. As of right now, this is done through nested dictionaries and lists.
4. Program then uses this "dictionary of dictionaries" to create a list ordered such that each piece of information lands in the correct form field. (This was the most challenging part since many fields may or may not be filled in, depending on the data itself).
5. The user is asked if they'd like to review the raw list in the terminal before proceeding.
6. Once the user is ready, a new csv file is written and saved to the folder (and filename) specified by the user. The first row is a pre-existing list of all fields in the form, and the second row is the list from Step 4. This simple format allows Foxit Reader to import the csv and autopopulate the blank form.

### What still needs to be done?
1. Shorten functions and make the code less verbose in general.
2. Utilize defined functions more efficiently
3. Enhance Readability
4. Make the program's parsing ability more robust... ideally all NJ facilities could use the program. (Right now it's very specific to our facility).
5. Figure out how to get the program to run without UAC interruption.
6. Is it possible for the program to run the reports and populate the blank form by itself?
7. General cleanup and better in-code commenting.
8. Better exception-handling.
