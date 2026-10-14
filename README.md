🚗 BMW Dataset — Excel Text Functions & Data Cleaning Practice

📌 Project Overview

This project is a hands-on Microsoft Excel Text Functions practice project using a BMW dataset covering the years 2010–2024.

The purpose of this project was to practice Excel text functions by solving specific data-cleaning and text-manipulation questions rather than simply learning the syntax of each function.

I worked with BMW vehicle data and applied functions such as "UPPER", "LOWER", "PROPER", "LEN", "LEFT", "RIGHT", "MID", "TRIM", "FIND", "SEARCH", "SUBSTITUTE", "REPLACE", "CONCATENATE", and "TEXTJOIN".

I also practiced combining functions to solve more realistic data-cleaning problems.

---

🎯 Questions & Tasks Solved

1️⃣ UPPER Function

📝 Question / Task

«Convert all the model names into uppercase.»

Function Used

=UPPER(Model_Cell)

What I Practiced

- Converting model names to uppercase
- Standardizing text format
- Applying the formula across the dataset

---

2️⃣ LOWER Function

📝 Question / Task

«Convert all the model names into lowercase.»

Function Used

=LOWER(Model_Cell)

What I Practiced

- Converting model names to lowercase
- Maintaining consistent text formatting

---

3️⃣ PROPER Function

📝 Question / Task

«Convert all the country names into proper case.»

Function Used

=PROPER(Country_Cell)

What I Practiced

- Standardizing country names
- Converting inconsistent capitalization into readable format

---

4️⃣ LEN Function

📝 Question / Task

«Find the length of all the model names.»

Function Used

=LEN(Model_Cell)

What I Practiced

- Counting the number of characters in model names
- Understanding how Excel calculates text length

---

5️⃣ LEFT Function

📝 Question / Task

«Using the LEFT function, extract the specified number of characters from the beginning of all the country names.»

Function Used

=LEFT(Country_Cell,Number_of_Characters)

What I Practiced

- Extracting characters from the beginning of text
- Working with a specified character count

---

6️⃣ RIGHT Function

📝 Question / Task

«Using the RIGHT function, extract the specified number of characters from the end of all the country names.»

Function Used

=RIGHT(Country_Cell,Number_of_Characters)

What I Practiced

- Extracting characters from the end of text
- Understanding text position from the right side

---

7️⃣ MID Function

📝 Question / Task

«Using the MID function, extract the middle portion of the combined Year, Model and Fuel Type data.»

First, I created a combined value using "TEXTJOIN".

Example:

2010-3567-Petrol

Then I used "MID" to extract the required middle portion.

Function Used

=MID(Combined_Cell,Start_Position,Number_of_Characters)

What I Practiced

- Extracting characters from the middle of a text string
- Understanding starting position and character count
- Using one function's output as another function's input

---

8️⃣ TRIM Function

📝 Question / Task

«Remove unnecessary spaces from the country names.»

Function Used

=TRIM(Country_Cell)

What I Practiced

- Removing unwanted spaces
- Cleaning inconsistent text data
- Preparing data for further analysis

---

9️⃣ FIND Function

📝 Question / Task

«Use the FIND function to find the position of the specified character in the country names.»

Function Used

=FIND("a",Country_Cell)

Important Concept Practiced

"FIND" is case-sensitive.

This exercise helped me understand how case sensitivity affects text searching in Excel.

---

🔟 SEARCH Function

📝 Question / Task

«Use the SEARCH function to find the position of the specified character in the country names.»

Function Used

=SEARCH("A",Country_Cell)

Important Concept Practiced

"SEARCH" is not case-sensitive.

I practiced the difference between:

FIND → Case-sensitive
SEARCH → Not case-sensitive

---

1️⃣1️⃣ SUBSTITUTE Function

📝 Question / Task

«The company made a mistake in the color pattern data. The color was entered as "Red", but it should be "Maroon". Use the SUBSTITUTE function to correct the data.»

Function Used

=SUBSTITUTE(Color_Cell,"Red","Maroon")

Nested Function Practice

I also practiced combining "TRIM" and "SUBSTITUTE":

=SUBSTITUTE(TRIM(Color_Cell),"Red","Maroon")

What I Practiced

- Correcting incorrect text values
- Replacing specific text
- Combining functions for data cleaning
- Handling inconsistent data

---

1️⃣2️⃣ REPLACE Function

📝 Question / Task

«Replace the model "3567" with "SRS" using the REPLACE function, where the series itself needs to be represented using capital letters "SRS".»

Function Used

=REPLACE(Model_Cell,Start_Position,Number_of_Characters,"SRS")

What I Practiced

The important concept here was understanding the difference between:

REPLACE → Replaces text based on position
SUBSTITUTE → Replaces specific matching text

---

1️⃣3️⃣ CONCATENATE Function

📝 Question / Task

«Using the CONCATENATE function, combine the Model and Fuel Type.»

Function Used

=CONCATENATE(Model_Cell," - ",Fuel_Type_Cell)

Example

3567 - Petrol

What I Practiced

- Combining information from multiple columns
- Creating a new text field from existing data

---

1️⃣4️⃣ TEXTJOIN Function

📝 Question / Task

«Using the TEXTJOIN function, join the Year, Model and Fuel Type.»

Function Used

=TEXTJOIN("-",TRUE,Year_Cell,Model_Cell,Fuel_Type_Cell)

Example

2010-3567-Petrol

What I Practiced

- Combining multiple columns
- Using delimiters
- Creating structured combined text
- Handling multiple values with one function

---

1️⃣5️⃣ TEXTSPLIT Function

📝 Question / Task

«Split the combined Year, Model and Fuel Type data using the TEXTSPLIT function.»

Status

⚠️ Not performed

The Excel version I am currently using does not provide the "TEXTSPLIT" function, so I skipped this exercise rather than treating an unavailable function as completed.

What I Learned

"TEXTSPLIT" can be used to split text based on delimiters, such as:

2010-3567-Petrol

into separate values.

---

🔥 Combined Function Challenge

After practicing the individual functions, I started combining functions to solve more practical problems.

📝 Task

«Combine the Year, Model and Fuel Type using TEXTJOIN, then use MID to extract the middle/model portion from the combined value.»

Step 1 — TEXTJOIN

=TEXTJOIN("-",TRUE,Year_Cell,Model_Cell,Fuel_Type_Cell)

Output:

2010-3567-Petrol

Step 2 — MID

=MID(Combined_Cell,6,4)

Output:

3567

This exercise helped me understand how multiple Excel functions can work together to solve a single data problem.

---

📊 Functions Covered

Function| Main Skill Practiced
"UPPER"| Uppercase conversion
"LOWER"| Lowercase conversion
"PROPER"| Proper-case formatting
"LEN"| Character counting
"LEFT"| Extracting characters from the beginning
"RIGHT"| Extracting characters from the end
"MID"| Extracting characters from the middle
"TRIM"| Removing unnecessary spaces
"FIND"| Case-sensitive text searching
"SEARCH"| Case-insensitive text searching
"SUBSTITUTE"| Replacing matching text
"REPLACE"| Replacing text based on position
"CONCATENATE"| Combining text
"TEXTJOIN"| Joining multiple values
"TEXTSPLIT"| Reviewed but not performed due to Excel version

---

💼 Data Analyst Skills Demonstrated

Through this practice project, I worked on:

- Excel data cleaning
- Text manipulation
- Data standardization
- Text extraction
- Data transformation
- Error correction
- Pattern searching
- Combining data fields
- Nested Excel functions
- Problem-solving using formulas
- Preparing data for analysis

---

🧠 Key Learning

This project helped me move beyond simply memorizing Excel functions.

For every function, I focused on:

Understanding the problem → Choosing the appropriate function → Writing the formula → Applying it to the dataset → Checking the result

The main takeaway was learning which Excel function is appropriate for a specific data-cleaning problem.

---

🗂️ Dataset

Dataset: BMW Vehicle Dataset
Data Period: 2010–2024
Tool: Microsoft Excel
Project Type: Excel Text Functions & Data Cleaning Practice

---

🚀 Future Practice

I plan to extend this project by practicing additional Excel functions and concepts, including:

- "IF"
- "IFS"
- "IFERROR"
- "COUNTIF"
- "SUMIF"
- "AVERAGEIF"
- "XLOOKUP"
- "VLOOKUP"
- "INDEX" + "MATCH"
- Date & Time functions
- Logical functions
- Statistical functions
- PivotTables
- Excel dashboards
- Data visualization

---

⭐ Project Goal

«Practice Excel by solving data problems, not just memorizing formulas.»

This project is part of my preparation for Data Analyst / Business Analyst roles and my ongoing development of practical data-cleaning and analytical skills using Microsoft Excel.
