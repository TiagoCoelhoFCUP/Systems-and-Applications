# Student Grade Management

In this project we implemented an Excel book to manage the students' grades in a course in which the assessment is made up of several components: a quiz, a project proposal, a final report project, a final report, the final presentation, and an exam. 
The figures presented are examples only.
The Excel book is structured in such waay that it uses one sheet for each assessment component and a "Students" sheet where all the information is aggregated, as shown in the figure below.

![image](https://user-images.githubusercontent.com/13381706/163296814-59b8464a-7bdd-42bf-8333-af464a30a031.png)

All the values in yellow cells are values calculated by formulas. So, for example example, the grade of Ana Durães (fictitious name) in the midterm (1.1) is shown in the "Students" sheet by a function that searches for it in the "Mini-Test" sheet. The Sum column represents the sum of the values from column C to column G.
In row 2, cells C2:G2 contain the weight of the respective evaluation component. So, for example example, in column G, the cell values can only be between 0 and 6.

![image](https://user-images.githubusercontent.com/13381706/163296837-8eafce57-146c-4872-87bb-aec17fac0a6b.png)

The figure above represents an example of the Mini-test, which consists of 10 questions. Column L represents the sum of the previous columns. Column M represents the values to the left but with higher decimal decimal precision, this being the column that is used to obtain the values for the grid in the "Students" sheet.
The "ProposalProject" sheet is something like illustrated in the figure below. As you can see this is a group evaluation component. In that sense column A is used to identify each group. In column B the values should appear in red if they are longer than 12 minutes.

![image](https://user-images.githubusercontent.com/13381706/163297137-26a80fe1-83d2-4828-99bb-af6ae16865df.png)

Columns C, E, G and I are for noting comments and columns D, F, H and J place values in these evaluation parameters. In column K the sum of the parameter values is calculated. The "Final Report" sheet looks like this. It is also a group activity. Being column F calculates the sum of the values in each of the parameters.

![image](https://user-images.githubusercontent.com/13381706/163297309-c2527b26-6fa0-4e31-971a-e44eb4f9d866.png)

The "FinalPresentation" sheet is very similar to the "ProjectProposal" sheet.

The "Exam" sheet includes a column with the students' names, a column for the time spent on the exam, and a column for each question on the exam. This is followed by a column that adds up the of the values for each question.

![image](https://user-images.githubusercontent.com/13381706/163297425-943377ab-74ac-454c-99c6-503a65595c16.png)

Finally, there is a "Summary" sheet that shows a histogram of the marks obtained and some other statistical information, as shown in the figure below.
The graph is prepared to receive the values coming from cells A2:B22, and the values in cells B2:B22 come from the "Students" sheet, making an integration between regular and appeal season grades. 

![image](https://user-images.githubusercontent.com/13381706/163297563-b68803f2-21dd-4a53-9778-810c73e7b729.png)
