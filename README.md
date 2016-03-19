# PowerPivotDrillButtons
This allows you to create custom buttons that will Drill Down, Drill Up and Drill to Top of your pivot hierarchy.

I am by no means an expert in VBA and am open to suggestions on ways to improve this. I found this code to be extremely useful for the product I was making, so I thought I would share in an effort to give something back to the community.

I designed code to be as simple as possible and to be able to reuse the code with minimal modifications; therefore I use a naming prefix of "Lvl" and numbered the levels 1-4 (however I coded it so that you could specify your own custom prefix as well). Given that you are able to rename your fields in the actual pivot table without affecting the back end, the hierarchy prefix will not cause any customization issues.

note: there are a few sections that need user input of your prefixes, table names, etc. and are marked with "User entry needed". Also, this was developed using the AdventureWorks SQL sample database (excel connected to SQL via power query and pulled the data into the Excel data model).
