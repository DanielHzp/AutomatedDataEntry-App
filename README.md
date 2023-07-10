# AutomatedDataEntryInterface

This is a desktop excel business solution that automates data entry of customers info. through a friendly user interface. This user form updates and sets new records in a local data set that should populate a SQl server data base using an ADO Visual Basic Connection.


i.e.
dbADOConnection = new ADODB.Connection 

dbADOConnection.Open ConnectionString, UserID, Password, OpenOptions






The user must load the form clicking on an excel action button,then the following attributes must be filled out and an option to automatically write data on the dataset becomes available:









![image](https://github.com/DanielHzp/AutomatedDataEntryInterface/assets/124480168/18982b83-b27f-4c4a-b5b0-0d8bb3966294)




When the user clicks on the 'save' option all new records are added to the following dataset which is pending to be synch. to the SQL business server:


![image](https://github.com/DanielHzp/AutomatedDataEntryInterface/assets/124480168/5ecde386-9c4d-4718-b09a-59c7abe2831e)


Additionally, the solution performs calculations in a Macro in order to update the following data subset:








![image](https://github.com/DanielHzp/AutomatedDataEntryInterface/assets/124480168/387a5b40-b10a-4ab5-a7ce-cbb1acc0bbb1)











