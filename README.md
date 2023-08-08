# Automated Data Entry 

This is a desktop business solution that automates customers data entry through a windows form. This user form updates and sets new records in a local data set that should populate a SQl server data base using an ADO Visual Basic Connection.

<br/>

Sample:

https://github.com/DanielHzp/AutomatedDataEntryInterface/blob/008e10f2cf1220aa50287400b1cd368b2515bd22/On%20Click%20Actions/InsertDataDB.bas#L4-L43


<br/>


The user must load the form clicking on an excel action button,then the following attributes must be filled out and an option to automatically write data on the dataset becomes available:

<br/>

<img src="https://github.com/DanielHzp/AutomatedDataEntryInterface/assets/124480168/18982b83-b27f-4c4a-b5b0-0d8bb3966294" width="510" height="510">

<br/>


<br/>

<br/>


The form contains a 'Add New Record' button and an option to clean all the user input fields. Additionally, a 'Delete Record' button lets the user select which record should be deleted. When the user clicks on the 'save' option, all new input fields are added to a new row in the following dataset which is synchronized with a SQL business server connection:

![image](https://github.com/DanielHzp/AutomatedDataEntryInterface/assets/124480168/5ecde386-9c4d-4718-b09a-59c7abe2831e)

<br/>

<br/>
Additionally, the solution performs calculations in a Macro in order to update the following data set:



![image](https://github.com/DanielHzp/AutomatedDataEntryInterface/assets/124480168/387a5b40-b10a-4ab5-a7ce-cbb1acc0bbb1)

<br/>

<br/>

## Usage

The form can be executed by importing the .frx and .bas files in a VB developer editor windows forms











