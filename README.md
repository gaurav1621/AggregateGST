# AggregateGST
This code helps to generate report with aggregate CGST calculated for each customer on all his orders

Prerequisite Python libraries:
1) sqllite3
2) pandas
3) xlsxwriter
4) PySimpleGUI

Steps to follow to generate aggregate GST sheet:
1) Run the Aggregate_GST_UI.py:

   ![image](https://user-images.githubusercontent.com/17623783/130358692-11b3b3b4-c439-473d-b5e9-bed0420cb9d8.png)

2) Create a excel file similar to attached "Sales Register.xls":

   ![image](https://user-images.githubusercontent.com/17623783/130358785-89816856-f344-456c-9d58-0a9a96d5bbe7.png)

3) Select the file in Python Aggregate GST GUI:
   
   ![image](https://user-images.githubusercontent.com/17623783/130358831-65a0d57e-ebfd-4b20-8449-443d1b2c533d.png)

4) Click "Generate":
   
   ![image](https://user-images.githubusercontent.com/17623783/130359036-ae6adc9e-8919-421d-8c01-cf9bb840cae5.png)

5) Check the generated file at destination, you should find new file generated as "output.xlsx":
   
   ![image](https://user-images.githubusercontent.com/17623783/130359082-e34d20bc-834e-4690-af4c-cc751c250d3a.png)
   
