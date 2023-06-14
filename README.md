# AutoExtract-Outlook-Email-Attachments-and-Create-Reports-using-PowerBI

## Problem Statement
In company people used to receive an mail every hour, it is company related to SEZ Services which receives Invoice data every hour from different clients via email and
there were team dedicated of like 10 resources their job was to download their data via email ,clean and process the data, understand what data problems are then post that they have to combine all the data what they receive every hour and then create a report in Excel and they used to manually send this cleaned data to their relevant team or stakeholders .. this process used to take lot of time and consume lot of resources in the organization and company used to spend huge amount of time and effort.. so the project is end-to-end project in power bi, How to get email attachments without opening the email those email attachments can be excel, csv, pdf etc.

## Traditional Report in organization using Excel 

![img](https://github.com/chetana-vasave3/AutoExtract-Outlook-Email-Attachments-and-Create-Reports-using-PowerBI-/blob/main/Screenshots/Traditiona.png)

Now we will see hoe to extract email attachments from outlook.

- Receiving Email from clients

![img]( https://github.com/chetana-vasave3/AutoExtract-Outlook-Email-Attachments-and-Create-Reports-using-PowerBI-/blob/main/Screenshots/Recieving%20email%20from%20client.png)


Here we can see that we received Invoice data from Allstate client with subject Allstate Monthly Invoice Data.. Now we will create Rule in Outlook to create a folder so that all the Attachments coming from Allstate client will be stored in that folder and after that we will Access that folder in Power bi..

-  Creating Rule in Outlook

  
![img]( https://github.com/chetana-vasave3/AutoExtract-Outlook-Email-Attachments-and-Create-Reports-using-PowerBI-/blob/main/Screenshots/creating%20rules.png?raw=true)
![img]( https://github.com/chetana-vasave3/AutoExtract-Outlook-Email-Attachments-and-Create-Reports-using-PowerBI-/blob/main/Screenshots/folder.png?raw=true)

Here we can see we have created folder named as Allstate Monthly Invoice Data..Now whenever we received the Attachments with Subject include “Allstate Monthly Invoice Data “ will automatically moved to this folder. This will be totally automatic there is no need of manual power to download and create folder.

 Now will connect to the Power Bi Desktop to get an folder with all attachments
 
-	Connecting to Power Bi

![img](https://github.com/chetana-vasave3/AutoExtract-Outlook-Email-Attachments-and-Create-Reports-using-PowerBI-/blob/main/Screenshots/getting%20data.png?raw=true)

After that we have to provide login credentials to the Power Bi . make sure that you have logged in with same Power Bi email and Outlook email.

Now we have done with connection after that we will load the data from outlook to Power Bi.
-	Loading Data

![img](https://github.com/chetana-vasave3/AutoExtract-Outlook-Email-Attachments-and-Create-Reports-using-PowerBI-/blob/main/Screenshots/powerbi%20data%20preview.png?raw=true)

We will see data in this format..after that we will click on mail coz we want mail data.


![img](https://github.com/chetana-vasave3/AutoExtract-Outlook-Email-Attachments-and-Create-Reports-using-PowerBI-/blob/main/Screenshots/loading%20data.png?raw=true)


Now Further we will  do the Data cleaning, Data validation and Data Transformation. after that we will do the visualization in Power Bi and Create a report.


-	To See Report Click Here - https://app.powerbi.com/groups/me/reports/dfbfbbe4-dc64-434f-8cc4-02192e17d22d/ReportSection?experience=power-bi


You can send this report to the Manager ..Here the Dashboard will keep on upgrading as new attachment will come from that particular client.



