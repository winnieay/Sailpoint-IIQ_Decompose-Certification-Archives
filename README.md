# Sailpoint-IIQ_Decompose-Certification-Archives
The Sailpoint IIQ  provides functionality to archive the certification campaign. But cannot decompose the archived certification. The PowerShell script is for generating the csv report based on the archived data export from the SailPoint DB.


Scenario:
MSSQL
SailPoint IdentityIQ


Step 1 - Export the archived certification form DB

1. Go the Table identityiq.spt._certification_achive.
2. Right click and select the option "Select Top 1000 Rows". And then Execute the sql script.
![alt text](https://github.com/[username]/[reponame]/blob/[branch]/image.jpg?raw=true)
3. In the result console, find and select the column named "archive"
4. Right click and sava the archived data as csv file. 


![alt text](https://github.com/[username]/[reponame]/blob/[branch]/image.jpg?raw=true)
