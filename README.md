# Sailpoint-IIQ_Decompose-Certification-Archives
The Sailpoint IIQ  provides functionality to archive the certification campaign. But cannot decompose the archived certification. The PowerShell script is for generating the csv report based on the archived data export from the SailPoint DB.

The Powershell decomposeArchivedCert.ps1 is for the sample archived certification (certificationData.csv). In that csv file, only one certification campaign is included.
And for the Powershell Complex_decomposeArchivedCert.ps1, it is used for the condiction that multiple certification campaign archived in signle csv file.


To be clear, the upload powershell code is the basis for processing archived certification activity data. If more customized reports are required, modifications will be required.

### Scenario:
* MSSQL
* SailPoint IdentityIQ

#### Step One - Export the archived certification form DB
1. Go the Table identityiq.spt._certification_achive.
2. Right click and select the option "Select Top 1000 Rows". And then Execute the sql script.
   <br /><br /><img src="images/1.png" width="300" ><br /><br />
4. In the result console, find and select the column named "archive"
5. Right click and sava the archived data as csv file. 
  <br /><br /><img src="images/2.png" width="350" ><br /><br />

#### Step Two - Change the file directory in the PowerShell script
1. Input CSV file directory
2. Output CSV file directory
  <br /><br /><img src="images/3.png" width="350" ><br /><br />
