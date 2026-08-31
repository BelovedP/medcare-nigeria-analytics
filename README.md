# MedCare Nigeria: Health Data Analytics Case Study

## Project Overview
MedCare Nigeria is a health data consultancy partnering with hospital network across major regional hubs ( Lagos, Abuja, Kano, Port Harcourt, and Enugu) to optimize patient data management and financial billing systems.
This project solves critical operational and analytical requests from multiple cross-functional departments ( Finance, Clinical Audit, Insurance Partnerships, and Infection control) using advanced spreadsheet formulars, data cleaning practices, and error-tolerant engineering frameworks.


## Data Architecture & Reference Materials
The analytics pipelines handles **200 patient rows** spanning 10 core fields in 'MedCare_Nigeria_Analytics.xlsx'. Due to upstream extraction inconsistencies, internal hospital names contain trailing and leading spaces which reuires robust trimming functions to avoid breaking reference pointers.

The lookup matrices span across two primary workbooks:
1. ## `MedCare_Nigeria_Analytics.xlsx` Sheet: `Readme Documentation`, `Doctor_Directory`, `Bill_Tier_Reference`, `Age_Bracket_Reference`, `Protocol_reference`, `State_Coordinator`, and `Insurance_Reimbursement_Matrix`.
2. ## `Hospital_Master.xlsx` Sheet: Managed by IT , holding hospital capacity and network infrastructure metrics across `Hospital_Details` and `Accereditation_Benchmark`.

## Analytics Breakdown & Advanced Formula Implementations
*Note: All formulars are tailored with commas (`,`) separatorsto match regional system localization standard*

### Doctor Responsibility Analysis (Clinical Audit Desk)
* **Objective:** Extract the attending doctor's full name from a separate directory using alpha-numeric code to remove manual validtaion  overhead.
* **Formular:**
* ```excel
  =IFNA(VLOOKUP(TRIM(E2),Doctor_Directory!$A$2:$D$31,2,FALSE),"No name")
  ```

  ### Financial Standardization & Master Data Integration ( Finance Desk )
