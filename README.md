# MedCare Nigeria: Health Data Analytics Case Study

## Project Overview
MedCare Nigeria is a health data consultancy partnering with hospital network across major regional hubs ( Lagos, Abuja, Kano, Port Harcourt, and Enugu) to optimize patient data management and financial billing systems.
This project solves critical operational and analytical requests from multiple cross-functional departments ( Finance, Clinical Audit, Insurance Partnerships, and Infection control ) using advanced spreadsheet formulars, data cleaning practices, and error-tolerant engineering frameworks.


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

### Financial Standardization & Master Data Integration (Finance Desk)
* **Objective:** Pull the official Network ID from an external master database file (`Hospital_Master.xlsx`) to normalize financial reporting layouts, resolving whitespace fragmentation.
* **Formular:**
* ```excel
  =VLOOKUP(TRIM(D2),[Hospital_Master.xlsx]Hospital_Details'!$A$2:$D$16,2,FALSE)
  ```
  
### Dynamic Revenue Tiering (Billing Desk)
* **Objective:** Automate pricing categoryassignments based on  sliding scales of overall bills without requiring hardlimits.
* **Formular:**
* ```excel
  =VLOOKUP(I2,Bill_Tier_Reference!$A$2:$B$5,2,TRUE)
  ```

### Demographics Cohort Analysis(Reporting Desk)
* **Objective:** Standardize raw age values into structural life-stage categories for quarterly dashboard consumption.
* **Formula:**
* ```excel
  =VLOOKUP(F2,Age_Bracket_Reference!$A$2:$B$6,2,TRUE)
  ```
  
### Account Management & Escalations (Insurance Desk)
* **Objective:** Identify the regional lead overseeing specific geographic domains to accelerate claim inquires.
* **Formular:**
* ```excel
  =VLOOKUP(C2,State_Coordinator!$A$2:$B$6,2,FALSE)
  ```

### Vendor Capacity Assessment(Partnerships Desk)
* **Objective:** Determine contract scaling strategies by extracting provider accreditation tiers stored externally.
* **Formular:**
* ```excel
  =VLOOKUP(TRIM(D2),'C:\Users\User\Downloads\WEEK 2\Assignment 2B CONTINUATION\[Hospital_Master.xlsx]Hospital_Details'!A$2:D$16,4,FALSE)
  ```

### Wildcard Medical Text Parsing (Infection Control Desk)
* **Objective:** Standardize unstructured diagnostic notes into actionable treatment guidelines by identifying text mention of specific disease keywords(e.g., TB, Malaria, Typhoid),falling back gracefully on standard protocol.
* **Formular:**
* ```excel
  =IFS(ISNUMBER(SEARCH("TB", G2)),Protocol_Reference!$B$2, ISNUMBER(SEARCH("Malaria",G2)),Protocol_Reference!$B$3, ISNUMBER(SEARCH("Typhoid",G2)),Protocol_Reference!$B$4, ISNUMBER(SEARCH("Hypertension", G2,)),Protocol_Reference!$B$5, ISNUMBER(SEARCH("Diabetes", G2)),Protocol_Reference!$B$6, ISNUMBER(SEARCH("Cholera", G2)),Protocol_Reference!$B$7, ISNUMBER(SEARCH("Pneumonia", G2)),Protocol_Reference!$B$8, TRUE, "Out of scope/ Routine Care")
  ```

### Fault-Tolerant Financial Auditing (Pharmacy Desk)
* **Objective:** Retrieve doctor consultation fees while handling non-existent Id strings gracefully with explicit administrative text feedback instead of standard `#N/A`breaks.
* **Formular:**
* ```excel
  = IFERROR(VLOOKUP(E2,Doctor_Directory!$A$2:$D$31,4,FALSE),"Check Doctor Code")
  ```

### Matrix Rate Extraction ()
