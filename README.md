Loan Eligibility Checker
========================

Introduction
------------

The Loan Eligibility Checker is a VBA macro that checks whether an individual or entity is eligible for a loan based on various criteria. The macro reads data from an Excel worksheet and performs eligibility checks for each row of data.

Requirements
------------

To use the Loan Eligibility Checker, you need:

-   Microsoft Excel (version 2010 or later)
-   Basic knowledge of Excel functions and macros

Installation
------------

To install the Loan Eligibility Checker, follow these steps:

1.  Download the `Loan Eligibility Checker.xlsm` file.
2.  Open Microsoft Excel and enable macros.
3.  Open the `Loan Eligibility Checker.xlsm` file.
4.  Input your loan data in columns A to O of the `Loan Data` worksheet.
5.  Click on the `Check Eligibility` button.

Usage
-----

The Loan Eligibility Checker performs two eligibility checks: a property eligibility check and a credit bureau eligibility check. The results are displayed in columns P, Q, and R.

### Property Eligibility Check

The macro checks the following criteria for the property eligibility check:

-   Loan amount should be between 500,000 and 3,500,000
-   Total property value should be less than 4,500,000 or the property should be sanctioned under SECO
-   Property should not be non-agricultural land
-   Property layout plan should be formal
-   Property should be sanctioned under Collector (Zilla Parishad) (ZP), Gram Panchayat (GP), or Municipality/Town Planning (TP)

If the property is eligible for the loan, the value in column P will be "Eligible". If not, the value will be "Not Eligible" and the reason for the denial will be provided in column R.

### Credit Bureau Eligibility Check

The macro checks the following criteria for the credit bureau eligibility check:

-   The applicant's CIBIL score should be greater than 675 or the score should be -1 (meaning no score is available)
-   The CIBIL scores of all co-applicants should be greater than 675 or the scores should be -1

If the applicant and all co-applicants are eligible for the loan, the value in column Q will be "Eligible". If not, the value will be "Not Eligible" and the reason for the denial will be provided in column R.

Troubleshooting
---------------

If the macro encounters an error, a pop-up window will appear with a brief message describing the issue. The macro also provides error handling to prevent the Excel application from crashing.

If you encounter any issues with the Loan Eligibility Checker, please contact the developer.

License
-------

The Loan Eligibility Checker is distributed under the [MIT License](https://opensource.org/licenses/MIT).

Acknowledgements
----------------

This macro was developed by Abhishek Gidde and is based on a sample loan eligibility checker provided, takes no liability of the result.
