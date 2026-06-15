# Automated-Bookkeeper
An bookkeeping automation tool that extracts transactions from messy Markdown bank statements, auto-categorizes merchants, and safely appends data to an Excel ledger without breaking formulas.

# Features and Workflow
1. Upload a markdown of credit card bank statements. This repo is optimized for RBC credit card statements. Note that this repo works best when pdfs are converted to markdowns with Microsoft's Markitdown: https://github.com/microsoft/markitdown
2. Run main.py
3. Open the bookkeeping excel sheet

# Limitations
This automation could easily be done with AI. However, due to the nature of my use cases (sensitive data), I wanted to create a program that could run locally and securely. This comes with its drawbacks:
1. keywords.json has to be edited manually to accommodate and categorize new transactions. On the bright side, this could be done with AI as the info is less sensitive.
2. Similarly, locations and numbers in transaction descriptions also have to be manually blacklisted.
3. This program is not flexible. It is likely that for each client, I would have to substantially modify the code--especially with different banks.
