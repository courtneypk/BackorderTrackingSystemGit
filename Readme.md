# Backorder Tracking System

## Overview and Purpose

DFCI pharmacy places orders weekly for drugs through RxScan from three different sources, ABC (the primary wholesale source), Drop Ship, and Buy Direct (for more expensive drugs). DFCI relies on suppliers of these medications for patient care. However there are situations occur in the industry that disrupt the supply chain and these produce drug shortages. ABC has a report of the status of supplies for each drug that it supplies that we can download as needed. This has been used as a reference for Pharmacy to evaluate any anticipated shortages/interruptions and to help understand the impact of these shortages/interruptions.

Hilary then uses this information to manually create a spreadsheet that reports on just the medications that DFCI uses from the ABC report. The effect of a shortage is measured by taking the current inventory at DFCI and dividing by the weekly rate of use over the previous 3-4 month period of time (using the RxScan Transaction Log as a data source) to establish how many days' supply is available on hand. This is then color-coded based on the number of days' supply with 0-30 Red, 31-60 Yellow and <60 Green. This report is used in meetings with leadership regarding current status of supply chain issues and their resolution. Sample of existing Excel worksheet usage:

![Sample Report](./images/Supply%20Issues%20Research%20Inventory%20Review.png)

## Comments

- The MBO file downloaded from ABC contains all drugs that ABC supplies, not just those supplied to DFCI. Therefore the drugs in this file must be matched to those in the DFCI inventory (taken from the RxScan system).
- The MBO file as downloaded has two problems:
  - The file is labeled on the ABC site as a CSV file, but when the download dialog box is created the file type is given as ".xls". The extension must be manually changed to ".csv". Otherwise Windows thinks the file is a corrupted Excel file (i.e. the formatting of the file does not comport with the file extension).
  - As downloaded, the file cannot be imported into Excel or Access as a CSV file because there are two lines of text at the head and one line at the tail that are descriptors only. They must be removed before the file can be treated as a CSV file.
