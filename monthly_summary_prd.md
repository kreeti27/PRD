**Product Requirements Document (PRD)**

**Title:** Monthly Summary Report Generator in Excel VBA

**Purpose:**
The purpose of this macro is to generate a monthly financial summary report based on a primary dataset ("Sheet1") and several auxiliary mapping and exclusion sheets. The macro reads transaction-level financial data and outputs a formatted, sorted, and filtered monthly summary with various mappings applied.

**Sheets Involved:**

1. **Sheet1 (Source Data)** - Contains raw transaction data with columns:
    | Column | Header           | Description                             |
    |--------|------------------|-----------------------------------------|
    | A      | Fiscal Year      | Fiscal year of the transaction          |
    | B      | Date             | Date of the transaction (used for month)|
    | C      | Parent Code      | Original parent code                    |
    | D      | Description      | Description of the line item            |
    | E      | Account Code     | Account number                          |
    | G      | Amount           | Monetary value                          |
    | I      | Fund Code        | Financial fund identifier               |

2. **MappingAccount** - Maps (Fund, Account) pair to SCO Account (Adjusted Parent Code)
    | Fund | Account | SCO Account |
    |------|---------|-------------|
    | 0044 | 12345   | 600         |

3. **MappingFund** - Maps original Fund to mapped Fund (no header)
    | Fund | Mapped Fund |
    |------|-------------|
    | 0044 | 1200        |

4. **ExcludeFund** - List of funds to be excluded (no header)
    | Fund |
    |------|
    | 3001 |

5. **ExcludeAccounts** - List of accounts to be excluded (no header)
    | Account |
    |---------|
    | 5555    |

6. **Order** - Custom order for funds (no header). Funds listed here will appear in order in the final report, others will appear after.
    | Fund |
    |------|
    | 1200 |
    | 1300 |

**Report Format (Output):**
- The report is generated in "Sheet2" with the following columns (dynamic based on available months):
    | Fund | Description | Parent Code | Adjusted Parent | Jan | Feb | ... | Dec | Total | FY |
- The month columns will appear only if those months exist in the source data.
- Amount columns (months and Total) will be formatted as numbers with two decimal places and comma separators.
- Rows where total is negative will be highlighted in light red (`RGB(255,199,206)`).

**Transformation Logic:**

1. **Fund and Account Mapping:**
   - First, check `MappingAccount` for a match on (Fund, Account) and use `SCO Account` as `Adjusted Parent` if found.
   - Then check `MappingFund` to override the fund value if applicable.

   **Example:**
   Source:
   | Fund | Account | Parent |
   |------|---------|--------|
   | 0044 | 12345   | 81141  |

   MappingAccount:
   | 0044 | 12345 | 600 |

   MappingFund:
   | 0044 | 1200  |

   Result:
   | Fund | Adjusted Parent |
   |------|------------------|
   | 1200 | 600              |

2. **Exclusion:**
   - Skip any row where the Fund (after mapping) is in `ExcludeFund`.
   - Skip any row where the Account is in `ExcludeAccounts`.

3. **Monthly Aggregation:**
   - Group data by: Mapped Fund, Description, Parent Code, Fiscal Year.
   - Sum values per month based on the Date field (column B).
   - Add a Total column to hold the row sum.

4. **Sorting:**
   - Sort first by the fund order as defined in the `Order` sheet.
   - Then by Fund alphabetically, and then by Parent Code numerically.
   - Funds not in `Order` will appear at the bottom of the report.

**Additional Notes:**
- If Sheet2 does not exist, it will be created.
- If Sheet2 exists, it will be cleared before writing the new report.
- Number formatting is applied only to columns starting from the first month to the Total column.
- The header row is bold.
- All columns are auto-fit at the end of the macro.

**Macro Trigger:**
- The macro is to be manually run via the `GenerateMonthlySummaryOrdered` subroutine.

**Future Enhancements (Not yet implemented):**
- Generating a new workbook with the filename `CL_REV_MMDDYYYY.xlsx` (currently this writes to Sheet2 only).
- Making exclusion/mapping dynamic via user prompt.

This PRD captures the complete logic of your macro and should enable accurate reproduction or expansion of the functionality by other teams or automated agents.

