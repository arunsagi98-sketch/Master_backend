# Excel Output Formatting & Styling Documentation

## Overview

The MPN Campaign Report Generator now includes complete cell-level formatting and styling using openpyxl. All output Excel files follow a precise color, font, and formula specification to ensure professional presentation and functional calculations.

## Color Palette

Three main colors are used throughout the report:

| Color | Hex Code | Usage | Applied To |
|-------|----------|-------|-----------|
| **Blue** | #0000FF | Section Title Background | Column B of section header rows |
| **Yellow** | #FFFF00 | Sub-section Label Background | Column B of subsection rows |
| **White** | #FFFFFF | Data & Header Background | Column headers and data cells with fills |
| **None** | N/A | Transparent/Formula Cells | Formula cells and label cells |

## Row Formatting Rules

### 1. Section Title Rows
**Examples:**
- "Overview" (Row 4)
- "Performance breakdown - by Audience - Vietnamese - Banner - Burst - 1"

**Formatting:**
- **Column B only**: Blue fill (#0000FF), Bold white font (#FFFFFF)
- **Other columns**: No special formatting

### 2. Sub-section Header Rows
**Examples:**
- "By Device", "By Creative", "By Age", "By Gender"

**Formatting:**
- **Column B (label cell)**: Yellow fill (#FFFF00), Bold black font
- **Columns C onwards (header cells)**: White fill (#FFFFFF), Bold black font

### 3. Overview Column Header Row
**Row:** Contains "Platform", "Format", "Audience", "Reporting Date", etc.

**Formatting:**
- **All cells (A-R)**: White fill (#FFFFFF), Bold black font

### 4. Data Rows
**Rows:** All rows with actual values (Mobile/Tablet/Desktop, age bands, genders, creatives)

**Formatting By Column:**

#### Overview Section Data Rows:
| Column | Content | Fill |
|--------|---------|------|
| 1-4 | Platform, Format, Audience, Reporting Date | White |
| 5-6 | Start Date, End Date | None |
| 7 | Booked Impressions | White |
| 8 | Actual Impressions | None |
| 9-10 | Campaign/Impression Pacing (formulas) | None |
| 11-12 | Reach, Frequency | White |
| 13 | Link Click | None |
| 14 | CTR (formula) | None |
| 15-16 | Complete Views, VCR | White |
| 17 | Amount Spent | None |
| 18 | Investment | None |

#### Breakdown Section Data Rows (By Device, Creative, Age, Gender):
| Column | Content | Fill |
|--------|---------|------|
| 2 | Label (Device/Creative/Age/Gender) | None |
| 3 | Actual Impressions | None |
| 4 | Reach | White |
| 5 | Frequency | White |
| 6 | Complete Views | White |
| 7 | Link Click | None |
| 8 | CTR (formula) | None |
| 9 | Amount Spent | White |

## Formula Implementation

All formulas are live Excel calculations that update automatically when source data changes.

### 1. CTR (Click-Through Rate)

**Overview Section:**
```
=N{row}/I{row}
```
- **N column**: Link Click
- **I column**: Actual Impressions
- **Example**: `=N6/I6`
- **Format**: 0.00% (percentage with 2 decimals)

**Breakdown Sections:**
```
=G{row}/C{row}
```
- **G column**: Link Click
- **C column**: Actual Impressions
- **Example**: `=G25/C25`
- **Format**: 0.00%

### 2. Campaign Pacing (Overview)

**Formula:**
```
=(E{row}-F{row})/(G{row}-F{row})
```
- **E column**: Reporting Date
- **F column**: Start Date
- **G column**: End Date
- **Example**: `=(E6-F6)/(G6-F6)`
- **Format**: 0.00%
- **Note**: Currently using placeholder value 0; requires date integration

### 3. Impression Pacing (Overview)

**Formula:**
```
=IFERROR(I{row}/H{row},0)
```
- **I column**: Actual Impressions
- **H column**: Booked Impressions
- **Example**: `=IFERROR(I6/H6,0)`
- **Format**: 0.00%
- **Feature**: Returns 0 if denominator is zero (error handling)

### 4. Amount Spent (Overview)

**Formula:**
```
=S{row}*K{row}
```
- **S column**: Investment
- **K column**: Impression Pacing
- **Example**: `=S6*K6`
- **Note**: To be implemented; currently uses hardcoded value 0

## Code Implementation Details

### Color Palette Constants
```python
BLUE_FILL = PatternFill(fill_type="solid", fgColor="0000FF")
YELLOW_FILL = PatternFill(fill_type="solid", fgColor="FFFF00")
WHITE_FILL = PatternFill(fill_type="solid", fgColor="FFFFFF")
NO_FILL = PatternFill(fill_type=None)
```

### Font Constants
```python
WHITE_BOLD_FONT = Font(bold=True, color="FFFFFF")
BLACK_BOLD_FONT = Font(bold=True, color="000000")
BLACK_REGULAR_FONT = Font(bold=False, color="000000")
```

### Number Format Constant
```python
PERCENTAGE_FORMAT = "0.00%"
```

### Styling Methods

#### `_apply_section_title(row)`
Applies blue background and white bold font to Column B.

#### `_apply_subsection_header(row, num_cols)`
Applies yellow fill and bold to Column B, white fill and bold to Columns C+.

#### `_apply_overview_header(row, num_cols)`
Applies white fill and bold font to all columns.

#### `_apply_formula_cell(row, col, formula, num_format)`
- Sets cell formula
- Applies number format if provided
- Sets fill to NO_FILL
- Sets font to BLACK_REGULAR_FONT

#### `_get_col_letter(col_num)`
Converts column number (1-26+) to Excel column letter (A-Z, AA-AZ).
```python
_get_col_letter(1)  # Returns 'A'
_get_col_letter(8)  # Returns 'H'
_get_col_letter(13) # Returns 'M'
```

## Cell Reference System

### Overview Section Columns (1-18)
```
 1: Platform              I: Actual Impressions
 2: Format               J: Campaign Pacing
 3: Audience             K: Impression Pacing
 4: Reporting Date       L: Reach
 5: Start Date           M: Frequency
 6: End Date             N: Link Click
 7: Booked Impressions   O: CTR
 8: Actual Impressions   P: Complete Views
 H: Booked Impressions   Q: VCR
 I: Actual Impressions   R: Amount Spent
                         S: Investment
```

### Breakdown Section Columns (1-9)
```
 1: Empty (padding)
 2: Label (Device/Creative/Age/Gender)
 3: Actual Impressions
 4: Reach
 5: Frequency
 6: Complete Views
 7: Link Click
 8: CTR
 9: Amount Spent
```

## Styling Application Flow

1. **Header rows** → `_apply_section_title()` for blue styling
2. **Column headers** → `_apply_overview_header()` or `_apply_subsection_header()`
3. **Data values** → `_apply_formula_cell()` for formulas, `NO_FILL` for labels/impressions
4. **Calculation cells** → `_apply_formula_cell()` with formula and percentage format

## Excel Display Examples

### Overview Section
```
Row 4:  [Overview with blue Column B]
Row 5:  [Platform] [Format] [Audience] ... [Investment]  [all white headers]
Row 6:  MPN        Banner   Vietnamese... 0              [data with specific fills]
```

### Breakdown Section
```
Row 25:               [By Device with yellow Column B, white C-onwards]
Row 26:               [Device Type] [Act Impr] [Reach] [Freq] ...
Row 27: (empty) Mobile 50000       16667     3     - 350 formula -
```

## Future Enhancements

1. **Campaign Pacing Formula**: Requires proper date handling
   - Current: Placeholder value (0)
   - Future: `=(E{row}-F{row})/(G{row}-F{row})`

2. **Amount Spent Formula**: Requires investment column implementation
   - Current: Hardcoded 0
   - Future: `=S{row}*K{row}`

3. **Dynamic Number Formatting**: Additional formats for specific columns
   - Impressions: Integer format
   - CTR: Percentage format (implemented)
   - Investment: Currency format

4. **Conditional Formatting**: Based on performance thresholds
   - High performing: Green
   - Low performing: Red
   - Target range: Yellow

## Testing & Validation

### Verify Formatting in Generated Excel:
1. **Section titles**: Blue background in Column B, white bold text
2. **Sub-section headers**: Yellow in Column B, white in other columns
3. **Formulas**: Click cells to verify formula bar shows calculation
4. **Number formats**: CTR values show as percentages (e.g., 0.50%)
5. **Color consistency**: Only 3 colors used throughout

### Common Issues & Resolution:
- **Colors not appearing**: Verify `fgColor` parameter uses correct format (#0000FF)
- **Formulas not calculating**: Ensure cell format is set before formula entry
- **Wrong column references**: Use `_get_col_letter()` to convert numbers to letters
- **Percentage format issues**: Verify cell.number_format = "0.00%"

## References

- openpyxl Documentation: https://openpyxl.readthedocs.io/
- PatternFill: Handles cell background colors
- Font: Handles text styling (bold, color, size)
- Number Formats: Excel-compatible format strings

---

**Last Updated**: May 7, 2026  
**Implementation Status**: Complete for styling and CTR/Impression Pacing formulas
**Pending**: Campaign Pacing formula, Amount Spent formula, additional number formats
