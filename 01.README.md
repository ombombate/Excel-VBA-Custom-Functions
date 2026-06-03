# Excel VBA Custom Functions

This repository contains custom Excel VBA functions:

## Functions Included

### SUMCOLOR()
Assigns values based on cell colors.

| Color | Value |
|---------|---------|
| Light Blue | 15 |
| Light Green | 10 |
| Yellow | 5 |
| Red | 0 |
| Black | -5 |

Example:

=SUMCOLOR(A1:A20)

---

### COUNTIFCOLOR()

Counts cells of a selected color.

Example:

=COUNTIFCOLOR(A1:A20,D1)

---

## Installation Guide

1. Open Excel
2. Press Alt + F11
3. Open PERSONAL.XLSB
4. Insert Module
5. Paste VBA Code
6. Save
