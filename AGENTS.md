# AGENTS.md

## Project Overview

This project analyzes Japanese credit card data.

It merges multiple CSV files and generates Excel reports with summaries and charts.

## Directory Structure

src/
main.py – main script

input/
credit card CSV files

output/
generated Excel reports

## Key Columns

date – transaction date  
category – merchant name  
amount – transaction amount  
group – merchant classification

## Merchant Groups

英希サブスク  
由利子サブスク  
光熱費  
その他

## Development Rules

When modifying this project:

- do not rewrite the entire script
- keep the function structure
- modify only necessary parts

## Workflow

1 place CSV files in input
2 run

python src/main.py

3 Excel report is generated in output
