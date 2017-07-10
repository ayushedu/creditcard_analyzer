# -*- coding: utf-8 -*-
"""
Created on Mon Jul  3 18:59:05 2017

@author: Ayush Vatsyayan
"""

import os
import sys
import pyparsing as pp
import re
import pandas as pd
import locale
import ConfigParser


# Change directory
#os.chdir("/Users/snowleopard/workspace/hack/creditcard_analyzer")

#==============================================================================
# Convert PDF to text using Apache PDFBox
#==============================================================================
def pdf_to_text():
    global config

    # Prepare command
    cmd = config.get("PDF","PDFBOX_COMMAND") + " " + config.get("PDF","PASSWORD")
    cmd += " " + config.get("PDF","PDF_FILE_PATH") + " tmp.txt"
    
    # Execute command
    resp = os.system(cmd)
    
    if resp != 0:
        print "Error converting PDF to txt"
        sys.exit(resp)
    
    # Read data
    textfile = open("tmp.txt")
    lines = textfile.readlines()
    textfile.close() #close file
    
    # Remove tmp.txt
    os.remove("tmp.txt")
    
    return lines

#==============================================================================
# Initialize
#==============================================================================
def init():
    #Read configuration
    global config
    config = ConfigParser.ConfigParser()
    config.read("config.ini")
    
    # Pattern to parse the statement
    transaction_id = pp.Word(pp.nums) + pp.Suppress(pp.White())
    balance = pp.Combine(pp.Word(pp.nums + ",") + "." + pp.Word(pp.nums, exact=2) + pp.Optional('CR'))
    date_pattern = pp.Word(pp.nums, exact=6)
    merchant = pp.restOfLine()

    global transactional_pattern
    global non_transactional_pattern
    transactional_pattern = transaction_id + balance + date_pattern + merchant
    non_transactional_pattern = balance + date_pattern + merchant

#==============================================================================
# Parse records that have transaction id
#==============================================================================
def parse_transactions(lines):
    global transactional_pattern
    global non_transactional_pattern
    
    # Initialize dictionary    
    stmt_dict = {'transaction_id':[], 'balance':[], 'date':[],'merchant':[] }
    
    for line in lines:
        try:
            result = pp.OneOrMore(transactional_pattern).parseString(line)
            
            # Append result to dictionary
            stmt_dict['transaction_id'].append(result[0])
            stmt_dict['balance'].append(result[1])
            stmt_dict['date'].append(result[2])
            stmt_dict['merchant'].append(result[3])
        except pp.ParseException:
            # In case of exception parse Non-transactional
            try:
                result = non_transactional_pattern.parseString(line)
                stmt_dict['transaction_id'].append('')
                stmt_dict['balance'].append(result[0])
                stmt_dict['date'].append(result[1])
                stmt_dict['merchant'].append(result[2])
            except pp.ParseException:
                # Skip, as lines aren't valid
                pass
    
    # Create DataFrame
    stmt_df = pd.DataFrame(stmt_dict)
    
    # Convert CR to negative balance
    stmt_df.balance = [ '-' + v.replace('CR','').strip() if v.endswith('CR') else v.strip() for v in stmt_df.balance]
    
    # Remove extra spaces from merchant
    stmt_df.merchant = [re.sub("\s\s+" , " ",v) for v in stmt_df.merchant]
    
    # Change balance to float
    locale.setlocale(locale.LC_ALL,'en_US.UTF-8')
    #locale.setlocale(locale.LC_NUMERIC, '')  
    stmt_df.balance = stmt_df.balance.apply(locale.atof)
    
    # Arrange column order
    stmt_df = stmt_df[['date','transaction_id','merchant','balance']]
    return stmt_df

#==============================================================================
# Write the results to Excel
#==============================================================================
def write_excel(stmt_df):
    # Group the data
    by_merchant = stmt_df.groupby(["merchant","date"])
    summary_df = by_merchant.sum()

    # Write result
    global config
    
    output_file_name = config.get("OUTPUT","FILE_NAME_PREFIX") + "_" +str(stmt_df.date[0][2:]) + ".xlsx"    
    writer = pd.ExcelWriter(output_file_name,engine='xlsxwriter')
    stmt_df.to_excel(writer, sheet_name='detailed')
    summary_df.to_excel(writer, sheet_name='summary')
    writer.save()
    
    print "Results written successfully as", output_file_name

#==============================================================================
# Main function
#==============================================================================
if __name__ == '__main__':
    # initialize
    print ("Initializing1..")
    init()

    # read data    
    print("Reading PDF...")
    lines = pdf_to_text()

    # Parse data    
    print("Parsing data...")
    stmt_df = parse_transactions(lines)
    
    # write results
    write_excel(stmt_df)
