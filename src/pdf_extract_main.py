import PyPDF2, re, csv, os, time, pandas, openpyxl

def fileReadErrorType (FileName, ErrorType):    # Function for recording any errors, files will not be processed as soon as error is encountered
    global ErrorFlag, TxtFileAppend, FileErrorCount
    PerformDict['Customer Number Errors Count'] = PerformDict['Customer Number Errors Count'] - FileErrorCount
    PerformDict['Hopeless Error Count'] = PerformDict['Hopeless Error Count'] - FileErrorCount
    f = open(TxtFileAppend+"_Man_Hand_Files.txt","a+")          # Filenames only for Powershell to move files
    f.write(FileName+"\n")
    f.close()
    f= open((TxtFileAppend+"_Manual_Handle_Required.txt"),"a+") # PDFs that errored and the reason why
    f.write("\n"+FileName+" "+ErrorType+"\n")
    f.close()
    MasterDict[PDFName].clear()                                 # Remove all entries from dictionary that were recorded prior to error
    PerformDict['File - Error'] = (PerformDict['File - Error'] + 1)
    ErrorFlag = True
    return

def amountErrorType (FileName, ErrorType):      # Function for recording any amount related errors, files will still process but go to Manual Handling
    global ErrorFlag, TxtFileAppend, FileErrorCount
    PerformDict['Customer Number Errors Count'] = PerformDict['Customer Number Errors Count'] - FileErrorCount
    PerformDict['Hopeless Error Count'] = PerformDict['Hopeless Error Count'] - FileErrorCount
    f = open(TxtFileAppend+"_Man_Hand_Files.txt","a+")          # Filenames only for Powershell to move files
    f.write(FileName+"\n")
    f.close()
    f= open((TxtFileAppend+"_Manual_Handle_Required.txt"),"a+") # PDFs that errored and the reason why
    f.write("\n"+FileName+" "+ErrorType+"\n")
    f.close()
    PerformDict['File - Error'] = (PerformDict['File - Error'] + 1)
    ErrorFlag = True
    return

def dollarFormatting(amount):                       # Function for removing '$' and ',' from any amounts
    amount = amount.replace("$","").replace(",", "")
    return float(amount)

def removeZeros(Extract):                           # Powered By Crystal PDFs only - removes all $0.00s from Extract
    Extract = Extract.replace("$0.00", "")
    return Extract

def extractAmount(amountString):                    # Function for finding last amount in Extract (the Total)
    amountIndex = amountString.rfind('$')           # Find last '$'
    amountString = amountString[amountIndex:]
    decIndex = amountString.find('.',1)             # Find where the decimal point is
    amountActual = dollarFormatting(amountString[1:decIndex+3])     # Sort formatting for float
    return amountActual

def settlementAmt(Extract):                         # Powered By Crystal PDFs only - function for finding settlement amount, can identify credits
    Extract = Extract.split("Settlement (")[0]      # This cuts off the statement totals so it finds the banked amt
    SettleAmount = extractAmount(Extract)           # Func extracts dollar amount from string provided
    return SettleAmount

def statementAmt(Extract):                          # Function for finding Statement total
    StateAmount = extractAmount(Extract)
    return StateAmount

def activeReportsCleanup(Extract, page):                # This cleans up specific ActiveReports files, removes header/footer
    global PDFName
    if "02-0192-0115055-02" in Extract:                 # Remove bank acc details
        Extract = Extract.replace("02-0192-0115055-02","")
    if page == 0:
        try:
            Extract = Extract.split("Property Address Amount")[1]
        except IndexError:
            fileReadErrorType(PDFName, "Customer has used non-standard naming conventions, please process manually")
    Extract  = Extract.split("Remittance Advice")[0]    # Separate Footer
    if "Total$" in Extract:
        Extract = Extract.split("Total$")[0]
    return Extract

def poweredByCrystalCleanUp(Extract, page):             # Powered by Crystal specific handling of Header/Footer removal
    # Extract = Extract.split("DebitCreditBalance")[1]    # Remove Header
    if "Settlement (" in Extract:                       # Remove Footer Last page
        Extract = Extract.split("Settlement (")[0]
        ExtractIndex = Extract.rfind('$')
        Extract = Extract[:ExtractIndex]
        return Extract
    elif "Totals at end of period" not in Extract:      # Remove Footer any page but last/settlement page
        ExtractIndex = Extract.rfind('Page ',int(page+1))
        Extract = Extract[:ExtractIndex]
        return Extract
    return None

def QTCleanup(Extract):                                 # This cleans up specific QT files, removes header/footer

    # GREYED OUT FOOT AT THE BOTTOM OF THE PDF (B&T Ltd..) APPEARS AT THE FRONT OF Qt 5.5.1 pdf_string, REMOVE 
    very_first_up_to_word = "By\t"                    # need tabspace (\t) afterwards to ensure you get word and not all occurances of "By"
    rx_to_very_first = r'^.*?{}'.format(re.escape(very_first_up_to_word))
    Extract = re.sub(rx_to_very_first, '', Extract, flags=re.DOTALL).strip()  
    up_to_word = "Amount"                                                   # Remove Header
    rx_to_first = r'^.*?{}'.format(re.escape(up_to_word))                   # Remove Header
    Extract = re.sub(rx_to_first, '', Extract, flags=re.DOTALL).strip()     # Remove Header
    rx_after_last = r'(Creditor.*)'.format(re.escape("Creditor"))           # Remove Header 
    Extract = re.sub(rx_after_last, '', Extract, flags=re.DOTALL).strip()   # Remove Header

    return Extract

def QTAmounts(Extract):                                         # QT specific handling for identifying amounts and removing running balance
    pattern = re.compile(r"\$+[0-9]+[,0-9]*\.[0-9]{2}")         # Identifies numbers after a $ up to 2 d.p. including with ','
    AllAmounts = re.findall(pattern, Extract)
    AllAmounts = [dollarFormatting(x) for x in AllAmounts]      # Format to float
    Amounts = AllAmounts[::2]                                   # Skip every second item [0, 2, 4, 6 etc]
    return Amounts

def crystalAmounts(Extract):        # Powered By Crystal specific handling for identifying amounts and removing running balance
    pattern = re.compile(r"\$+[0-9]+[,0-9]*\.[0-9]{2}")         # Identifies numbers after a $ up to 2 d.p. including with ','
    AllAmounts = re.findall(pattern, Extract)
    AllAmounts = [dollarFormatting(x) for x in AllAmounts]      # Format to float
    Amounts = AllAmounts[::2]                                   # Skip every second item [0, 2, 4, 6 etc]
    return Amounts

def validationCustNum(custnum):                         # Function for validating customer number
    global FileErrorCount
    if custnum in ValidityMaster:                       # Happy path, if customer number in Master Data
        if ValidityMaster[custnum] not in ['Active (A)', 'ACTIVE']:
            PerformDict['Customer Number Errors Count'] = (PerformDict['Customer Number Errors Count'] + 1)
            FileErrorCount = FileErrorCount+1
        return custnum, ValidityMaster[custnum]

    # Else, check if first digit has been truncated
    TruncCustNum  = custnum.replace('-','')
    TruncCustNum = TruncCustNum[0:6]+'-'+TruncCustNum[6:8]
    result = [((key, value)) for key, value in ValidityMaster.items() if TruncCustNum in key]
    if len(result) == 1:                            # if only one possible result returned
        PerformDict['Customer Number Errors Count'] = (PerformDict['Customer Number Errors Count']+1)
        PerformDict['Customer Truncation Fix'] = (PerformDict['Customer Truncation Fix']+1)
        FileErrorCount = FileErrorCount +1
        return result[0][0], (result[0][1]+" Replaced using Truncation Wildcard")   # Identify replace items
    elif len(result) == 0:                            # customer number not in Master Data
        PerformDict['Customer Number Errors Count'] = (PerformDict['Customer Number Errors Count']+1)
        FileErrorCount = FileErrorCount +1
        return custnum, [x for item in result for x in item]
    else:                                           # if more than one possible result returned
        PerformDict['Customer Number Errors Count'] = (PerformDict['Customer Number Errors Count'] + 1)
        PerformDict['Customer Possibles Provided'] = (PerformDict['Customer Possibles Provided'] + 1)
        FileErrorCount = FileErrorCount +1
        return "Cust Num Error",[x for item in result for x in item]

def sortCustomerNumFormat(refDetails):                  # Function for identifying customer number in ref details / ref cleanup
    global FileErrorCount
    HypPattern = re.compile(r"[0-9]{7}\-[0-9]{2}")      # Cust Num in proper format
    NoHypPattern = re.compile(r"[0-9]{9}")              # Cust Num all digits
    CustNums = []
    for item in refDetails:                 # Extra clean up
        if '\n' in item:                    # Remove any '\n'
            item = item.replace('\n','')
        if '.' in item[0:10]:               # Look for amount at beginning of ref item and remove
            DecIndex = item.find('.',1)
            item = item[DecIndex+3:]
        CustFind = re.findall(HypPattern, item)                     # Look for proper format first
        if len(CustFind) > 0:
            if len(CustFind) == 1:
                CustFind = CustFind[0]
            else:
                CustFind = 'Cust Num Error'                     # Will error if more than one possible cust num
                PerformDict['Customer Number Errors Count'] = (PerformDict['Customer Number Errors Count'] + 1)
                PerformDict['Hopeless Error Count'] = (PerformDict['Hopeless Error Count']+1)
                FileErrorCount =FileErrorCount+1
        else:   # Check more than one pattern of nums
            CheckList = re.findall(NoHypPattern, item)
            if len(CheckList) > 1:
                CustFind = 'Cust Num Error'                     # Will error if more than one possible cust num
                PerformDict['Customer Number Errors Count'] = (PerformDict['Customer Number Errors Count'] + 1)
                PerformDict['Hopeless Error Count'] = (PerformDict['Hopeless Error Count']+1)
                FileErrorCount = FileErrorCount+1
            else:
                try:
                    CustFind = re.findall(NoHypPattern, item)[0]    # try find 9 digits in a row
                    CustFind = CustFind[:7]+"-"+CustFind[7:]        # format cust num correctly
                except IndexError:
                    CustFind = 'Cust Num Error'                     # Will error if Cust Ref missing altogether
                    PerformDict['Customer Number Errors Count'] = (PerformDict['Customer Number Errors Count'] + 1)
                    PerformDict['Hopeless Error Count'] = (PerformDict['Hopeless Error Count']+1)
                    FileErrorCount =FileErrorCount+1
        if ValidFlag:                                               # If Validity Master file read in
            if CustFind == 'Cust Num Error':
                CustStatus = 'Please Review'
            else:
                CustFind, CustStatus = validationCustNum(CustFind)  # Call validation func
        else:                                                       # Else set Status to indicate not checked
            CustStatus = "No Validity Master file"
        CustNums.append((CustFind, CustStatus, item))
    return CustNums

def ziplist(refDetails, billDetails):       # Function for zipping reference details and remit amounts together
    global FileRemitCount
    FileRemitCount = FileRemitCount + len(billDetails)
    ZipList = []
    for x, y in zip(refDetails, billDetails):
        if y != 0:                          # Don't include zero balances
            ZipList.append((x[0], y, x[1], x[2]))
    return ZipList

def zipstatement(ziplist, page, balance):               # Function for updating the PDF Dictionary with the zipped ref/amount list
    if page > 0:                                        # Middle / Last page handling (where key already exists)
        CurrentContents = MasterDict[PDFName]['Zipped']
        for item in ziplist:
            CurrentContents.append(item)
        balance = round(balance + MasterDict[PDFName]['Sum Total'],2)
        MasterDict[PDFName].update({'Zipped': CurrentContents})
        MasterDict[PDFName].update(({'Sum Total':balance}))
    else:                                               # First page handling (where the key doesn't already exist)
        MasterDict[PDFName].update({'Zipped': ziplist})
        MasterDict[PDFName].update({"Sum Total": balance})
    return

def creditIdentify(Sum, Total):     # Function for identifying credit, only works on singular, will error if more than 1 credit
    global PossCredit, PDFName
    PossCredit = True                                                                                # Flag possible credit
    CredAmt = round(((Sum - Total)/2),2)                                                            # Identify credit amount
    for item in MasterDict[PDFName]['Zipped']:                                                      # Check amounts to try find match
        if str(CredAmt) in str(item[1]):                                                            # If match found
            NewItem = (item[0], -item[1], item[2], item[3])                                         # Change amount to a negative
            MasterDict[PDFName]['Zipped']= [NewItem if x == item else x for x in MasterDict[PDFName]['Zipped']]     # Replace item with updated negative
            UpdateSumTotal = []                                                                      # Update the sum total for the PDF Master
            for item in MasterDict[PDFName]['Zipped']:
                UpdateSumTotal.append(item[1])
                MasterDict[PDFName].update({'Sum Total': round(sum(UpdateSumTotal),2)})
    if MasterDict[PDFName]['Sum Total'] == Total:                                                   # If Sum and Total match now
        return True                                                                                # Move on, else Error
    else:
        amountErrorType(PDFName, "amount mismatch in this file, please process manually and review for credits")
        return False

def amountChecking():                               # Function to error check sum total, settlement total (Powered By Crystal) and Statement Total
    global PDFName, CredFlag, PossCredit
    Total = MasterDict[PDFName]['Statement Total']
    Sum = MasterDict[PDFName]['Sum Total']
    if pdf_info['/Producer'] == 'Powered By Crystal':           # Code related to Powered By Crystal PDFs
        Total = MasterDict[PDFName]['Settlement Total']         # Use Settlement instead of Statement Total
        if MasterDict[PDFName]['Statement Total'] == MasterDict[PDFName]['Settlement Total']:
            pass
        else:                                                   # Possible Error Handling for Powered By Crystal only
            if MasterDict[PDFName]['Statement Total'] < MasterDict[PDFName]['Settlement Total']:
                amountErrorType(PDFName, "Totals don't match, error reading this file, please process manually")
                PossCredit = False
                return True
    if Sum == Total:        # If Totals Match
        PossCredit = False
        return False
    elif Sum > Total:       # If Sum is more, indicates credit
        PossCredit = True
        CredFlag = creditIdentify(Sum, Total)       # Call function to identify credit
        if CredFlag:                                # if successfully identified, no error
            return False
        else:                                       # Else return error
            return True
    else:                   # Else Error
        amountErrorType(PDFName, "Totals don't match, error reading this file, please process manually")
        return True

start_time = time.time()    # Variable for capturing run time
TxtFileAppend = time.strftime("%Y_%m_%d_%H%M", time.localtime(time.time()))     # Date / Time for file name convention

try:    # Try to read in validity master and set validity flag if successful
    #CustMasterList = pandas.read_csv('Book1.csv')
    CustMasterList = pandas.read_excel('MasterListAccNum.xlsx')
    ValidityMaster = dict(zip(CustMasterList['Account Number'], CustMasterList['Account Status']))
    ValidityMaster_keys = ValidityMaster.keys()
    ValidFlag = True
except:
    ValidFlag = False
Validity_time = (time.time() - start_time)

#Global Variables
PerformDict= {}                                                                         # Dictionary for storing performance data
RecognisedProducer = ["Powered By Crystal","ActiveReports Developer","Qt 5.5.1"]        # Running list of identified software that the code works on
MasterDict = {}                                                                         # Dictionary for storing each PDFs details
writeString = []                                                                        # Variable used for writing the dictionary details to before writing to CSV
ManualWriteString = []                                                                  # Variable for capturing manual handle files data

# Performance Stats
PerformDict['Validation File Read in [secs]'] = round(Validity_time,3)  # Time taken to read in validity master data file
PerformDict['File Count'] = 0                                           # Total count of files
PerformDict['File - Success'] = 0                                       # Count of files processed
PerformDict['File - Error'] = 0                                         # Count of files that couldn't be processed
PerformDict['Remittances Count'] = 0                                    # Count of all remittances
PerformDict['Customer Number Errors Count'] = 0                         # All errors identified
PerformDict['Customer Truncation Fix'] = 0                              # Errors fixed by truncation code
PerformDict['Customer Possibles Provided'] = 0                          # Truncation errors that returned more than one possible customer number
PerformDict['Hopeless Error Count'] = 0                                 # Errors that require manual handling
PerformDict['Estimated Time Saved'] = 0                                 # Calculated based on processing time and number of errors corrected

#os.chdir('c:\\Temp\\Receipting')
## MAIN

# get PDFs in dir
FileDir = []                                            # Identify directory and create list of file names to read
for file in os.listdir():                               # Read in all PDFs in directory
    if file.endswith(".pdf") or file.endswith(".PDF"):
        Newfile = file.replace('.pdf','').replace('.PDF','')
        NewFile = Newfile.replace('[','').replace(']','').replace('#',' ').replace('$','').replace(',','').replace(':','') # Removes illegal char from filename for Powershell exe
        NewFile = NewFile[:67]+'.pdf'                   # Truncate File name to be 68 char + '.pdf'
        os.rename(file,NewFile)                         # Renames file in directory
        FileDir.append(NewFile)

for file in FileDir:        # For each file in the directory, read and extract relevant information

    print(f"file: {file}")

    PerformDict['File Count'] = (PerformDict['File Count'] +1)  # Keeping count of number of files
    FileRemitCount = 0                                  # Variable for keeping count of number of remittances per file
    FileErrorCount = 0                                  # Variable for keeping count of number of errors per file
    FileScore = 0                                       # Scorecard for statement
    PDFName = file                                      # Variable for filename
    ErrorFlag = False                                   # Flag for error identified
    PossCredit = False                                  # Flag for credit identified
    CredFlag = False                                    # Used when credit positively identified
    pdfFileObj = open(PDFName, 'rb')                    # Open PDF
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)        # Read PDF
    pdf_info = pdfReader.getDocumentInfo()              # Variable for PDF Properties
    NumPages = pdfReader.numPages                       # Num of pages, required to help identify first, last, or middle pages and appropriate handling required
    MasterDict[PDFName]= {'Pages' : NumPages}           # Set amount of pages in dictionary details
    if '/Producer' in pdf_info:
        if pdf_info['/Producer'] in RecognisedProducer: # If file isn't recognised type, Error and move on to next
            pass
        else:
            fileReadErrorType(PDFName, "isn't a recognised file type, please process manually")
            continue
    else:
        fileReadErrorType(PDFName, "isn't a recognised file type, please process manually")
        continue
    for page in range(0,NumPages):                      # Read Page(s) - each page in the PDF is read in sequence in this loop
        pageObj = pdfReader.getPage(page)               # Get page one at a time to read
        PDFExtract = pageObj.extractText()              # Variable for extracted page
        if len(PDFExtract) == 0:                        # If file can't be read - scanned images will fail here (though they should have already failed due to not recognised type)
            fileReadErrorType(PDFName, "couldn't be read, please process manually. If image file can be converted to reg PDF, save and try again")
            break
        Branch = PDFName                                            # Use PDF name as Branch ID
        MasterDict[PDFName]['Branch']= Branch                       # Update PDF Dict with Branch name
        if pdf_info['/Producer'] == 'ActiveReports Developer':      # Code for ActiveReports file type
            if 'Total$' in PDFExtract or 'Total $' in PDFExtract:                              # Identify Statement Total
                StatementTotal = statementAmt(PDFExtract)
                MasterDict[PDFName]['Statement Total'] = StatementTotal
                MasterDict[PDFName]['Settlement Total'] = StatementTotal
            PDFExtract = activeReportsCleanup(PDFExtract, page)
            if len(PDFExtract) > 0:     # Find remittance amounts and create list
                SplitPattern = re.compile(r"\$+[\d]+[,\d]+\.[0-9][0-9]|\$+[\d]+\.[0-9][0-9]")
                SplitAmounts = re.findall(SplitPattern,PDFExtract)
                BillAmounts = []
                for item in SplitAmounts:
                    NewItem = dollarFormatting(item)
                    BillAmounts.append(NewItem)

                RefList = re.split(SplitPattern,PDFExtract) # Split list into references
                if '' in RefList:   # Remove blank item from end of list
                    RefList.pop(RefList.index(''))
                if '\nTotal' in RefList:
                    RefList.pop(RefList.index('\nTotal'))
                if '\nTotal ' in RefList:
                    RefList.pop(RefList.index('\nTotal '))
                DatePattern = re.compile(r'[0-9]{2}/[0-9]{2}/[0-9]{4}') # Remove dates from refs to reduce confusion pre cust num ID
                for item in RefList:
                    DateFound = re.findall(DatePattern, item)
                    if len(DateFound) > 0:
                        DateInd = item.find(DateFound[0])
                        NewItem = item[:DateInd]
                        RefList = [NewItem if x == item else x for x in RefList]
                if len(RefList) != len(BillAmounts):        # Error checking for mismatch of amounts and references
                    fileReadErrorType(PDFName, "amount mismatch, please process this file manually")
                    break
                RefList = sortCustomerNumFormat(RefList)
                ZipRefList = ziplist(RefList, BillAmounts)  # Zip everything together
                zipstatement(ZipRefList, page, round(sum(BillAmounts),2))   # Update PDF Master dictionary

        if pdf_info['/Producer'] == 'Powered By Crystal':           # Code related to Powered By Crystal PDFs
            PDFExtract = removeZeros(PDFExtract)                    # Clean up zeros
            if 'Settlement (' in PDFExtract:                        # Find Settlement Amount
                SettlementTotal = settlementAmt(PDFExtract)
                MasterDict[PDFName]['Settlement Total']=SettlementTotal
            if page == NumPages -1:                                 # Last Page, find Statement Total
                StatementTotal = statementAmt(PDFExtract)
                MasterDict[PDFName]['Statement Total']=StatementTotal
            PDFExtract = poweredByCrystalCleanUp(PDFExtract, page)  # Remove Header, Footer
            if PDFExtract is None:      # If last page has nothing on it except Settlement or Total
                pass
            else:
                BillAmounts = crystalAmounts(PDFExtract)    # Func to collect list of amounts
                try:                                        # Split extract into separate references
                    PDFExtract = PDFExtract.split("$",1)[1] # Ensure any unnecessary text before first amount is removed
                    PDFExtract = PDFExtract.split("$")      # Split by '$'
                    RefList = PDFExtract[1::2]              # Add only reference strings to RefList
                except IndexError:                          # Happens when only Settlement on last page
                    break

                if len(RefList) != len(BillAmounts):        # Error checking for mismatch of amounts and references
                    fileReadErrorType(PDFName, "amount mismatch, please process this file manually")
                    break
                RefList = sortCustomerNumFormat(RefList)    # Sort ref list into correct customer number format
                ZipRefList = ziplist(RefList, BillAmounts)  # Zip everything together
                zipstatement(ZipRefList, page, round(sum(BillAmounts),2))   # Update PDF Master dictionary

        if pdf_info['/Producer'] == 'Qt 5.5.1':                     # Code for QT file type
            if 'Total' in PDFExtract:                               # Identify Statement Total
                StatementTotal = statementAmt(PDFExtract)
                MasterDict[PDFName]['Statement Total'] = StatementTotal
                # NO Settlement Total In Qt 5.5.1, REPLACING WITH Statement Total MEANS THE OUTPUT CSV IS PROPERLY ALIGNED
                MasterDict[PDFName]['Settlement Total'] = StatementTotal
            PDFExtract = QTCleanup(PDFExtract)                      # Remove Header, Footer
            if len(PDFExtract) > 0:                                 # Find remittance amounts and create list
                SplitPattern = re.compile(r"\$+[\d]+[,\d]+\.[0-9][0-9]|\$+[\d]+\.[0-9][0-9]")
                SplitAmounts = re.findall(SplitPattern,PDFExtract)
                BillAmounts = []
                for item in SplitAmounts:
                    NewItem = dollarFormatting(item)
                    BillAmounts.append(NewItem)

                RefList = re.split(SplitPattern,PDFExtract)         # Split list into references
                if '' in RefList:                                   # Remove blank item from end of list
                    RefList.pop((RefList.index('')))

                if '\nTotal\n' in RefList:
                    RefList.remove('\nTotal\n')
                    del BillAmounts[-1]

                DatePattern = re.compile(r'[0-9]{2}/[0-9]{2}/[0-9]{2}') # Remove dates from refs to reduce confusion pre cust num ID
                updated_ref_list = []
                for item in RefList:
                    invoice_no_pattern_full = re.compile(r'\D(\d{9})\D') 
                    invoice_no_pattern_hyphen = re.compile(r'[0-9]{7}-[0-9]{2}')
                    if len(re.findall(invoice_no_pattern_full, item)) != 0:
                        if len(re.findall(invoice_no_pattern_full, item)) > 1:
                            invoice_no_parsed = "Cust Num Error"
                            PerformDict['Customer Number Errors Count'] = (PerformDict['Customer Number Errors Count'] + 1)
                            PerformDict['Hopeless Error Count'] = (PerformDict['Hopeless Error Count']+1)
                            FileErrorCount =FileErrorCount+1
                        else:
                            invoice_no_parsed = re.findall(invoice_no_pattern_full, item)[0]
                            invoice_no_parsed = invoice_no_parsed[:7]+"-"+invoice_no_parsed[7:]
                        updated_ref_list.append(invoice_no_parsed)
                    elif len(re.findall(invoice_no_pattern_hyphen, item)) != 0:
                        updated_ref_list.append(re.findall(invoice_no_pattern_hyphen, item)[0])
                    else:
                        updated_ref_list.append(item)
                RefList = updated_ref_list
                RefList = sortCustomerNumFormat(RefList)
                ZipRefList = ziplist(RefList, BillAmounts)  # Zip everything together
                zipstatement(ZipRefList, page, round(sum(BillAmounts),2))   # Update PDF Master dictionary
    if ErrorFlag == True:          # If error in file, continue to next file
        continue

    if amountChecking() == True:   # If error, continue to next file
        continue

    if ErrorFlag == False:         # If no errors in file, write to writeString to add to output
        try:
            for item in MasterDict[PDFName]['Zipped']:
                writeString.append([PDFName, item[0], item[1], item[2], item[3], MasterDict[PDFName]['Statement Total'], MasterDict[PDFName]['Settlement Total']])
            PerformDict['File - Success'] = (PerformDict['File - Success'] + 1)
            PerformDict['Remittances Count'] = (PerformDict['Remittances Count'] + FileRemitCount)
        except:
            fileReadErrorType(PDFName, "error writing output file, please process manually")
            continue

# Write output CSV
with open("Receipting_Output_" + TxtFileAppend + ".csv", "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Filename", "Customer Number", "Remittance Amount", "Customer Status", "Ref Detail", "Statement Total", "Settlement Total"])
    writer.writerows(writeString)
print(f"Finished Receipting - {time.strftime('%X', time.localtime(time.time()))}")

# Print summary stats
print(PerformDict)
