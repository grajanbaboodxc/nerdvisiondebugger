import sys
import subprocess
import argparse
import xlrd, xlsxwriter

def comparatorparser():
    parser = argparse.ArgumentParser(description=__doc__, add_help=True)
    parser.add_argument("files", type=str, default=None, help="File to process.")
    parser.add_argument("-i", "--inxlsx", type=str, default=None, help="Input excel")
    parser.add_argument("-s", "--sheetname", type=str, default=None, help="Sheet name")
    parser.add_argument("-r", "--cellrange", type=str, default=None, help="cell range")
    return parser

def breakstrimmer(strinp):
    if "\r\n" not in strinp:
        return strinp
    else:
        #introduce a separator in place of breaks.
        #remove all leading and trailing separators while retaining the interleaving ones
        #this is so that strings separated by interleaving breaks will not inadvertently be concatenated.
        return strinp.replace("\r\n", "<_sep_>").strip("<_sep_>")

def processtext(bytestring):
    #decode and split bytestring ignoring non "utf-8" characters
    splt = bytestring.decode("utf-8", "ignore").split('  ')
    #remove leading and trailing line breaks in the contents of decoded list
    strlist = []
    for i in range(len(splt)):
        outstr = breakstrimmer(splt[i]).split("<_sep_>")
        for item in outstr:
            strlist.append(item.strip())
    return strlist

def col_to_num(col_str):
    """ Convert base26 column string to number. """
    expn = 0
    col_num = 0
    for char in reversed(col_str):
        col_num += (ord(char) - ord('A') + 1) * (26 ** expn)
        expn += 1
    return col_num

def checker(inplist, checkvalue):
    return checkvalue.strip() in inplist
    
def compareexcel(pdftexts, inxlsx, sheetname, cellrange):
    #Open the workbook and define the worksheet
    book = xlrd.open_workbook(inxlsx)
    sheet = book.sheet_by_name(sheetname)
    startcell = cellrange.split(":")[0]
    endcell   = cellrange.split(":")[1]
    #column of the start cell
    startcolumnstr  = startcell.rstrip('0123456789')
    startcolumn     = col_to_num(startcolumnstr)
    startcolumn    -= 1
    #row of the start cell
    startrow    = int(startcell[len(startcolumnstr):])
    startrow   -= 1
    #column of the end cell
    endcolumnstr  = endcell.rstrip('0123456789')
    endcolumn     = col_to_num(endcolumnstr)
    endcolumn    -= 1
    #row of the end cell
    endrow    = int(endcell[len(endcolumnstr):])
    endrow   -= 1


    #Create output workbook
    outbook  = xlsxwriter.Workbook('compare.xlsx')
    outsheet = outbook.add_worksheet()

    outcolumn = 0
    #work column by column in the input range
    for i in range(startcolumn, endcolumn + 1):
        outcolumn += startcolumn
        colvalues = sheet.col_values(i, startrow, endrow + 1)
        outcolpair = {}
        rowcount = 0
        for colvalue in colvalues:
            if colvalue != "":
                outcolpair[(startrow + rowcount, i)] = [colvalue, checker(pdftexts, colvalue)]
                rowcount += 1
        for k, v in outcolpair.items():
            outrow = k[0] 
            outsheet.write(outrow, outcolumn, v[0])
            outsheet.write(outrow, outcolumn + 1, v[1])
        outcolumn += 2    
    outbook.close()
        
    
def main(args=None):
    P = comparatorparser()
    A = P.parse_args(args=args)
    s_out = subprocess.check_output([sys.executable, "pdf2txt.py", A.files])
    compareexcel(processtext(s_out), A.inxlsx, A.sheetname, A.cellrange)

if __name__ == "__main__":
    main()
