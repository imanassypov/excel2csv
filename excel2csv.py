import click
import os
import pandas as pd
import glob
import re
import pprint

# Print iterations progress
def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█'):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print('\r%s |%s| %s%% %s' % (prefix, bar, percent, suffix), end = '\r')
    # Print New Line on Complete
    if iteration == total: 
        print()

def searchcsv (indir, search):
    pattern = re.compile (search)
    filelist = glob.glob(os.path.join(indir, '*.csv'), recursive=False)
    filecount = len(filelist)
    fileindex = 0
    #dictionary of lists, <filename>: <index><string>
    hit_dict = {}
    #pretty printer for dict
    pp = pprint.PrettyPrinter(indent=4)
    printProgressBar(0, filecount, prefix = 'Search Progress:', suffix = 'Complete', length = 50)
    for file in filelist:
        strng = open(file)
        fileindex = fileindex + 1
        rowindex = 0
        for lines in strng.readlines():
            rowindex = rowindex + 1
            if re.search(pattern, lines): 
                #hit_dict[file + '|L' + str(rowindex) +'|'] = lines
                if file in hit_dict.keys():
                    hit_dict[file][rowindex] = lines
                else:
                    hit_dict[file]={}
                    hit_dict[file][rowindex] = lines
            printProgressBar(fileindex, filecount, prefix = 'Search Progress:', suffix = 'Complete', length = 50)
    #pp.pprint (hit_dict)
    return hit_dict

def searchxls (indir, search):
    pattern = re.compile (search)
    filelist = glob.glob(os.path.join(indir, '*.xlsx'), recursive=False)
    filecount = len(filelist)
    fileindex = 0
    #dictionary of lists, <filename>: <index><string>
    hit_dict = {}
    printProgressBar(0, filecount, prefix = 'XLS Search Progress:', suffix = 'Complete', length = 50)
    for file in filelist:
        fileindex = fileindex + 1
        df1 = pd.ExcelFile(file)
        for sheet in df1.sheet_names:
            df2 = pd.read_excel(file, sheetname=sheet)
            for row_index, row_series in df2.iterrows():
                for col_index, col_series in row_series.iteritems():
                    if re.search(pattern, str(col_series)):
                        lines = ','.join(map(str, row_series.values)) 
                        filename = os.path.join(indir, os.path.basename(file) + '_' + sheet + '.' + 'xlsx')

                        if filename in hit_dict.keys():
                            hit_dict[filename][row_index] = lines
                        else:
                            hit_dict[filename]={}
                            hit_dict[filename][row_index] = lines
        printProgressBar(fileindex, filecount, prefix = 'XLS Search Progress:', suffix = 'Complete', length = 50)
    return hit_dict

def printres (hit_dict):
    #pretty printer for dict
    pp = pprint.PrettyPrinter(indent=4)
    for file in hit_dict:
        print ('\n'+file)
        for row in hit_dict[file]:
            print ('>\tR'+str(row)+'|', hit_dict[file][row].rstrip('\n'))

def dumpcsv(indir, outdir):
    path = os.path.join(indir, '*.xlsx')
    filelist = glob.glob(path, recursive=False)
    filecount = len(filelist)
    fileindex = 0
    printProgressBar(0, filecount, prefix = 'Dump Progress:', suffix = 'Complete', length = 50)
    for file in filelist:
        fileindex = fileindex + 1
        df1 = pd.ExcelFile(file)
        for sheet in df1.sheet_names:
            df2 = pd.read_excel(file, sheetname=sheet)
            filename = os.path.join(outdir, os.path.basename(file) + '_' + sheet + '.' + 'csv')
            df2.to_csv(filename, index=False)
        printProgressBar(fileindex, filecount, prefix = 'Dump Progress:', suffix = 'Complete', length = 50)

CONTEXT_SETTINGS = dict(help_option_names=['-h', '--help'])


@click.command(context_settings=CONTEXT_SETTINGS)
@click.argument('indir', type=click.Path(exists=True))
@click.argument('outdir', type=click.Path(exists=True))
@click.option('-d', '--dump', type=click.Choice([
    'yes', 'no']), default='no', help='Should I dump fresh csv out of my xlsx files from src directory to dst')
@click.option('-ss', '--searchsource', type=click.Choice([
    'yes', 'no']), default='no', help='Search in the source(xls) or destination(csv)')
@click.option('--search', prompt=True, help='Search string, must be in single quotes if contains whitespace. Regex friendly - if searching for multiple values, use | separator')

def cli(indir, outdir, dump, searchsource, search):
    '''This is a simple utility written in Python3 to search for a string (regex expressions supported) in a given directory
    containing XLSX spreadsheets. The script will go through every XLSX workbook in the source directory, and every worksheet of that workbook looking for a match.
    There is also an option to bulk export all of the worksheets out of all of the workbooks to a destination directory. 
    Search can be done both on the source XLSX directory or on the destination CSV directory (line-by-line)

    Examples:

    \b
        excel2csv src dst

    \b
        excel2csv src dst --dump yes

    \b
        excel2csv.py src dst --dump yes --searchsource no --search "Igor|Danik"
    '''
    #resulting search dict of dict's
    hit_dict = {}

    #only dump csv if required, may take a long time if a lot of large files
    if dump == 'yes':
        dumpcsv(indir, outdir)

    #lets search
    if searchsource == 'no':
        hit_dict = searchcsv(outdir,search)
    else:
        hit_dict = searchxls(indir,search)

    printres(hit_dict)

if __name__ == "__main__":
    cli()