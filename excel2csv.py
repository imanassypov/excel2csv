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

def searchfor (path, search):
    pattern = re.compile (search)
    filelist = glob.glob(os.path.join(path, '*.csv'), recursive=False)
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

def printres (hit_dict):
    #pretty printer for dict
    pp = pprint.PrettyPrinter(indent=4)
    for file in hit_dict:
        print ('\n'+file)
        for row in hit_dict[file]:
            print ('>\tR'+str(row)+'|', hit_dict[file][row].rstrip('\n'))

def getsheets(indir, outdir):
    path = os.path.join(indir, '*.xlsx')
    filelist = glob.glob(path, recursive=False)
    filecount = len(filelist)
    fileindex = 0
    printProgressBar(0, filecount, prefix = 'Dump Progress:', suffix = 'Complete', length = 50)
    for inputfile in filelist:
        fileindex = fileindex + 1
        df1 = pd.ExcelFile(inputfile)
        for x in df1.sheet_names:
            df2 = pd.read_excel(inputfile, sheetname=x)
            filename = os.path.join(outdir, os.path.basename(inputfile) + '_' + x + '.' + 'csv')
            df2.to_csv(filename, index=False)
        printProgressBar(fileindex, filecount, prefix = 'Dump Progress:', suffix = 'Complete', length = 50)

CONTEXT_SETTINGS = dict(help_option_names=['-h', '--help'])


@click.command(context_settings=CONTEXT_SETTINGS)
@click.argument('indir', type=click.Path(exists=True))
@click.argument('outdir', type=click.Path(exists=True))
@click.option('-d', '--dump', type=click.Choice([
    'yes', 'no']), default='no', help='Should I dump fresh csv out of my xlsx files from src directory to dst')
@click.option('--search', prompt=True, help='Search string, must be in single quotes if contains whitespace. Regex friendly - if searching for multiple values, use | separator')

def cli(indir, outdir, dump, search):
    '''Dump every sheet from a workbook to a separate csv file and search for specified value in all resulting files

    Examples:

    \b
        getsheets src dst

    \b
        getsheets src dst --dump
    '''
    #resulting search dict of dict's
    hit_dict = {}

    #only dump csv if required, may take a long time if a lot of large files
    if dump == 'yes':
        getsheets(indir, outdir)

    #lets search
    hit_dict = searchfor(outdir,search)
    printres(hit_dict)

if __name__ == "__main__":
    cli()