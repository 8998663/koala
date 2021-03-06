# cython: profile=True

'''
Python equivalents of various excel functions
'''

# source: https://github.com/dgorissen/pycel/blob/master/src/pycel/excellib.py

from __future__ import division
import numpy as np
from datetime import datetime
from math import log
from decimal import Decimal, ROUND_HALF_UP
import re
from excelutils import (
    flatten, 
    split_address, 
    col2num, 
    num2col,
    index2addres,
    is_number,
    is_range,
    date_from_int,
    normalize_year,
    is_leap_year,
    get_max_days_in_month,
    find_corresponding_index,
    check_length,
    extract_numeric_values,
    resolve_range
)

from ..ast.Range import RangeCore as Range
from ExcelError import ExcelError, ErrorCodes

CELL_REF_RE = re.compile(r"\!?(\$?[A-Za-z]{1,3})(\$?[1-9][0-9]{0,6})$")

######################################################################################
# A dictionary that maps excel function names onto python equivalents. You should
# only add an entry to this map if the python name is different to the excel name
# (which it may need to be to  prevent conflicts with existing python functions 
# with that name, e.g., max).

# So if excel defines a function foobar(), all you have to do is add a function
# called foobar to this module.  You only need to add it to the function map,
# if you want to use a different name in the python code. 

# Note: some functions (if, pi, atan2, and, or, array, ...) are already taken care of
# in the FunctionNode code, so adding them here will have no effect.
FUNCTION_MAP = {
      "ln":"xlog",
      "min":"xmin",
      "min":"xmin",
      "max":"xmax",
      "sum":"xsum",
      "gammaln":"lgamma",
      "round": "xround"
      }

IND_FUN = [        
    "SUM",        
    "MIN",        
    "MAX",        
    "SUMPRODUCT",     
    "IRR",        
    "COUNT",      
    "COUNTA",     
    "COUNTIF",        
    "COUNTIFS",       
    "MATCH",      
    "LOOKUP",     
    "INDEX",      
    "AVERAGE",        
    "SUMIF"       
]

######################################################################################
# List of excel equivalent functions
# TODO: needs unit testing

def value(text):
    # make the distinction for naca numbers
    if text.find('.') > 0:
        return float(text)
    else:
        return int(text)


def xlog(a):
    if isinstance(a,(list,tuple,np.ndarray)):
        return [log(x) for x in flatten(a)]
    else:
        #print a
        return log(a)


def xmax(*args): # Excel reference: https://support.office.com/en-us/article/MAX-function-e0012414-9ac8-4b34-9a47-73e662c08098
    # ignore non numeric cells and boolean cells
    values = extract_numeric_values(*args)

    # however, if no non numeric cells, return zero (is what excel does)
    if len(values) < 1:
        return 0
    else:
        return max(values)


def xmin(*args): # Excel reference: https://support.office.com/en-us/article/MIN-function-61635d12-920f-4ce2-a70f-96f202dcc152
    # ignore non numeric cells and boolean cells
    values = extract_numeric_values(*args)

    # however, if no non numeric cells, return zero (is what excel does)
    if len(values) < 1:
        return 0
    else:
        return min(values)


def xsum(*args): # Excel reference: https://support.office.com/en-us/article/SUM-function-043e1c7d-7726-4e80-8f32-07b23e057f89
    # ignore non numeric cells and boolean cells

    values = extract_numeric_values(*args)

    # however, if no non numeric cells, return zero (is what excel does)
    if len(values) < 1:
        return 0
    else:
        return sum(values)

def sumif(range, criteria, sum_range = None): # Excel reference: https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b

    # WARNING: 
    # - wildcards not supported
    # - doesn't really follow 2nd remark about sum_range length

    if not isinstance(range, Range):
        return TypeError('%s must be a Range' % str(range))

    if isinstance(criteria, Range) and not isinstance(criteria , (str, bool)): # ugly... 
        return 0

    indexes = find_corresponding_index(range.values, criteria)

    if sum_range:
        if isinstance(sum_range, Range):
            return TypeError('%s must be a Range' % str(sum_range))

        def f(x):
            return sum_range.values[x] if x < sum_range.length else 0
        
        return sum(map(f, indexes))

    else:
        return sum(map(lambda x: range.values[x], indexes))
        

def average(*args): # Excel reference: https://support.office.com/en-us/article/AVERAGE-function-047bac88-d466-426c-a32b-8f33eb960cf6
    # ignore non numeric cells and boolean cells
    values = extract_numeric_values(*args)

    return sum(values) / len(values)


def right(text,n):
    #TODO: hack to deal with naca section numbers
    if isinstance(text, unicode) or isinstance(text,str):
        return text[-n:]
    else:
        # TODO: get rid of the decimal
        return str(int(text))[-n:]
        

def index(my_range, row, col = None): # Excel reference: https://support.office.com/en-us/article/INDEX-function-a5dcf0dd-996d-40a4-a822-b56b061328bd

    if isinstance(my_range, Range):
        cells = my_range.addresses
        nr = my_range.nrows
        nc = my_range.ncols
    else:
        cells, nr, nc = my_range
        cells = list(flatten(cells))
    
    if type(cells) != list:
        return ExcelError('#VALUE!', '%s must be a list' % str(cells))

    if not is_number(row):
        return ExcelError('#VALUE!', '%s must be a number' % str(row))

    if row == 0 and col == 0:
        return ExcelError('#VALUE!', 'No index asked for Range')

    if row > nr:
        return ExcelError('#VALUE!', 'Index %i out of range' % row)

    if nr == 1:
        return cells[col - 1]

    if nc == 1:
        return cells[row - 1]
        
    else: # could be optimised
        if col is None:
            return ExcelError('#VALUE!', 'Range is 2 dimensional, can not reach value with col = None')

        if not is_number(col):
            return ExcelError('#VALUE!', '%s must be a number' % str(col))

        if col > nc:
            return ExcelError('#VALUE!', 'Index %i out of range' % col)

        indices = range(len(cells))

        if row == 0: # get column
            filtered_indices = filter(lambda x: x % nc == col - 1, indices)
            filtered_cells = map(lambda i: cells[i], filtered_indices)

            return filtered_cells

        elif col == 0: # get row
            filtered_indices = filter(lambda x: int(x / nc) == row - 1, indices)
            filtered_cells = map(lambda i: cells[i], filtered_indices)

            return filtered_cells

        else:
            return cells[(row - 1)* nc + (col - 1)]    


def lookup(value, lookup_range, result_range = None): # Excel reference: https://support.office.com/en-us/article/LOOKUP-function-446d94af-663b-451d-8251-369d5e3864cb
    
    # TODO
    if not isinstance(value,(int,float)):
        return Exception("Non numeric lookups (%s) not supported" % value)
    
    # TODO: note, may return the last equal value
    
    # index of the last numeric value
    lastnum = -1
    for i,v in enumerate(lookup_range.values):
        if isinstance(v,(int,float)):
            if v > value:
                break
            else:
                lastnum = i

    output_range = result_range.values if result_range is not None else lookup_range.values

    if lastnum < 0:
        return ExcelError('#VALUE!', 'No numeric data found in the lookup range')
    else:
        if i == 0:
            return ExcelError('#VALUE!', 'All values in the lookup range are bigger than %s' % value)
        else:
            if i >= len(lookup_range)-1:
                # return the biggest number smaller than value
                return output_range[lastnum]
            else:
                return output_range[i-1]

# NEEDS TEST 
def linest(*args, **kwargs): # Excel reference: https://support.office.com/en-us/article/LINEST-function-84d7d0d9-6e50-4101-977a-fa7abf772b6d

    Y = args[0].values()
    X = args[1].values()
    
    if len(args) == 3:
        const = args[2]
        if isinstance(const,str):
            const = (const.lower() == "true")
    else:
        const = True
        
    degree = kwargs.get('degree',1)
    
    # build the vandermonde matrix
    A = np.vander(X, degree+1)
    
    if not const:
        # force the intercept to zero
        A[:,-1] = np.zeros((1,len(X)))
    
    # perform the fit
    (coefs, residuals, rank, sing_vals) = np.linalg.lstsq(A, Y)
        
    return coefs

# NEEDS TEST
def npv(*args): # Excel reference: https://support.office.com/en-us/article/NPV-function-8672cb67-2576-4d07-b67b-ac28acf2a568 
    discount_rate = args[0]
    cashflow = args[1]
    return sum([float(x)*(1+discount_rate)**-(i+1) for (i,x) in enumerate(cashflow)])


def match(lookup_value, lookup_range, match_type=1): # Excel reference: https://support.office.com/en-us/article/MATCH-function-e8dffd45-c762-47d6-bf89-533f4a37673a
    
    def type_convert(value):
        if type(value) == str:
            value = value.lower()
        elif type(value) == int:
            value = float(value)

        return value;

    lookup_value = type_convert(lookup_value)
    range_length = lookup_range.length
    range_values = lookup_range.values

    if match_type == 1:
        # Verify ascending sort
        posMax = -1
        for i in range(range_length):
            current = type_convert(range_values[i])

            if i is not range_length-1 and current > type_convert(range_values[i+1]):
                return ExcelError('#VALUE!', 'for match_type 0, lookup_range must be sorted ascending')
            if current <= lookup_value:
                posMax = i 
        if posMax == -1:
            return ('no result in lookup_range for match_type 0')
        return posMax +1 #Excel starts at 1

    elif match_type == 0:
        # No string wildcard
        try:
            return [type_convert(x) for x in range_values].index(lookup_value) + 1
        except:
            return ExcelError('#VALUE!', '%s not found' % lookup_value)

    elif match_type == -1:
        # Verify descending sort
        posMin = -1
        for i in range((range_length)):
            current = type_convert(range_values[i])

            if i is not range_length-1 and current < type_convert(range_values[i+1]):
               return ('for match_type 0, lookup_range must be sorted descending')
            if current >= lookup_value:
               posMin = i 
        if posMin == -1:
            return ExcelError('#VALUE!', 'no result in lookup_range for match_type 0')
        return posMin +1 #Excel starts at 1


def mod(nb, q): # Excel Reference: https://support.office.com/en-us/article/MOD-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3
    if not isinstance(nb, (int, long)):
        return ExcelError('#VALUE!', '%s is not an integer' % str(nb))
    elif not isinstance(q, (int, long)):
        return ExcelError('#VALUE!', '%s is not an integer' % str(q))
    else:
        return nb % q


def count(*args): # Excel reference: https://support.office.com/en-us/article/COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c
    l = list(args)

    total = 0

    for arg in l:
        if isinstance(arg, Range):
            total += len(filter(lambda x: is_number(x) and type(x) is not bool, arg.values)) # count inside a list
        elif is_number(arg): # int() is used for text representation of numbers
            total += 1

    return total

def counta(range):
    if isinstance(range, ExcelError):
        if range.value == '#NULL':
            return 0
        else:
            raise Exception('ExcelError other than #NULL passed to excellib.counta()')
    else:
        return len(filter(lambda x: x != None, range.values))

def countif(range, criteria): # Excel reference: https://support.office.com/en-us/article/COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34
    
    # WARNING: 
    # - wildcards not supported
    # - support of strings with >, <, <=, =>, <> not provided

    valid = find_corresponding_index(range.values, criteria)

    return len(valid)


def countifs(*args): # Excel reference: https://support.office.com/en-us/article/COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842

    arg_list = list(args)
    l = len(arg_list)

    if l % 2 != 0:
        return ExcelError('#VALUE!', 'excellib.countifs() must have a pair number of arguments, here %d' % l)


    if l >= 2:
        indexes = find_corresponding_index(args[0].values, args[1]) # find indexes that match first layer of countif

        remaining_ranges = [elem for i, elem in enumerate(arg_list[2:]) if i % 2 == 0] # get only ranges
        remaining_criteria = [elem for i, elem in enumerate(arg_list[2:]) if i % 2 == 1] # get only criteria

        # verif that all Ranges are associated COULDNT MAKE THIS WORK CORRECTLY BECAUSE OF RECURSION
        # association_type = None

        # temp = [args[0]] + remaining_ranges

        # for index, range in enumerate(temp): # THIS IS SHIT, but works ok
        #     if type(range) == Range and index < len(temp) - 1:
        #         asso_type = range.is_associated(temp[index + 1])

        #         print 'asso', asso_type
        #         if association_type is None:
        #             association_type = asso_type
        #         elif associated_type != asso_type:
        #             association_type = None
        #             break

        # print 'ASSO', association_type

        # if association_type is None:
        #     return ValueError('All items must be Ranges and associated')

        filtered_remaining_ranges = []

        for range in remaining_ranges: # filter items in remaining_ranges that match valid indexes from first countif layer
            filtered_remaining_cells = []
            filtered_remaining_range = []

            for index, item in enumerate(range.values):
                if index in indexes:
                    filtered_remaining_cells.append(range.addresses[index]) # reconstructing cells from indexes
                    filtered_remaining_range.append(item) # reconstructing values from indexes

            # WARNING HERE
            filtered_remaining_ranges.append(Range(filtered_remaining_cells, filtered_remaining_range))

        new_tuple = ()

        for index, range in enumerate(filtered_remaining_ranges): # rebuild the tuple that will be the argument of next layer
            new_tuple += (range, remaining_criteria[index])

        return min(countifs(*new_tuple), len(indexes)) # only consider the minimum number across all layer responses

    else:
        return float('inf')



def xround(number, num_digits = 0): # Excel reference: https://support.office.com/en-us/article/ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c

    if not is_number(number):
        return ExcelError('#VALUE!', '%s is not a number' % str(number))
    if not is_number(num_digits):
        return ExcelError('#VALUE!', '%s is not a number' % str(num_digits))

    if num_digits >= 0: # round to the right side of the point
        return float(Decimal(repr(number)).quantize(Decimal(repr(pow(10, -num_digits))), rounding=ROUND_HALF_UP))
        # see https://docs.python.org/2/library/functions.html#round
        # and https://gist.github.com/ejamesc/cedc886c5f36e2d075c5

    else:
        return round(number, num_digits)


def mid(text, start_num, num_chars): # Excel reference: https://support.office.com/en-us/article/MID-MIDB-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028
    
    text = str(text)

    if type(start_num) != int:
        return ExcelError('#VALUE!', '%s is not an integer' % str(start_num))
    if type(num_chars) != int:
        return ExcelError('#VALUE!', '%s is not an integer' % str(num_chars))

    if start_num < 1:
        return ExcelError('#VALUE!', '%s is < 1' % str(start_num))
    if num_chars < 0:
        return ExcelError('#VALUE!', '%s is < 0' % str(num_chars))

    return text[start_num:num_chars]


def date(year, month, day): # Excel reference: https://support.office.com/en-us/article/DATE-function-e36c0c8c-4104-49da-ab83-82328b832349

    if type(year) != int:
        return ExcelError('#VALUE!', '%s is not an integer' % str(year))

    if type(month) != int:
        return ExcelError('#VALUE!', '%s is not an integer' % str(month))

    if type(day) != int:
        return ExcelError('#VALUE!', '%s is not an integer' % str(day))

    if year < 0 or year > 9999:
        return ExcelError('#VALUE!', 'Year must be between 1 and 9999, instead %s' % str(year))

    if year < 1900:
        year = 1900 + year

    year, month, day = normalize_year(year, month, day) # taking into account negative month and day values

    date_0 = datetime(1900, 1, 1)
    date = datetime(year, month, day)

    result = (datetime(year, month, day) - date_0).days + 2

    if result <= 0:
        return ExcelError('#VALUE!', 'Date result is negative')
    else:
        return result


def yearfrac(start_date, end_date, basis = 0): # Excel reference: https://support.office.com/en-us/article/YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8
    
    def actual_nb_days_ISDA(start, end): # needed to separate days_in_leap_year from days_not_leap_year
        y1, m1, d1 = start
        y2, m2, d2 = end

        days_in_leap_year = 0
        days_not_in_leap_year = 0

        year_range = range(y1, y2 + 1)

        for y in year_range:

            if y == y1 and y == y2:
                nb_days = date(y2, m2, d2) - date(y1, m1, d1)
            elif y == y1:
                nb_days = date(y1 + 1, 1, 1) - date(y1, m1, d1)
            elif y == y2:
                nb_days = date(y2, m2, d2) - date(y2, 1, 1)
            else:
                nb_days = 366 if is_leap_year(y) else 365

            if is_leap_year(y):
                days_in_leap_year += nb_days
            else:
                days_not_in_leap_year += nb_days

        return (days_not_in_leap_year, days_in_leap_year)

    def actual_nb_days_AFB_alter(start, end): # http://svn.finmath.net/finmath%20lib/trunk/src/main/java/net/finmath/time/daycount/DayCountConvention_ACT_ACT_YEARFRAC.java
        y1, m1, d1 = start
        y2, m2, d2 = end

        delta = date(*end) - date(*start)

        if delta <= 365:
            if is_leap_year(y1) and is_leap_year(y2):
                denom = 366
            elif is_leap_year(y1) and date(y1, m1, d1) <= date(y1, 2, 29):
                denom = 366
            elif is_leap_year(y2) and date(y2, m2, d2) >= date(y2, 2, 29):
                denom = 366
            else:
                denom = 365
        else:
            year_range = range(y1, y2 + 1)
            nb = 0

            for y in year_range:
                nb += 366 if is_leap_year(y) else 365

            denom = nb / len(year_range)

        return delta / denom

    if not is_number(start_date):
        return ExcelError('#VALUE!', 'start_date %s must be a number' % str(start_date))
    if not is_number(end_date):
        return ExcelError('#VALUE!', 'end_date %s must be number' % str(end_date))
    if start_date < 0:
        return ExcelError('#VALUE!', 'start_date %s must be positive' % str(start_date))
    if end_date < 0:
        return ExcelError('#VALUE!', 'end_date %s must be positive' % str(end_date))

    if start_date > end_date: # switch dates if start_date > end_date
        temp = end_date
        end_date = start_date
        start_date = temp 

    y1, m1, d1 = date_from_int(start_date)
    y2, m2, d2 = date_from_int(end_date)

    if basis == 0: # US 30/360
        d2 = 30 if d2 == 31 and (d1 == 31 or d1 == 30) else min(d2, 31)
        d1 = 30 if d1 == 31 else d1

        count = 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)
        result = count / 360

    elif basis == 1: # Actual/actual
        result = actual_nb_days_AFB_alter((y1, m1, d1), (y2, m2, d2))

    elif basis == 2: # Actual/360
        result = (end_date - start_date) / 360

    elif basis == 3: # Actual/365
        result = (end_date - start_date) / 365

    elif basis == 4: # Eurobond 30/360
        d2 = 30 if d2 == 31 else d2
        d1 = 30 if d1 == 31 else d1

        count = 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)
        result = count / 360

    else:
        return ExcelError('#VALUE!', '%d must be 0, 1, 2, 3 or 4' % basis)


    return result


def isNa(value):
    # This function might need more solid testing
    try:
        eval(value)
        return False
    except:
        return True

def isblank(value):
    return value is None

def offset(reference, rows, cols, height=None, width=None): # Excel reference: https://support.office.com/en-us/article/OFFSET-function-c8de19ae-dd79-4b9b-a14e-b4d906d11b66
    # This function accepts a list of addresses
    # Maybe think of passing a Range as first argument

    for i in [rows, cols, height, width]:
        if type(i) == ExcelError:
            return i

    # get first cell address of reference
    if is_range(reference):
        ref = list(flatten(resolve_range(reference)[0]))[0]
    else:
        ref = reference
    ref_sheet = ''
    end_address = ''

    if '!' in ref:
        ref_sheet = ref.split('!')[0] + '!'
        ref_cell = ref.split('!')[1]
    else:
        ref_cell = ref

    found = re.search(CELL_REF_RE, ref)
    new_col = col2num(found.group(1)) + cols
    new_row = int(found.group(2)) + rows

    if new_row <= 0 or new_col <= 0:
        return ExcelError('#VALUE!', 'Offset is out of bounds')

    start_address = str(num2col(new_col)) + str(new_row)

    if (height is not None and width is not None):
        if type(height) != int:
            return ExcelError('#VALUE!', '%d must not be integer' % height)
        if type(width) != int:
            return ExcelError('#VALUE!', '%d must not be integer' % width)

        if height > 0:
            end_row = new_row + height - 1
        else:
            return ExcelError('#VALUE!', '%d must be strictly positive' % height)
        if width > 0:
            end_col = new_col + width - 1
        else:
            return ExcelError('#VALUE!', '%d must be strictly positive' % width)

        end_address = ':' + str(num2col(end_col)) + str(end_row)
    elif height and not width or not height and width:
        return ExcelError('Height and width must be passed together')

    return ref_sheet + start_address + end_address

def sumproduct(*ranges): # Excel reference: https://support.office.com/en-us/article/SUMPRODUCT-function-16753e75-9f68-4874-94ac-4d2145a2fd2e
    range_list = list(ranges)
    
    reduce(check_length, range_list) # check that all ranges have the same size

    return reduce(lambda X, Y: X + Y, reduce(lambda x, y: Range.apply_all('multiply', x, y), range_list).values)

def iferror(value, value_if_error): # Excel reference: https://support.office.com/en-us/article/IFERROR-function-c526fd07-caeb-47b8-8bb6-63f3e417f611

    if isinstance(value, ExcelError) or value in ErrorCodes:
        return value_if_error
    else:
        return value

def irr(values, guess = None): # Excel reference: https://support.office.com/en-us/article/IRR-function-64925eaa-9988-495b-b290-3ad0c163c1bc
                               # Numpy reference: http://docs.scipy.org/doc/numpy-1.10.0/reference/generated/numpy.irr.html
    if (isinstance(values, Range)):
        values = values.values

    if guess is not None and guess != 0:
        raise ValueError('guess value for excellib.irr() is %s and not 0' % guess)
    else:
        try:
            return np.irr(values)
        except Exception as e:
            return ExcelError('#NUM!', e)

def vlookup(lookup_value, table_array, col_index_num, range_lookup = True): # https://support.office.com/en-us/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1
    
    if not isinstance(table_array, Range):
        return ExcelError('#VALUE', 'table_array should be a Range')

    if col_index_num > table_array.ncols:
        return ExcelError('#VALUE', 'col_index_num is greater than the number of cols in table_array')

    first_column = table_array.get(0, 1)
    result_column = table_array.get(0, col_index_num)

    list = zip(first_column.order, first_column.values)
    
    if not range_lookup:
        if lookup_value not in first_column.values:
            return ExcelError('#N/A', 'lookup_value not in first column of table_array')
        else:
            i = first_column.values.index(lookup_value)
            ref = first_column.order[i]
    else:
        i = None
        for v in first_column.values:
            if lookup_value >= v:
                i = first_column.values.index(v)
                ref = first_column.order[i]
            else:
                break

        if i is None:
            return ExcelError('#N/A', 'lookup_value smaller than all values of table_array')

    return Range.find_associated_value(ref, result_column)

if __name__ == '__main__':
    pass

