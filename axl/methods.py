import numpy as np
import pandas as pd

from .converters import to_excel

ToExcel = to_excel


def TupleMat(x):
    '''Converts the Excel input into a 2D tuple of tuples.

    Input:
        x: the data to be converted. It is assumed that this is a
        valid output of FromExcel(); e.g., a scalar or tuple.
    Output:
        the converted data, wrapped in additional tuple layers as
        necessary to yield a 2D array.'''
    if type(x) is not tuple:
        x = ((x, ), )
    if len(x) == 0 or type(x[0]) is not tuple:
        x = (x, )
    return x


def TupleVec(x, arg=0, flatten=False):
    '''Converts the Excel input into a 1D tuple. If the input is
    two dimensional, and both the number of rows and columns is
    greater than 1, an exception is raised.

    Input:
        x: the data to be converted. It is assumed that this is a
        valid output of FromExcel(); e.g., a scalar or tuple.
    Keyword:
        flatten: if True, 2-D arrays will be converted to 1-D arrays
        even if they do not have single rows or columns.
    Output:
        the converted data, always a 1D tuple.'''
    if type(x) is not tuple:
        return (x, )
    elif len(x) == 0 or type(x[0]) is not tuple:
        return x
    elif len(x) == 1:
        return x[0]
    elif len(x[0]) == 1:
        return tuple(y[0] for y in x)
    elif flatten:
        return sum(x, ())
    elif type(arg) is str:
        raise ValueError("Argument '{}' must be an Excel row or column, got this: {}".format(arg, str(x)))
    else:
        raise ValueError("Argument {} must be an Excel row or column, got this: {}".format(arg, str(x)))


def Transpose(arr):
    '''Transposes an Excel input.

    Input:
        arr: the data to convert.
    Output:
        a 2-D Excel array representing the transposed data.'''
    return tuple(map(tuple, zip(*TupleMat(arr))))


def ColDF(arr, labels=None):
    '''Converts an Excel input into a Pandas DataFrame, assuming that the
    columns correspond to distinct Pandas columns.

    Input:
        arr: the data to convert.
    Keywords:
        labels: a vector of column labels. If supplied, it must be a vector
        of valid labels (e.g., strings). If not, the first row of arr is
        assumed to contain the labels.
    Output:
        a Pandas dataframe.'''
    arr = TupleMat(arr)
    if labels is None:
        if len(arr) == 0:
            return pd.DataFrame()
        elif len(arr) == 1:
            return pd.DataFrame(columns=arr[0])
        else:
            return pd.DataFrame(list(arr[1:]), columns=arr[0])
    labels = TupleVec(labels, 'labels')
    if len(labels) == 0:
        return pd.DataFrame(list(arr))
    elif len(arr) == 0:
        return pd.DataFrame(columns=labels)
    elif len(arr[0]) != len(labels):
        raise ValueError("Number of labels mismatch")
    else:
        return pd.DataFrame(list(arr), columns=labels)


def RowDF(arr, labels=None):
    '''Converts an Excel input into a Pandas DataFrame, assuming that the
    *rows* correspond to distinct Pandas columns. This is precisely equivalent
    to ColDF(Transpose(arr), labels)

    Input:
        arr: the data to convert.
    Keywords:
        labels: a vector of column labels. If supplied, it must be a vector
        of valid labels (e.g., strings). If not, the first column of arr is
        assumed to contain the labels.
    Output:
        a Pandas dataframe.'''
    arr = Transpose(arr)
    return ColDF(arr, labels)


def VecDF(*args, **kwargs):
    '''Converts the input to a Pandas DataFrame. Each argument is assumed to
    be a vector corresponding to a distinct Pandas column. The column labels
    are supplied either as the last argument, or using the "labels" keyword.

    Inputs:
        *args: the data to convert. Each argument must be a vector (row or
            column). A mixture of rows and columns may be supplied. Their
            lengths must be identical, except that scalar values may be
            supplied; they will be broadcast to all rows of the DataFrame.
    Keywords:
        labels: a vector of column labels. If supplied, it must be a vector
        of valid labels (e.g., strings). If not, the last argument is assumed
        to contain the labels.
    Output:
        a Pandas dataframe.'''
    n = len(args)
    labels = kwargs.pop('labels', None)
    if labels:
        raise TypeError('Unexpected keyword arguments: %s' % ', '.join(kwargs))
    if labels is None and n > 0 and type(args[-1]) is tuple and not any(type(y) is tuple for y in args[-1]):
        labels = args[-1]
        args = args[:-1]
        n = n - 1
    if n == 0:
        return ColDF([], labels=labels, test=False)
    lx = 1
    any_one = False
    args = list(args)
    for k in range(n):
        arg = TupleVec(args[k])
        args[k] = arg
        ly = len(arg)
        if ly == 1:
            any_one = True
        elif lx == 1:
            lx = ly
        elif lx != ly:
            raise ValueError("Vectors are not the same length")
    if any_one and lx != 1:
        for k in range(n):
            arg = args[k]
            if len(arg) == 1:
                args[k] = arg * lx
    return RowDF(tuple(args), labels=labels)


def MatDF(*args, **kwargs):
    '''Converts the input to a Pandas DataFrame. Each argument is assumed to
    be a matrix corresponding to a distinct Pandas column. The column labels
    are supplied either as the last argument, or using the "labels" keyword.

    Inputs:
        *args: the data to convert. Each argument must be a matrix, whose
            sizes are identical, *except* for the application of broadcasting
            rules. So, for instance, if one argument is a column, its elements
            will be duplicated across all columns.
    Keywords:
        labels: a vector of column labels. If supplied, it must be a vector
        of valid labels (e.g., strings). If not, the last argument is assumed
        to contain the labels.
    Output:
        a Pandas dataframe.'''
    n = len(args)
    labels = kwargs.pop('labels', None)
    if labels:
        raise TypeError('Unexpected keyword arguments: %s' % ', '.join(kwargs))
    if labels is None and n > 0 and type(args[-1]) is tuple and not any(type(y) is tuple for y in args[-1]):
        labels = args[-1]
        args = args[:-1]
        n = n - 1
    if n == 0:
        return ColDF([], labels=labels, test=False)
    r = c = 1
    args = list(args)
    for k in range(n):
        arg = TupleMat(args[k])
        args[k] = arg
        lr = len(arg)
        lc = len(arg[0])
        if lr == 1:
            pass
        elif r == 1:
            r = lr
        elif r != lr:
            raise ValueError("Matrices must have the same number of rows")
        if lc == 1:
            pass
        elif c == 1:
            c = lc
        elif c != lc:
            raise ValueError("Matrices must have the same number of columns")
    for k in range(n):
        arg = args[k]
        if c != len(arg[0]):
            arg = tuple(x * c for x in arg)
        if r != len(arg):
            arg = arg[0] * r
        else:
            arg = sum(arg, ())
        args[k] = arg
    return RowDF(tuple(args), labels=labels)


def Dict(*args):
    '''Converts the input to a Python dictionary. The conversion behavior
    depends on the number of arguments:
    --- 1 argument: the argument is required to have exactly two rows or
        two columns, with keys in the first column and values in the second.
    --- 2 arguments: the arguments are assumed to have exactly the same
        number of elements, with keys in the first and values in the second.
        The shape of the arguments is ignored (and may be different).
    --- 2*N arguments: an alternating tuple of key/value pairs is assumed.

    Inputs:
        The data to be loaded into the dictionary.
    Output:
        a dictionary.'''
    n = len(args)
    if n == 1:
        arg = TupleMat(args[0])
        if len(arg[0]) != 2:
            raise ValueError("Expected a 2-column matrix")
        return dict(arg)
    elif n == 2:
        keys = TupleVec(args[0], flatten=True)
        vals = TupleVec(args[1], flatten=True)
    else:
        keys = args[::2]
        vals = args[1::2]
    if len(keys) != len(vals):
        raise ValueError("Expected the same number of keys and values")
    return dict(zip(keys, vals))


def List(*args):
    '''Converts the input to a Python list.

    Inputs:
        The data to be loaded into the list.
    Output:
        a list containing the arguments.'''
    return list(args)


def Tuple(*args):
    '''Converts the input to a Python tuple.

    Inputs:
        The data to be loaded into the tuple.
    Output:
        a tuple containing the arguments.'''
    return args


def Slice(*args):
    '''Converts the input to a Python slice.

    Inputs:
        The data to be loaded into the slice.
    Output:
        a slice constructed from the arguments.'''
    return slice(*args)


def Array(arr, dtype=None):
    '''Converts the input to a 2-D NumPy array.

    Inputs:
        The data to be loaded into the array.
    Keywords:
        dtype: A string representing a valid NumPy dtype.
            If not supplied, the dtype will be inferred in the usual manner.
    Output:
        a NumPy array.'''
    return np.array(TupleMat(arr), dtype=dtype)


def Matrix(arr, dtype=None):
    '''Converts the input to a 2-D NumPy matrix.

    Inputs:
        The data to be loaded into the matrix.
    Keywords:
        dtype: A string representing a valid NumPy dtype.
            If not supplied, the dtype will be inferred in the usual manner.
    Output:
        a NumPy matrix.'''
    return np.matrix(TupleMat(arr), dtype=dtype)


def Vector(arr, flatten=False, dtype=None):
    '''Converts the input to a 2-D NumPy vector.

    Inputs:
        The data to be loaded into the vector.
    Keywords:
        flatten: If True, accept a matrix input but flatten it
            into a vector. If False (


default), raise an Exception if
            the input is not a vector.
        dtype: A string representing a valid NumPy dtype.
            If not supplied, the dtype will be inferred in the usual manner.
    Output:
        a NumPy vector.'''
    return np.array(TupleVec(arr, arg=0, flatten=flatten), dtype=dtype)


def Row(arr, flatten=False, dtype=None):
    '''Converts the input to a 2-D NumPy row vector (shape=(1, n)).

    Inputs:
        The data to be loaded into the vector.
    Keywords:
        flatten: If True, accept a matrix input but flatten it
            into a vector. If False (default), raise an Exception if
            the input is not a vector.
        dtype: A string representing a valid NumPy dtype.
            If not supplied, the dtype will be inferred in the usual manner.
    Output:
        a NumPy vector.'''
    return Vector(arr, flatten, dtype)[None, :]


def Column(arr, flatten=False, dtype=None):
    '''Converts the input to a 2-D NumPy column vector (shape=(n, 1)).

    Inputs:
        The data to be loaded into the vector.
    Keywords:
        flatten: If True, accept a matrix input but flatten it
            into a vector. If False (default), raise an Exception if
            the input is not a vector.
        dtype: A string representing a valid NumPy dtype.
            If not supplied, the dtype will be inferred in the usual manner.
    Output:
        a NumPy vector.'''
    return Vector(arr, flatten, dtype)[:, None]


def Echo(*args):
    '''Returns a string representation of its input argument tuple.
    Useful for debugging purposes.

    Inputs:
        An arbitrary number of inputs.
    Outputs:
        A string representation of the input; specifically str(args).'''
    return str(args)


def Grab(arg):
    '''Returns its argument. Sometimes used by the Excel connector.''
    Inputs:
        Something.
    Outputs:
        That thing.'''
    return arg

def Extract(obj, *cols, **kwargs):
    '''Extract values from a single row of a DataFrame. The last argument specifies
    the column or columns to be extracted; the previous arguments specify the row.

    Inputs:
        obj: a Pandas DataFrame.
        args[:-1]: values to be matched to the elements of the index. The number of
            elements in args[:-1] must match the number of index levels.
        args[-1]: a string representing a single column, or a tuple/list of same.
            If args[-1] is a tuple, it is converted to a list.

    Outputs:
        the selected values of the DataFrame and their values
    '''
    if type(obj) is not pd.DataFrame:
        raise ValueError('First argument must be a DataFrame')
    obj = util.filter(obj, **kwargs)
    if obj.shape[0] != 1:
        raise ValueError('Filter must result in exactly one row'.format(obj.index.nlevels))
    if len(cols) == 1:
        cols = cols[0]
        if type(cols) is tuple:
            cols = list(TupleVec(cols))
    else:
        cols = list(cols)
    obj = obj[cols]
    return tuple(obj) if type(cols) is list else obj


def DFColNames(obj):
    '''return the column names for a dataframe
    Inputs:
        obj: a pandas dataframe
    Outputs:
        the column names of the inpute dataframe
    '''
    if type(obj) is not pd.DataFrame:
        raise ValueError('First argument must be a DataFrame')
    return obj.columns


def DFCols(obj, columns=None, exclude=None, sortby=None, ascending=True):
    '''Retrieves a portion of a Pandas DataFrame for display in Excel, including
    the ability to sort by one or more columns.

    Inputs:
        obj: a Pandas DataFrame.
    Keywords:
        columns: a vector of strings denoting the columns to select. These
            may be data columns *or* index columns. If omitted, all columns
            are considered (except those supplied by the "exclude" keyword.)
        exclude: a vector of strings denoting the columns to exclude. These
            may be data columns *or* index columns. If omitted, no columns
            are excluded from consideration.
        sortby: a string or vector of strings denoting a sort order.
        ascending: a boolean or vector of booleans denoting sort directions.
            if a vector is supplied, it must be the same length as sortby.
    Outputs:
        the sorted/selected DataFrame.
    '''
    if type(obj) is not pd.DataFrame:
        raise ValueError('First argument must be a DataFrame object')
    ndx = obj.index
    if type(ndx) is pd.MultiIndex:
        inames = ndx.names
    elif ndx.name is None:
        inames = []
    else:
        inames = [ndx.name]
    if columns is None:
        columns = list(inames) + list(obj.columns)
    else:
        columns = list(TupleVec(columns, arg='columns'))
    if exclude is not None:
        exclude = set(TupleVec(exclude, arg='exclude'))
        columns = [x for x in columns if x not in exclude]
    dropped = False
    if sortby is not None:
        sortby = TupleVec(sortby, arg='sortby', flatten=True)
        ascending = TupleVec(ascending, arg='ascending', flatten=True)
        if any(x in columns for x in inames):
            obj = obj.reset_index()
            dropped = True
        obj = obj.sort(list(sortby), ascending=list(ascending))
    if not dropped:
        obj = obj.reset_index()
    return obj[columns]
