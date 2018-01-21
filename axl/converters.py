__all__ = ['from_excel', 'to_excel']

import pandas as pd
import numpy as np

from pywintypes import TimeType
from datetime import datetime
from functools import singledispatch
from win32timezone import TimeZoneInfo

winUTC = TimeZoneInfo('GMT Standard Time', True)


# A recursive routine to convert Windows dates to naive
# datetime objects, including those found in tuples
@singledispatch
def cleanin(arg):
    return arg


@cleanin.register(TimeType)
def cleanin_timetype(arg):
    return datetime(*arg.timetuple()[:6])


@cleanin.register(tuple)
def cleanin_tuple(arg):
    return tuple(map(cleanin, arg))


# A routine to quickly prune data that will be truncated by Excel.
# There's no need to be exact here, we're just trying to reduce the
# cost of subsequent conversion and communication steps
@singledispatch
def trim(var, nr, nc, headers):
    return var


@trim.register(tuple)
@trim.register(list)
def trim_tuple(var, nr, nc, headers):
    if len(var) == 0:
        return None
    elif type(var[0]) in {list, tuple}:
        return tuple(tuple(x[:nc]) for x in var[:nr])
    elif nc == 1 and nr != 1:
        return tuple((x,) for x in var[:nr])
    else:
        return (tuple(var[:nc]),)


@trim.register(np.ndarray)
def trim_array(var, nr, nc, headers):
    if var.size == 0:
        return None
    elif var.size == 1:
        return np.asscalar(var)
    elif var.ndim == 1:
        return var[:nr, None] if nc == 1 else var[None, :nc]
    elif var.ndim == 2:
        return var[:nr, :nc]
    else:
        var = var.reshape(var.shape[0], -1)
        return var[:nr, :nc]


@trim.register(pd.DataFrame)
@trim.register(pd.Series)
def trim_dataframe(var, nr, nc, headers):
    if type(var.index) is pd.MultiIndex:
        nc = max(0, nc - len(var.index.names))
    elif var.index.name is not None:
        nc = max(0, nc - 1)
    if headers:
        nr = max(0, nr - 1)
    return pd.DataFrame(var).iloc[:nr, :nc]


# A recursive routine to convert scalars to Windows-friendly
# types, and all containers to tuples of tuples
@singledispatch
def cleanout(arg):
    if arg is not None:
        raise RuntimeError("Unexpected object of type {} found: {}".format(type(arg), arg))


@cleanout.register(int)
@cleanout.register(float)
@cleanout.register(bool)
def cleanout_base(arg):
    return arg


@cleanout.register(str)
def cleanout_str(arg):
    return arg[:32767] if len(arg) > 32767 else arg


@cleanout.register(tuple)
@cleanout.register(list)
def cleanout_tuple(arg):
    return tuple(map(cleanout, arg))


@cleanout.register(np.int64)
@cleanout.register(np.float64)
@cleanout.register(np.bool_)
def cleanout_np(arg):
    return arg.item()


@cleanout.register(pd.Timestamp)
@cleanout.register(datetime)
def cleanout_dt(arg):
    return TimeType(*arg.utctimetuple()[:6], tzinfo=winUTC)


@cleanout.register(np.datetime64)
def cleanout_npdt(arg):
    return cleanout(pd.Timestamp(arg))


@cleanout.register(dict)
def cleanout_dict(arg):
    return tuple(zip(map(str, arg.keys()), map(cleanout, arg.values())))


@cleanout.register(np.ndarray)
def cleanout_array(arg):
    return cleanout(arg.tolist())


@cleanout.register(pd.DataFrame)
@cleanout.register(pd.Series)
def cleanout_dataframe(arg):
    index = type(arg.index) is pd.MultiIndex or arg.index.name is not None
    for col in arg.columns[arg.dtypes.values == np.dtype('<M8[ns]')]:
        arg[col] = arg[col].astype(datetime)
    arg = pd.DataFrame(arg).to_records(index, False)
    return (arg.dtype.names,) + cleanout(arg.tolist())


def from_excel(val):
    return cleanin(val)


def to_excel(val, nr, nc, headers=True):
    nr = int(nr)
    nc = int(nc)
    if nr < 0 or nc < 0:
        return cleanout(val)
    val = trim(val, nr, nc, headers)
    isdf = type(val) in {pd.Series, pd.DataFrame}
    val = cleanout(val)
    if isdf and not headers:
        val = val[1:]
    # Pad with spaces if necessary to fill out the array.
    if nr == 1 and nc == 1:
        while type(val) is tuple:
            if len(val) == 0:
                return ''
            val = val[0]
        return val
    if type(val) is not tuple:
        val = ((val,),)
    elif len(val) and type(val[0]) is not tuple:
        val = (val,)
    else:
        val = val[:nr]
    tc = len(val[0])
    if tc < nc:
        tmp = ('',) * (nc-tc)
        val = tuple(x + tmp for x in val)
        tc = nc
    tr = len(val)
    if tr < nr:
        val = val + ((('',) * tc),) * (nr-tr)
    return val
