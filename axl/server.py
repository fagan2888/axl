from sys import argv, exc_info
from traceback import format_tb
from pywintypes import IID

import win32com.client
import win32com.server.util
import win32com.server.dispatcher
import win32com.server.policy
import win32api
import pythoncom

from .converters import from_excel
from .imports import _imports

class CommandLoop(object):
    _public_methods_ = ['Call']
    _local_methods = ['Log', 'Save', 'Load']

    log_ = []
    dolog_ = True
    cache_ = {}

    def Log(self, active=None):
        '''Activates/deactivates logging, and returns the log output.

        Inputs:
            active: If True, activates logging. If False, deactivates logging.
            If None or omitted, preserves the current logging state.
        Outputs:
            A list of all commands issued since the last call to Log(),
            represented as a string. If logging has not been taking place,
            the string 'None' is returned.'''
        if self.dolog_:
            ans = "\n".join(self.log_)
            self.log_ = []
        else:
            ans = 'None'
        if active is not None:
            self.dolog_ = bool(active)
        return ans

    def validate_name_(name):
        if type(name) is not str:
            raise RuntimeError("Name must be a string, got this instead: "+str(name))
        elif name == '' or name == '0':
            raise RuntimeError("Name must not be empty or '0'")

    def Save(self, name, obj):
        '''Saves the specified object to the Excel cache for later use.

        Inputs:
            name: a string containing the name of the object to save. The
                string must not be empty, nor must it be the value '0'.
            obj: the Python object to save.
        Outputs:
            The name string. In Excel, the cell that stores this string
            becomes a placeholder for this Python object. Whenever the
            data used to construct the object is changed, Excel will
            automatically recalculate the object's value (unless automatic
            recalculation is disabled).'''
        validate_name_(name)
        self.cache_[name] = obj
        return name

    def Load(self, name):
        '''Loads the specified object from the Excel cache.

        Inputs:
            name: a string containing the name of the object to load.
        Outputs:
            The object stored under the given name. If there is no such
            object, None is returned. No errors are returned in this case,
            because a failure to find an object may be due to an out-of-order
            calculation by Excel. By silently failing, Excel allows the full
            recalculation to eventually complete.'''
        validate_name_(name)
        return self.cache_.get(str(name))

    def Call(self, queue):
        # Process the calls in order, pushing the results onto a
        # result stack for potential later use. The top of the stack
        # will be returned by the function.
        vals = []
        if self.dolog_:
            vnames = []
            for val in queue:
                if val[0] == 'Save':
                    vname = "'" + val[1] + "'"
                    vnames[int(val[2][2:])] = val[1]
                elif val[0] == 'Load':
                    vname = val[1]
                else:
                    vname = '_' + str(len(vnames))
                vnames.append(vname)
        for val in queue:
            fname = val[0]
            args = []
            xargs = []
            kwargs = {}
            key = None
            for k2 in range(1, len(val)):
                arg = from_excel(val[k2])
                if self.dolog_:
                    xarg = repr(arg)
                tstr = type(arg) is str
                if tstr and arg[:2] == '!$':
                    arg = int(arg[2:])
                    if self.dolog_:
                        xarg = vnames[arg]
                    arg = vals[arg]
                    tstr = False
                if key is not None:
                    if self.dolog_:
                        xargs.append(key + '=' + str(xarg))
                    kwargs[key] = arg
                    key = None
                elif tstr and arg[-1:] == '=':
                    key = arg[:-1]
                elif kwargs:
                    return '''#PYTHON?
                        Error encountered parsing Python function "{}":
                        Expected a keyword string, found this: {}'''.format(fname, str(arg))
                else:
                    args.append(arg)
                    if self.dolog_:
                        xargs.append(str(xarg))
            if key is not None:
                return '''#PYTHON?
                    Error encountered parsing Python function "{}":
                    Missing argument value for keyword "{}"'''.format(fname, key)
            try:
                if fname in self._local_methods:
                    obj = self
                elif fname.startswith('.'):
                    oname = xargs[0] + '.'
                    fname = fname[1:]
                    xargs = xargs[1:]
                    obj = getattr(args[0], fname)
                else:
                    obj = _imports
                    oname = ''
                if self.dolog_ and obj is not self:
                    self.log_.append('{} = {}{}({})'.format(vnames[len(vals)], oname, fname, ','.join(xargs)))
                for ftok in fname.split('.'):
                    obj = getattr(obj, ftok)
                val = obj(*args, **kwargs)
            except:
                estr = exc_info()
                tt = "   ".join(format_tb(estr[2]))
                return '''#PYTHON?
                    Error encountered executing Python function {}:
                        {}: {}
                        {}'''.format(fname, estr[0].__name__, str(estr[1]), tt)
            if val is None:
                break
            vals.append(val)

        # Wrap in an extra tuple so the calling function does not
        # attempt to unpack it.
        return (val,) if type(val) is tuple else val


def execute(clsid):
    clsid = IID(clsid)
    BaseDefaultPolicy = win32com.server.policy.DefaultPolicy

    class MyPolicy(BaseDefaultPolicy):
        def _CreateInstance_(self, reqClsid, reqIID):
            if reqClsid == clsid:
                return win32com.server.util.wrap(CommandLoop(), reqIID)
            else:
                return BaseDefaultPolicy._CreateInstance_(self, clsid, reqIID)
    win32com.server.policy.DefaultPolicy = MyPolicy
    factory = pythoncom.MakePyFactory(clsid)
    clsctx = pythoncom.CLSCTX_LOCAL_SERVER
    flags = pythoncom.REGCLS_MULTIPLEUSE | pythoncom.REGCLS_SUSPENDED
    try:
        revokeId = pythoncom.CoRegisterClassObject(clsid, factory, clsctx, flags)
        pythoncom.EnableQuitMessage(win32api.GetCurrentThreadId())
        pythoncom.CoResumeClassObjects()
        pythoncom.PumpMessages()
    finally:
        pythoncom.CoRevokeClassObject(revokeId)
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    clsid = argv[1]
    print('Starting a COM server with CLSID={}'.format(clsid))
    execute(clsid)
