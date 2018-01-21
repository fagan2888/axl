from sys import argv, exc_info
from traceback import format_tb
from pywintypes import IID

import win32com.client
import win32com.server.util
import win32com.server.dispatcher
import win32com.server.policy
import win32api
import pythoncom
import re

from .converters import from_excel
from .imports import _imports
from . import methods

class CommandLoop(object):
    _public_methods_ = ['Call']
    _local_methods = ['Log', 'Save', 'Load']

    log_ = []
    dolog_ = False
    cache_ = {}

    def Log(self, *args):
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
            ans = ''
        if args and args[0] is not None:
            self.dolog_ = bool(args[0])
        return ans

    def Save(self, addr, obj):
        '''Saves the specified object to the Excel cache for later use.

        Inputs:
            addr: the address of the Excel cell where this data is "saved".
            obj: the Python object to save.'''
        self.cache_[addr] = obj

    def Load(self, addr):
        '''Loads the specified object from the Excel cache.

        Inputs:
            addr: the address of the Excel cell where this data is "saved".
        Outputs:
            The object stored under the given name. If there is no such
            object, None is returned. No errors are returned in this case,
            because a failure to find an object may be due to an out-of-order
            calculation by Excel. By silently failing, Excel allows the full
            recalculation to eventually complete.'''
        return self.cache_.get(addr)

    @staticmethod
    def range2var(rng):
        return rng.rsplit(']', 1)[-1].replace('!','_').replace(' ','').replace('$','').replace(':','_').replace('Sheet1_','')

    def Call(self, queue):
        # Process the calls in order, pushing the results onto a
        # result stack for potential later use. The top of the stack
        # will be returned by the function.
        commands, arguments = queue
        output_values = []
        output_range = commands[0]
        if self.dolog_:
            output_reprs = []
            output_repr = ''
        for cmd_name, cmd_args in zip(commands[1:], arguments[1:]):
            dolog = self.dolog_
            args = []
            kwargs = {}
            key = None
            if self.dolog_:
                arg_reprs = []
            for arg in map(from_excel, cmd_args):
                if self.dolog_:
                    arg_repr = repr(arg)
                tstr = type(arg) is str
                if tstr and arg[:2] == '!$':
                    arg = int(arg[2:])
                    if dolog:
                        arg_repr = output_reprs[arg]
                    arg = output_values[arg]
                    tstr = False
                if key is not None:
                    if self.dolog_:
                        arg_reprs.append('{}={}'.format(key, arg_repr))
                    kwargs[key] = arg
                    key = None
                elif tstr and arg[-1:] == '=':
                    key = arg[:-1]
                elif kwargs:
                    return '''#PYTHON?
                        Error encountered parsing Python function "{}":
                        Expected a keyword string, found this: {}'''.format(cmd_name, repr(arg))
                else:
                    args.append(arg)
                    if dolog:
                        arg_reprs.append(arg_repr)
            if key is not None:
                return '''#PYTHON?
                    Error encountered parsing Python function "{}":
                    Missing argument value for keyword "{}"'''.format(cmd_name, key)
            try:
                if cmd_name.startswith('%'):
                    cmd_name = cmd_name[1:]
                    obj = getattr(self, cmd_name)
                    if self.dolog_:
                        if cmd_name == 'Load':
                            output_repr = self.range2var(args[0])
                        elif cmd_name == 'Save' and not (type(cmd_args[1]) is str and cmd_args[1].startswith('!$')):
                            output_repr = arg_reprs[1]
                elif cmd_name.startswith('@'):
                    cmd_name = cmd_name[1:]
                    obj = getattr(methods, cmd_name)
                    if self.dolog_:
                        output_repr = 'axlm.{}({})'.format(cmd_name, ', '.join(arg_reprs))
                else:
                    cmd_parts = cmd_name.split('.')
                    if cmd_parts[0] == '':
                        obj = args[0]
                        cmd_name = cmd_name[1:]
                        if self.dolog_:
                            output_repr = '{}.{}({})'.format(arg_reprs[0], cmd_name, ', '.join(arg_reprs[1:]))
                    else:
                        obj = _imports if hasattr(_imports, cmd_parts[0]) else __builtins__
                        if self.dolog_:
                            output_repr = '{}({})'.format(cmd_name, ', '.join(arg_reprs))
                    for ftok in cmd_parts:
                        obj = getattr(obj, ftok)
                output_value = obj(*args, **kwargs)
            except:
                estr = exc_info()
                tt = "   ".join(format_tb(estr[2]))
                return '''#PYTHON?
                    Error encountered executing Python function {}:
                        {}: {}
                        {}'''.format(cmd_name, estr[0].__name__, str(estr[1]), tt)
            output_values.append(output_value)
            if dolog:
                output_reprs.append(output_repr)

        # Wrap in an extra tuple so the calling function does not
        # attempt to unpack it.
        if dolog and output_repr:
            final_name = self.range2var(output_range) if output_range else '_Out'
            final_line = '{} = {}'.format(final_name, output_repr)
            self.log_.append(final_line)
        return (output_value,) if type(output_value) is tuple else output_value


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
