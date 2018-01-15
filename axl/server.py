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
            ans = 'None'
        if args and args[0] is not None:
            self.dolog_ = bool(args[0])
        return ans

    def Save(self, name, addr, obj):
        '''Saves the specified object to the Excel cache for later use.

        Inputs:
            name: a label to be assigned to this saved data. It can be
                  any data type that can be rendered in a single Excel
                  cell, usually a string.
            addr: the address of the Excel cell where this data is "saved".
            obj: the Python object to save.
        Outputs:
            The name string. In Excel, the cell that stores this string
            becomes a placeholder for this Python object. Whenever the
            data used to construct the object is changed, Excel will
            automatically recalculate the object's value (unless automatic
            recalculation is disabled).'''
        self.cache_[addr] = obj
        return name

    def Load(self, name, addr):
        '''Loads the specified object from the Excel cache.

        Inputs:
            name: a string containing the name of the object to load.
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
        output_values = []
        output_range = queue[-1]
        if self.dolog_:
            output_names = []
            output_lines = []
        for ndx, cmd in enumerate(queue[:-1]):
            cmd_name = cmd[0]
            cmd_args = cmd[1:]
            dolog = self.dolog_
            if dolog:
                output_name = '_%d' % ndx
                output_repr = None
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
                        arg_repr = output_names[arg]
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
                output_repr = None
                if cmd_name.startswith('%'):
                    obj = self
                    cmd_name = cmd_name[1:]
                    if self.dolog_:
                        if cmd_name == 'Load':
                            output_name = self.range2var(args[1]) + '_sav'
                        elif cmd_name == 'Save':
                            new_name = self.range2var(args[1]) + '_sav'
                            if type(cmd_args[2]) is str and cmd_args[2].startswith('!$'):
                                old_reg = '_%d(?!\d)' % int(cmd_args[2][2:])
                                output_lines = [re.sub(old_reg, new_name, line)
                                                for line in output_lines]
                            else:
                                output_lines.append('{} = {}'.format(new_name, arg_reprs[2]))
                            output_repr = arg_reprs[0]
                elif cmd_name.startswith('@'):
                    obj = methods
                    cmd_name = cmd_name[1:]
                    if self.dolog_:
                        output_repr = 'axlm.{}({})'.format(cmd_name, ','.join(arg_reprs))
                elif cmd_name.startswith('.'):
                    obj = args[0]
                    cmd_name = cmd_name[1:]
                    if self.dolog_:
                        output_repr = '{}.{}({})'.format(arg_reprs[0], cmd_name, ','.join(arg_reprs[1:]))
                else:
                    obj = _imports
                    if self.dolog_:
                        output_repr = '{}({})'.format(cmd_name, ','.join(arg_reprs))
                for ftok in cmd_name.split('.'):
                    obj = getattr(obj, ftok)
                output_value = obj(*args, **kwargs)
                if self.dolog_ and output_repr:
                    output_lines.append('{} = {}'.format(output_name, output_repr))
            except:
                estr = exc_info()
                tt = "   ".join(format_tb(estr[2]))
                return '''#PYTHON?
                    Error encountered executing Python function {}:
                        {}: {}
                        {}'''.format(cmd_name, estr[0].__name__, str(estr[1]), tt)
            if output_value is None:
                break
            if dolog:
                output_values.append(output_value)
                output_names.append(output_name)

        # Wrap in an extra tuple so the calling function does not
        # attempt to unpack it.
        if dolog:
            if output_lines and output_lines[-1].startswith('_'):
                final_name = self.range2var(output_range) if output_range else '_Out'
                output_lines[-1] = final_name + ' ' + output_lines[-1].split(' ', 1)[-1]
            self.log_.extend(output_lines)
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
