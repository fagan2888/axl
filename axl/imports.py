import re, importlib, types, sys, os, glob

_imports = types.ModuleType('axl_imports')
_modules = {}


def add_symbol(module_name, symbol_name, value):
    idict = _imports.__dict__
    if symbol_name not in idict:
        idict[symbol_name] = value
        _modules[symbol_name] = module_name
    elif value is not idict[symbol_name]:
        other_module = _modules[symbol_name]
        other_symbol = '{}.{}'.format(other_module, symbol_name) if other_module else symbol_name
        symbol_name = '{}.{}'.format(module_name, symbol_name) if module_name else symbol_name
        raise ImportError('Conflict between symbol names:\n  - {}\n  - {}'.format(other_symbol, symbol_name))


EOL = r'\s*(?:[#].*)?$'
EMPTY = r'^' + EOL
SYMBOL = r'[^\d\W][\d\w]*'
AS_SYMBOL = r'(?:\s+as\s+' + SYMBOL + r')?'
SYMBOL_AS = SYMBOL + AS_SYMBOL
SYMBOL_LIST = SYMBOL_AS + r'(?:\s*,\s*' + SYMBOL_AS + r')*'
MODULE = SYMBOL + r'(?:[.]' + SYMBOL + r')*'
MODULE_AS = MODULE + AS_SYMBOL
MODULE_LIST = MODULE_AS + r'(?:\s*,\s*' + MODULE_AS + r')*' 
IMPORT_LINE = '^import\s+(' + MODULE_LIST + r')' + EOL
FROM_LINE = '^from\s+(' + MODULE + r')\s+import\s+(' + SYMBOL_LIST + r'|[*])' + EOL


def parse_input_line(line):
    if re.match(EMPTY, line):
        return (None, ())
    match = re.match(IMPORT_LINE, line)
    if match:
        module, symbols = (None,) + match.groups()
    else:
        match = re.match(FROM_LINE, line)
        if match:
            module, symbols = match.groups()
            if module.startswith('.'):
                raise RuntimeError('Relative imports are not allowed: {}'.format(line))
            mod = importlib.import_module(module)
        else:
            raise RuntimeError('Invalid import string: {}'.format(line))
    if symbols != '*':
        symbols = re.split(',\s*', symbols)
    elif hasattr(mod, '__all__'):
        symbols = mod.__all__
    else:
        symbols = [name for name in dir(mod) if not name.startswith('_')]
    imports = {}
    for iname in symbols:
        nparts = re.split('\s+as\s+', iname)
        if len(nparts) == 2:
            iname, oname = nparts
        else:
            oname = iname.rsplit('.', 1)[-1]
        if module is None:
            if iname.startswith('.'):
                raise RuntimeError('Relative imports are not allowed: {}'.format(line))
            value = importlib.import_module(iname)
        elif not hasattr(mod, iname):
            raise ImportError('Cannot import name {!r} from module {!r}'.format(iname, module))
        else:
            value = getattr(mod, iname)
        add_symbol(module, oname, value)


parse_input_line('from axl.methods import *')
import_glob = os.path.join(sys.prefix, 'Tools', 'axl', 'imports.*')
for fname in glob.glob(import_glob):
    with open(fname) as fp:
        for imp in fp.read().splitlines():
            parse_input_line(imp)
